import time as _time
from pathlib import Path

import pandas as pd
from dagster import (
    ConfigurableResource,
    DefaultSensorStatus,
    Definitions,
    OpExecutionContext,
    RunRequest,
    SkipReason,
    job,
    op,
    sensor,
)
from kafka import (
    KafkaConsumer as _KafkaConsumer,
    KafkaProducer as _KafkaProducer,
    TopicPartition,
)

from pipeline import (
    SUPPORTED_EXTS,
    archive_source_files,
    collect_source_files,
    merge_anonymized,
    read_and_anonymize,
    save_merged,
)

KAFKA_BROKERS        = "localhost:9092"
KAFKA_COMMANDS_TOPIC = "etl-commands"
KAFKA_RESULTS_TOPIC  = "etl-results"

# ── SHARED RESOURCE ────────────────────────────────────────────────────────────

class RetailPaths(ConfigurableResource):
    """Folder paths and config for the retail ETL pipeline."""
    input_folder:          str  = "data"
    archive_folder:        str  = "archive"
    output_folder:         str  = "merged"
    output_filename:       str  = ""        # "" → merged_YYYYMMDD.xlsx
    categorize:            bool = True       # AI-категоризация в целевую систему
    name_match:            bool = True       # canonical_name из name_matches.json
    model:                 str  = "qwen2.5:3b"
    aggregate_after_merge: bool = False      # мэтчинг сохраняет все строки продаж


# ── OPS ────────────────────────────────────────────────────────────────────────

@op
def op_collect_files(
    context: OpExecutionContext,
    paths: RetailPaths,
) -> list[str]:
    folder = Path(paths.input_folder).resolve()
    if not folder.exists():
        raise FileNotFoundError(
            f"input_folder не существует: {folder}\n"
            f"Укажи реальный путь в run_config.yaml:\n"
            f"  resources:\n"
            f"    paths:\n"
            f"      config:\n"
            f"        input_folder: \"/путь/к/папке/с/файлами\""
        )

    files = collect_source_files(folder)
    if not files:
        raise ValueError(
            f"Нет поддерживаемых файлов в {folder}/ "
            f"(форматы: {', '.join(sorted(SUPPORTED_EXTS))})"
        )

    context.log.info(f"Найдено файлов: {len(files)} в {folder}/")
    for f in files:
        kb = max(f.stat().st_size // 1024, 1)
        context.log.info(f"  {f.name} ({kb} КБ)")

    return [str(f) for f in files]


@op
def op_read_anonymize(
    context: OpExecutionContext,
    paths: RetailPaths,
    file_paths: list[str],
) -> list[tuple[str, pd.DataFrame]]:
    files  = [Path(p) for p in file_paths]
    frames = read_and_anonymize(
        files,
        log_fn=context.log.info,
        categorize=paths.categorize,
        name_match=paths.name_match,
        model=paths.model,
    )

    if not frames:
        raise ValueError("Ни один файл не удалось прочитать/обработать.")

    total_rows = sum(len(df) for _, df in frames)
    context.log.info(
        f"Обработано файлов: {len(frames)}, итого строк: {total_rows:,}"
    )
    return frames


@op
def op_merge(
    context: OpExecutionContext,
    paths: RetailPaths,
    frames: list[tuple[str, pd.DataFrame]],
) -> pd.DataFrame:
    merged = merge_anonymized(
        frames,
        log_fn=context.log.info,
        aggregate=paths.aggregate_after_merge,
    )
    context.log.info(
        f"Объединено: {len(merged):,} строк × {len(merged.columns)} столбцов"
    )
    return merged


@op
def op_save(
    context: OpExecutionContext,
    paths: RetailPaths,
    merged: pd.DataFrame,
) -> str:
    out = save_merged(
        merged,
        paths.output_folder,
        paths.output_filename,
        log_fn=context.log.info,
    )
    context.log.info(f"Файл сохранён: {out}")
    return str(out)


@op
def op_archive(
    context: OpExecutionContext,
    paths: RetailPaths,
    file_paths: list[str],
    saved_path: str,
) -> None:
    context.log.info(f"Результат сохранён в: {saved_path}")
    files = [Path(p) for p in file_paths]
    n = archive_source_files(files, paths.archive_folder, log_fn=context.log.info)
    context.log.info(f"Архивировано: {n}/{len(files)} файлов → {paths.archive_folder}/")


@op
def op_notify(context: OpExecutionContext, saved_path: str) -> str:
    """Публикует в Kafka результат после успешного сохранения."""
    try:
        p = _KafkaProducer(
            bootstrap_servers=KAFKA_BROKERS,
            value_serializer=lambda v: v.encode("utf-8"),
        )
        p.send(KAFKA_RESULTS_TOPIC, f"✅ Готово: {saved_path}")
        p.flush(); p.close()
        context.log.info("Уведомление отправлено в Kafka")
    except Exception as e:
        context.log.warning(f"Kafka недоступна, уведомление пропущено: {e}")
    return saved_path


# ── JOB ────────────────────────────────────────────────────────────────────────

@job
def retail_etl_job():
    files   = op_collect_files()
    frames  = op_read_anonymize(files)
    merged  = op_merge(frames)
    saved   = op_save(merged)
    noticed = op_notify(saved)
    op_archive(files, noticed)


# ── SENSOR ─────────────────────────────────────────────────────────────────────

@sensor(
    job=retail_etl_job,
    minimum_interval_seconds=10,
    default_status=DefaultSensorStatus.RUNNING,
)
def kafka_merge_sensor(context):
    consumer = None
    try:
        consumer = _KafkaConsumer(
            bootstrap_servers=KAFKA_BROKERS,
            value_deserializer=lambda m: m.decode("utf-8"),
        )
        tp = TopicPartition(KAFKA_COMMANDS_TOPIC, 0)
        consumer.assign([tp])

        if context.cursor is None:
            consumer.seek_to_end(tp)
            context.update_cursor(str(consumer.position(tp)))
            consumer.close()
            consumer = None
            yield SkipReason("Первый запуск: курсор установлен на конец топика")
            return

        consumer.seek(tp, int(context.cursor))
    except Exception as e:
        if consumer is not None:
            consumer.close()
        yield SkipReason(f"Kafka недоступна: {e}")
        return

    commands = []
    new_offset = int(context.cursor)
    try:
        records = consumer.poll(timeout_ms=3000)
        for msgs in records.values():
            for msg in msgs:
                new_offset = msg.offset + 1
                if msg.value.strip() == "/merge":
                    commands.append(msg)
    finally:
        consumer.close()

    context.update_cursor(str(new_offset))

    if not commands:
        yield SkipReason("Нет команды /merge")
        return

    for i in range(len(commands)):
        yield RunRequest(run_key=f"merge-{int(_time.time()*1000)}-{i}")


retail_paths_resource = RetailPaths(
    input_folder="data",
    archive_folder="archive",
    output_folder="merged",
    output_filename="",
    categorize=True,
    name_match=True,
    model="qwen2.5:3b",
    aggregate_after_merge=False,
)

defs = Definitions(
    jobs=[retail_etl_job],
    resources={"paths": retail_paths_resource},
    sensors=[kafka_merge_sensor],
)