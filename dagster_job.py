"""
dagster_job.py — Dagster integration for retail_etl pipeline
─────────────────────────────────────────────────────────────
Wraps pipeline.py in Dagster ops so the pipeline can be
scheduled, monitored, and re-executed from the Dagster UI.

Required: pip install dagster dagster-webserver

── HOW TO RUN ──────────────────────────────────────────────

  # Launch the Dagster UI (development mode):
  dagster dev -f dagster_job.py

  # Execute once from CLI with default config:
  dagster job execute -f dagster_job.py -j retail_etl_job

  # Execute with a custom run config:
  dagster job execute -f dagster_job.py -j retail_etl_job \\
      --config run_config.yaml

── HOW TO ADD TO AN EXISTING PROJECT ───────────────────────

  In your existing definitions.py (or wherever you build Definitions):

    from dagster_job import retail_etl_job, retail_paths_resource

    defs = Definitions(
        jobs=[..., retail_etl_job],          # add the job
        resources={
            ...,
            "paths": retail_paths_resource,  # add the resource
        },
    )

── SAMPLE run_config.yaml ──────────────────────────────────

  resources:
    paths:
      config:
        input_folder:   "data/raw"
        archive_folder: "data/archive"
        output_folder:  "data/output"
        output_filename: ""          # empty → merged_YYYYMMDD.xlsx
"""

from pathlib import Path

import pandas as pd
from dagster import (
    ConfigurableResource,
    Definitions,
    OpExecutionContext,
    job,
    op,
)

# Import all business logic from pipeline.py (no Dagster in pipeline.py)
from pipeline import (
    SUPPORTED_EXTS,
    archive_source_files,
    collect_source_files,
    merge_anonymized,
    read_and_anonymize,
    save_merged,
)

from kafka import TopicPartition


from kafka import KafkaConsumer as _KafkaConsumer, KafkaProducer as _KafkaProducer
import time as _time

KAFKA_BROKERS        = "localhost:9092"
KAFKA_COMMANDS_TOPIC = "etl-commands"
KAFKA_RESULTS_TOPIC  = "etl-results"

# ── SHARED RESOURCE ────────────────────────────────────────────────────────────

class RetailPaths(ConfigurableResource):
    """Folder paths and matching config for the retail ETL pipeline."""
    input_folder:            str   = "data"
    archive_folder:          str   = "archive"
    output_folder:           str   = "merged"
    output_filename:         str   = ""        # "" → merged_YYYYMMDD.xlsx
    product_match:           bool  = True      # deduplicate product names before anon
    product_match_threshold: float = 0.82      # 0–1; lower = more aggressive grouping
    aggregate_after_merge:   bool  = True      # sum metrics for identical rows


# ── OPS ────────────────────────────────────────────────────────────────────────

@op
def op_collect_files(
    context: OpExecutionContext,
    paths: RetailPaths,
) -> list[str]:
    """Scan input_folder and return a list of absolute file paths (as strings).

    Serialised as strings so Dagster's default IO manager handles them
    without needing a custom type loader.
    """
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
    """Read every source file and anonymize it.

    Returns list of (filename, anonymized_df).
    The shared mappings cache (mappings.json) is loaded once and saved after
    all files are processed — so the same real entity always gets the same
    fake name across files.
    """
    files  = [Path(p) for p in file_paths]
    frames = read_and_anonymize(
        files,
        log_fn=context.log.info,
        product_match=paths.product_match,
        product_match_threshold=paths.product_match_threshold,
    )

    if not frames:
        raise ValueError("Ни один файл не удалось прочитать/анонимизировать.")

    total_rows = sum(len(df) for _, df in frames)
    context.log.info(
        f"Анонимизировано файлов: {len(frames)}, итого строк: {total_rows:,}"
    )
    return frames


@op
def op_merge(
    context: OpExecutionContext,
    paths: RetailPaths,
    frames: list[tuple[str, pd.DataFrame]],
) -> pd.DataFrame:
    """Find column intersection across all frames and concatenate, then aggregate."""
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
    """Write merged DataFrame to a formatted xlsx file.

    Returns the absolute output path as a string (used by op_archive
    to guarantee sequencing: archive only runs after save succeeds).
    """
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
    file_paths: list[str],  # from op_collect_files (fan-in)
    saved_path: str,        # from op_save (ordering guarantee)
) -> None:
    """Move original source files to archive_folder.

    Receives saved_path only to ensure this op runs AFTER op_save —
    we never want to archive sources if the output failed to write.
    Handles filename collisions automatically (appends _1, _2, …).
    """
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
    noticed = op_notify(saved)   # ← новое
    op_archive(files, noticed)   # принимает str вместо saved

from dagster import sensor, RunRequest, SkipReason

from kafka import TopicPartition

@sensor(job=retail_etl_job, minimum_interval_seconds=10)
def kafka_merge_sensor(context):
    try:
        consumer = _KafkaConsumer(
            bootstrap_servers=KAFKA_BROKERS,
            value_deserializer=lambda m: m.decode("utf-8"),
        )
        tp = TopicPartition(KAFKA_COMMANDS_TOPIC, 0)
        consumer.assign([tp])
        last_offset = int(context.cursor) if context.cursor else 0
        consumer.seek(tp, last_offset)
    except Exception as e:
        yield SkipReason(f"Kafka недоступна: {e}")
        return

    commands = []
    new_offset = last_offset
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
    product_match=True,
    product_match_threshold=0.82,
    aggregate_after_merge=True,
)

defs = Definitions(
    jobs=[retail_etl_job],
    resources={"paths": retail_paths_resource},
    sensors=[kafka_merge_sensor],
)