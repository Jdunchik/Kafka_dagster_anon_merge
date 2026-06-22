#!/usr/bin/env python3
"""
review_cli.py — real-time Qwen3-ассистент для разбора пограничных пар.

Что делает:
  - перед каждой парой показывает мнение ИИ + причину + предложение
  - алёрты срабатывают немедленно когда паттерн повторяется > порога
  - ai_context.json — память между сессиями (знает что уже добавлено)
  - финальный анализ с /think-режимом

Требования:
    brew install ollama
    ollama pull qwen3:4b    # или qwen3:30b-a3b если RAM позволяет
    ollama serve

Использование:
    python review_cli.py
    python review_cli.py --model qwen3:8b
    python review_cli.py --auto             # LLM анализирует match_decisions без интерактива
    python review_cli.py --no-llm           # только ручной разбор
"""
import argparse, json, re, sys, time
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

REVIEW_PATH    = Path("review_pairs.json")
DECISIONS_PATH = Path("match_decisions.jsonl")
GOLDEN_PATH    = Path("golden_pairs.json")
PATCH_PATH     = Path("suggested_patch.json")
CONTEXT_PATH   = Path("ai_context.json")

OLLAMA_URL      = "http://localhost:11434"
DEFAULT_MODEL   = "qwen3:14b"
ALERT_THRESHOLD = 2   # алёрт если токен встречается в > N подтверждённых FP

_TTY = sys.stdout.isatty()

def _c(t, code): return f"\033[{code}m{t}\033[0m" if _TTY else t
def red(t):    return _c(t, "91")
def green(t):  return _c(t, "92")
def yellow(t): return _c(t, "93")
def cyan(t):   return _c(t, "96")
def bold(t):   return _c(t, "1")
def dim(t):    return _c(t, "2")


# ── OLLAMA ─────────────────────────────────────────────────────────────────────

def _chat(messages: list, model: str, think: bool = False, timeout: int = 45) -> str | None:
    import urllib.request, urllib.error
    msgs = [m.copy() for m in messages]
    directive = "/think\n" if think else "/no_think\n"
    for m in reversed(msgs):
        if m["role"] == "user":
            m["content"] = directive + m["content"]
            break
    data = json.dumps({"model": model, "messages": msgs, "stream": False,
                        "options": {"temperature": 0.05}}).encode()
    try:
        req = urllib.request.Request(
            f"{OLLAMA_URL}/api/chat", data=data,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read())["message"]["content"].strip()
    except Exception:
        return None


def _extract_json(raw: str) -> dict | None:
    if "<think>" in raw:
        raw = raw.split("</think>")[-1].strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
    m = re.search(r'\{[^{}]+\}', raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group())
        except json.JSONDecodeError:
            pass
    return None


def check_ollama(model: str) -> bool:
    import urllib.request, urllib.error
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=3) as r:
            tags = json.loads(r.read())
        names = [m["name"] for m in tags.get("models", [])]
        base  = model.split(":")[0]
        if not any(base in n for n in names):
            print(yellow(f"  ⚠ Модель {model} не найдена в Ollama"))
            print(yellow(f"    Запусти: ollama pull {model}"))
            return False
        return True
    except Exception:
        print(red(f"  Ollama недоступен на {OLLAMA_URL}"))
        print(red(f"  Запусти: ollama serve"))
        return False


# ── CONTEXT (память между сессиями) ───────────────────────────────────────────

def load_context() -> dict:
    if CONTEXT_PATH.exists():
        try:
            return json.loads(CONTEXT_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"known_patterns": {}, "applied_patches": {}, "sessions": []}


def save_context(ctx: dict):
    CONTEXT_PATH.write_text(json.dumps(ctx, ensure_ascii=False, indent=2), encoding="utf-8")


def _pipeline_dicts() -> dict:
    from pipeline import _SUBBRAND_DISC, _SYNONYMS, _PRODUCT_TYPES
    return {
        "subbrand_disc": sorted(_SUBBRAND_DISC),
        "synonyms":      _SYNONYMS,
        "product_types": sorted(_PRODUCT_TYPES),
    }


def _build_system(ctx: dict) -> str:
    dicts   = _pipeline_dicts()
    known   = ctx.get("known_patterns", {})
    applied = ctx.get("applied_patches", {})
    return (
        "Ты — NLP-эксперт по матчингу товаров российского ретейла.\n"
        "Правило: лучше пропустить склейку, чем склеить разные товары.\n\n"
        f"ТЕКУЩИЕ СЛОВАРИ (не предлагай то что уже есть):\n{json.dumps(dicts, ensure_ascii=False)}\n\n"
        f"ПРИМЕНЁННЫЕ ПАТЧИ: {json.dumps(applied, ensure_ascii=False)}\n\n"
        f"ИЗВЕСТНЫЕ ПАТТЕРНЫ: {json.dumps(known, ensure_ascii=False) if known else 'нет'}\n\n"
        "Для каждой пары отвечай ТОЛЬКО одной строкой JSON без пояснений:\n"
        '{"verdict":"match"|"non-match"|"uncertain",'
        '"confidence":0-100,'
        '"reason":"1-2 предложения",'
        '"suggest":{"dict":"_SUBBRAND_DISC","add":"токен"}|null}'
    )


# ── PER-PAIR ASSESSMENT ────────────────────────────────────────────────────────

def assess_pair(p: dict, system: str, model: str) -> dict | None:
    a, b   = p["a"], p["b"]
    diff   = p.get("diff_tokens", [])
    gate   = p.get("gate", "?")
    score  = p.get("score", 0)
    # думаем только на сложных: высокий sim + бренд-гейт
    think  = score > 0.87 and "brand" in gate
    user   = f'A: "{a}"\nB: "{b}"\ndiff_tokens: {diff}\ngate: {gate}\nsim: {score:.3f}'

    t0  = time.time()
    raw = _chat([{"role": "system", "content": system},
                 {"role": "user",   "content": user}], model, think=think)
    elapsed = round(time.time() - t0, 2)
    if not raw:
        return None
    result = _extract_json(raw)
    if result:
        result["_elapsed"] = elapsed
        result["_think"]   = think
    return result


# ── ALERT SYSTEM ───────────────────────────────────────────────────────────────

def check_alerts(token_counter: Counter, alerted: set, ctx: dict) -> list[str]:
    dicts    = _pipeline_dicts()
    existing = set(dicts["subbrand_disc"]) | set(dicts["product_types"]) | set(dicts["synonyms"])
    alerts   = []
    for tok, cnt in token_counter.most_common():
        if cnt <= ALERT_THRESHOLD or tok in alerted or tok in existing or len(tok) < 3:
            continue
        alerts.append(tok)
        alerted.add(tok)
        ctx["known_patterns"].setdefault(tok, {"fp_count": 0})
        ctx["known_patterns"][tok].update({"fp_count": cnt, "status": "alert"})
    return alerts


# ── MAIN REVIEW LOOP ───────────────────────────────────────────────────────────

def interactive_review(pairs: list, model: str, ctx: dict) -> tuple[list, list, list]:
    system        = _build_system(ctx)
    labeled       = []
    fp_confirmed  = []
    overrides     = []
    token_counter = Counter()
    alerted:set   = set()
    patch_acc     = {"subbrand_disc_add": [], "synonyms_add": {}, "product_types_add": []}
    n = len(pairs)

    print(f"\n{bold(f'{n} пар на разбор')}  |  m=match  n=non-match  s=пропустить  q=стоп\n")

    for i, p in enumerate(pairs):
        a, b   = p["a"], p["b"]
        score  = p.get("score", 0)
        gate   = p.get("gate", "?")
        diff   = p.get("diff_tokens", [])

        print(f"{cyan(f'[{i+1}/{n}]')}  gate={yellow(gate)}  sim={score:.3f}")
        if diff:
            print(f"  diff: {', '.join(diff[:10])}{'…' if len(diff)>10 else ''}")
        print(f"  A: {a}")
        print(f"  B: {b}")

        # LLM assessment до решения человека
        print(f"  {dim('ИИ...')}", end="\r")
        ai = assess_pair(p, system, model)
        if ai:
            v       = ai.get("verdict", "?")
            conf    = ai.get("confidence", 0)
            reason  = ai.get("reason", "")
            suggest = ai.get("suggest")
            mode    = "🧠 " if ai.get("_think") else ""
            t_str   = dim(f"{ai.get('_elapsed', 0)}s")
            vcol    = green if v == "match" else red if v == "non-match" else yellow
            print(f"  {mode}{vcol(f'ИИ: {v}')} {conf}%  {t_str}  {reason}")
            if suggest:
                print(f"  {dim(f'→ {suggest[\"dict\"]} += \"{suggest[\"add\"]}\"')}")
        else:
            print(f"  {dim('ИИ не ответил              ')}")
            ai = None

        while True:
            raw = input("  > ").strip().lower()
            if raw and raw[0] in ("m", "n", "s", "q"):
                raw = raw[0]; break
            print("  m / n / s / q")
        print()

        if raw == "q":
            break
        if raw == "s":
            continue

        label = "match" if raw == "m" else "non-match"
        labeled.append({"a": a, "b": b, "label": label})

        # расхождение с ИИ
        if ai:
            ai_says = ai.get("verdict")
            if raw == "n" and ai_says == "match":
                overrides.append({"pair": p, "ai": ai, "human": "non-match"})
                print(f"  {yellow('↩ human override (ИИ думал match)')}")
            elif raw == "m" and ai_says == "non-match":
                overrides.append({"pair": p, "ai": ai, "human": "match"})
                print(f"  {yellow('↩ human override (ИИ думал non-match)')}")

        if raw == "n":
            fp_confirmed.append(p)
            for tok in diff:
                token_counter[tok] += 1

            # копим предложение ИИ в патч
            if ai and ai.get("suggest"):
                s = ai["suggest"]
                tok, d = s.get("add", ""), s.get("dict", "")
                if tok and d == "_SUBBRAND_DISC" and tok not in patch_acc["subbrand_disc_add"]:
                    patch_acc["subbrand_disc_add"].append(tok)
                elif tok and d == "_PRODUCT_TYPES" and tok not in patch_acc["product_types_add"]:
                    patch_acc["product_types_add"].append(tok)

            # алёрты по частоте токенов
            for tok in check_alerts(token_counter, alerted, ctx):
                print(f"\n  {bold(red(f'⚠ ALERT: \"{tok}\" встречается в {token_counter[tok]} FP'))}")
                print(f"  {red('→ скорее всего системная дыра')}")
                ans = input(f"  добавить \"{tok}\" в _SUBBRAND_DISC? [y/n] ").strip().lower()
                if ans == "y" and tok not in patch_acc["subbrand_disc_add"]:
                    patch_acc["subbrand_disc_add"].append(tok)
                    print(green(f"  ✓ добавлено в патч"))
                print()

        # промежуточное сохранение каждые 10 FP
        if fp_confirmed and len(fp_confirmed) % 10 == 0:
            _save_patch(patch_acc)
            print(green(f"\n  [промежуточный патч → {PATCH_PATH}]\n"))

    _save_patch(patch_acc)
    return labeled, fp_confirmed, overrides


# ── ФИНАЛЬНЫЙ БАТЧ-АНАЛИЗ ─────────────────────────────────────────────────────

def final_analysis(fp_confirmed: list, overrides: list, model: str, ctx: dict) -> dict | None:
    if not fp_confirmed:
        return None

    system = (
        _build_system(ctx) + "\n\n"
        "Режим: финальный анализ всей сессии.\n"
        "Отвечай ТОЛЬКО JSON (без пояснений):\n"
        '{"subbrand_disc_add":[],"synonyms_add":{},"product_types_add":[],'
        '"type_canon_add":{},"analysis":"...","code_snippets":"..."}'
    )
    payload = {
        "fp_confirmed":     fp_confirmed[:50],
        "overrides_count":  len(overrides),
        "override_examples": [o["pair"] for o in overrides[:5]],
    }

    print(f"\n  Финальный анализ {len(fp_confirmed)} FP…", end="", flush=True)
    raw = _chat(
        [{"role": "system", "content": system},
         {"role": "user",   "content": json.dumps(payload, ensure_ascii=False)}],
        model, think=True, timeout=90,
    )
    print(" готово")
    if not raw:
        return None
    if "<think>" in raw:
        raw = raw.split("</think>")[-1].strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]
    try:
        return json.loads(raw)
    except Exception:
        return _extract_json(raw)


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _enrich(pairs: list) -> list:
    try:
        from pipeline import _match_tokens
        return [{**p, "diff_tokens": sorted(_match_tokens(p["a"]) ^ _match_tokens(p["b"]))} for p in pairs]
    except Exception:
        return pairs


def _save_patch(acc: dict):
    PATCH_PATH.write_text(json.dumps(acc, ensure_ascii=False, indent=2), encoding="utf-8")


def load_json(path: Path, default):
    if not path.exists(): return default
    return json.loads(path.read_text(encoding="utf-8"))


def save_to_golden(pairs: list, path: Path):
    existing = load_json(path, [])
    seen = {(p["a"], p["b"]) for p in existing}
    add  = [p for p in pairs if (p["a"], p["b"]) not in seen]
    if add:
        path.write_text(json.dumps(existing + add, ensure_ascii=False, indent=2), encoding="utf-8")
        print(green(f"  +{len(add)} пар → {path}"))


def load_decisions() -> tuple[list, list]:
    if not DECISIONS_PATH.exists(): return [], []
    accepted, review = [], []
    for line in DECISIONS_PATH.read_text(encoding="utf-8").splitlines():
        if line.strip():
            d = json.loads(line)
            v = d.get("verdict")
            if v == "match":   accepted.append(d)
            elif v == "review": review.append(d)
    return accepted, review


def _plain_review(pairs: list) -> tuple[list, list, list]:
    labeled, fps = [], []
    n = len(pairs)
    print(f"\n{bold(f'{n} пар')}  |  m=match  n=non-match  s=пропустить  q=стоп\n")
    for i, p in enumerate(pairs):
        diff = p.get("diff_tokens", [])
        print(f"{cyan(f'[{i+1}/{n}]')}  gate={yellow(p.get('gate','?'))}  sim={p.get('score',0):.3f}")
        if diff: print(f"  diff: {', '.join(diff[:10])}")
        print(f"  A: {p['a']}\n  B: {p['b']}")
        while True:
            raw = input("  > ").strip().lower()
            if raw and raw[0] in ("m", "n", "s", "q"): raw = raw[0]; break
        print()
        if raw == "q": break
        if raw == "s": continue
        labeled.append({"a": p["a"], "b": p["b"], "label": "match" if raw=="m" else "non-match"})
        if raw == "n": fps.append(p)
    return labeled, fps, []


def print_patch(patch: dict):
    print(f"\n{bold('══════ Патч ══════')}")
    if patch.get("analysis"):
        print(f"\n{patch['analysis']}")
    for key, label in [
        ("subbrand_disc_add", "_SUBBRAND_DISC"),
        ("product_types_add", "_PRODUCT_TYPES"),
        ("synonyms_add",      "_SYNONYMS     "),
        ("type_canon_add",    "_TYPE_CANON   "),
    ]:
        val = patch.get(key)
        if not val: continue
        print(f"\n  {yellow(label)}:")
        if isinstance(val, list): print("    " + ", ".join(f'"{v}"' for v in val))
        else:
            for k, v in val.items(): print(f'    "{k}": "{v}"')
    if patch.get("code_snippets"):
        print(f"\n{bold('Вставить в pipeline.py:')}\n{patch['code_snippets']}")
    print()


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pairs",  default=str(REVIEW_PATH))
    ap.add_argument("--golden", default=str(GOLDEN_PATH))
    ap.add_argument("--model",  default=DEFAULT_MODEL,    help="Ollama модель (default: qwen3:4b)")
    ap.add_argument("--auto",   action="store_true",       help="LLM анализирует match_decisions без интерактива")
    ap.add_argument("--no-llm", action="store_true",       help="только ручной разбор")
    args = ap.parse_args()

    golden_path = Path(args.golden)
    ctx         = load_context()
    accepted, _ = load_decisions()

    llm_ok = False if args.no_llm else check_ollama(args.model)
    if not llm_ok and not args.no_llm:
        print(yellow("  Продолжаю без LLM\n"))

    # ── auto: только LLM без интерактива ──────────────────────────────────────
    if args.auto:
        if not llm_ok:
            print(red("  --auto требует Ollama")); sys.exit(1)
        fp_for_llm = _enrich(accepted[:60])
        patch = final_analysis(fp_for_llm, [], args.model, ctx)
        if patch:
            _save_patch(patch)
            print_patch(patch)
            print(green(f"  Патч → {PATCH_PATH}"))
        _update_session(ctx, 0, len(fp_for_llm), 0)
        return

    # ── интерактивный разбор ──────────────────────────────────────────────────
    pairs = load_json(Path(args.pairs), [])
    if not pairs:
        print(f"Нет пар в {args.pairs}. Запусти pipeline сначала.")
        sys.exit(0)
    pairs = _enrich(pairs)

    if llm_ok:
        labeled, fp_confirmed, overrides = interactive_review(pairs, args.model, ctx)
    else:
        labeled, fp_confirmed, overrides = _plain_review(pairs)

    if labeled:
        save_to_golden(labeled, golden_path)

    # финальный анализ
    if fp_confirmed and llm_ok:
        patch = final_analysis(fp_confirmed, overrides, args.model, ctx)
        if patch:
            existing = load_json(PATCH_PATH, {})
            for key in ("subbrand_disc_add", "product_types_add"):
                patch[key] = list(set(existing.get(key, []) + patch.get(key, [])))
            patch["synonyms_add"] = {**existing.get("synonyms_add", {}), **patch.get("synonyms_add", {})}
            _save_patch(patch)
            print_patch(patch)
            print(green(f"  Патч → {PATCH_PATH}"))

    # summary
    if overrides:
        print(f"\n{bold('Расхождений с ИИ:')} {len(overrides)}")
        for o in overrides[:3]:
            ai_v = o["ai"].get("verdict", "?")
            print(f"  ИИ={ai_v} → человек={o['human']}  |  {o['pair']['a'][:55]}")
        if len(overrides) > 3: print(f"  … ещё {len(overrides)-3}")

    _update_session(ctx, len(labeled), len(fp_confirmed), len(overrides))
    print(f"\n{dim(f'Контекст сохранён → {CONTEXT_PATH}  (сессий: {len(ctx[\"sessions\"])})') }")


def _update_session(ctx: dict, reviewed: int, fps: int, overrides: int):
    ctx["sessions"].append({
        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "reviewed": reviewed, "fp": fps, "overrides": overrides,
    })
    save_context(ctx)


if __name__ == "__main__":
    main()