#!/usr/bin/env python3
"""
eval_matching.py — измеряет precision/recall/F1 матчинга на golden_pairs.json

Использование:
    python eval_matching.py
    python eval_matching.py --threshold 0.78
    python eval_matching.py --golden my_pairs.json --threshold 0.86

Формат golden_pairs.json:
    [
      {"a": "Название товара А", "b": "Название товара Б", "label": "match"},
      {"a": "...", "b": "...", "label": "non-match"}
    ]

Добавляй новые пары: приоритет — товары с одинаковым размером
(разные размеры и так блокирует size-гейт, они не влияют на калибровку порога).
"""
import argparse
import json
import sys
from pathlib import Path

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
except ImportError:
    sys.exit("Нужен scikit-learn: pip install scikit-learn")

sys.path.insert(0, str(Path(__file__).parent))
# заменить импорт из pipeline:
from pipeline import (
    _product_type_conflict,
    _product_type_sig,
    _dye_codes,
    _dye_conflict,
    _norm_for_match,
    _norm_for_tfidf,
    _extract_size_signature,
    _extract_variants,
    _variant_conflict,
    _match_tokens,
    _doc_freq,
    _extract_ply_rolls,
    _spec_shade_sig,      # новое
    _spec_count_sig,      # новое
    _SUBBRAND_DISC,       # новое
    RARE_DF,
)


def _score(i: int, j: int, mat) -> float:
    return float(cosine_similarity(mat[[i]], mat[[j]])[0, 0])


def _predict(a: str, b: str, threshold: float, idx: dict, mat, df, rare_df: int) -> tuple[bool, float]:
    if _product_type_conflict(_product_type_sig(a), _product_type_sig(b)):
        return False, 0.0
    if _dye_conflict(_dye_codes(a), _dye_codes(b)):
        return False, 0.0
    if _extract_size_signature(a) != _extract_size_signature(b):
        return False, 0.0
    if _extract_ply_rolls(a) != _extract_ply_rolls(b):
        return False, 0.0
    if _spec_shade_sig(a) != _spec_shade_sig(b):
        return False, 0.0
    if _spec_count_sig(a) != _spec_count_sig(b):
        return False, 0.0
    if _variant_conflict(_extract_variants(a), _extract_variants(b)):
        return False, 0.0
    diff = _match_tokens(a) ^ _match_tokens(b)
    if any(df[t] <= rare_df for t in diff) or any(t in _SUBBRAND_DISC for t in diff):
        return False, 0.0
    sim = _score(idx[a], idx[b], mat)
    return sim >= threshold, sim


def _run(pairs: list, threshold: float, rare_df: int = RARE_DF):
    names = sorted({p["a"] for p in pairs} | {p["b"] for p in pairs})
    idx = {n: i for i, n in enumerate(names)}
    norms = [_norm_for_tfidf(n) for n in names]

    vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=1)
    mat = vec.fit_transform(norms)
    df = _doc_freq(names)   # ⚠ DF на golden-наборе, не на полных данных — для грубой оценки

    tp = fp = tn = fn = 0
    fps, fns = [], []

    for p in pairs:
        a, b, truth = p["a"], p["b"], p["label"] == "match"
        pred, sim = _predict(a, b, threshold, idx, mat, df, rare_df)
        va, vb = _extract_variants(a), _extract_variants(b)
        sa, sb = _extract_size_signature(a), _extract_size_signature(b)

        if truth and pred:         tp += 1
        elif not truth and pred:   fp += 1; fps.append((a, b, sim, sa, sb, va, vb))
        elif truth and not pred:   fn += 1; fns.append((a, b, sim, sa, sb, va, vb))
        else:                      tn += 1

    precision = tp / (tp + fp) if (tp + fp) else 0.0
    recall    = tp / (tp + fn) if (tp + fn) else 0.0
    f1 = 2 * precision * recall / (precision + recall) if (precision + recall) else 0.0

    return tp, fp, tn, fn, precision, recall, f1, fps, fns


def _fmt_pair(a, b, sim, sa, sb, va, vb):
    print(f"    A size={sa} var={set(va) or '—'}  {a}")
    print(f"    B size={sb} var={set(vb) or '—'}  {b}")
    print(f"    sim={sim:.3f}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--golden",    default="golden_pairs.json", help="путь к файлу с разметкой")
    ap.add_argument("--threshold", type=float, default=0.82,    help="порог схожести (default 0.82)")
    ap.add_argument("--sweep",     action="store_true",          help="прогнать пороги 0.70..0.92")
    args = ap.parse_args()

    path = Path(args.golden)
    if not path.exists():
        sys.exit(f"Файл не найден: {path}\nСоздай его в формате golden_pairs.json (пример в README)")

    pairs = json.loads(path.read_text(encoding="utf-8"))
    n_match    = sum(1 for p in pairs if p["label"] == "match")
    n_nonmatch = sum(1 for p in pairs if p["label"] == "non-match")
    print(f"Golden set: {len(pairs)} пар  ({n_match} match / {n_nonmatch} non-match)")

    if args.sweep:
        print(f"\n{'Порог':>7}  {'P':>6}  {'R':>6}  {'F1':>6}  {'FP':>4}  {'FN':>4}")
        print("─" * 44)
        for t in [round(0.70 + i * 0.02, 2) for i in range(12)]:
            tp, fp, tn, fn, p, r, f1, *_ = _run(pairs, t)
            marker = " ◄" if abs(t - args.threshold) < 0.001 else ""
            print(f"  {t:.2f}   {p:6.3f}  {r:6.3f}  {f1:6.3f}  {fp:>4}  {fn:>4}{marker}")
        print()

    tp, fp, tn, fn, precision, recall, f1, fps, fns = _run(pairs, args.threshold)

    print(f"\nПорог: {args.threshold}")
    print(f"TP={tp}  FP={fp}  TN={tn}  FN={fn}")
    print(f"Precision: {precision:.3f}   Recall: {recall:.3f}   F1: {f1:.3f}")

    if fps:
        print(f"\n⚠  False Positives ({len(fps)}) — склеил лишнее:")
        for item in fps:
            _fmt_pair(*item)

    if fns:
        print(f"\n⚠  False Negatives ({len(fns)}) — не склеил нужное:")
        for item in fns:
            _fmt_pair(*item)

    if not fps and not fns:
        print("\n✅ Ни одного FP и FN на текущем golden set")


if __name__ == "__main__":
    main()