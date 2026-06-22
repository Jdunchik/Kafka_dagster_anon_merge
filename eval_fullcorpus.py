#!/usr/bin/env python3
"""
eval_fullcorpus.py — precision/recall с TF-IDF и DF на реальном корпусе данных.

В отличие от eval_matching.py (DF из ~87 пар → P=1.0 даже на 0.70),
здесь матрица и DF строятся на всём корпусе. Это даёт честные метрики.

Использование:
    python eval_fullcorpus.py --corpus data/ --sweep
    python eval_fullcorpus.py --corpus data/
    python eval_fullcorpus.py --corpus product_names.txt --sweep
    python eval_fullcorpus.py --corpus data/ --threshold 0.84

--corpus:
    папка  — читает все xlsx/csv/tsv через pipeline (ищет колонку item_name)
    .txt   — одно название в строке
"""
import argparse, json, sys
from pathlib import Path

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
except ImportError:
    sys.exit("pip install scikit-learn")

sys.path.insert(0, str(Path(__file__).parent))
from pipeline import (
    _norm_for_tfidf, _extract_size_signature, _extract_variants,
    _variant_conflict, _match_tokens, _real_diff, _doc_freq,
    _extract_ply_rolls, _spec_shade_sig, _spec_count_sig,
    _product_type_sig, _product_type_conflict,
    _dye_codes, _dye_conflict,
    _brand_sig, _brand_conflict,
    _SUBBRAND_DISC, RARE_DF,
    SUPPORTED_EXTS, read_table, get_col_map,
)


def load_corpus(path_str: str) -> list[str]:
    p = Path(path_str)
    if p.is_dir():
        names = []
        for f in sorted(p.iterdir()):
            if f.suffix.lower() not in SUPPORTED_EXTS:
                continue
            try:
                df = read_table(f)
                df.columns = [str(c).strip() for c in df.columns]
                col = get_col_map(df)
                ic = col.get("item_name")
                if ic and ic in df.columns:
                    batch = df[ic].dropna().astype(str).str.strip().tolist()
                    names.extend(batch)
                    print(f"  {f.name}: {len(batch):,} названий")
            except Exception as e:
                print(f"  warn {f.name}: {e}", file=sys.stderr)
        return list({n for n in names if n and n != "nan"})

    return [l.strip() for l in p.read_text(encoding="utf-8").splitlines() if l.strip()]


def _sim(i, j, mat):
    return float(cosine_similarity(mat[[i]], mat[[j]])[0, 0])


def _predict(a, b, threshold, idx, mat, df, rare_df):
    # HARD: любое несовпадение = разные товары
    if _extract_size_signature(a)   != _extract_size_signature(b):   return False, 0.0
    if _extract_ply_rolls(a)        != _extract_ply_rolls(b):        return False, 0.0
    if _spec_shade_sig(a)           != _spec_shade_sig(b):           return False, 0.0
    if _spec_count_sig(a)           != _spec_count_sig(b):           return False, 0.0
    if _dye_conflict(_dye_codes(a),    _dye_codes(b)):                return False, 0.0
    if _variant_conflict(_extract_variants(a), _extract_variants(b)): return False, 0.0
    if _product_type_conflict(_product_type_sig(a), _product_type_sig(b)): return False, 0.0
    # SOFT: в продакшене → review, здесь трактуем как "не матч" (нижняя граница recall)
    ta, tb = _match_tokens(a), _match_tokens(b)
    if _brand_conflict(_brand_sig(a), _brand_sig(b)):                return False, 0.0
    if any(df[t] <= rare_df for t in _real_diff(ta, tb)):            return False, 0.0
    if any(t in _SUBBRAND_DISC for t in ta ^ tb):                    return False, 0.0

    if a not in idx or b not in idx:
        return False, 0.0
    s = _sim(idx[a], idx[b], mat)
    return s >= threshold, s


def _run(pairs, corpus, threshold, rare_df=RARE_DF):
    # golden пары гарантированно в матрице (union с корпусом)
    all_names = list({*corpus, *(p["a"] for p in pairs), *(p["b"] for p in pairs)})
    idx   = {n: i for i, n in enumerate(all_names)}
    norms = [_norm_for_tfidf(n) for n in all_names]

    vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=1)
    mat = vec.fit_transform(norms)
    df  = _doc_freq(all_names)

    tp = fp = tn = fn = 0
    fps, fns = [], []

    for p in pairs:
        a, b, truth = p["a"], p["b"], p["label"] == "match"
        pred, s = _predict(a, b, threshold, idx, mat, df, rare_df)
        va, vb  = _extract_variants(a), _extract_variants(b)
        sa, sb  = _extract_size_signature(a), _extract_size_signature(b)
        if   truth and pred:       tp += 1
        elif not truth and pred:   fp += 1; fps.append((a, b, s, sa, sb, va, vb))
        elif truth and not pred:   fn += 1; fns.append((a, b, s, sa, sb, va, vb))
        else:                      tn += 1

    prec = tp/(tp+fp) if (tp+fp) else 0.0
    rec  = tp/(tp+fn) if (tp+fn) else 0.0
    f1   = 2*prec*rec/(prec+rec) if (prec+rec) else 0.0
    return tp, fp, tn, fn, prec, rec, f1, fps, fns


def _fmt(a, b, sim, sa, sb, va, vb):
    print(f"    A size={sa} var={set(va) or '—'}  {a}")
    print(f"    B size={sb} var={set(vb) or '—'}  {b}")
    print(f"    sim={sim:.3f}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--corpus",    required=True,            help="папка с данными или txt-файл с именами")
    ap.add_argument("--golden",    default="golden_pairs.json")
    ap.add_argument("--threshold", type=float, default=0.86)
    ap.add_argument("--sweep",     action="store_true")
    args = ap.parse_args()

    gpath = Path(args.golden)
    if not gpath.exists():
        sys.exit(f"Файл не найден: {gpath}")
    pairs = json.loads(gpath.read_text(encoding="utf-8"))
    n_m = sum(1 for p in pairs if p["label"] == "match")

    print(f"Загрузка корпуса: {args.corpus}")
    corpus = load_corpus(args.corpus)
    print(f"Корпус: {len(corpus):,} имён  |  Golden: {len(pairs)} пар  ({n_m} match / {len(pairs)-n_m} non-match)")

    if args.sweep:
        print(f"\n{'Порог':>7}  {'P':>6}  {'R':>6}  {'F1':>6}  {'FP':>4}  {'FN':>4}")
        print("─" * 44)
        for t in [round(0.70 + i*0.02, 2) for i in range(12)]:
            tp, fp, tn, fn, p, r, f1, *_ = _run(pairs, corpus, t)
            mark = " ◄" if abs(t - args.threshold) < 0.001 else ""
            print(f"  {t:.2f}   {p:6.3f}  {r:6.3f}  {f1:6.3f}  {fp:>4}  {fn:>4}{mark}")
        print()

    tp, fp, tn, fn, prec, rec, f1, fps, fns = _run(pairs, corpus, args.threshold)
    print(f"\nПорог: {args.threshold}")
    print(f"TP={tp}  FP={fp}  TN={tn}  FN={fn}")
    print(f"Precision: {prec:.3f}   Recall: {rec:.3f}   F1: {f1:.3f}")

    if fps:
        print(f"\n⚠  False Positives ({len(fps)}):")
        for x in fps: _fmt(*x)
    if fns:
        print(f"\n⚠  False Negatives ({len(fns)}):")
        for x in fns: _fmt(*x)
    if not fps and not fns:
        print("\n✅ Ни одного FP и FN")


if __name__ == "__main__":
    main()