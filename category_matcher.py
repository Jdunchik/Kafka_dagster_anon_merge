"""
category_matcher.py — keyword-based categorization for Russian retail products.

No LLM required. Two-pass classification:

  Pass 1 (src_cats — high accuracy, zero maintenance):
      Source category strings (category_1..4 from the supplier's table) are
      tokenized and matched against target category names by weighted Jaccard.
      Because both vocabularies are the same Russian retail domain, this works
      without any hand-crafted rules: add a new target category to the JSON and
      it is found automatically.

  Pass 2 (product name keywords — fallback when src_cats is absent):
      A static keyword dictionary maps product-name substrings to semantic
      clusters, and clusters are mapped to target category keywords.
      This covers cases where only the product name is available.

API is fully backward-compatible:
    classify_batch / apply_categories / load_target_categories /
    load_cache / load_overrides / save_cache / _key
"""
import json
import re
from pathlib import Path

import pandas as pd

TARGET_CATS_PATH = Path("target_categories.json")
CACHE_PATH       = Path("category_cache.json")
OVERRIDES_PATH   = Path("category_overrides.json")


# ── TEXT NORMALIZATION ─────────────────────────────────────────────────────────

_SIZE_RE  = re.compile(
    r'\b\d+[.,]?\d*\s*(мл|л|г|гр|кг|шт|уп|пак|рул|ml|l|g|kg|pcs)\b', re.I
)
_JUNK_RE  = re.compile(r'[^\w\s]', re.U)
_SPACE_RE = re.compile(r'\s+')

# Common stop-words that are too frequent to be discriminating
_STOP = {
    "средство", "средства", "для", "и", "в", "на", "с", "по", "от",
    "уход", "за", "а", "не", "при", "из",
}


def _norm(text: str) -> str:
    s = str(text).lower().strip()
    s = _SIZE_RE.sub(' ', s)
    s = _JUNK_RE.sub(' ', s)
    return _SPACE_RE.sub(' ', s).strip()


def _tokens(text: str) -> set[str]:
    """Normalized tokens, stop-words removed, min length 3."""
    return {t for t in _norm(text).split() if len(t) >= 3 and t not in _STOP}


# ── PASS-2 KEYWORD FALLBACK ────────────────────────────────────────────────────
# Only used when src_cats is empty. Covers the real category_1 values found in
# the actual data (merged_202606191700.xlsx).

_PRODUCT_CLUSTERS: dict[str, list[str]] = {
    # ── Стирка ────────────────────────────────────────────────────────────────
    "стирка": [
        "для стирк", "стирал", "стиральн", "прач",
        "ariel", "ариэль", "tide", "тайд", "losk", "лоск",
        "persil", "персил", "bingo стир",
    ],
    "ополаскиватель_белья": [
        "ополаскив", "кондиционер для бель",
        "lenor", "ленор", "vernel", "вернель", "downy",
    ],
    "пятновыводитель": [
        "пятновывод", "vanish", "ваниш", "антипятн",
    ],
    "отбеливатель": [
        "отбелив", "белизн", "хлор жидк", "персоль",
    ],
    "уход_за_одеждой": [
        "антистатик", "для одежд", "для обув",
    ],
    # ── Посуда ────────────────────────────────────────────────────────────────
    "посуда_ручная": [
        "для мытья посуд", "мытья посуд", "посуд",
        "fairy", "фейри", "aos", "аос", "gala посуд",
        "palmolive посуд", "frosch посуд", "капля", "бальзам для посуд",
    ],
    "посуда_пмм": [
        "для посудомоечн", "посудомоечн", "для посудомойк",
        "finish", "финиш", "таблетк для пм", "гель для пмм", "порошок для пмм",
    ],
    # ── Уборка / чистящие ─────────────────────────────────────────────────────
    "чистящее": [
        "чистящ",
        "cif", "сиф", "comet", "комет", "ajax", "аякс",
        "mr proper", "мистер пропер", "пемолюкс", "pemolux", "sanita",
    ],
    "накипь": [
        "антинакип", "накипь", "antikalk", "calg", "кальг", "антикальц",
    ],
    "ржавчина": [
        "от ржавчин", "ржавчин", "антиржав",
    ],
    "стекло": [
        "для стекл", "стекл", "для зеркал",
    ],
    "пол": [
        "для пол", "мытья пол", "паркет", "ламинат",
    ],
    "мебель_ковры": [
        "для мебел", "ковр", "обивк",
    ],
    "трубы": [
        "для труб", "прочист труб", "стоков",
    ],
    # ── Санузел ───────────────────────────────────────────────────────────────
    "унитаз": [
        "для унитаз", "блок для унитаз", "унитаз",
        "domestos", "доместос", "duck", "дак", "bref", "бреф",
    ],
    "сантехника": [
        "для сантехник", "сантехник", "для ванн", "ванн", "раковин",
        "дезинфицир", "дезинфект сантех",
    ],
    # ── Воздух ────────────────────────────────────────────────────────────────
    "освежитель_воздуха": [
        "освежит воздух", "ароматиз воздух", "air freshener",
        "febreze", "фабриз", "glade", "глейд", "амбиантор",
        "аромадиффузор", "диффузор", "гелев освежит", "интерьерн освежит",
    ],
    # ── Насекомые ─────────────────────────────────────────────────────────────
    "насекомые": [
        "инсектицид", "репелент", "от насеком", "от мух", "от комар",
        "от моли", "от тараканов", "от муравьев", "ленты липкие",
    ],
    # ── Личная гигиена ────────────────────────────────────────────────────────
    "шампунь": [
        "шампун", "shampoo",
    ],
    "кондиционер_волос": [
        "бальзам для волос", "кондиционер для волос", "ополаскиватель для волос",
        "маска для волос",
    ],
    "средства_душ": [
        "гель для душ", "пена для душ", "пена для ванн", "соль для ванн",
        "скраб для тел", "соль для душ",
    ],
    "мыло": [
        "мыло жидк", "мыло куск", "мыло хоз", "мыло туалетн",
        "мыло дет",
    ],
    "зубы": [
        "зубн паст", "зубн щетк", "ополаскив для рт", "для полост рта",
        "нить зубн",
    ],
    "дезодорант": [
        "дезодорант", "антиперспирант",
    ],
    "депиляция": [
        "депиляц", "эпилятор", "воск для волос",
    ],
    "бритье": [
        "для бритья", "пена для бритья", "гель для бритья", "после бритья",
        "бритв", "лезвия",
    ],
    "уход_кожа": [
        "крем для кож", "лосьон для кож", "молочко для тел", "сыворотк",
        "маска для лиц", "скраб для лиц", "тоник",
    ],
    "детская_гигиена": [
        "детск крем", "детск гель", "детск шампун", "детск мыло",
        "подгузник", "трусик подгузник", "присыпка детск",
        "детск лосьон", "детск масло",
    ],
    "бумага": [
        "туалетн бумаг", "бумажн полотенц", "бумажн салфетк",
        "носов платок", "влажн салфетк",
    ],
    "женская_гигиена": [
        "прокладк", "тампон", "ежедневн прокладк", "урологическ",
    ],
    "загар": [
        "для загара", "автозагар", "после загара", "спф", "spf", "защит от солнц",
    ],
    "универсальное": [
        "универсальн",
    ],
}

# Cluster → indicator substrings in TARGET CATEGORY NAMES
_CAT_CLUSTER_KEYWORDS: dict[str, list[str]] = {
    "стирка":              ["стирк"],
    "ополаскиватель_белья":["ополаскив", "кондиционер для бель"],
    "пятновыводитель":     ["пятновывод", "пятн"],
    "отбеливатель":        ["отбелив", "хлор"],
    "уход_за_одеждой":     ["одежд", "обув"],
    "посуда_ручная":       ["посуд"],
    "посуда_пмм":          ["посудомоечн", "посудомойк", "пмм"],
    "чистящее":            ["чистящ"],
    "накипь":              ["накипь", "накип"],
    "ржавчина":            ["ржавчин"],
    "стекло":              ["стекл", "зеркал"],
    "пол":                 ["пол", "паркет"],
    "мебель_ковры":        ["мебел", "ковр"],
    "трубы":               ["труб", "стоков"],
    "унитаз":              ["унитаз", "туалет"],
    "сантехника":          ["сантехник", "ванн"],
    "освежитель_воздуха":  ["освежит", "ароматиз", "воздух"],
    "насекомые":           ["насеком", "инсектицид", "моль", "репелент"],
    "шампунь":             ["шампун"],
    "кондиционер_волос":   ["волос", "бальзам", "кондиционер"],
    "средства_душ":        ["душ", "ванн", "скраб"],
    "мыло":                ["мыло"],
    "зубы":                ["зуб", "полост рта"],
    "дезодорант":          ["дезодорант", "антиперспирант"],
    "депиляция":           ["депиляц"],
    "бритье":              ["бритье", "бритв"],
    "уход_кожа":           ["кожей", "кожи", "лиц", "тел", "рук", "ног"],
    "детская_гигиена":     ["детск", "подгузник", "ребенк"],
    "бумага":              ["бумаг", "салфетк", "платок"],
    "женская_гигиена":     ["прокладк", "тампон", "гигиен"],
    "загар":               ["загар", "солнц"],
    "универсальное":       ["универсальн"],
}


# ── CLASSIFICATION LOGIC ───────────────────────────────────────────────────────

def _build_cat_meta(target_cats: list[str]) -> dict[str, dict]:
    """
    Precompute for each target category:
      - 'tokens': set of discriminating tokens
      - 'norm':   normalized string (for substring search)
      - 'clusters': set of clusters this category belongs to
    """
    meta: dict[str, dict] = {}
    for cat in target_cats:
        n = _norm(cat)
        toks = _tokens(cat)
        clusters: set[str] = set()
        for cluster, keywords in _CAT_CLUSTER_KEYWORDS.items():
            for kw in keywords:
                if kw in n:
                    clusters.add(cluster)
                    break
        # Also scan product-side keywords (first few, most distinctive)
        for cluster, keywords in _PRODUCT_CLUSTERS.items():
            if cluster not in clusters:
                for kw in keywords[:3]:
                    if kw in n:
                        clusters.add(cluster)
                        break
        meta[cat] = {"norm": n, "tokens": toks, "clusters": clusters}
    return meta


def _jaccard_weighted(a: set[str], b: set[str]) -> float:
    """Weighted Jaccard: longer tokens get more weight (more discriminating)."""
    if not a or not b:
        return 0.0
    inter = a & b
    union = a | b
    w_inter = sum(len(t) for t in inter)
    w_union = sum(len(t) for t in union)
    return w_inter / w_union if w_union else 0.0


def _classify_from_src(
    src_cats: list[str],
    cat_meta: dict[str, dict],
) -> tuple[str, float]:
    """
    Pass 1: match source category strings against target categories.
    Returns (best_category, confidence).
    """
    src_combined = " ".join(src_cats)
    src_norm     = _norm(src_combined)
    src_toks     = _tokens(src_combined)

    best_cat, best_score = "", 0.0

    for cat, m in cat_meta.items():
        # Substring shortcut: target category name appears verbatim in src string
        if m["norm"] and m["norm"] in src_norm:
            return cat, 0.95

        score = _jaccard_weighted(src_toks, m["tokens"])
        if score > best_score:
            best_score, best_cat = score, cat

    if best_score >= 0.20:
        conf = round(min(0.55 + 0.40 * best_score, 0.93), 2)
        return best_cat, conf

    return "", 0.0


def _classify_from_name(
    norm_name: str,
    cat_meta: dict[str, dict],
) -> tuple[str, float]:
    """
    Pass 2: keyword cluster matching on product name.
    Used when src_cats is absent or Pass 1 found nothing.
    """
    prod_clusters: set[str] = set()
    for cluster, keywords in _PRODUCT_CLUSTERS.items():
        for kw in keywords:
            if kw in norm_name:
                prod_clusters.add(cluster)
                break

    if prod_clusters:
        scores: dict[str, float] = {}
        for cat, m in cat_meta.items():
            overlap = prod_clusters & m["clusters"]
            if overlap:
                scores[cat] = len(overlap) / max(len(prod_clusters), len(m["clusters"]), 1)
        if scores:
            best = max(scores, key=scores.get)
            conf = round(min(0.50 + 0.45 * scores[best], 0.93), 2)
            return best, conf
        # Known cluster but no matching category → don't guess
        return "", 0.0

    # Last resort: token overlap between product name and category names
    name_toks = _tokens(norm_name)
    tok_scores: dict[str, float] = {}
    for cat, m in cat_meta.items():
        s = _jaccard_weighted(name_toks, m["tokens"])
        if s:
            tok_scores[cat] = s
    if tok_scores:
        best = max(tok_scores, key=tok_scores.get)
        return best, 0.30

    return "", 0.0


# ── JSON HELPERS ───────────────────────────────────────────────────────────────

def _load_json(path, default):
    p = Path(path)
    if not p.exists():
        return default
    try:
        return json.loads(p.read_text("utf-8"))
    except Exception:
        return default


def _key(name: str, brand: str) -> str:
    return f"{str(name).strip().lower()}|||{str(brand).strip().lower()}"


def load_target_categories(path=TARGET_CATS_PATH) -> list:
    raw = _load_json(path, [])
    if isinstance(raw, dict):
        raw = raw.get("categories", [])
    return [str(c).strip() for c in raw if str(c).strip()]


def load_cache():      return _load_json(CACHE_PATH, {})
def load_overrides():  return _load_json(OVERRIDES_PATH, {})


def save_cache(cache: dict):
    CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), "utf-8")


# ── PUBLIC API ─────────────────────────────────────────────────────────────────

def classify_batch(rows: list, target_cats: list, model: str = "") -> list:
    """
    Classify a batch of product rows (instant, no LLM).

    `model` is accepted for API compatibility but ignored.
    rows: [{name, brand, src_cats:[...]}]
    Returns: [{...row, target_category: str, category_confidence: float}]
    """
    cat_meta = _build_cat_meta(target_cats)
    result = []
    for r in rows:
        src = [s for s in r.get("src_cats", []) if s and s != "nan"]
        if src:
            cat, conf = _classify_from_src(src, cat_meta)
            if cat:
                result.append({**r, "target_category": cat, "category_confidence": conf})
                continue
        # src_cats absent or gave no match → use product name
        norm_name = _norm(r.get("name", ""))
        cat, conf = _classify_from_name(norm_name, cat_meta)
        result.append({**r, "target_category": cat, "category_confidence": conf})
    return result


def apply_categories(
    df: pd.DataFrame,
    col: dict,
    target_cats: list,
    model: str = "",
    batch_size: int = 0,
    log_fn=print,
) -> pd.DataFrame:
    """
    Add target_category / category_confidence columns to df.

    Checks cache and overrides first; classifies only unique (name, brand) pairs
    not yet cached. Results are written back to category_cache.json.
    """
    ic = col.get("item_name")
    if not ic or ic not in df.columns:
        log_fn("  [Категории] нет item_name — пропуск")
        return df
    if not target_cats:
        log_fn("  [Категории] target_categories пуст — пропуск")
        return df

    bc       = col.get("brand")
    cat_cols = [c for c in (col.get(f"category_{k}") for k in range(5))
                if c and c in df.columns]

    cache, overrides = load_cache(), load_overrides()
    cat_meta = _build_cat_meta(target_cats)

    need_cols = list(dict.fromkeys(
        [ic] + ([bc] if bc and bc in df.columns else []) + cat_cols
    ))
    new_keys = 0
    for _, row in df[need_cols].drop_duplicates().iterrows():
        name  = str(row[ic]).strip()
        brand = str(row.get(bc, "")).strip() if bc else ""
        k = _key(name, brand)
        if k in overrides or k in cache:
            continue
        src = [str(row.get(cc, "")).strip() for cc in cat_cols
               if str(row.get(cc, "")).strip() not in ("", "nan")]
        if src:
            cat, conf = _classify_from_src(src, cat_meta)
        else:
            cat, conf = "", 0.0
        if not cat:
            cat, conf = _classify_from_name(_norm(name), cat_meta)
        cache[k] = {"target_category": cat, "category_confidence": conf}
        new_keys += 1

    if new_keys:
        save_cache(cache)

    names  = df[ic].astype(str).str.strip()
    brands = (df[bc].astype(str).str.strip() if bc and bc in df.columns
              else pd.Series([""] * len(df), index=df.index))

    def _lookup(name: str, brand: str) -> tuple[str, float]:
        k = _key(name, brand)
        if k in overrides:
            o = overrides[k]
            return o.get("target_category", ""), float(o.get("category_confidence", 1.0))
        c = cache.get(k, {})
        return c.get("target_category", ""), float(c.get("category_confidence", 0.0))

    pairs = [_lookup(n, b) for n, b in zip(names, brands)]
    df["target_category"]     = [p[0] for p in pairs]
    df["category_confidence"] = [p[1] for p in pairs]

    log_fn(
        f"  [Категории] {len(df)} строк классифицировано "
        f"(src_cats+keywords, {new_keys} новых в кэш)"
    )
    return df
