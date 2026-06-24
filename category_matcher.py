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
    "посуда_ручная":       ["посуд", "кухн", "ручн"],    # кухн/ручн → «КУХНЯ (ручн)»
    "посуда_пмм":          ["посудомоечн", "посудомойк", "пмм"],
    "чистящее":            ["чистящ", "убор"],            # убор → «УБОРКА»
    "накипь":              ["накипь", "накип"],
    "ржавчина":            ["ржавчин"],
    "стекло":              ["стекл", "зеркал"],
    "пол":                 ["пол", "паркет"],
    "мебель_ковры":        ["мебел", "ковр"],
    "трубы":               ["труб", "стоков"],
    "унитаз":              ["унитаз", "туалет"],
    "сантехника":          ["сантехник", "ванн"],
    "освежитель_воздуха":  ["освежит", "ароматиз", "воздух", "дом"],  # дом → «Товары для дома»
    "насекомые":           ["насеком", "инсектицид", "моль", "репелент", "дом"],
    "шампунь":             ["шампун"],
    "кондиционер_волос":   ["волос", "бальзам", "кондиционер"],
    "средства_душ":        ["душ", "скраб", "тел"],       # тел → «ТЕЛО»
    "мыло":                ["мыло"],
    "зубы":                ["зуб", "полост рта", "oral"], # oral → «ORAL CARE»
    "дезодорант":          ["дезодорант", "антиперспирант", "тел"],
    "депиляция":           ["депиляц", "тел"],
    "бритье":              ["бритв", "тел"],               # тел → «ТЕЛО»
    "уход_лицо":           ["лиц"],                        # отдельный кластер для «ЛИЦО»
    "уход_тело":           ["тел"],                        # отдельный кластер для «ТЕЛО»
    "детская_гигиена":     ["детск", "детств", "подгузник", "ребенк"],
    "бумага":              ["бумаг", "салфетк", "платок"],
    "женская_гигиена":     ["прокладк", "тампон", "гигиен", "тел"],
    "загар":               ["загар", "солнц", "тел"],
    "универсальное":       ["универсальн"],
    "антисептики":         ["антисептик", "антибактер"],
}

# ── SOURCE CATEGORY → CLUSTER MAP ─────────────────────────────────────────────
# Сопоставляет подстроки нормализованных source-категорий (category_1..4)
# с именами кластеров; используется в Pass 1 когда Jaccard даёт низкий score.
# Порядок важен: более специфичные паттерны — выше.
_SRC_CAT_MAP: list[tuple[str, str]] = [
    # ── Посуда ──────────────────────────────────────────────────────────────
    ("автоматического мытья",  "посуда_пмм"),
    ("посудомоечн",            "посуда_пмм"),
    ("ручного мытья",          "посуда_ручная"),
    ("мытья посуд",            "посуда_ручная"),
    # ── Стирка ──────────────────────────────────────────────────────────────
    ("для стирк",              "стирка"),
    ("стиральн",               "стирка"),
    ("ополаскив",              "ополаскиватель_белья"),
    ("пятновывод",             "пятновыводитель"),
    ("отбелив",                "отбеливатель"),
    ("ухода за одежд",         "уход_за_одеждой"),
    ("ухода за обув",          "уход_за_одеждой"),
    # ── Уборка ──────────────────────────────────────────────────────────────
    ("уборки",                 "чистящее"),
    ("чистящ",                 "чистящее"),
    ("накип",                  "накипь"),
    ("ржавчин",                "ржавчина"),
    ("для стекл",              "стекло"),
    ("для пол",                "пол"),
    ("для ковр",               "мебель_ковры"),
    ("прочист труб",           "трубы"),
    # ── Санузел ─────────────────────────────────────────────────────────────
    ("для сантехник",          "сантехника"),
    ("для унитаз",             "унитаз"),
    ("блок унитаз",            "унитаз"),
    ("сантехник",              "сантехника"),
    # ── Воздух / насекомые ──────────────────────────────────────────────────
    ("освежит",                "освежитель_воздуха"),
    ("ароматиз",               "освежитель_воздуха"),
    ("насеком",                "насекомые"),
    ("инсектицид",             "насекомые"),
    # ── Волосы ──────────────────────────────────────────────────────────────
    ("уход за волос",          "кондиционер_волос"),
    ("шампун",                 "шампунь"),
    ("волос",                  "кондиционер_волос"),
    # ── Зубы / рот ──────────────────────────────────────────────────────────
    ("полости рта",            "зубы"),
    ("зубами",                 "зубы"),
    ("зубн",                   "зубы"),
    # ── Кожа лица (специфично выше тела) ───────────────────────────────────
    ("кожи лица",              "уход_лицо"),
    ("лица",                   "уход_лицо"),
    # ── Кожа тела / личная гигиена ──────────────────────────────────────────
    ("кожи тела",              "уход_тело"),
    ("кожи рук",               "уход_тело"),
    ("тела",                   "уход_тело"),
    ("для душа",               "средства_душ"),
    ("дезодорант",             "дезодорант"),
    ("антиперспирант",         "дезодорант"),
    ("брить",                  "бритье"),           # брить → бритье/бриться
    ("бритв",                  "бритье"),           # бритвы/бритвенн
    ("депиляц",                "депиляция"),
    ("загар",                  "загар"),
    ("женск",                  "женская_гигиена"),  # женская/женских (любая форма)
    ("прокладк",               "женская_гигиена"),
    ("тампон",                 "женская_гигиена"),
    # ── Мыло ────────────────────────────────────────────────────────────────
    ("жидкого мыла",           "мыло"),
    ("личной гигиен",          "мыло"),
    ("мыл",                    "мыло"),
    # ── Антисептики ─────────────────────────────────────────────────────────
    ("антисептик",             "антисептики"),
    ("антибактер",             "антисептики"),
    # ── Дети ────────────────────────────────────────────────────────────────
    ("подгузник",              "детская_гигиена"),
    ("детьм",                  "детская_гигиена"),  # «ухода за детьми»
    ("для детей",              "детская_гигиена"),
    ("детских",                "детская_гигиена"),
    ("детской",                "детская_гигиена"),
    # ── Бумага / салфетки ───────────────────────────────────────────────────
    ("бумажн",                 "бумага"),
    ("салфетк",                "бумага"),
    ("носов платк",            "бумага"),
]


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

    1a. Substring check: target norm is verbatim in src norm (e.g. "пмм").
    1b. SRC_CAT_MAP: source text contains a known domain phrase → cluster →
        best matching target category (handles morphological mismatches).
    1c. Weighted Jaccard on tokens as fallback.
    """
    src_combined = " ".join(src_cats)
    src_norm     = _norm(src_combined)
    src_toks     = _tokens(src_combined)

    # ── 1a. verbatim substring of target name ────────────────────────────────
    for cat, m in cat_meta.items():
        if m["norm"] and m["norm"] in src_norm:
            return cat, 0.95

    # ── 1b. SRC_CAT_MAP: domain phrase → cluster → best target ───────────────
    matched_clusters: set[str] = set()
    for phrase, cluster in _SRC_CAT_MAP:
        if phrase in src_norm:
            matched_clusters.add(cluster)

    if matched_clusters:
        scores: dict[str, float] = {}
        for cat, m in cat_meta.items():
            overlap = matched_clusters & m["clusters"]
            if overlap:
                scores[cat] = len(overlap) / max(len(matched_clusters), len(m["clusters"]), 1)
        if scores:
            best = max(scores, key=scores.get)
            conf = round(min(0.60 + 0.35 * scores[best], 0.93), 2)
            return best, conf

    # ── 1c. weighted Jaccard on tokens ───────────────────────────────────────
    best_cat, best_score = "", 0.0
    for cat, m in cat_meta.items():
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
        if k in overrides:
            continue
        # Skip cache only when it has a non-empty result; empty = try again
        if k in cache and cache[k].get("target_category", ""):
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
