#!/usr/bin/env python3
"""
pipeline.py — Anonymize + merge retail tables in one pass
──────────────────────────────────────────────────────────
Reads ALL source files from input_folder, anonymizes each,
merges into one xlsx, then moves originals to archive_folder.

Required:  pip install pandas openpyxl
Optional:  pip install faker    (better name generation)
           pip install xlrd     (.xls support)
           pip install pyxlsb   (.xlsb support)

CLI:     python pipeline.py
Dagster: see dagster_job.py — imports core functions from this module
"""

import sys
import json
import time
import hashlib
import random
import difflib
import shutil
import re
from collections import Counter, defaultdict
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

_Faker = None  # type: ignore[assignment]
try:
    from faker import Faker as _Faker  # type: ignore[assignment]
    HAS_FAKER = True
except ImportError:
    HAS_FAKER = False


# ── CONSTANTS ──────────────────────────────────────────────────────────────────

CONFIG_PATH    = Path("config.json")
MAPPINGS_PATH  = Path("mappings.json")
PRODUCT_MATCHES_PATH = Path("product_matches.json")   # ← кэш матчинга товаров
PRICE_COEF     = 1.117
QTY_SHIFT      = 2
SUPPORTED_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".tsv"}

CANONICAL = {
    "week":                "keep",
    "month":               "keep",
    "year":                "keep",
    "store_code":          "drop",
    "store_name":          "anon_store",
    "address":             "anon_store",
    "distribution_center": "anon_company",
    "region":              "keep",
    "store_format":        "keep",
    "store_subformat":     "keep",
    "category_0":          "keep",
    "category_1":          "keep",
    "category_2":          "keep",
    "category_3":          "keep",
    "category_4":          "keep",
    "item_code":           "drop",
    "item_name":           "anon_product",
    "brand":               "anon_brand",
    "manufacturer":        "anon_company",
    "supplier":            "anon_company",
    "weight":              "keep",
    "unit":                "keep",
    "barcode":             "drop",
    "qty_sold":            "add_qty",
    "sales_rub":           "mul_price",
    "cost_rub":            "mul_price",
}

DESIRED_ORDER = [
    "week", "month", "year",
    "store_name", "address", "region", "store_format", "store_subformat",
    "distribution_center",
    "category_0", "category_1", "category_2", "category_3", "category_4",
    "item_name", "brand", "manufacturer", "supplier", "weight", "unit",
    "qty_sold", "sales_rub", "cost_rub",
]

ALIASES = {
    "week":                ["неделя", "нед", "week", "wk", "неделя продаж"],
    "month":               ["месяц", "мес", "month", "mo"],
    "year":                ["год", "year", "yr", "год продаж"],
    "store_code":          ["код", "код магазина", "код тт", "store_code",
                            "id магазина", "код торговой точки"],
    "store_name":          ["название магазина", "магазин", "наим магазина",
                            "тт", "store", "store_name", "торговая точка"],
    "address":             ["адрес", "address", "addr", "адрес магазина"],
    "distribution_center": ["основной рц", "рц", "distribution center",
                            "distribution_center", "распределительный центр",
                            "основной распределительный центр", "склад", "dc"],
    "region":              ["регион", "region", "reg", "регион продаж",
                            "территория", "регион тт", "область"],
    "store_format":        ["формат", "format", "тип магазина",
                            "формат магазина", "формат тт"],
    "store_subformat":     ["субформат", "subformat", "sub-format",
                            "подформат", "суб-формат", "тип тт"],
    "category_0":          ["уровень 0", "уровень0", "level 0", "category_0",
                            "cat0", "lvl0", "категория 0"],
    "category_1":          ["уровень 1", "уровень1", "категория 1", "level 1",
                            "category_1", "cat1", "lvl1"],
    "category_2":          ["уровень 2", "уровень2", "категория 2", "level 2",
                            "category_2", "cat2", "lvl2"],
    "category_3":          ["уровень 3", "уровень3", "категория 3", "level 3",
                            "category_3", "cat3", "lvl3"],
    "category_4":          ["уровень 4", "уровень4", "категория 4", "level 4",
                            "category_4", "cat4", "lvl4"],
    "item_code":           ["код позиции", "код товара", "sku", "item_code",
                            "артикул", "арт", "код поз"],
    "item_name":           ["наименование", "название товара", "товар",
                            "позиция", "item_name", "item", "наим"],
    "brand":               ["бренд", "brand", "марка", "торговая марка"],
    "manufacturer":        ["производитель", "пр-ль", "manufacturer", "mfr",
                            "producer", "изготовитель"],
    "supplier":            ["поставщик", "supplier", "spl", "поставщ",
                            "дистрибьютор"],
    "weight":              ["вес", "weight", "масса", "вес нетто",
                            "net weight", "вес брутто", "масса нетто"],
    "unit":                ["единица измерения", "ед. изм.", "ед.изм.", "unit",
                            "units", "ед изм", "единица", "ед"],
    "barcode":             ["штриховой код", "штрихкод", "ean", "barcode",
                            "bar code", "код штрих", "баркод"],
    "qty_sold":            ["продажи в шт", "продажи в шт.", "qty", "qty_sold",
                            "шт", "количество", "кол-во", "продажи в штуках"],
    "sales_rub":           ["продажи в руб", "продажи в руб.", "sales_rub",
                            "выручка", "оборот", "продажи в рублях"],
    "cost_rub":            ["себестоимость", "себестоимсть",
                            "себестоимость в руб", "себестоимсть в руб",
                            "cost_rub", "cost", "себес", "закупка",
                            "закупочная цена"],
}


# ── WORD POOLS ─────────────────────────────────────────────────────────────────

_BRAND_PREFIXES = [
    "АРК","БЛЕ","ВЕЛ","ГРИ","ДИВ","ЗЕН","КЛИ","ЛЮМ","МИР","НОР",
    "ОМЕ","ПРИ","РАД","СВЕ","ТИП","УЛЬ","ФОР","ХОР","ЮГА","ЯРО",
    "АЛЬ","БОР","ВОЛ","ГАЛ","ДОН","ЖАР","ЗОР","ИВА","КАМ","ЛАД",
    "МАГ","НАД","ОРА","ПАЛ","РЯД","СИЛ","ТАЛ","УРА","ФАЛ","ЦЕН",
    "ЧИС","ШАЛ","ЩЕД","ЭЛЬ","ЮНА","ЯСЕ","АВА","БАЛ","ВАЛ","ГОР",
]
_BRAND_SUFFIXES = [
    "А","ЕКС","ИТА","ОН","ОС","УМ","ЕТ","АН","ЕЛ","ЕВА",
    "ИН","АР","ОР","ЕЖ","ОВА","ИМ","ЕР","АЛ","ИС","ЮМ",
]
_STORE_WORDS = [
    "Рябина","Берёза","Клён","Сосна","Ясень","Дуб","Каштан","Ива","Кедр","Тополь",
    "Заря","Рассвет","Весна","Маяк","Радуга","Горизонт","Меридиан","Сфера","Орбита",
    "Вектор","Импульс","Прогресс","Полюс","Сигма","Феникс","Атлант","Аврора","Ладья",
    "Родник","Исток","Утёс","Бриз","Прибой","Янтарь","Топаз","Рубин","Агат","Оникс",
    "Кварц","Малахит","Опал","Жемчуг","Сапфир","Изумруд","Гранит","Базальт","Мрамор",
    "Ветер","Буря","Гроза","Туман","Иней","Метель","Поток","Якорь","Штурвал","Компас",
    "Фрегат","Бриг","Колос","Росток","Лепесток","Бутон","Нептун","Борей","Зефир",
    "Дубрава","Роща","Поляна","Опушка","Апекс","Зенит","Вершина","Долина","Луга",
    "Курган","Равнина","Перекрёсток","Рубеж","Форт","Цитадель","Твердыня",
    "Авангард","Арьергард","Резерв","Запас","Клад","Сокровище","Копилка",
]
_CO_FIRST  = [
    "Центр","Север","Юг","Восток","Запад","Главный","Новый","Первый",
    "Элит","Прайм","Мега","Гранд","Альфа","Бета","Омега","Кристалл",
    "Союз","Стандарт","Профи","Базис","Контур","Спектр","Горизонт","Форс",
    "Пром","Агро","Техно","Логис","Дата","Инфо","Медиа","Рапид",
]
_CO_SECOND = [
    "Торг","Пром","Продукт","Ресурс","Групп","Логистик","Поставк",
    "Холдинг","Инвест","Сервис","Маркет","Снаб","Трейд","Партнер",
    "Дистриб","Агент","Брокер","Посред","Интер","Экспо","Конс","Финанс",
    "Опт","Ритейл","Дилер","Транс","Экспрес","Систем","Решен","Технол",
]
_CO_LEGAL = ["ООО","АО","ЗАО","ПАО"]
_REGIONS = [
    "Ростовская обл.","Краснодарский край","Воронежская обл.",
    "Волгоградская обл.","Белгородская обл.","Ставропольский край",
    "Тамбовская обл.","Саратовская обл.","Липецкая обл.",
    "Курская обл.","Орловская обл.","Тульская обл.","Рязанская обл.",
    "Пензенская обл.","Ульяновская обл.","Самарская обл.","Нижегородская обл.",
]
_CITIES = [
    "Таганрог","Шахты","Батайск","Волгодонск","Азов","Зверево","Гуково",
    "Каменск-Шахтинский","Сальск","Морозовск","Цимлянск","Константиновск",
    "Аксай","Кропоткин","Ейск","Тихорецк","Армавир","Белая Калитва","Миллерово",
    "Семикаракорск","Новоалександровск","Невинномысск","Будённовск","Лермонтов",
    "Михайловск","Благодарный","Апшеронск","Горячий Ключ","Тимашевск","Темрюк",
    "Геленджик","Анапа","Туапсе","Абинск","Славянск-на-Кубани","Кореновск",
]
_STREET_TYPES = ["улица","переулок","проспект","бульвар","площадь","набережная","шоссе"]
_STREET_NAMES = [
    "Садовая","Московская","Ленинская","Центральная","Советская","Строительная",
    "Молодёжная","Школьная","Лесная","Полевая","Мирная","Победы","Комсомольская",
    "Гагарина","Пушкина","Чехова","Горького","Кирова","Дружбы","Заречная",
    "Набережная","Рабочая","Колхозная","Крестьянская","Революционная","Октябрьская",
    "Первомайская","Красная","Зелёная","Берёзовая","Сосновая","Луговая","Степная",
]
_DESCRIPTORS = {
    "посуд":     ["Средство для мытья посуды","Гель для посуды","Жидкость для посуды","Бальзам для посуды"],
    "чистящ":    ["Крем чистящий","Средство чистящее","Порошок чистящий","Паста чистящая"],
    "стирк":     ["Гель для стирки","Жидкость для стирки","Порошок стиральный","Капсулы для стирки"],
    "унитаза":   ["Блок для унитаза","Таблетка для унитаза","Гель для унитаза","Диск для унитаза"],
    "сантехник": ["Средство для сантехники","Очиститель сантехники","Гель для ванной"],
    "уборк":     ["Средство для уборки","Спрей для уборки","Универсальный очиститель"],
    "освеж":     ["Освежитель воздуха","Спрей-освежитель","Ароматизатор воздуха"],
    "ополаск":   ["Ополаскиватель","Кондиционер для белья","Бальзам-ополаскиватель"],
    "дезинфект": ["Дезинфицирующее средство","Антисептик","Дезинфектант"],
    "пятновы":   ["Пятновыводитель","Отбеливатель","Средство от пятен"],
}
_VARIANTS = [
    "Лаванда","Лимон","Морозная свежесть","Зелёный чай","Яблоко","Хвоя",
    "Ромашка","Океан","Мята","Роза","Кедр","Цитрус","Лайм","Ваниль",
    "Жасмин","Ландыш","Морской бриз","Альпийский луг","Утренняя роса","Хлопок",
]
_SIZES = [
    "500мл","750мл","1000мл","1.3кг","2.4кг","600г","3х50г","450мл","800мл",
    "1.5кг","400мл","2кг","1.2л","350мл","1800г","900мл","5л","200мл","300мл",
]


# ── GENERATORS ─────────────────────────────────────────────────────────────────

def _seed(key: str) -> int:
    return int(hashlib.md5(str(key).encode("utf-8")).hexdigest()[:8], 16)

def _rng(key: str) -> random.Random:
    return random.Random(_seed(key))

def _faker(key: str):
    if not HAS_FAKER:
        return None
    f = _Faker("ru_RU")
    f.seed_instance(_seed(key))
    return f

def gen_store(key: str, used_names: set) -> dict:
    r = _rng(key)
    f = _faker(key)
    base = r.choice(_STORE_WORDS)
    name, n = base, 1
    while name in used_names:
        name = base + str(n); n += 1
    address = ""  # always initialized; overwritten below
    if f:
        try:
            address = f"{f.region()}, {f.city()}, {f.street_name()}, {f.building_number()}"
        except Exception:
            f = None
    if not f or not address:
        address = (f"{r.choice(_REGIONS)}, {r.choice(_CITIES)}, "
                   f"{r.choice(_STREET_NAMES)} {r.choice(_STREET_TYPES)}, {r.randint(1, 200)}")
    return {"name": name, "address": address}

def gen_brand(key: str, used: set) -> str:
    r = _rng(key)
    base = r.choice(_BRAND_PREFIXES) + r.choice(_BRAND_SUFFIXES)
    cand, n = base, 1
    while cand in used:
        cand = base + str(n); n += 1
    return cand

def gen_company(key: str, used: set = None) -> str:
    f = _faker(key)
    if f:
        try:
            base = f.company()
            if used is None:
                return base
            cand, n = base, 1
            while cand in used:
                cand = f"{base} {n}"; n += 1
            return cand
        except Exception:
            pass
    r = _rng(key)
    base = f"{r.choice(_CO_FIRST)}{r.choice(_CO_SECOND)} {r.choice(_CO_LEGAL)}"
    if used is None:
        return base
    cand, n = base, 1
    while cand in used:
        cand = base + str(n); n += 1
    return cand

def gen_product(key: str, category: str, fake_brand: str) -> str:
    r = _rng(key)
    cat = str(category).lower()
    descriptor = next(
        (r.choice(opts) for kw, opts in _DESCRIPTORS.items() if kw in cat),
        "Средство",
    )
    return f"{fake_brand} {descriptor} {r.choice(_VARIANTS)} {r.choice(_SIZES)}"


# ── PRODUCT MATCHING (runs BEFORE anonymization) ──────────────────────────────
#
# Problem: store A calls it "Fairy Ориджинал 500мл", store B calls it
#          "Фейри Original 0.5л посуда" → after anonymization they get
#          different fake names → impossible to aggregate.
#
# Solution: two-pass pipeline.
#   Pass 1 – read all raw files, collect item_names.
#   Matching – map variants to a single canonical real name:
#     Stage A – barcode-based (exact EAN match, zero false-positives).
#     Stage B – TF-IDF char n-gram cosine similarity on normalized names
#               (strips sizes/units so "500мл" vs "0.5л" stops mattering).
#   Pass 2 – replace item_name in every DF with canonical, then anonymize.
#             Now all variants share ONE fake name and can be grouped.
#
# Cache: product_matches.json stores {original → canonical} for non-trivial
#        pairs only.  Human-reviewable and editable between runs.
# ──────────────────────────────────────────────────────────────────────────────

_SIZE_RE = re.compile(
    r'\b\d+[.,]?\d*\s*'
    r'(мл|л|г|гр|кг|шт|уп|мг|таб|капс|мм|см|ml|l|g|kg|pcs|pc|oz|fl\.oz)\b',
    re.IGNORECASE,
)
_NUM_RE  = re.compile(r'\b\d+\b')
_JUNK_RE = re.compile(r'[^\w\s]')
_WS_RE   = re.compile(r'\s+')


_SIZE_EXTRACT_RE = re.compile(
    r'(\d+[,.]?\d*)\s*(мл|л|г|гр|кг|шт|уп|мг|ml|l|g|kg|pcs|pc)\b',
    re.IGNORECASE,
)


def _norm_for_match(name: str) -> str:
    """Normalize product name for similarity comparison."""
    s = str(name).lower().strip()
    s = _SIZE_RE.sub(' ', s)      # "500мл" → " "
    s = _NUM_RE.sub(' ', s)       # orphan numbers
    s = _JUNK_RE.sub(' ', s)      # punctuation
    s = _WS_RE.sub(' ', s).strip()
    return s


def _extract_primary_size(name: str) -> str:
    """
    Extract and normalize the first size/volume/count token from a product name.
    Returns canonical form: '500ml', '1000g', '20pcs', etc.
    Returns '' if no size found.

    Used as a gate: products with different sizes are never merged,
    even when their names are otherwise similar (different pack = different SKU).

    Examples:
        "Fairy 500мл"       → '500ml'
        "Fairy 0.5л"        → '500ml'   # same as above → allowed to match
        "Fairy 1л"          → '1000ml'  # different → blocked
        "Прокладки 10шт"    → '10pcs'
        "Прокладки 20шт"    → '20pcs'   # different → blocked
    """
    m = _SIZE_EXTRACT_RE.search(name)
    if not m:
        return ''
    try:
        val = float(m.group(1).replace(',', '.'))
    except ValueError:
        return ''
    unit = m.group(2).lower()
    if unit in ('л', 'l'):
        return f"{int(round(val * 1000))}ml"
    if unit in ('мл', 'ml'):
        return f"{int(round(val))}ml"
    if unit in ('кг', 'kg'):
        return f"{int(round(val * 1000))}g"
    if unit in ('г', 'гр', 'g'):
        return f"{int(round(val))}g"
    if unit in ('мг',):
        return f"{int(round(val))}mg"
    if unit in ('шт', 'уп', 'pcs', 'pc'):
        return f"{int(round(val))}pcs"
    return f"{int(round(val))}{unit}"


def load_product_matches() -> dict:
    """Load cached {original_name: canonical_name} from product_matches.json."""
    if not PRODUCT_MATCHES_PATH.exists():
        return {}
    return json.loads(PRODUCT_MATCHES_PATH.read_text(encoding="utf-8"))


def save_product_matches(matches: dict):
    """Persist only non-trivial matches (where canonical differs from original)."""
    non_trivial = {k: v for k, v in matches.items() if k != v}
    PRODUCT_MATCHES_PATH.write_text(
        json.dumps(non_trivial, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _union_find(n: int, pairs: list[tuple[int, int]]) -> dict[int, list[int]]:
    """Union-Find: returns {root_idx: [member_indices]}."""
    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i, j in pairs:
        pi, pj = find(i), find(j)
        if pi != pj:
            parent[pi] = pj

    clusters: dict[int, list[int]] = defaultdict(list)
    for i in range(n):
        clusters[find(i)].append(i)
    return dict(clusters)


def _barcode_matches(
    raw_frames: list,
    log_fn: Callable[[str], None] = print,
) -> dict:
    """
    Stage A: barcode → canonical name.
    Same EAN in different stores = same product, zero ambiguity.
    Returns {variant_name: canonical_name} for multi-name barcodes only.
    """
    all_names: list[str] = []
    barcode_to_names: dict = defaultdict(list)

    for _, df, col in raw_frames:
        ic: str | None = col.get("item_name")
        bc: str | None = col.get("barcode")
        if ic and ic in df.columns:
            all_names.extend(df[ic].dropna().astype(str).str.strip().tolist())
        if bc and ic and bc in df.columns and ic in df.columns:
            ic_col, bc_col = str(ic), str(bc)   # narrow str|None → str for type checker
            sub = df[[bc_col, ic_col]].dropna()
            for _, row in sub.iterrows():
                bv = str(row[bc_col]).strip()
                nv = str(row[ic_col]).strip()
                if bv and bv not in ("nan", "0", ""):
                    barcode_to_names[bv].append(nv)

    freq = Counter(all_names)
    result: dict[str, str] = {}
    n_groups = 0

    for names in barcode_to_names.values():
        unique = list(set(names))
        if len(unique) > 1:
            canonical = max(unique, key=lambda nm: freq.get(nm, 0))
            for name in unique:
                if name != canonical:
                    result[name] = canonical
            n_groups += 1

    if n_groups:
        log_fn(f"  [Штрихкод] объединено групп: {n_groups}")
    return result


def _fuzzy_matches(
    item_names: list[str],
    existing: dict,
    threshold: float,
    log_fn: Callable[[str], None] = print,
) -> dict:
    """
    Stage B: fuzzy name similarity for products not covered by barcode matching.
    Priority: sklearn TF-IDF (best) → rapidfuzz token_set_ratio → difflib (built-in).
    """
    all_unique = list({str(n).strip() for n in item_names
                       if pd.notna(n) and str(n).strip()})
    new_names  = [n for n in all_unique if n not in existing]

    if not new_names:
        log_fn(f"  [Fuzzy] все {len(all_unique)} товаров уже в кэше")
        return existing

    if len(new_names) < 2:
        result = dict(existing)
        result[new_names[0]] = new_names[0]
        return result

    log_fn(f"  [Fuzzy] матчинг {len(new_names)} новых товаров (порог={threshold})")
    norms = [_norm_for_match(n) for n in new_names]
    freq  = Counter(str(n).strip() for n in item_names if pd.notna(n))
    pairs: list[tuple[int, int]] = []
    method = "difflib"  # default; overwritten below when better lib found

    # ── Detect available library once, then run ──────────────────────────────
    _sklearn_ok   = False
    _rapidfuzz_ok = False
    try:
        import sklearn.feature_extraction.text  # noqa: F401
        import numpy                            # noqa: F401
        _sklearn_ok = True
    except ImportError:
        pass

    if not _sklearn_ok:
        try:
            import rapidfuzz  # noqa: F401
            _rapidfuzz_ok = True
        except ImportError:
            pass

    if _sklearn_ok:
        # ── sklearn TF-IDF (recommended) ─────────────────────────────────────
        from sklearn.feature_extraction.text import TfidfVectorizer
        import numpy as np

        vec   = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=1)
        mat   = vec.fit_transform(norms)
        n_items = len(new_names)
        batch   = 500

        for start in range(0, n_items, batch):
            end   = min(start + batch, n_items)
            block = (mat[start:end] @ mat.T).toarray()
            for local_i in range(end - start):
                global_i = start + local_i
                js: list[int] = [int(j) for j in np.where(block[local_i] >= threshold)[0]
                                 if int(j) > global_i]
                pairs.extend((global_i, j) for j in js)

        method = "TF-IDF/sklearn"

    elif _rapidfuzz_ok:
        # ── rapidfuzz ────────────────────────────────────────────────────────
        from rapidfuzz import fuzz as _fuzz

        for i in range(len(new_names)):
            for j in range(i + 1, len(new_names)):
                if _fuzz.token_set_ratio(norms[i], norms[j]) / 100.0 >= threshold:
                    pairs.append((i, j))

        method = "rapidfuzz"

    else:
        # ── difflib fallback (built-in, O(N²), slow on large sets) ───────────
        for i in range(len(new_names)):
            for j in range(i + 1, len(new_names)):
                if difflib.SequenceMatcher(None, norms[i], norms[j]).ratio() >= threshold:
                    pairs.append((i, j))

        method = "difflib (медленно; рекомендуется: pip install scikit-learn)"

    log_fn(f"  [Fuzzy] метод: {method}, найдено пар до фильтра: {len(pairs)}")

    # ── Size gate: разный объём/количество = разный SKU, не матчить ──────────
    # Normalizes "0.5л" == "500мл" so unit variants still match,
    # but "500мл" ≠ "1000мл" and "10шт" ≠ "20шт" are blocked.
    sizes = [_extract_primary_size(n) for n in new_names]
    pairs = [(i, j) for i, j in pairs if sizes[i] == sizes[j]]
    log_fn(f"  [Fuzzy] после фильтра по объёму: {len(pairs)} пар")

    clusters = _union_find(len(new_names), pairs)
    result   = dict(existing)
    n_merged = 0

    for indices in clusters.values():
        group     = [new_names[i] for i in indices]
        canonical = max(group, key=lambda nm: freq.get(nm, 0))
        for nm in group:
            result[nm] = canonical
        if len(group) > 1:
            others = [nm for nm in group if nm != canonical]
            log_fn(f"    ✓ [{canonical}] ← {others}")
            n_merged += 1

    log_fn(f"  [Fuzzy] объединено групп: {n_merged}")
    return result


def run_product_matching(
    raw_frames: list,
    threshold: float = 0.82,
    log_fn: Callable[[str], None] = print,
) -> dict:
    """
    Full two-stage product matching.
    Returns {original_name: canonical_name} (complete mapping).
    Loads existing cache, processes only new names, saves back.
    """
    matches = load_product_matches()

    # Stage A: barcode
    bc_m = _barcode_matches(raw_frames, log_fn)
    matches.update(bc_m)

    # Collect all item_names (applying barcode fixes already)
    all_names: list[str] = []
    for _, df, col in raw_frames:
        ic = col.get("item_name")
        if ic and ic in df.columns:
            names = df[ic].dropna().astype(str).str.strip().tolist()
            all_names.extend(matches.get(n, n) for n in names)

    # Stage B: fuzzy name
    if all_names:
        matches = _fuzzy_matches(all_names, matches, threshold, log_fn)

    save_product_matches(matches)
    return matches


def apply_product_matches(df: pd.DataFrame, col: dict, matches: dict) -> pd.DataFrame:
    """Replace item_name values with their canonical names in-place."""
    ic = col.get("item_name")
    if ic and ic in df.columns:
        df[ic] = df[ic].astype(str).str.strip().map(lambda n: matches.get(n, n))
    return df


# ── FILE I/O ───────────────────────────────────────────────────────────────────

def read_table(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(path)
    elif ext == ".xls":
        return pd.read_excel(path, engine="xlrd")
    elif ext == ".xlsb":
        return pd.read_excel(path, engine="pyxlsb")
    elif ext == ".csv":
        sample = path.read_text(encoding="utf-8", errors="replace")[:4096]
        sep = ";" if sample.count(";") > sample.count(",") else ","
        return pd.read_csv(path, sep=sep)
    elif ext == ".tsv":
        return pd.read_csv(path, sep="\t")
    raise ValueError(f"Неподдерживаемый формат: {ext}")

def _format_xlsx(path: Path):
    wb = load_workbook(path)
    ws = wb.active
    if ws is None:
        return
    header_fill = PatternFill("solid", start_color="1F4E79")
    for cell in ws[1]:
        cell.font      = Font(bold=True, name="Arial", size=10, color="FFFFFF")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 38
    ws.freeze_panes = "A2"
    for ci in range(1, ws.max_column + 1):
        col_letter = get_column_letter(ci)
        maxlen = max(
            (len(str(ws.cell(row=r, column=ci).value or ""))
             for r in range(1, min(ws.max_row + 1, 200))),
            default=8,
        )
        ws.column_dimensions[col_letter].width = min(maxlen + 3, 52)
    wb.save(path)

def write_xlsx(df: pd.DataFrame, path: Path):
    df.to_excel(path, index=False)
    _format_xlsx(path)


# ── MAPPINGS ───────────────────────────────────────────────────────────────────

def load_mappings() -> dict:
    defaults = {"stores": {}, "brands": {}, "companies": {}, "products": {}}
    if not MAPPINGS_PATH.exists():
        return defaults
    data = json.loads(MAPPINGS_PATH.read_text(encoding="utf-8"))
    if "companies" not in data:
        data["companies"] = {**data.pop("manufacturers", {}), **data.pop("suppliers", {})}
    return {**defaults, **data}

def save_mappings(m: dict):
    MAPPINGS_PATH.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")

def build_mappings(df: pd.DataFrame, col: dict, m: dict, log_fn: Callable[[str], None]) -> dict:
    """col: {canonical_name → actual_df_column_name}"""
    key_col = col.get("store_code") or col.get("store_name")
    used_names = {v["name"] for v in m["stores"].values()}
    new_s = 0
    if key_col and key_col in df.columns:
        for val in sorted(df[key_col].dropna().unique()):
            k = str(int(val)) if isinstance(val, float) else str(val).strip()
            if k not in m["stores"]:
                store = gen_store(k, used_names)
                m["stores"][k] = store
                used_names.add(store["name"])
                new_s += 1
    log_fn(f"  Магазины: +{new_s} новых (кэш: {len(m['stores'])})")

    bc = col.get("brand")
    used_brands = set(m["brands"].values())
    new_b = 0
    if bc and bc in df.columns:
        for brand in sorted(df[bc].dropna().unique()):
            k = str(brand).strip()
            if k not in m["brands"]:
                m["brands"][k] = gen_brand(k, used_brands)
                used_brands.add(m["brands"][k])
                new_b += 1
    log_fn(f"  Бренды: +{new_b} новых (кэш: {len(m['brands'])})")

    new_c = 0
    used_co = set(m["companies"].values())
    for ckey in ["manufacturer", "supplier", "distribution_center"]:
        c = col.get(ckey)
        if c and c in df.columns:
            for val in sorted(df[c].dropna().unique()):
                k = str(val).strip()
                if k not in m["companies"]:
                    m["companies"][k] = gen_company(k, used_co)
                    used_co.add(m["companies"][k])
                    new_c += 1
    log_fn(f"  Компании: +{new_c} новых (кэш: {len(m['companies'])})")

    ic   = col.get("item_name")
    cat2 = col.get("category_2")
    bc2  = col.get("brand")
    new_p = 0
    if ic and ic in df.columns:
        needed = [c for c in [ic, bc2, cat2] if c and c in df.columns]
        subset = df[needed].drop_duplicates().dropna(subset=[ic])
        subset = subset[~subset[ic].astype(str).isin(m["products"])]
        for _, row in subset.iterrows():
            k  = str(row[ic]).strip()
            fb = m["brands"].get(str(row.get(bc2, "")).strip(), "БРЕНД") if bc2 else "БРЕНД"
            cat = str(row.get(cat2, "")) if cat2 else ""
            m["products"][k] = gen_product(k, cat, fb)
            new_p += 1
    log_fn(f"  Товары: +{new_p} новых (кэш: {len(m['products'])})")
    return m


# ── COLUMN MAPPING (fuzzy only, no external API) ───────────────────────────────

def _normalize(col: str) -> str:
    s = col.lower().strip()
    for suffix in [" с ндс", " без ндс", ". с ндс", " с nds", " incl. vat", " excl. vat"]:
        if s.endswith(suffix):
            s = s[:-len(suffix)]
    return s.rstrip(". ,")

def fuzzy_map(actual_cols: list) -> dict:
    """Returns {actual_col: canonical_name}"""
    alias_to_canon = {
        alias: cname
        for cname, aliases in ALIASES.items()
        for alias in aliases
    }
    result = {}
    for col in actual_cols:
        norm = _normalize(col)
        if norm in alias_to_canon:
            result[col] = alias_to_canon[norm]
            continue
        matches = difflib.get_close_matches(norm, alias_to_canon.keys(), n=1, cutoff=0.72)
        if matches:
            result[col] = alias_to_canon[matches[0]]
    return result

def get_col_map(df: pd.DataFrame) -> dict:
    """Returns {canonical_name → actual_df_column}  (fuzzy matching only)"""
    raw = fuzzy_map(list(df.columns))          # {actual: canonical}
    canon_to_actual: dict = {}
    for act_col, cname in raw.items():
        if cname and cname not in canon_to_actual:
            canon_to_actual[cname] = act_col
    return canon_to_actual


# ── TRANSFORM ──────────────────────────────────────────────────────────────────

def transform(df: pd.DataFrame, col: dict, m: dict) -> pd.DataFrame:
    """Apply all transforms and rename columns to English canonical names."""
    sc   = col.get("store_code")
    sn_c = col.get("store_name")
    sa_c = col.get("address")
    key_col = sc if (sc and sc in df.columns) else sn_c

    if key_col and key_col in df.columns:
        store_name_map: dict = {}
        store_addr_map: dict = {}
        for k, v in m["stores"].items():
            try:
                store_name_map[int(k)] = v["name"]
                store_addr_map[int(k)] = v["address"]
            except ValueError:
                store_name_map[k] = v["name"]
                store_addr_map[k] = v["address"]
        if sn_c and sn_c in df.columns:
            df[sn_c] = df[key_col].map(store_name_map)
        if sa_c and sa_c in df.columns:
            df[sa_c] = df[key_col].map(store_addr_map)

    for cname, action in CANONICAL.items():
        actual = col.get(cname)
        if not actual or actual not in df.columns:
            continue
        if action in ("anon_store", "drop"):
            continue
        elif action == "anon_brand":
            df[actual] = df[actual].astype(str).str.strip().map(m["brands"])
        elif action == "anon_company":
            df[actual] = df[actual].astype(str).str.strip().map(m["companies"])
        elif action == "anon_product":
            df[actual] = df[actual].astype(str).str.strip().map(m["products"])
        elif action == "mul_price":
            df[actual] = (pd.to_numeric(df[actual], errors="coerce") * PRICE_COEF).round(2)
        elif action == "add_qty":
            df[actual] = pd.to_numeric(df[actual], errors="coerce") + QTY_SHIFT

    drop_cols = [
        col[k] for k in CANONICAL
        if CANONICAL[k] == "drop" and col.get(k) in df.columns
    ]
    df = df.drop(columns=drop_cols, errors="ignore")

    rename = {
        actual: cname
        for cname, actual in col.items()
        if actual in df.columns and CANONICAL.get(cname) != "drop"
    }
    df = df.rename(columns=rename)

    ordered = [c for c in DESIRED_ORDER if c in df.columns]
    extras  = [c for c in df.columns if c not in ordered]
    return df[ordered + extras]


# ── COLLECT & ARCHIVE ──────────────────────────────────────────────────────────

def collect_source_files(folder: Path) -> list[Path]:
    """Return sorted list of supported files in folder."""
    return sorted(
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS
    )

def archive_source_files(
    files: list[Path],
    archive_folder,
    log_fn: Callable[[str], None] = print,
) -> int:
    """Move source files to archive_folder. Returns count of moved files."""
    archive_folder = Path(archive_folder)
    archive_folder.mkdir(parents=True, exist_ok=True)
    moved = 0
    for f in files:
        try:
            dest = archive_folder / f.name
            if dest.exists():
                stem, suffix, i = f.stem, f.suffix, 1
                while dest.exists():
                    dest = archive_folder / f"{stem}_{i}{suffix}"
                    i += 1
            shutil.move(str(f), dest)
            log_fn(f"  Архивирован: {f.name} → {archive_folder.name}/{dest.name}")
            moved += 1
        except Exception as e:
            log_fn(f"  ОШИБКА архивирования {f.name}: {e}")
    return moved


# ── CORE PIPELINE FUNCTIONS (importable by Dagster) ────────────────────────────

def read_and_anonymize(
    files: list[Path],
    log_fn: Callable[[str], None] = print,
    product_match: bool = True,
    product_match_threshold: float = 0.82,
) -> list[tuple[str, pd.DataFrame]]:
    """
    Read and anonymize every file.

    If product_match=True (default), runs two-pass processing:
      Pass 1 – read all files into memory, build raw col_maps.
      Match  – deduplicate product names across all files (barcode + fuzzy).
               Writes/updates product_matches.json.
      Pass 2 – apply canonical names, build anonymization mappings, transform.

    This guarantees that product variants from different stores get the
    same fake name and can be aggregated in merge_anonymized.

    Returns list of (filename, anonymized_df) tuples.
    Files that fail to read are skipped (logged as errors).
    """
    m = load_mappings()
    log_fn(
        f"Маппинги загружены: магазины={len(m['stores'])}, бренды={len(m['brands'])}, "
        f"компании={len(m['companies'])}, товары={len(m['products'])}"
    )

    # ── Pass 1: read all files ─────────────────────────────────────────────────
    raw_frames: list = []   # (path, df, col_map)
    for path in files:
        try:
            log_fn(f"📖 {path.name}")
            df = read_table(path)
            df.columns = [str(c).strip() for c in df.columns]
            log_fn(f"   Строк: {len(df):,}  Столбцов: {len(df.columns)}")

            col_map = get_col_map(df)
            unmapped = [c for c in df.columns if c not in col_map.values()]
            log_fn(
                f"   Сопоставлено: {len(col_map)}  "
                f"Не определены (пройдут без изменений): {unmapped or '—'}"
            )
            raw_frames.append((path, df, col_map))

        except Exception as exc:
            log_fn(f"   ✗ ОШИБКА ({path.name}): {exc}")

    if not raw_frames:
        return []

    # ── Product matching (before anonymization) ────────────────────────────────
    if product_match:
        log_fn("🔍 Дедупликация товаров (до анонимизации)...")
        matches = run_product_matching(
            raw_frames,
            threshold=product_match_threshold,
            log_fn=log_fn,
        )
        total_mapped = sum(1 for k, v in matches.items() if k != v)
        log_fn(f"   Всего нетривиальных матчей в кэше: {total_mapped}")

        for _, df, col in raw_frames:
            apply_product_matches(df, col, matches)
    else:
        log_fn("ℹ️  Дедупликация товаров отключена (product_match=False)")

    # ── Pass 2: anonymize ──────────────────────────────────────────────────────
    frames: list[tuple[str, pd.DataFrame]] = []
    for path, df, col_map in raw_frames:
        try:
            build_mappings(df, col_map, m, log_fn)
            df_out = transform(df.copy(), col_map, m)
            frames.append((path.name, df_out))
            log_fn(f"   ✓ Анонимизирован: {len(df_out.columns)} столбцов")

        except Exception as exc:
            log_fn(f"   ✗ ОШИБКА ({path.name}): {exc}")

    save_mappings(m)
    log_fn(
        f"Маппинги сохранены: магазины={len(m['stores'])}, бренды={len(m['brands'])}, "
        f"компании={len(m['companies'])}, товары={len(m['products'])}"
    )
    return frames


def merge_anonymized(
    frames: list[tuple[str, pd.DataFrame]],
    log_fn: Callable[[str], None] = print,
    aggregate: bool = True,
) -> pd.DataFrame:
    """
    Find column intersection across all frames, concat, then aggregate.

    Aggregation (aggregate=True by default):
      group by all dimension columns, sum qty_sold / sales_rub / cost_rub.
      This collapses rows that refer to the same product/store/week after
      product matching unified their names.
    """
    file_cols = {fname: list(df.columns) for fname, df in frames}

    col_sets = [set(cols) for cols in file_cols.values()]
    common: set[str] = col_sets[0].copy()
    for s in col_sets[1:]:
        common &= s

    any_dropped = False
    for fname, cols in file_cols.items():
        dropped = [c for c in cols if c not in common]
        if dropped:
            log_fn(f"  ⚠ {fname} — удалены уникальные столбцы: {dropped}")
            any_dropped = True
    if not any_dropped:
        log_fn("  ✓ Все файлы имеют одинаковый набор столбцов")

    ordered    = [c for c in DESIRED_ORDER if c in common]
    extras     = sorted(c for c in common if c not in ordered)
    final_cols = ordered + extras
    log_fn(f"  Итоговых столбцов: {len(final_cols)} → {', '.join(final_cols)}")

    merged = pd.concat([df[final_cols] for _, df in frames], ignore_index=True)
    log_fn(f"  Строк после объединения: {len(merged):,}")

    # ── Aggregation ────────────────────────────────────────────────────────────
    if aggregate:
        metric_candidates = ("qty_sold", "sales_rub", "cost_rub")
        metric_cols = [c for c in metric_candidates if c in merged.columns]
        group_cols  = [c for c in merged.columns if c not in metric_cols]

        if metric_cols and group_cols:
            # Ensure metrics are numeric before groupby
            for mc in metric_cols:
                numeric: pd.Series = pd.to_numeric(merged[mc], errors="coerce")
                merged[mc] = numeric.fillna(0)

            merged = (
                merged
                .groupby(group_cols, as_index=False, dropna=False)
                .agg({mc: "sum" for mc in metric_cols})
            )

            # Restore DESIRED_ORDER after groupby
            ordered2 = [c for c in DESIRED_ORDER if c in merged.columns]
            extras2  = [c for c in merged.columns if c not in ordered2]
            merged   = merged[ordered2 + extras2]

            log_fn(
                f"  Строк после агрегации: {len(merged):,}  "
                f"(схлопнуто по: {', '.join(metric_cols)})"
            )
        else:
            log_fn("  ⚠ Агрегация пропущена: нет числовых метрик или нет группирующих столбцов")

    log_fn(f"  Итого строк: {len(merged):,}")
    return merged


def save_merged(
    df: pd.DataFrame,
    output_folder,
    filename: str = "",
    log_fn: Callable[[str], None] = print,
) -> Path:
    """Save merged DataFrame as formatted xlsx. Returns output path."""
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    if not filename:
        filename = f"merged_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
    if not Path(filename).suffix:
        filename += ".xlsx"
    out = output_folder / filename
    write_xlsx(df, out)
    log_fn(f"  Сохранено → {out.resolve()}")
    return out


# ── CONFIG ─────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    defaults = {
        "input_folder": "",
        "archive_folder": "",
        "output_folder": "",
        "product_match": True,
        "product_match_threshold": 0.82,
    }
    if CONFIG_PATH.exists():
        try:
            existing = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            existing = {}
        return {**defaults, **existing}
    return defaults

def save_config(cfg: dict):
    """Merge cfg into config.json, preserving unrelated keys."""
    existing: dict = {}
    if CONFIG_PATH.exists():
        try:
            existing = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    existing.update(cfg)
    CONFIG_PATH.write_text(
        json.dumps(existing, indent=2, ensure_ascii=False), encoding="utf-8"
    )


# ── CLI HELPERS ────────────────────────────────────────────────────────────────

def _hr(char: str = "━", n: int = 52):
    print(char * n)

def select_folder(label: str, current: str) -> str:
    cwd     = Path(".")
    folders = sorted(d for d in cwd.iterdir() if d.is_dir() and not d.name.startswith("."))
    cur_hint = f"текущая: {current}" if current else "не задана"
    print(f"\n{label}  ({cur_hint})")
    if current:
        print(f"  {'↵':>3}  оставить текущую")
    for i, d in enumerate(folders, 1):
        print(f"  {i:>3}  {d.name}/")
    if not folders and not current:
        val = input("  Нет папок. Введите путь вручную: ").strip()
        return val or "."
    while True:
        raw = input("  > ").strip()
        if not raw and current:
            return current
        if raw.isdigit() and 1 <= int(raw) <= len(folders):
            return str(folders[int(raw) - 1])
        if not raw and not current:
            print("  Текущей папки нет — выберите из списка.")
        else:
            print(f"  Введите номер 1–{len(folders)} или Enter.")


# ── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    print()
    _hr()
    print("  🔀  Retail Pipeline — анонимизация и объединение таблиц")
    _hr()

    cfg = load_config()

    # 1. Input folder
    cfg["input_folder"] = select_folder(
        "Папка с исходными файлами", cfg["input_folder"]
    )
    save_config(cfg)

    input_dir = Path(cfg["input_folder"])
    files = collect_source_files(input_dir)
    if not files:
        print(f"\n  ⚠  Нет поддерживаемых файлов в {input_dir}/")
        print(f"     Форматы: {', '.join(sorted(SUPPORTED_EXTS))}")
        sys.exit(1)

    print(f"\n  Найдено файлов: {len(files)}")
    for i, f in enumerate(files, 1):
        kb = max(f.stat().st_size // 1024, 1)
        print(f"  {i:>3}  {f.name}  ({kb} КБ)")

    # 2. Archive folder (for processed source files)
    cfg["archive_folder"] = select_folder(
        "Папка для архива обработанных исходников", cfg["archive_folder"]
    )
    save_config(cfg)

    # 3. Output folder
    cfg["output_folder"] = select_folder(
        "Папка для сохранения результата", cfg["output_folder"]
    )
    save_config(cfg)

    # 4. Output filename
    default_name = f"merged_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
    print(f"\nИмя файла результата  [Enter → {default_name}]:")
    raw_name = input("  > ").strip() or default_name
    if not Path(raw_name).suffix:
        raw_name += ".xlsx"

    # ── Pipeline ───────────────────────────────────────────────────────────────
    print()
    _hr("─")
    t0 = time.time()

    def log(msg: str):
        print(f"  [{time.time() - t0:5.1f}s] {msg}")

    log(f"📖  Чтение и анонимизация {len(files)} файлов...")
    frames = read_and_anonymize(
        files,
        log_fn=log,
        product_match=cfg.get("product_match", True),
        product_match_threshold=cfg.get("product_match_threshold", 0.82),
    )

    if not frames:
        log("❌  Ни один файл не прочитан. Выход.")
        sys.exit(1)

    log("🔀  Объединение таблиц...")
    merged = merge_anonymized(frames, log_fn=log)

    log("💾  Сохранение...")
    out_path = save_merged(merged, cfg["output_folder"], raw_name, log_fn=log)

    log(f"📦  Архивирование исходников → {cfg['archive_folder']}/")
    n_archived = archive_source_files(files, cfg["archive_folder"], log_fn=log)

    elapsed = time.time() - t0
    _hr("─")
    print()
    print("  📊  Итог по файлам:")
    for fname, df in frames:
        print(f"       {fname:<46} {len(df):>8,} строк")
    errors = [f.name for f in files if f.name not in {n for n, _ in frames}]
    for fname in errors:
        print(f"       {fname:<46}  — ошибка чтения")
    print()
    print(f"  Обработано файлов:   {len(frames)}")
    print(f"  Итого строк:         {len(merged):,}")
    print(f"  Архивировано:        {n_archived} файлов → {cfg['archive_folder']}/")
    print()
    _hr()
    print(f"\n  ✅  Готово за {elapsed:.2f} сек")
    print(f"  →  {out_path.resolve()}\n")


if __name__ == "__main__":
    main()