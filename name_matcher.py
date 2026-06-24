"""
name_matcher.py — унификация названий товаров без LLM.

Алгоритм:
  1. Блокировка по (бренд + объём/кол-во) — пары сравниваются только внутри
     одного блока (одинаковый бренд + фасовка/шт).
  2. Нормализация: нижний регистр, удаление суффиксов поставщика «(код):N/N»,
     очистка пунктуации, транслитерация кирилл.→лат.
  3. Попарная схожесть внутри блока:
       • token-set Jaccard  — порядок слов не важен
       • SequenceMatcher    — опечатки/сокращения
     score = max(jaccard, seq_ratio × 0.9), порог 0.90
  4. Union-Find кластеризация пар с score ≥ THRESHOLD.
  5. Canonical = самое частое название в кластере, при равенстве — самое длинное.

Стратегия: КОНСЕРВАТИВНАЯ — лучше пропустить правильное слияние,
чем склеить разные товары (разные линейки, вкусы, оттенки, типы).

Публичный API совместим со старой LLM-версией (model= принимается, игнорируется).
"""
import re
import json
import difflib
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

NAME_MATCHES_PATH = Path("name_matches.json")
# Порог 0.90: выше максимального score среди всех известных ложных пар (0.884)
THRESHOLD = 0.90

# Объём И штучные единицы в ключе блокировки:
# «Mach3 2шт» и «Fusion 5шт» → разные блоки → никогда не сравниваются
_VOL_RE = re.compile(
    r'(\d+[.,]?\d*)\s*(мл|л|г|гр|кг|шт|уп|пак|рул|ml|l|g|kg|pcs)\b', re.I
)

# Суффикс поставщика «(КодПост):4/20» или «:6/12» в конце строки —
# одинаков для ВСЕХ товаров одного листа, раздувает seq_ratio при сравнении
_META_RE = re.compile(
    r'\s*\([^)]+\)\s*:\s*\d[\d/]*\s*$|\s*:\s*\d+/\d+\s*$'
)

_CYR = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
_LAT = ["a","b","v","g","d","e","e","zh","z","i","y","k","l","m","n","o","p",
        "r","s","t","u","f","kh","ts","ch","sh","sch","","y","","e","yu","ya"]
_CYR_MAP = dict(zip(_CYR, _LAT))


# ── helpers ────────────────────────────────────────────────────────────────────

def _translit(s: str) -> str:
    return "".join(_CYR_MAP.get(c, c) for c in s)

def _normalize(name: str) -> str:
    s = str(name).lower().strip()
    s = _META_RE.sub("", s)               # убрать «(Код):4/20» в конце
    s = re.sub(r"\+", " plus ", s)        # «Super+» → «Super plus» (отдельный токен)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return _translit(s)

def _jaccard(a: str, b: str) -> float:
    ta, tb = set(a.split()), set(b.split())
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)

def _similarity(a: str, b: str) -> float:
    na, nb = _normalize(a), _normalize(b)
    j   = _jaccard(na, nb)
    seq = difflib.SequenceMatcher(None, na, nb).ratio()
    return max(j, seq * 0.9)


# ── Union-Find ─────────────────────────────────────────────────────────────────

class _UF:
    def __init__(self, n: int):
        self.p = list(range(n))

    def find(self, x: int) -> int:
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a: int, b: int):
        self.p[self.find(a)] = self.find(b)

    def groups(self, n: int) -> list:
        buckets: dict = defaultdict(list)
        for i in range(n):
            buckets[self.find(i)].append(i)
        return list(buckets.values())


# ── volume key ────────────────────────────────────────────────────────────────

def _norm_vol(num: str, unit: str) -> str:
    v = float(num.replace(",", "."))
    u = unit.lower()
    if u in ("л", "l"):              return f"{int(round(v * 1000))}ml"
    if u in ("мл", "ml"):            return f"{int(round(v))}ml"
    if u in ("кг", "kg"):            return f"{int(round(v * 1000))}g"
    if u in ("г", "гр", "g"):        return f"{int(round(v))}g"
    return f"{int(round(v))}pcs"     # шт / уп / пак / рул / pcs

def extract_volume(name: str) -> str:
    vols = {_norm_vol(m.group(1), m.group(2)) for m in _VOL_RE.finditer(str(name))}
    return "+".join(sorted(vols))


# ── core matching ──────────────────────────────────────────────────────────────

def match_group(names: list, threshold: float = THRESHOLD, **_) -> dict:
    """Группирует схожие названия без LLM. model= принимается но игнорируется."""
    n = len(names)
    if n < 2:
        return {"groups": [[0]] if names else []}
    uf = _UF(n)
    for i in range(n):
        for j in range(i + 1, n):
            if _similarity(names[i], names[j]) >= threshold:
                uf.union(i, j)
    return {"groups": uf.groups(n)}


def build_name_matches(
    items,
    model: str = "",          # оставлен для совместимости API
    threshold: float = THRESHOLD,
    log_fn=print,
) -> dict:
    """
    items: [(name, brand), ...] → {orig_name: canonical_name}

    Canonical = самое частое название в кластере (при равной частоте — самое длинное).
    """
    freq   = Counter(n for n, _ in items)
    blocks: dict = defaultdict(list)
    for name, brand in items:
        key = (str(brand).strip().lower(), extract_volume(name))
        blocks[key].append(name)

    multi = [v for v in blocks.values() if len(set(v)) > 1]
    log_fn(f"  [Мэтчинг] блоков с кандидатами: {len(multi)}")

    matches: dict = {}
    n_groups = 0
    for names in multi:
        uniq = list(dict.fromkeys(names))
        for g in match_group(uniq, threshold)["groups"]:
            if len(g) < 2:
                continue
            members   = [uniq[i] for i in g]
            canonical = max(members, key=lambda nm: (freq[nm], len(nm)))
            for nm in members:
                if nm != canonical:
                    matches[nm] = canonical
            n_groups += 1

    log_fn(f"  [Мэтчинг] групп: {n_groups}, переименовано названий: {len(matches)}")
    return matches


# ── pipeline helpers ───────────────────────────────────────────────────────────

def items_from_frames(raw_frames) -> list:
    """Собирает (name, brand) из raw_frames pipeline."""
    out = []
    for _, df, col in raw_frames:
        ic, bc = col.get("item_name"), col.get("brand")
        if not ic or ic not in df.columns:
            continue
        names  = df[ic].astype(str).str.strip()
        brands = (df[bc].astype(str).str.strip() if bc and bc in df.columns
                  else pd.Series([""] * len(df), index=df.index))
        out.extend(zip(names.tolist(), brands.tolist()))
    return out


def load_name_matches() -> dict:
    """Загружает ручные правки / кэш из name_matches.json."""
    if not NAME_MATCHES_PATH.exists():
        return {}
    try:
        return json.loads(NAME_MATCHES_PATH.read_text("utf-8"))
    except Exception:
        return {}


def save_name_matches(matches: dict):
    NAME_MATCHES_PATH.write_text(
        json.dumps(matches, ensure_ascii=False, indent=2), "utf-8"
    )


def apply_name_matches(df, col, matches: dict):
    """Дописывает canonical_name в df; без матча — canonical = само исходное имя."""
    ic = col.get("item_name")
    if not ic or ic not in df.columns:
        return df
    df["canonical_name"] = df[ic].astype(str).str.strip().map(
        lambda n: matches.get(n, n)
    )
    return df
