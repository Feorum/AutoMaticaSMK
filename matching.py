# -*- coding: utf-8 -*-
"""
matching.py — словарь соответствий фраз заказа → названий в базе.  v2

Улучшения v2:
  - Удаление шумовых слов из строки заказа перед матчингом.
  - Нормализация аббревиатур (Колбаса→К-са, Сосиски→С-ки и т.д.).
  - Исправление типовых опечаток (салцем→сальцем, Свинн→Свин и т.д.).
  - naiti_cherez_slovar теперь возвращает 3 значения: (fraza, naideny, score).
"""

import csv
import re
from pathlib import Path


def _find_matching_csv() -> Path:
    candidates = [
        Path(__file__).parent / "baza" / "matching.csv",
        Path(__file__).parent / "matching.csv",
    ]
    for p in candidates:
        if p.exists():
            return p
    return candidates[0]


MATCHING_CSV = _find_matching_csv()

# ── Шумовые фразы (удаляются из строки заказа перед сравнением) ───
_NOISE_WORDS = [
    "групповая упаковка", "модифицированная атмосфера", "полиамидная оболочка",
    "натуральная оболочка", "натур оболочка", "натур. оболочка",
    "нат/об", "мол/атм", "вакуумная упаковка", "синюга",
    "слонимский мк", "слонимский", "слоним",
    "изделие колбасное", "колб. изд.", "кол. изд.",
    "изд. к/б.", "изд. колб.", "изд. кол.", "изд. к/б", "изд. колб",
    "пр-т из шп. сол.", "пр-ты из шп. сол.",
    "1 кг", "0,5 кг", "0.5 кг",
    "вес,", " вес", "мяк", "короб",
]

# ── Нормализация аббревиатур ──────────────────────────────────────
_ABBR_MAP = [
    (r"\bколбаса\b",   "к-са"),
    (r"\bколбасы\b",   "к-са"),
    (r"\bсосиски\b",   "с-ки"),
    (r"\bсосисок\b",   "с-ки"),
    (r"\bсардельки\b", "сард"),
    (r"\bсарделек\b",  "сард"),
    (r"\bливерная\b",  "лив"),
    (r"\bпродукт\b",   "пр-т"),
    (r"\bпродукты\b",  "пр-т"),
]

# ── Типовые опечатки ─────────────────────────────────────────────
_TYPOS = [
    ("салцем",      "сальцем"),
    ("мортделла",   "мортаделла"),
    ("свинные",     "свиные"),
    ("дымчатая",    "дымница"),
    ("волковы ",    "волковыс "),
]


def _norm(s: str) -> str:
    """Расширенная нормализация v2: опечатки → аббревиатуры → шум → пунктуация."""
    s = s.lower().strip()
    s = re.sub(r'[«»""„\']', " ", s)

    for wrong, correct in _TYPOS:
        s = s.replace(wrong, correct)

    for pattern, repl in _ABBR_MAP:
        s = re.sub(pattern, repl, s)

    for noise in sorted(_NOISE_WORDS, key=len, reverse=True):
        s = s.replace(noise.lower(), " ")

    s = re.sub(r'[^\w\s/\-]', " ", s)
    return re.sub(r'\s+', " ", s).strip()


# ── Загрузка и сохранение словаря ────────────────────────────────

def zagruzit_slovar() -> list[tuple[str, str]]:
    if not MATCHING_CSV.exists():
        return []
    rezultat = []
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            with open(MATCHING_CSV, encoding=enc, newline="") as f:
                for row in csv.reader(f, delimiter=";"):
                    if len(row) >= 2 and row[0].strip() and row[1].strip():
                        rezultat.append((_norm(row[0].strip()), row[1].strip()))
            break
        except UnicodeDecodeError:
            continue
    rezultat.sort(key=lambda x: len(x[0]), reverse=True)
    return rezultat


def dobavit_v_slovar(fraza: str, baza_podstroka: str) -> bool:
    MATCHING_CSV.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(MATCHING_CSV, "a", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerow([fraza.strip(), baza_podstroka.strip()])
        return True
    except OSError:
        return False


def udalit_iz_slovarya(fraza: str) -> bool:
    if not MATCHING_CSV.exists():
        return False
    fraza_norm = _norm(fraza)
    try:
        stroki = []
        for enc in ("utf-8-sig", "utf-8", "cp1251"):
            try:
                with open(MATCHING_CSV, encoding=enc, newline="") as f:
                    stroki = list(csv.reader(f, delimiter=";"))
                break
            except UnicodeDecodeError:
                continue
        novye = [r for r in stroki if not (len(r) >= 1 and _norm(r[0]) == fraza_norm)]
        with open(MATCHING_CSV, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerows(novye)
        return True
    except OSError:
        return False


# ── Скор совпадения ───────────────────────────────────────────────

def _match_score(fraza_norm: str, zapros_norm: str) -> float:
    """Возвращает 0.0–1.0: доля токенов фразы, найденных в запросе."""
    if fraza_norm == zapros_norm:
        return 1.0
    ftok = set(fraza_norm.split())
    ztok = set(zapros_norm.split())
    if not ftok:
        return 0.0
    return round(len(ftok & ztok) / len(ftok), 2)


# ── Основная функция поиска ───────────────────────────────────────

def naiti_cherez_slovar(
    zapros: str,
    tovary: list[dict],
    slovar: list[tuple[str, str]] | None = None,
) -> tuple[str | None, list[tuple[int, dict]], float]:
    """
    Найти товары в базе через словарь соответствий.

    Возвращает: (sovpavshaya_fraza, [(idx, tovar), ...], score)
      score: 1.0 = точное совпадение, <1.0 = частичное
      (None, [], 0.0) — фраза не найдена в словаре
      (fraza, [], score) — фраза есть, но в базе не найдено

    ИЗМЕНЕНИЕ v2: сигнатура расширена — добавлен третий элемент score.
    В main_gui.py заменить:
        fraza, naideny = naiti_cherez_slovar(...)
    на:
        fraza, naideny, score = naiti_cherez_slovar(...)
    """
    if slovar is None:
        slovar = zagruzit_slovar()

    zn = _norm(zapros)
    sovp_fraza_orig = None
    baza_podstroka  = None
    score           = 0.0

    for fraza_norm, baza in slovar:
        if fraza_norm in zn:
            sovp_fraza_orig = fraza_norm
            baza_podstroka  = baza
            score = _match_score(fraza_norm, zn)
            break

    if baza_podstroka is None:
        return None, [], 0.0

    baza_norm = _norm(baza_podstroka)
    naideny: list[tuple[int, dict]] = [
        (idx, t) for idx, t in enumerate(tovary)
        if baza_norm in _norm(t["naimenovanie"])
    ]

    return sovp_fraza_orig, naideny, score
