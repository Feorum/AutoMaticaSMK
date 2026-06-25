# -*- coding: utf-8 -*-
"""
main_gui.py  —  единый GUI для авто-ввода продукции  (v1.0)

Три вкладки:
  1. Заказ   — загрузить Excel/TXT, увидеть план, запустить ввод
  2. Словарь — таблица соответствий фраз заказа → товаров базы
  3. Журнал  — история всех операций ввода

Зависимости (должны лежать рядом):
  auto_vvod_v3.py   (или auto_vvod.py с патчем matching)
  matching.py
  oformlenie_slovar.py
  baza/tovary.csv
  baza/matching.csv  (создаётся автоматически)

pip install pandas openpyxl
"""
from __future__ import annotations

import os
import sys
import csv
import re
import threading
import subprocess
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ocr_scanner import OCRScanner, ScanResult

# ── Путь базы ────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

BAZA_DIR     = BASE_DIR / "baza"
ZADANIYA_DIR = BASE_DIR / "zadaniya"
ARHIV_DIR    = ZADANIYA_DIR / "_arhiv"
TOVARY_CSV   = BAZA_DIR / "tovary.csv"
ISTORIYA_CSV = BAZA_DIR / "istoriya.csv"
MATCHING_CSV = BAZA_DIR / "matching.csv"
ZAKAZY_DIR   = BASE_DIR / "zakazy"   # сохранённые именованные заказы

# ── Настройки ввода ───────────────────────────────────────────────
REZHIM_VVODA     = "vboxmanage"
IMYA_VM          = "Xp"
VBOXMANAGE_PATH  = r"C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
SBROS_ZAPAS      = 5         # запас нажатий PageUp при сбросе курсора наверх
RAZMER_STRANICY  = 17        # сколько строк сдвигает PageUp/PageDown (регулируется в программе)
PAUZA_KLAVISHA   = 0.05
PAUZA_MEZHDU     = 0.15
SORTIROVAT       = True
KLAVISHA_PODTV   = ""

# ── Цвета статусов ────────────────────────────────────────────────
CLR_OK      = "#1a7a1a"
CLR_WARN    = "#b86000"
CLR_ERR     = "#aa1111"
CLR_INFO    = "#003388"
CLR_BG      = "#f4f4f4"
CLR_STRIPE1 = "#ffffff"
CLR_STRIPE2 = "#eef2f8"
CLR_MISS    = "#fff0f0"   # фон строки «не найдено»
CLR_MULTI   = "#fff8e0"   # фон строки «несколько вариантов»
CLR_REZERV  = "#e8f0ff"   # фон строки «резерв»
CLR_TOV_ODD = "#ffffff"
CLR_TOV_EVN = "#f0f4f8"

# ╔══════════════════════════════════════════════════════════════════╗
# ║  БИЗНЕС-ЛОГИКА (упрощённые копии из auto_vvod + matching)       ║
# ╚══════════════════════════════════════════════════════════════════╝

def prochitat_tekst(path) -> str:
    for enc in ("utf-8-sig", "cp1251", "utf-8"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, OSError):
            continue
    return ""


# ── База товаров ──────────────────────────────────────────────────

def zagruzit_bazu() -> list[dict]:
    if not TOVARY_CSV.exists():
        return []
    tovary = []
    tekst = prochitat_tekst(TOVARY_CSV)
    reader = csv.DictReader(tekst.splitlines(), delimiter=";")
    for row in reader:
        name = (row.get("naimenovanie") or "").strip()
        if not name:
            continue
        try:
            ost = int((row.get("ostatok") or "0").strip() or 0)
        except ValueError:
            ost = 0
        rezerv_raw = (row.get("rezerv") or "0").strip().lower()
        active_raw = (row.get("active")  or "1").strip().lower()
        tovary.append({
            "naimenovanie": name,
            "oformlenie":   (row.get("oformlenie") or "").strip(),
            "massa":        (row.get("massa") or "").strip(),
            "ostatok":      ost,
            "rezerv":       rezerv_raw in ("1", "yes", "да", "true"),
            "active":       active_raw not in ("0", "no", "нет", "false"),
        })
    return tovary


def sohranit_bazu(tovary: list[dict]) -> None:
    BAZA_DIR.mkdir(parents=True, exist_ok=True)
    with open(TOVARY_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["pozitsiya", "naimenovanie", "oformlenie", "massa",
                    "ostatok", "rezerv", "active"])
        for i, t in enumerate(tovary, 1):
            w.writerow([i, t["naimenovanie"], t.get("oformlenie",""), t.get("massa",""),
                        t.get("ostatok",0),
                        "1" if t.get("rezerv") else "0",
                        "1" if t.get("active", True) else "0"])


def zapisat_istoriyu(zapisi: list) -> None:
    BAZA_DIR.mkdir(parents=True, exist_ok=True)
    novyi = not ISTORIYA_CSV.exists()
    with open(ISTORIYA_CSV, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        if novyi:
            w.writerow(["data_vremya", "zadanie", "stroka", "naimenovanie",
                        "oformlenie", "massa", "kolichestvo",
                        "ostatok_do", "ostatok_posle", "rezerv"])
        for z in zapisi:
            w.writerow(z)


def format_tovar(t: dict) -> str:
    parts = [t["naimenovanie"]]
    if t.get("oformlenie"):
        parts.append(t["oformlenie"])
    if t.get("massa"):
        parts.append(t["massa"])
    return " | ".join(parts)


# ── Словарь соответствий ──────────────────────────────────────────

def _norm(s: str) -> str:
    s = s.replace("«", " ").replace("»", " ").replace('"', " ").replace("'", " ")
    s = re.sub(r"[^\w\s]", " ", s.lower())
    return re.sub(r"\s+", " ", s).strip()


def zagruzit_slovar() -> list[tuple[str, str]]:
    if not MATCHING_CSV.exists():
        return []
    result = []
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            with open(MATCHING_CSV, encoding=enc, newline="") as f:
                for row in csv.reader(f, delimiter=";"):
                    if len(row) >= 2 and row[0].strip() and row[1].strip():
                        result.append((row[0].strip(), row[1].strip()))
            break
        except UnicodeDecodeError:
            continue
    result.sort(key=lambda x: len(_norm(x[0])), reverse=True)
    return result


def sohranit_slovar(pairs: list[tuple[str, str]]) -> None:
    BAZA_DIR.mkdir(parents=True, exist_ok=True)
    with open(MATCHING_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        for fraza, baza in pairs:
            w.writerow([fraza.strip(), baza.strip()])


def _bigrams(s: str) -> set:
    return {s[i:i+2] for i in range(len(s) - 1)}


def _idf_weights(tovary: list[dict]) -> dict[str, float]:
    """
    Вычислить IDF-вес каждого токена по всей базе товаров.
    Токены, встречающиеся редко (специфичные) → высокий вес.
    Токены-«шум» (к/б, вар, в/с, изд ...) → низкий вес.
    Кэшируется снаружи, пересчитывается при перезагрузке базы.
    """
    import math
    N = len(tovary)
    if N == 0:
        return {}
    df: dict[str, int] = {}
    for t in tovary:
        for tok in set(_norm(t["naimenovanie"]).split()):
            df[tok] = df.get(tok, 0) + 1
    return {tok: math.log((N + 1) / (cnt + 1)) for tok, cnt in df.items()}


# Глобальный кэш IDF (заполняется при первом вызове fuzzy_podbor)
_IDF_CACHE: dict[str, float] = {}
_IDF_TOVARY_LEN: int = 0


def _get_idf(tovary: list[dict]) -> dict[str, float]:
    global _IDF_CACHE, _IDF_TOVARY_LEN
    if len(tovary) != _IDF_TOVARY_LEN:
        _IDF_CACHE = _idf_weights(tovary)
        _IDF_TOVARY_LEN = len(tovary)
    return _IDF_CACHE


def _fuzzy_score(
    query: str,
    candidate: str,
    idf: dict[str, float] | None = None,
) -> float:
    """
    Нечёткое сходство query ↔ candidate (0.0 … 1.0).

    Три составляющих:
      1. IDF-взвешенный Jaccard по токенам — «важные» слова вносят бо́льший вклад.
      2. Символьные биграммы — ловит опечатки и сокращения.
      3. Sliding-window по подстрокам candidate — сравниваем query со всеми
         окнами шириной len(q_tok) в candidate, берём лучшее окно.
         Это позволяет находить query внутри более длинного названия.
    """
    qn = _norm(query)
    cn = _norm(candidate)
    if not qn:
        return 0.0
    if qn == cn:
        return 1.0

    q_tok = qn.split()
    c_tok = cn.split()
    if not q_tok or not c_tok:
        return 0.0

    q_set = set(q_tok)
    c_set = set(c_tok)
    w = idf or {}

    # ── 1. IDF-взвешенный Jaccard по токенам ─────────────────────
    def idf_w(tok):
        return w.get(tok, 1.0)

    inter_w  = sum(idf_w(t) for t in q_set & c_set)
    union_w  = sum(idf_w(t) for t in q_set | c_set)
    tok_score = inter_w / union_w if union_w else 0.0

    # ── 2. Символьные биграммы ────────────────────────────────────
    qb = _bigrams(qn)
    cb = _bigrams(cn)
    if qb and cb:
        bi_score = len(qb & cb) / max(len(qb), len(cb))
    else:
        bi_score = 0.0

    # ── 3. Скользящее окно по токенам candidate ───────────────────
    # Сравниваем query с каждым подмножеством c_tok длиной len(q_tok)
    win_score = 0.0
    wsize = len(q_tok)
    if wsize <= len(c_tok):
        for start in range(len(c_tok) - wsize + 1):
            window = set(c_tok[start: start + wsize])
            w_inter = sum(idf_w(t) for t in q_set & window)
            w_union = sum(idf_w(t) for t in q_set | window)
            sc = w_inter / w_union if w_union else 0.0
            if sc > win_score:
                win_score = sc

    # Итоговая оценка: максимум из токенного Jaccard и окна,
    # плюс биграммы как дополнительный сигнал.
    base = max(tok_score, win_score)
    score = 0.55 * base + 0.30 * bi_score + 0.15 * max(tok_score, win_score)

    # Бонус за точное вхождение (query-строка целиком внутри candidate)
    if qn in cn:
        score = min(1.0, score + 0.25)

    return min(1.0, score)


def fuzzy_podbor(
    query: str,
    tovary: list[dict],
    top_n: int = 8,
    threshold: float = 0.12,
) -> list[tuple[float, int, dict]]:
    """
    Вернуть top_n товаров, наиболее похожих на query, с учётом IDF и подстрок.
    Возвращает [(score, idx, tovar), ...] по убыванию score.
    score == 1.0 означает точное нормализованное совпадение.
    """
    if not query.strip():
        return []
    idf = _get_idf(tovary)
    results = []
    for idx, t in enumerate(tovary):
        sc = _fuzzy_score(query, t["naimenovanie"], idf)
        if sc >= threshold:
            results.append((sc, idx, t))
    results.sort(key=lambda x: -x[0])
    return results[:top_n]


def _match_score(fraza_norm: str, zapros_norm: str) -> float:
    """Доля слов фразы из словаря, попавших в запрос. 1.0 = точно."""
    ftok = set(fraza_norm.split())
    ztok = set(zapros_norm.split())
    if not ftok:
        return 0.0
    if ftok == ztok:
        return 1.0
    return round(len(ftok & ztok) / len(ftok), 2)


def naiti_cherez_slovar(
    zapros: str,
    tovary: list[dict],
    slovar: list[tuple[str, str]],
) -> tuple[str | None, list[tuple[int, dict]], float]:
    zn = _norm(zapros)
    baza_podstroka = None
    sovp_fraza = None
    score = 0.0
    for fraza, baza in slovar:
        fn = _norm(fraza)
        if fn in zn:
            sovp_fraza = fraza
            baza_podstroka = baza
            score = _match_score(fn, zn)
            break
    if baza_podstroka is None:
        return None, [], 0.0
    bn = _norm(baza_podstroka)
    naideny = [(i, t) for i, t in enumerate(tovary) if bn in _norm(t["naimenovanie"])]
    return sovp_fraza, naideny, score


# ── Разбор задания ────────────────────────────────────────────────

def razobrat_zadanie_txt(path: str) -> tuple[list, list]:
    pozicii, oshibki = [], []
    tekst = prochitat_tekst(path)
    for nomer, raw in enumerate(tekst.splitlines(), 1):
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        parts = [p.strip() for p in s.split(";")]
        if len(parts) < 2:
            oshibki.append((nomer, raw, "нет количества"))
            continue
        kol_str = parts[-1]
        if not kol_str.isdigit():
            oshibki.append((nomer, raw, "количество не число"))
            continue
        name = parts[0]
        if not name:
            oshibki.append((nomer, raw, "пустое название"))
            continue
        oformlenie = parts[1] if len(parts) >= 3 else ""
        massa      = parts[2] if len(parts) >= 4 else ""
        pozicii.append((name, oformlenie, massa, kol_str))
    return pozicii, oshibki


def razobrat_zadanie_xlsx(path: str) -> tuple[list, list]:
    try:
        import pandas as pd
    except ImportError:
        return [], [(0, path, "pandas не установлен: pip install pandas openpyxl")]

    NAME_KW = ["наименование", "название", "продукция", "товар"]
    QTY_KW  = ["требуемое", "план", "заявк", "количество", "кол-во"]

    xl  = pd.ExcelFile(path)
    df_raw = pd.read_excel(path, sheet_name=xl.sheet_names[0], header=None, dtype=str)

    # Найти строку заголовков
    best_row, best_sc = 0, 0
    for i in range(min(20, len(df_raw))):
        vals = [str(v).lower() for v in df_raw.iloc[i].values]
        sc = sum(1 for kw in NAME_KW + QTY_KW if any(kw in v for v in vals))
        if sc > best_sc:
            best_sc, best_row = sc, i

    df = pd.read_excel(path, sheet_name=xl.sheet_names[0], header=best_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    def find_col(kws):
        for kw in kws:
            for col in df.columns:
                if kw in col.lower():
                    return col
        return None

    name_col = find_col(NAME_KW)
    qty_col  = find_col(QTY_KW)
    if not name_col or not qty_col:
        return [], [(0, path, f"Не найдены столбцы. Колонки: {list(df.columns)}")]

    SKIP = ["итого", "всего", "total", "наименование", "продукция"]
    pozicii, oshibki = [], []
    for _, row in df.iterrows():
        name = str(row.get(name_col, "")).strip()
        qty_raw = str(row.get(qty_col, "")).strip()
        if not name or name.lower() in ("nan", "") or any(s in name.lower() for s in SKIP):
            continue
        qty_clean = re.sub(r"[^\d]", "", qty_raw.split(".")[0])
        if not qty_clean or int(qty_clean) <= 0:
            continue
        pozicii.append((name, "", "", qty_clean))
    return pozicii, oshibki


# ── VBoxManage ввод ───────────────────────────────────────────────

SC_MAKE = {"0":0x0B,"1":0x02,"2":0x03,"3":0x04,"4":0x05,
           "5":0x06,"6":0x07,"7":0x08,"8":0x09,"9":0x0A,
           "enter":0x1C,"tab":0x0F,"esc":0x01}
SC_EXT  = {"up":0x48,"down":0x50,"pageup":0x49,"pagedown":0x51}

def _vbox_codes(key: str) -> list[str]:
    if key in SC_EXT:
        m = SC_EXT[key]
        return ["e0", f"{m:02x}", "e0", f"{(m|0x80):02x}"]
    m = SC_MAKE[key]
    return [f"{m:02x}", f"{(m|0x80):02x}"]

def _vbox_send(codes: list[str]) -> None:
    subprocess.run(
        [VBOXMANAGE_PATH, "controlvm", IMYA_VM, "keyboardputscancode"] + codes,
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)

def nazhat_klavishu(key: str) -> None:
    _vbox_send(_vbox_codes(key))
    time.sleep(PAUZA_KLAVISHA)

def peremestit_kursor(tek: int, cel: int) -> None:
    """Перемещает курсор: крупные прыжки PageUp/PageDown
    (1 страница = RAZMER_STRANICY строк), остаток — стрелками.
    RAZMER_STRANICY регулируется прямо в программе (поле «Строк в странице»).
    """
    delta = cel - tek
    if delta == 0:
        return
    pg  = "pagedown" if delta > 0 else "pageup"
    arr = "down"     if delta > 0 else "up"
    stranica = RAZMER_STRANICY if RAZMER_STRANICY > 0 else 17
    for _ in range(abs(delta) // stranica):
        nazhat_klavishu(pg)
    for _ in range(abs(delta) % stranica):
        nazhat_klavishu(arr)

def vvesti_kolichestvo(kol: str) -> None:
    for ch in kol:
        nazhat_klavishu(ch)
    if KLAVISHA_PODTV:
        nazhat_klavishu(KLAVISHA_PODTV)

def proverit_vm() -> tuple[bool, str]:
    if not os.path.isfile(VBOXMANAGE_PATH):
        return False, f"VBoxManage не найден: {VBOXMANAGE_PATH}"
    try:
        out = subprocess.run([VBOXMANAGE_PATH, "list", "runningvms"],
                             capture_output=True, text=True, check=True)
        if f'"{IMYA_VM}"' not in out.stdout:
            return False, f'VM "{IMYA_VM}" не запущена'
        return True, f'VM "{IMYA_VM}" запущена — готово к вводу'
    except Exception as e:
        return False, f"Ошибка VBoxManage: {e}"


# ╔══════════════════════════════════════════════════════════════════╗
# ║  GUI                                                             ║
# ╚══════════════════════════════════════════════════════════════════╝

class App(tk.Tk):
    def _build_tab_tovary(self, btn_frame=None):
        ...
        # Кнопка сканирования (добавить в панель кнопок)
        ttk.Button(btn_frame, text="🔍 Сканировать ассортимент",
                   command=self._ocr_scan).pack(side="left", padx=4)
        self.lbl_scan_stat = ttk.Label(btn_frame, text="", foreground=CLR_INFO)
        self.lbl_scan_stat.pack(side="left", padx=8)

    def _ocr_scan(self):
        from ocr_scanner import OCRScanner
        scanner = OCRScanner(
            vm_name=IMYA_VM,
            vboxmanage=VBOXMANAGE_PATH,
            tovary_csv=TOVARY_CSV,
            page_size=RAZMER_STRANICY,
            on_progress=self._ocr_progress,
            on_done=self._ocr_done,
            on_error=self._ocr_error,
        )
        ok, msg = scanner.check_deps()
        if not ok:
            messagebox.showerror("OCR: ошибка", msg); return
        self.lbl_scan_stat.config(text="⏳ Сканирование…", foreground=CLR_WARN)
        self._ocr_scanner = scanner
        scanner.scan_async()

    def _ocr_progress(self, msg: str, step: int, total: int):
        self.after(0, lambda: self.lbl_scan_stat.config(
            text=f"⏳ {msg}", foreground=CLR_WARN))

    def _ocr_done(self, result: ScanResult):
        self.after(0, lambda: self._ocr_show_result(result))

    def _ocr_error(self, err: str):
        self.after(0, lambda: (
            messagebox.showerror("OCR: ошибка", err),
            self.lbl_scan_stat.config(text="❌ Ошибка", foreground=CLR_ERR),
        ))

    def _ocr_show_result(self, result: ScanResult):
        from ocr_scanner import ScanResult
        if not result.has_changes:
            self.lbl_scan_stat.config(
                text=f"✓ Без изменений ({len(result.unchanged)} поз.)",
                foreground=CLR_OK)
            messagebox.showinfo("OCR: синхронизация", result.summary())
            return

        msg = (f"{result.summary()}\n\n"
            f"Новые:\n{result.new_names_str() or '  —'}\n\n"
            f"Исчезли из DOS:\n{result.removed_names_str() or '  —'}\n\n"
            "Применить изменения?")
        if not messagebox.askyesno("OCR: найдены изменения", msg):
            self.lbl_scan_stat.config(text="Отменено", foreground=CLR_WARN)
            return

        added, deact = self._ocr_scanner.apply_changes(result)
        self._reload_baza()
        self._tov_refresh()
        self.lbl_scan_stat.config(
            text=f"✓ +{added} новых, деакт. {deact}", foreground=CLR_OK)
        self._log(f"OCR-синхронизация: +{added} новых, деактивировано {deact}")
        messagebox.showinfo("OCR: готово",
                        f"Добавлено новых позиций:  {added}\n"
                        f"Деактивировано (исчезли): {deact}")
    def __init__(self):
        super().__init__()
        self.title("Авто-ввод продукции  v1.0")
        self.geometry("1100x720")
        self.minsize(900, 600)
        self.configure(bg=CLR_BG)

        self.tovary:  list[dict]         = []
        self.slovar:  list[tuple]        = []
        # Строки плана: {"name", "oform", "massa", "kol",
        #                "status": "ok"|"multi"|"miss",
        #                "idx": int|None, "t": dict|None,
        #                "варианты": list[(int,dict)]}
        self.plan_rows: list[dict] = []
        self._vvod_thread: threading.Thread | None = None
        self._stop_vvod = threading.Event()
        self._current_zakaz_name: str = ""   # имя текущего заказа
        self._current_path: str = ""

        self._build_ui()
        self._reload_baza()
        self.after(100, self._setup_keyboard_nav)

    def _build_tab_tovary(self, btn_frame=None):
        ...
        # Кнопка сканирования (добавить в панель кнопок)
        ttk.Button(btn_frame, text="🔍 Сканировать ассортимент",
                   command=self._ocr_scan).pack(side="left", padx=4)
        self.lbl_scan_stat = ttk.Label(btn_frame, text="", foreground=CLR_INFO)
        self.lbl_scan_stat.pack(side="left", padx=8)

    def _ocr_scan(self):
        """Запустить OCR-сканирование ассортимента DOS-программы."""
        from ocr_scanner import OCRScanner
        scanner = OCRScanner(
            vm_name=IMYA_VM,
            vboxmanage=VBOXMANAGE_PATH,
            tovary_csv=TOVARY_CSV,
            page_size=RAZMER_STRANICY,
            on_progress=self._ocr_progress,
            on_done=self._ocr_done,
            on_error=self._ocr_error,
        )
        ok, msg = scanner.check_deps()
        if not ok:
            messagebox.showerror("OCR: ошибка", msg); return
        self.lbl_scan_stat.config(text="⏳ Сканирование…", foreground=CLR_WARN)
        self._ocr_scanner = scanner
        scanner.scan_async()

    def _ocr_progress(self, msg: str, step: int, total: int):
        self.after(0, lambda: self.lbl_scan_stat.config(
            text=f"⏳ {msg}", foreground=CLR_WARN))

    def _ocr_done(self, result: ScanResult):
        self.after(0, lambda: self._ocr_show_result(result))

    def _ocr_error(self, err: str):
        self.after(0, lambda: (
            messagebox.showerror("OCR: ошибка", err),
            self.lbl_scan_stat.config(text="❌ Ошибка", foreground=CLR_ERR),
        ))

    def _ocr_show_result(self, result: ScanResult):
        from ocr_scanner import ScanResult
        if not result.has_changes:
            self.lbl_scan_stat.config(
                text=f"✓ Без изменений ({len(result.unchanged)} поз.)",
                foreground=CLR_OK)
            messagebox.showinfo("OCR: синхронизация", result.summary())
            return

        msg = (f"{result.summary()}\n\n"
            f"Новые:\n{result.new_names_str() or '  —'}\n\n"
            f"Исчезли из DOS:\n{result.removed_names_str() or '  —'}\n\n"
            "Применить изменения?")
        if not messagebox.askyesno("OCR: найдены изменения", msg):
            self.lbl_scan_stat.config(text="Отменено", foreground=CLR_WARN)
            return

        added, deact = self._ocr_scanner.apply_changes(result)
        self._reload_baza()
        self._tov_refresh()
        self.lbl_scan_stat.config(
            text=f"✓ +{added} новых, деакт. {deact}", foreground=CLR_OK)
        self._log(f"OCR-синхронизация: +{added} новых, деактивировано {deact}")
        messagebox.showinfo("OCR: готово",
                        f"Добавлено новых позиций:  {added}\n"
                        f"Деактивировано (исчезли): {deact}")
    # ── BUILD UI ──────────────────────────────────────────────────

    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb

        self.tab_zakaz   = ttk.Frame(nb)
        self.tab_slovar  = ttk.Frame(nb)
        self.tab_zhurnal = ttk.Frame(nb)

        self.tab_ruchnoy  = ttk.Frame(nb)
        self.tab_tovary    = ttk.Frame(nb)

        nb.add(self.tab_zakaz,   text="  📋 Заказ  ")
        nb.add(self.tab_ruchnoy, text="  ✏️ Ввод заказа  ")
        nb.add(self.tab_tovary,  text="  📦 Список товаров  ")
        nb.add(self.tab_slovar,  text="  📖 Словарь  ")
        nb.add(self.tab_zhurnal, text="  📜 Журнал  ")

        self._build_tab_zakaz()
        self._build_tab_ruchnoy()
        self._build_tab_tovary()
        self._build_tab_slovar()
        self._build_tab_zhurnal()

    # ── ВКЛАДКА ЗАКАЗ ─────────────────────────────────────────────

    def _build_tab_zakaz(self):
        f = self.tab_zakaz

        # Панель загрузки файла
        top = ttk.LabelFrame(f, text="Файл задания", padding=6)
        top.pack(fill="x", padx=6, pady=(6, 2))

        ttk.Button(top, text="📂 Excel…",    command=self._otkryt_excel).grid(row=0, column=0, padx=(0,3))
        ttk.Button(top, text="📂 TXT…",      command=self._otkryt_txt).grid(row=0, column=1, padx=(0,3))
        ttk.Button(top, text="📂 zadaniya…", command=self._otkryt_iz_papki).grid(row=0, column=2, padx=(0,3))
        ttk.Button(top, text="📁 Открыть заказ…", command=self._otkryt_sohranenny).grid(row=0, column=3, padx=(0,8))

        self.lbl_file = ttk.Label(top, text="Файл не выбран", foreground="#777", width=35)
        self.lbl_file.grid(row=0, column=4, sticky="ew")
        top.columnconfigure(4, weight=1)

        # Название заказа
        ttk.Label(top, text="Название заказа:").grid(row=1, column=0, sticky="w", pady=(4,0))
        self.zakaz_name_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.zakaz_name_var, width=40).grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=(0,4), pady=(4,0))
        ttk.Button(top, text="💾 Сохранить заказ", command=self._sohranit_zakaz).grid(
            row=1, column=4, sticky="w", pady=(4,0))

        # Статус базы и VM
        info = ttk.Frame(f)
        info.pack(fill="x", padx=6, pady=2)
        self.lbl_baza_stat = ttk.Label(info, text="База: не загружена", foreground=CLR_ERR)
        self.lbl_baza_stat.pack(side="left", padx=(0, 20))
        self.lbl_vm_stat = ttk.Label(info, text="VM: неизвестно", foreground=CLR_WARN)
        self.lbl_vm_stat.pack(side="left")
        ttk.Button(info, text="↺ Проверить VM", command=self._check_vm).pack(side="left", padx=8)
        ttk.Button(info, text="↺ Перезагрузить базу", command=self._reload_baza).pack(side="left")

        # Регулировка размера страницы (сколько строк сдвигает PageUp/PageDown)
        ttk.Label(info, text="Строк в странице:").pack(side="left", padx=(20, 2))
        self.stranica_var = tk.IntVar(value=RAZMER_STRANICY)
        sp = ttk.Spinbox(info, from_=1, to=99, width=4,
                         textvariable=self.stranica_var,
                         command=self._primenit_stranicu)
        sp.pack(side="left")
        sp.bind("<FocusOut>", lambda e: self._primenit_stranicu())
        sp.bind("<Return>",   lambda e: self._primenit_stranicu())

        # Таблица плана
        tbl_frame = ttk.LabelFrame(f, text="План ввода  (строка заказа → товар в базе)", padding=4)
        tbl_frame.pack(fill="both", expand=True, padx=6, pady=2)

        cols = ("num", "zakaz_name", "kol", "arrow", "baza_name", "stroka", "ost", "status")
        self.tree_plan = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="browse")
        heads = {"num":"№","zakaz_name":"Строка заказа","kol":"Кол",
                 "arrow":"","baza_name":"Товар в базе","stroka":"Стр",
                 "ost":"Остаток","status":"Статус"}
        widths = {"num":32,"zakaz_name":340,"kol":50,"arrow":24,
                  "baza_name":260,"stroka":42,"ost":65,"status":90}
        for c in cols:
            self.tree_plan.heading(c, text=heads[c])
            self.tree_plan.column(c, width=widths[c],
                                  stretch=(c in ("zakaz_name","baza_name")),
                                  anchor="center" if c in ("num","kol","arrow","stroka","ost") else "w")

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical",   command=self.tree_plan.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.tree_plan.xview)
        self.tree_plan.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree_plan.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        # Теги цветов строк
        self.tree_plan.tag_configure("ok",     background=CLR_STRIPE1)
        self.tree_plan.tag_configure("ok2",    background=CLR_STRIPE2)
        self.tree_plan.tag_configure("multi",  background=CLR_MULTI)
        self.tree_plan.tag_configure("miss",   background=CLR_MISS)
        self.tree_plan.tag_configure("warn",   background="#fff0cc")
        self.tree_plan.tag_configure("rezerv", background=CLR_REZERV)

        self.tree_plan.bind("<Double-1>", self._on_plan_dblclick)

        # Нижняя панель
        bot = ttk.Frame(f)
        bot.pack(fill="x", padx=6, pady=(2, 6))

        self.lbl_plan_stat = ttk.Label(bot, text="", foreground=CLR_INFO)
        self.lbl_plan_stat.pack(side="left")

        self.btn_start = ttk.Button(bot, text="▶  Запустить ввод",
                                    command=self._zapustit_vvod, state="disabled")
        self.btn_start.pack(side="right", padx=(8, 0))
        ttk.Button(bot, text="⏹ Стоп", command=self._stop_vvod_fn).pack(side="right")

        # Лог ввода
        log_frame = ttk.LabelFrame(f, text="Лог ввода", padding=4)
        log_frame.pack(fill="x", padx=6, pady=(0, 6))
        self.txt_log = tk.Text(log_frame, height=5, state="disabled",
                               font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
                               insertbackground="white")
        self.txt_log.pack(fill="x")
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sb.set)

    # ── ВКЛАДКА СЛОВАРЬ ───────────────────────────────────────────

    def _build_tab_slovar(self):
        f = self.tab_slovar

        ttk.Label(f, text="Каждая строка: фраза из заказа  →  подстрока названия в базе товаров.",
                  foreground="#444").pack(anchor="w", padx=8, pady=(6, 2))

        # Поиск
        sf = ttk.Frame(f)
        sf.pack(fill="x", padx=8, pady=2)
        ttk.Label(sf, text="🔍 Поиск:").pack(side="left")
        self.slovar_filter = tk.StringVar()
        self.slovar_filter.trace_add("write", lambda *_: self._filter_slovar())
        ttk.Entry(sf, textvariable=self.slovar_filter, width=30).pack(side="left", padx=4)
        self.lbl_slovar_count = ttk.Label(sf, text="Записей: 0", foreground="#555")
        self.lbl_slovar_count.pack(side="right")

        # Таблица
        tbl = ttk.LabelFrame(f, text="Словарь соответствий", padding=4)
        tbl.pack(fill="both", expand=True, padx=8, pady=2)

        cols = ("fraza", "baza", "sovpad")
        self.tree_sl = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        self.tree_sl.heading("fraza",  text="Фраза из заказа")
        self.tree_sl.heading("baza",   text="Подстрока в базе товаров")
        self.tree_sl.heading("sovpad", text="Совпадений в базе")
        self.tree_sl.column("fraza",  width=380, stretch=True)
        self.tree_sl.column("baza",   width=280, stretch=True)
        self.tree_sl.column("sovpad", width=130, anchor="center", stretch=False)

        vsb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree_sl.yview)
        self.tree_sl.configure(yscrollcommand=vsb.set)
        self.tree_sl.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tbl.rowconfigure(0, weight=1)
        tbl.columnconfigure(0, weight=1)

        self.tree_sl.tag_configure("odd",  background=CLR_STRIPE1)
        self.tree_sl.tag_configure("even", background=CLR_STRIPE2)
        self.tree_sl.tag_configure("warn", background="#fff3cc")  # нет совпадений в базе

        # Панель редактирования
        edit = ttk.LabelFrame(f, text="Добавить / изменить соответствие", padding=6)
        edit.pack(fill="x", padx=8, pady=(2, 4))

        ttk.Label(edit, text="Фраза из заказа:").grid(row=0, column=0, sticky="w")
        self.sl_fraza_var = tk.StringVar()
        fraza_entry = ttk.Entry(edit, textvariable=self.sl_fraza_var, width=42)
        fraza_entry.grid(row=0, column=1, padx=4, sticky="ew")

        ttk.Label(edit, text="Подстрока в базе:").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.sl_baza_var = tk.StringVar()
        self.sl_baza_var.trace_add("write", lambda *_: (self._preview_baza(), self._sl_update_hints()))
        ent_baza = ttk.Entry(edit, textvariable=self.sl_baza_var, width=30)
        ent_baza.grid(row=0, column=3, padx=4, sticky="ew")

        self.lbl_preview = ttk.Label(edit, text="", foreground=CLR_INFO, width=38)
        self.lbl_preview.grid(row=1, column=1, columnspan=3, sticky="w", pady=(2, 0))

        btn_frame = ttk.Frame(edit)
        btn_frame.grid(row=0, column=4, padx=(8, 0))
        ttk.Button(btn_frame, text="➕ Добавить",  command=self._sl_dobavit).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="✏️ Изменить",  command=self._sl_izmenit).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="🗑 Удалить",   command=self._sl_udalit).pack(side="left", padx=2)

        edit.columnconfigure(1, weight=1)
        edit.columnconfigure(3, weight=1)

        # ── Панель нечётких подсказок для поля «Подстрока в базе» ───
        hints_outer = ttk.LabelFrame(f, text="💡 Подсказки из базы (нечёткий поиск — выберите двойным кликом)", padding=4)
        hints_outer.pack(fill="x", padx=8, pady=(0, 6))

        h_cols = ("hsc", "hnaim", "hof", "hmassa")
        self.tree_hints = ttk.Treeview(hints_outer, columns=h_cols, show="headings",
                                       selectmode="browse", height=4)
        self.tree_hints.heading("hsc",    text="Схожесть")
        self.tree_hints.heading("hnaim",  text="Наименование в базе")
        self.tree_hints.heading("hof",    text="Оформление")
        self.tree_hints.heading("hmassa", text="Масса")
        self.tree_hints.column("hsc",    width=70,  anchor="center", stretch=False)
        self.tree_hints.column("hnaim",  width=320, stretch=True)
        self.tree_hints.column("hof",    width=120, stretch=False)
        self.tree_hints.column("hmassa", width=60,  anchor="center", stretch=False)
        self.tree_hints.tag_configure("h_fuzzy", background="#fffbe6")
        self.tree_hints.tag_configure("h_exact", background="#e8f8e8")

        hsb_h = ttk.Scrollbar(hints_outer, orient="vertical", command=self.tree_hints.yview)
        self.tree_hints.configure(yscrollcommand=hsb_h.set)
        self.tree_hints.pack(side="left", fill="both", expand=True)
        hsb_h.pack(side="right", fill="y")

        def _hints_dblclick(_e):
            """Двойной клик по подсказке → заполнить поле «Подстрока в базе»."""
            sel = self.tree_hints.focus()
            if not sel:
                return
            naim = self.tree_hints.item(sel, "values")[1]
            self.sl_baza_var.set(naim)
        self.tree_hints.bind("<Double-1>", _hints_dblclick)

        # Обновить подсказки и при смене фразы заказа (полезно при выборе из таблицы)
        self.sl_fraza_var.trace_add("write", lambda *_: self._sl_update_hints())

        self.tree_sl.bind("<<TreeviewSelect>>", self._sl_on_select)
        self._reload_slovar()

    # ── ВКЛАДКА ЖУРНАЛ ────────────────────────────────────────────

    def _build_tab_zhurnal(self):
        f = self.tab_zhurnal

        top = ttk.Frame(f)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Button(top, text="↺ Обновить", command=self._reload_zhurnal).pack(side="left")
        self.lbl_zh_count = ttk.Label(top, text="", foreground="#555")
        self.lbl_zh_count.pack(side="left", padx=8)

        cols = ("dt","zadanie","stroka","naim","of","massa","kol","ost_do","ost_posle")
        self.tree_zh = ttk.Treeview(f, columns=cols, show="headings")
        heads = {"dt":"Дата/время","zadanie":"Задание","stroka":"Стр",
                 "naim":"Наименование","of":"Оформление","massa":"Масса",
                 "kol":"Кол","ost_do":"Ост до","ost_posle":"Ост после"}
        widths = {"dt":130,"zadanie":140,"stroka":40,"naim":200,
                  "of":100,"massa":55,"kol":45,"ost_do":55,"ost_posle":60}
        for c in cols:
            self.tree_zh.heading(c, text=heads[c])
            self.tree_zh.column(c, width=widths[c], stretch=(c in ("naim","zadanie")))

        vsb = ttk.Scrollbar(f, orient="vertical",   command=self.tree_zh.yview)
        hsb = ttk.Scrollbar(f, orient="horizontal", command=self.tree_zh.xview)
        self.tree_zh.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree_zh.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        vsb.pack(side="right", fill="y", pady=(0, 8), padx=(0, 8))
        hsb.pack(side="bottom", fill="x")

        self._reload_zhurnal()

    # ── ЗАГРУЗКА ДАННЫХ ───────────────────────────────────────────

    def _reload_baza(self):
        self.tovary = zagruzit_bazu()
        self.slovar = zagruzit_slovar()
        n = len(self.tovary)
        if n:
            self.lbl_baza_stat.config(
                text=f"База: {n} товаров", foreground=CLR_OK)
        else:
            self.lbl_baza_stat.config(
                text=f"База не найдена: {TOVARY_CSV}", foreground=CLR_ERR)
        self._reload_slovar()

    def _check_vm(self):
        ok, msg = proverit_vm()
        self.lbl_vm_stat.config(text=f"VM: {msg}",
                                foreground=CLR_OK if ok else CLR_ERR)

    def _primenit_stranicu(self):
        """Применить размер страницы из поля GUI (PageUp/PageDown)."""
        global RAZMER_STRANICY
        try:
            v = int(self.stranica_var.get())
        except (tk.TclError, ValueError):
            return
        if v < 1:
            v = 1
            self.stranica_var.set(1)
        RAZMER_STRANICY = v
        self._log(f"Размер страницы: PageUp/PageDown = {v} строк.")

    # ── ОТКРЫТИЕ ФАЙЛА ────────────────────────────────────────────

    def _otkryt_excel(self):
        path = filedialog.askopenfilename(
            title="Файл заказа (Excel)",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Все файлы", "*.*")])
        if path:
            self._zagruzit_zadanie(path, "xlsx")

    def _otkryt_txt(self):
        path = filedialog.askopenfilename(
            title="Файл задания (TXT)",
            filetypes=[("TXT", "*.txt"), ("Все файлы", "*.*")])
        if path:
            self._zagruzit_zadanie(path, "txt")

    def _otkryt_iz_papki(self):
        ZADANIYA_DIR.mkdir(parents=True, exist_ok=True)
        path = filedialog.askopenfilename(
            title="Выберите файл из папки zadaniya",
            initialdir=str(ZADANIYA_DIR),
            filetypes=[("TXT", "*.txt"), ("Excel", "*.xlsx *.xls"), ("Все", "*.*")])
        if path:
            ext = Path(path).suffix.lower()
            self._zagruzit_zadanie(path, "xlsx" if ext in (".xlsx", ".xls", ".xlsm") else "txt")

    def _zagruzit_zadanie(self, path: str, fmt: str):
        self.lbl_file.config(text=Path(path).name, foreground="#222")
        if fmt == "xlsx":
            pozicii, oshibki = razobrat_zadanie_xlsx(path)
        else:
            pozicii, oshibki = razobrat_zadanie_txt(path)

        if oshibki:
            msg = "\n".join(f"  строка {n}: {raw!r} — {p}" for n,raw,p in oshibki[:10])
            messagebox.showwarning("Ошибки разбора", f"Пропущены строки:\n{msg}")

        self._postroit_plan(pozicii)
        self._current_path = path

    def _postroit_plan(self, pozicii: list):
        """Сопоставить позиции задания с базой через словарь → заполнить таблицу."""
        self.plan_rows = []
        seen_idx: set[int] = set()

        for name, oform, massa, kol in pozicii:
            fraza, naideny, score = naiti_cherez_slovar(name, self.tovary, self.slovar)

            if fraza is None:
                # Не в словаре
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "miss", "idx": None, "t": None, "varianty": [], "score": 0.0,
                })
            elif not naideny:
                # В словаре, но в базе нет
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "miss", "idx": None, "t": None, "varianty": [], "score": 0.0,
                })
            elif len(naideny) == 1:
                idx, t = naideny[0]
                # Дедупликация: суммируем
                if idx in seen_idx:
                    for p in self.plan_rows:
                        if p["idx"] == idx:
                            p["kol"] = str(int(p["kol"]) + int(kol))
                            break
                    continue
                seen_idx.add(idx)
                # Отключённый товар — пропускаем полностью
                if not t.get("active", True):
                    continue
                # Резерв: не вводим автоматически, помечаем отдельно
                st_val = "rezerv" if t.get("rezerv") else "ok"
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": st_val, "idx": idx, "t": t, "varianty": [], "score": score,
                })
            else:
                # Несколько вариантов — нужен выбор
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "multi", "idx": None, "t": None, "varianty": naideny, "score": score,
                })

        if SORTIROVAT:
            def sort_key(p):
                idx = p.get("idx")
                return (0, idx) if idx is not None else (1, 0)
            self.plan_rows.sort(key=sort_key)

        self._refresh_plan_tree()
        self._update_plan_stat()
        self._update_start_btn()

    def _refresh_plan_tree(self):
        self.tree_plan.delete(*self.tree_plan.get_children())
        ok_n = 0
        for i, p in enumerate(self.plan_rows, 1):
            st = p["status"]
            if st == "rezerv":
                tag        = "rezerv"
                baza_label = format_tovar(p["t"]) + "  [РЕЗЕРВ]"
                stroka     = str(p["idx"] + 1)
                ost        = str(p["t"]["ostatok"])
                stat_label = "📌 резерв"
            elif st == "ok":
                ok_n += 1
                if p.get("score", 1.0) < 0.7:
                    tag = "warn"
                else:
                    tag = "ok" if ok_n % 2 == 1 else "ok2"
                baza_label = format_tovar(p["t"])
                stroka     = str(p["idx"] + 1)
                ost        = str(p["t"]["ostatok"])
                sc_pct = int(p.get("score", 1.0) * 100)
                stat_label = f"✓ {sc_pct}%" if sc_pct < 100 else "✓ найдено"
            elif st == "multi":
                tag = "multi"
                baza_label = f"[{len(p['varianty'])} варианта — двойной клик]"
                stroka = "?"
                ost    = "?"
                stat_label = "⚠ выбрать"
            else:
                tag = "miss"
                baza_label = "— не найдено —"
                stroka = "—"
                ost    = "—"
                stat_label = "✗ нет в базе"

            self.tree_plan.insert("", "end", iid=str(i), tags=(tag,),
                values=(i, p["name"], p["kol"], "→", baza_label, stroka, ost, stat_label))

    def _update_plan_stat(self):
        ok    = sum(1 for p in self.plan_rows if p["status"] == "ok")
        multi = sum(1 for p in self.plan_rows if p["status"] == "multi")
        miss  = sum(1 for p in self.plan_rows if p["status"] == "miss")
        rez   = sum(1 for p in self.plan_rows if p["status"] == "rezerv")
        total = len(self.plan_rows)
        self.lbl_plan_stat.config(
            text=(f"Всего: {total}  |  ✓ {ok}  |  📌 резерв: {rez}"
                  f"  |  ⚠ выбрать: {multi}  |  ✗ не найдено: {miss}"),
            foreground=CLR_OK if miss == 0 and multi == 0 else CLR_WARN)

    def _update_start_btn(self):
        ok = sum(1 for p in self.plan_rows if p["status"] == "ok")
        self.btn_start.config(state="normal" if ok > 0 else "disabled")

    # ── ДВОЙНОЙ КЛИК ПО СТРОКЕ ПЛАНА ─────────────────────────────

    def _on_plan_dblclick(self, event):
        iid = self.tree_plan.focus()
        if not iid:
            return
        idx_plan = int(iid) - 1
        p = self.plan_rows[idx_plan]

        if p["status"] == "multi":
            self._dialog_vybrat_variant(idx_plan, p)
        elif p["status"] == "miss":
            self._dialog_dobavit_sootvetstvie(p["name"])

    def _dialog_vybrat_variant(self, idx_plan: int, p: dict):
        """Диалог выбора варианта из нескольких."""
        dlg = tk.Toplevel(self)
        dlg.title("Выберите товар")
        dlg.geometry("560x320")
        dlg.grab_set()

        ttk.Label(dlg, text=f"Заказ: {p['name']}  (кол: {p['kol']})",
                  font=("Arial", 10, "bold")).pack(padx=12, pady=(10, 4))
        ttk.Label(dlg, text="Выберите товар из базы двойным кликом:").pack(padx=12)

        tree = ttk.Treeview(dlg, columns=("stroka", "naim", "ost"), show="headings",
                            selectmode="browse", height=8)
        tree.heading("stroka", text="Стр"); tree.column("stroka", width=45, anchor="center")
        tree.heading("naim",   text="Товар в базе"); tree.column("naim", width=340)
        tree.heading("ost",    text="Остаток"); tree.column("ost", width=70, anchor="center")
        tree.pack(fill="both", expand=True, padx=12, pady=6)

        for ci, ct in p["varianty"]:
            tree.insert("", "end", iid=str(ci),
                        values=(ci+1, format_tovar(ct), ct["ostatok"]))

        def apply(event=None):
            sel = tree.focus()
            if not sel:
                return
            ci  = int(sel)
            ct  = next(t for i, t in p["varianty"] if i == ci)
            self.plan_rows[idx_plan].update(
                status="ok", idx=ci, t=ct, varianty=[])
            if SORTIROVAT:
                self.plan_rows.sort(key=lambda p: (0, p["idx"]) if p.get("idx") is not None else (1,0))
            self._refresh_plan_tree()
            self._update_plan_stat()
            self._update_start_btn()
            dlg.destroy()

        tree.bind("<Double-1>", apply)
        ttk.Button(dlg, text="Выбрать", command=apply).pack(pady=6)

    def _dialog_dobavit_sootvetstvie(self, fraza: str):
        """Диалог добавления нового соответствия прямо из плана.
        При открытии сразу показывает нечёткие подсказки из базы;
        при вводе в поле поиска переключается на точное совпадение.
        """
        dlg = tk.Toplevel(self)
        dlg.title("Добавить соответствие")
        dlg.geometry("680x520")
        dlg.grab_set()

        ttk.Label(dlg, text="Строка из заказа (фраза для словаря):",
                  font=("Arial", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 0))
        fraza_var = tk.StringVar(value=fraza)
        ttk.Entry(dlg, textvariable=fraza_var, width=80).pack(padx=12, pady=(2, 6), fill="x")

        # Подсказка о режиме отображения
        self._dlg_mode_lbl = ttk.Label(
            dlg,
            text="💡 Показаны нечёткие совпадения с фразой заказа. Введите текст в поиск для точного фильтра.",
            foreground=CLR_INFO, wraplength=640)
        self._dlg_mode_lbl.pack(anchor="w", padx=12)

        sf = ttk.Frame(dlg)
        sf.pack(fill="x", padx=12, pady=(4, 2))
        ttk.Label(sf, text="🔍 Поиск по базе:").pack(side="left")
        filter_var = tk.StringVar()
        filter_entry = ttk.Entry(sf, textvariable=filter_var, width=40)
        filter_entry.pack(side="left", padx=4)
        ttk.Button(sf, text="✖ Сбросить",
                   command=lambda: filter_var.set("")).pack(side="left", padx=2)

        res_frame = ttk.Frame(dlg)
        res_frame.pack(fill="both", expand=True, padx=12, pady=4)

        cols = ("score", "stroka", "naim", "of", "massa")
        tree2 = ttk.Treeview(res_frame, columns=cols, show="headings",
                              selectmode="browse", height=10)
        tree2.heading("score",  text="Схожесть")
        tree2.heading("stroka", text="Стр")
        tree2.heading("naim",   text="Наименование")
        tree2.heading("of",     text="Оформление")
        tree2.heading("massa",  text="Масса")
        tree2.column("score",  width=70,  anchor="center", stretch=False)
        tree2.column("stroka", width=42,  anchor="center", stretch=False)
        tree2.column("naim",   width=260, stretch=True)
        tree2.column("of",     width=120, stretch=False)
        tree2.column("massa",  width=60,  anchor="center", stretch=False)
        tree2.tag_configure("fuzzy",  background="#fffbe6")   # нечёткий режим — светло-жёлтый
        tree2.tag_configure("exact",  background=CLR_STRIPE1) # точный — белый
        tree2.tag_configure("exact2", background=CLR_STRIPE2)
        vsb2 = ttk.Scrollbar(res_frame, orient="vertical", command=tree2.yview)
        tree2.configure(yscrollcommand=vsb2.set)
        tree2.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")

        def refresh_tree(*_):
            kw = filter_var.get().strip()
            tree2.delete(*tree2.get_children())

            if kw:
                # Точный режим: подстрока входит в название
                kn = _norm(kw)
                self._dlg_mode_lbl.config(
                    text="🔍 Точный поиск по введённому тексту.",
                    foreground="#444")
                n = 0
                for i, t in enumerate(self.tovary):
                    if kn not in _norm(t["naimenovanie"]):
                        continue
                    tag = "exact" if n % 2 == 0 else "exact2"
                    tree2.insert("", "end", iid=str(i), tags=(tag,),
                                 values=("—", i+1, t["naimenovanie"],
                                         t["oformlenie"], t["massa"]))
                    n += 1
                if n == 0:
                    self._dlg_mode_lbl.config(
                        text="⚠ Точных совпадений нет. Попробуйте другое слово.",
                        foreground=CLR_ERR)
            else:
                # Нечёткий режим: показать подсказки по фразе заказа
                self._dlg_mode_lbl.config(
                    text="💡 Нечёткие совпадения с фразой заказа. Введите текст для точного поиска.",
                    foreground=CLR_INFO)
                hits = fuzzy_podbor(fraza_var.get(), self.tovary, top_n=15)
                if hits:
                    for sc, i, t in hits:
                        pct = f"{int(sc*100)}%"
                        tree2.insert("", "end", iid=str(i), tags=("fuzzy",),
                                     values=(pct, i+1, t["naimenovanie"],
                                             t["oformlenie"], t["massa"]))
                else:
                    self._dlg_mode_lbl.config(
                        text="⚠ Нечётких совпадений не найдено. Введите текст в поиск.",
                        foreground=CLR_ERR)

        filter_var.trace_add("write", refresh_tree)
        refresh_tree()  # сразу показать нечёткие подсказки

        self.lbl_dlg_stat = ttk.Label(dlg, text="", foreground=CLR_INFO)
        self.lbl_dlg_stat.pack(anchor="w", padx=12)

        def apply(auto: bool = False):
            sel = tree2.focus()
            if not sel:
                if not auto:
                    messagebox.showwarning("Выберите товар", "Кликните на строку в таблице.", parent=dlg)
                return
            ci   = int(sel)
            ct   = self.tovary[ci]
            baza = ct["naimenovanie"]
            frz  = fraza_var.get().strip()
            if not frz:
                if not auto:
                    messagebox.showwarning("Пустая фраза", "Введите фразу из заказа.", parent=dlg)
                return
            pairs = zagruzit_slovar()
            if not any(_norm(f) == _norm(frz) for f, _ in pairs):
                pairs.append((frz, baza))
                sohranit_slovar(pairs)
            self.slovar = zagruzit_slovar()
            self._reload_slovar()
            for rp in self.plan_rows:
                if rp["name"] == fraza and rp["status"] == "miss":
                    rp.update(status="ok", idx=ci, t=ct, varianty=[])
            if SORTIROVAT:
                self.plan_rows.sort(key=lambda p: (0, p["idx"]) if p.get("idx") is not None else (1,0))
            self._refresh_plan_tree()
            self._update_plan_stat()
            self._update_start_btn()
            action = "Авто-добавлено (100%)" if auto else "Добавлено в словарь"
            self._log(f"{action}: {frz!r} → {baza!r}")
            dlg.destroy()

        # ── Авто-применение при единственном 100%-м совпадении ───
        # Выполняется ПОСЛЕ построения UI диалога, через after(0),
        # чтобы dlg успел инициализироваться перед destroy().
        def _check_auto():
            hits = fuzzy_podbor(fraza, self.tovary, top_n=2)
            if not hits or hits[0][0] < 1.0:
                return
            # Ровно одно 100%-е совпадение (второй хит либо отсутствует, либо <100%)
            if len(hits) > 1 and hits[1][0] >= 1.0:
                return
            _sc, _ci, _ct = hits[0]
            frz = fraza.strip()
            pairs = zagruzit_slovar()
            if not any(_norm(f) == _norm(frz) for f, _ in pairs):
                pairs.append((frz, _ct["naimenovanie"]))
                sohranit_slovar(pairs)
            self.slovar = zagruzit_slovar()
            self._reload_slovar()
            for rp in self.plan_rows:
                if rp["name"] == fraza and rp["status"] == "miss":
                    rp.update(status="ok", idx=_ci, t=_ct, varianty=[])
            self._refresh_plan_tree()
            self._update_plan_stat()
            self._update_start_btn()
            self._log(f"Авто-добавлено (100%): {frz!r} → {_ct['naimenovanie']!r}")
            if dlg.winfo_exists():
                dlg.destroy()

        dlg.after(0, _check_auto)

        tree2.bind("<Double-1>", lambda _e: apply())
        ttk.Button(dlg, text="✅ Добавить в словарь и применить",
                   command=apply).pack(pady=8)

    # ── СЛОВАРЬ ───────────────────────────────────────────────────

    def _reload_slovar(self):
        self.slovar = zagruzit_slovar()
        self._filter_slovar()

    def _filter_slovar(self):
        kw = self.slovar_filter.get().lower().strip()
        self.tree_sl.delete(*self.tree_sl.get_children())
        n = 0
        for i, (fraza, baza) in enumerate(self.slovar):
            if kw and kw not in fraza.lower() and kw not in baza.lower():
                continue
            bn = _norm(baza)
            sovpad = sum(1 for t in self.tovary if bn in _norm(t["naimenovanie"]))
            tag = "warn" if sovpad == 0 else ("odd" if i % 2 == 0 else "even")
            sv_text = str(sovpad) if sovpad > 0 else "⚠ нет в базе"
            self.tree_sl.insert("", "end", iid=str(i), tags=(tag,),
                                values=(fraza, baza, sv_text))
            n += 1
        self.lbl_slovar_count.config(text=f"Записей: {n}")

    def _sl_on_select(self, _=None):
        iid = self.tree_sl.focus()
        if not iid:
            return
        i = int(iid)
        if 0 <= i < len(self.slovar):
            fraza, baza = self.slovar[i]
            self.sl_fraza_var.set(fraza)
            self.sl_baza_var.set(baza)

    def _sl_update_hints(self):
        """Обновить панель нечётких подсказок под полем «Подстрока в базе»."""
        if not hasattr(self, "tree_hints"):
            return
        self.tree_hints.delete(*self.tree_hints.get_children())
        query = self.sl_baza_var.get().strip()

        if not query:
            # Если поле пустое — подсказки по фразе из заказа
            query = self.sl_fraza_var.get().strip()
            if not query:
                return
            hits = fuzzy_podbor(query, self.tovary, top_n=8)
            tag = "h_fuzzy"
        else:
            # Комбо: сначала точные (подстрока), затем нечёткие
            kn = _norm(query)
            exact = [(1.0, i, t) for i, t in enumerate(self.tovary)
                     if kn in _norm(t["naimenovanie"])]
            fuzzy_hits = fuzzy_podbor(query, self.tovary, top_n=8)
            # Слить, точные первыми, без дублей
            seen = {i for _, i, _ in exact}
            combined = exact + [(sc, i, t) for sc, i, t in fuzzy_hits if i not in seen]
            hits = combined[:10]
            tag = "h_exact" if exact else "h_fuzzy"

        for sc, i, t in hits:
            pct = f"{int(sc * 100)}%"
            row_tag = "h_exact" if sc >= 0.99 else tag
            self.tree_hints.insert("", "end", iid=str(i), tags=(row_tag,),
                                   values=(pct, t["naimenovanie"],
                                           t["oformlenie"], t["massa"]))

    def _preview_baza(self):
        kw = _norm(self.sl_baza_var.get())
        if not kw:
            self.lbl_preview.config(text="", foreground=CLR_INFO)
            return
        sovpad = [t for t in self.tovary if kw in _norm(t["naimenovanie"])]
        if not sovpad:
            self.lbl_preview.config(text="⚠ Нет совпадений в базе!", foreground=CLR_ERR)
        elif len(sovpad) == 1:
            self.lbl_preview.config(text=f"✓ 1 товар: {sovpad[0]['naimenovanie']}", foreground=CLR_OK)
        else:
            names = ", ".join(t["naimenovanie"] for t in sovpad[:3])
            self.lbl_preview.config(
                text=f"⚠ {len(sovpad)} товаров: {names}{'…' if len(sovpad)>3 else ''}",
                foreground=CLR_WARN)

    def _sl_dobavit(self):
        fraza = self.sl_fraza_var.get().strip()
        baza  = self.sl_baza_var.get().strip()
        if not fraza or not baza:
            messagebox.showwarning("Пусто", "Заполните обе строки."); return
        pairs = zagruzit_slovar()
        if any(_norm(f) == _norm(fraza) for f, _ in pairs):
            messagebox.showinfo("Уже есть", f"Фраза {fraza!r} уже в словаре."); return
        pairs.append((fraza, baza))
        sohranit_slovar(pairs)
        self._reload_slovar()
        self._log(f"Словарь: добавлено {fraza!r} → {baza!r}")

    def _sl_izmenit(self):
        iid = self.tree_sl.focus()
        if not iid:
            messagebox.showwarning("Не выбрано", "Кликните строку в таблице."); return
        i     = int(iid)
        fraza = self.sl_fraza_var.get().strip()
        baza  = self.sl_baza_var.get().strip()
        if not fraza or not baza:
            messagebox.showwarning("Пусто", "Заполните обе строки."); return
        pairs = zagruzit_slovar()
        if 0 <= i < len(pairs):
            pairs[i] = (fraza, baza)
            sohranit_slovar(pairs)
            self._reload_slovar()
            self._log(f"Словарь: изменено → {fraza!r} → {baza!r}")

    def _sl_udalit(self):
        iid = self.tree_sl.focus()
        if not iid:
            messagebox.showwarning("Не выбрано", "Кликните строку."); return
        i = int(iid)
        pairs = zagruzit_slovar()
        if not (0 <= i < len(pairs)):
            return
        fraza, baza = pairs[i]
        if not messagebox.askyesno("Удалить?", f"Удалить:\n{fraza!r} → {baza!r}?"):
            return
        pairs.pop(i)
        sohranit_slovar(pairs)
        self._reload_slovar()
        self._log(f"Словарь: удалено {fraza!r}")

    # ── ЖУРНАЛ ────────────────────────────────────────────────────

    def _reload_zhurnal(self):
        self.tree_zh.delete(*self.tree_zh.get_children())
        if not ISTORIYA_CSV.exists():
            self.lbl_zh_count.config(text="Журнал пуст")
            return
        tekst = prochitat_tekst(str(ISTORIYA_CSV))
        rows  = list(csv.reader(tekst.splitlines(), delimiter=";"))
        if len(rows) <= 1:
            self.lbl_zh_count.config(text="Журнал пуст"); return
        data_rows = rows[1:][-500:]  # последние 500
        for i, row in enumerate(reversed(data_rows)):
            while len(row) < 9:
                row.append("")
            tag = "odd" if i % 2 == 0 else "even"
            self.tree_zh.insert("", "end", tags=(tag,),
                                values=row[:9])
        self.tree_zh.tag_configure("odd",  background=CLR_STRIPE1)
        self.tree_zh.tag_configure("even", background=CLR_STRIPE2)
        self.lbl_zh_count.config(text=f"Записей: {len(data_rows)}")

    # ── ВВОД ─────────────────────────────────────────────────────

    def _log(self, msg: str):
        def _do():
            self.txt_log.config(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            self.txt_log.insert("end", f"[{ts}] {msg}\n")
            self.txt_log.see("end")
            self.txt_log.config(state="disabled")
        self.after(0, _do)

    def _zapustit_vvod(self):
        self._primenit_stranicu()   # взять актуальный размер страницы из поля
        ok_rows = [p for p in self.plan_rows if p["status"] == "ok"]
        if not ok_rows:
            messagebox.showwarning("Нет данных", "Нет готовых строк для ввода."); return

        multi = sum(1 for p in self.plan_rows if p["status"] == "multi")
        miss  = sum(1 for p in self.plan_rows if p["status"] == "miss")
        rez_rows = [p for p in self.plan_rows if p["status"] == "rezerv"]
        if self.rezerv_vvod_var.get():
            ok_rows = sorted(ok_rows + rez_rows, key=lambda p: p.get("idx", 9999))
        rez = len(rez_rows)
        if multi > 0 or miss > 0 or (rez > 0 and not self.rezerv_vvod_var.get()):
            msg = []
            if rez and not self.rezerv_vvod_var.get():
                detail = "\n".join(
                    f"  стр {p['idx']+1}: {p['t']['naimenovanie']} = {p['kol']}"
                    for p in rez_rows if p.get("t"))
                msg.append(f"📌 {rez} РЕЗЕРВ — добавить вручную:\n{detail}")
            if multi: msg.append(f"⚠ {multi} строк требуют выбора")
            if miss:  msg.append(f"✗ {miss} не найдено в базе")
            if not messagebox.askyesno(
                "Внимание перед запуском",
                "\n".join(msg) + f"\n\nЗапустить ввод {len(ok_rows)} строк?"):
                return

        vm_ok, vm_msg = proverit_vm()
        if not vm_ok:
            messagebox.showerror("VM недоступна", vm_msg); return

        self._stop_vvod.clear()
        self.btn_start.config(state="disabled")
        self._log(f"Запуск ввода: {len(ok_rows)} позиций")

        def run():
            try:
                self._vvod_loop(ok_rows)
            except Exception as e:
                self._log(f"ОШИБКА: {e}")
            finally:
                self.after(0, lambda: self.btn_start.config(state="normal"))

        self._vvod_thread = threading.Thread(target=run, daemon=True)
        self._vvod_thread.start()

    def _stop_vvod_fn(self):
        self._stop_vvod.set()
        self._log("⏹ Остановка запрошена...")

    def _vvod_loop(self, ok_rows: list[dict]):
        """Основной цикл ввода — выполняется в отдельном потоке."""
        z_name = getattr(self, "_current_path", "задание")
        z_name = Path(z_name).name if z_name else "задание"

        self._log("Начинаю ввод...")

        tek_poz  = 0
        istoriya = []
        vvedeno  = 0

        # Записать резервные позиции в историю как ручные
        teper = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for p in self.plan_rows:
            if p["status"] == "rezerv" and p["t"]:
                t = p["t"]
                istoriya.append([
                    teper, z_name, p["idx"] + 1,
                    t["naimenovanie"], t["oformlenie"], t["massa"],
                    p["kol"], t["ostatok"], t["ostatok"],
                    "РЕЗЕРВ — добавить вручную",
                ])
                self._log(f"  📌 РЕЗЕРВ: {t['naimenovanie']} = {p['kol']} — добавить вручную!")

        for n, p in enumerate(ok_rows, 1):
            if self._stop_vvod.is_set():
                self._log("⏹ Ввод остановлен пользователем.")
                break

            idx = p["idx"]
            t   = p["t"]
            kol = int(p["kol"])

            try:
                peremestit_kursor(tek_poz, idx)
                vvesti_kolichestvo(str(kol))
                tek_poz = idx
            except subprocess.CalledProcessError as e:
                err = (e.stderr or b"").decode("utf-8", "ignore") if isinstance(e.stderr, bytes) else str(e.stderr)
                self._log(f"  ! VBoxManage ошибка: {err.strip()}")
                break

            ost_do    = t["ostatok"]
            ost_posle = ost_do - kol
            t["ostatok"] = ost_posle

            istoriya.append([
                teper, z_name, idx + 1,
                t["naimenovanie"], t["oformlenie"], t["massa"],
                kol, ost_do, ost_posle, "1" if t["rezerv"] else "0",
            ])
            vvedeno += 1

            label = format_tovar(t)
            self._log(f"  [{n}/{len(ok_rows)}] стр {idx+1}: {label} = {kol}  (ост {ost_do}→{ost_posle})  ✓")

            if PAUZA_MEZHDU > 0:
                time.sleep(PAUZA_MEZHDU)

        # Сохраняем результаты
        if istoriya:
            sohranit_bazu(self.tovary)
            zapisat_istoriyu(istoriya)
            self._log(f"✓ Введено {vvedeno}/{len(ok_rows)} позиций. База и журнал обновлены.")
            self.after(0, self._reload_zhurnal)
        else:
            self._log("Ничего не введено.")


    # ╔══════════════════════════════════════════════════════════════╗
    # ║  СОХРАНЕНИЕ / ЗАГРУЗКА ИМЕНОВАННЫХ ЗАКАЗОВ                  ║
    # ╚══════════════════════════════════════════════════════════════╝



    def _build_tab_ruchnoy(self):
        f = self.tab_ruchnoy

        ttk.Label(f, text="Введите позиции заказа вручную. Формат строки:  Название;количество",
                  foreground="#444").pack(anchor="w", padx=8, pady=(6, 2))

        # Поле ввода названия
        nf = ttk.LabelFrame(f, text="Название заказа", padding=4)
        nf.pack(fill="x", padx=8, pady=(2, 4))
        self.ruchnoy_name_var = tk.StringVar()
        ttk.Entry(nf, textvariable=self.ruchnoy_name_var, width=50).pack(side="left", fill="x", expand=True)

        # Таблица позиций
        tbl = ttk.LabelFrame(f, text="Позиции заказа", padding=4)
        tbl.pack(fill="both", expand=True, padx=8, pady=2)

        cols = ("num","pos_b","naim","oform","massa","kol","status")
        self.tree_ruchnoy = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        hd = {"num":"№","pos_b":"Стр.","naim":"Наименование","oform":"Оформление","massa":"Масса","kol":"Кол","status":"Статус"}
        wd = {"num":36,"pos_b":52,"naim":280,"oform":110,"massa":58,"kol":52,"status":90}
        an = {"num":"center","pos_b":"center","massa":"center","kol":"center","status":"center"}
        for c in cols:
            self.tree_ruchnoy.heading(c, text=hd[c])
            self.tree_ruchnoy.column(c, width=wd[c], stretch=(c=="naim"), anchor=an.get(c,"w"))
        vsb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree_ruchnoy.yview)
        self.tree_ruchnoy.configure(yscrollcommand=vsb.set)
        self.tree_ruchnoy.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tbl.rowconfigure(0, weight=1); tbl.columnconfigure(0, weight=1)

        self.tree_ruchnoy.tag_configure("odd",  background="#ffffff")
        self.tree_ruchnoy.tag_configure("even", background="#eef2f8")
        self.tree_ruchnoy.bind("<Double-1>", self._ruchnoy_edit_row)

        # Панель добавления/редактирования
        add_frame = ttk.LabelFrame(f, text="Добавить / изменить позицию", padding=6)
        add_frame.pack(fill="x", padx=8, pady=2)

        ttk.Label(add_frame, text="Наименование:").grid(row=0, column=0, sticky="w")
        self.ruchnoy_naim_var = tk.StringVar()
        naim_entry = ttk.Entry(add_frame, textvariable=self.ruchnoy_naim_var, width=46)
        naim_entry.grid(row=0, column=1, sticky="ew", padx=4)

        # Автодополнение из базы
        self.ruchnoy_hint_lb = tk.Listbox(add_frame, height=4, font=("Arial", 9))
        self.ruchnoy_hint_lb.grid(row=1, column=1, sticky="ew", padx=4, pady=(0, 2))
        self.ruchnoy_hint_lb.bind("<ButtonRelease-1>", self._ruchnoy_pick_hint)
        self.ruchnoy_naim_var.trace_add("write", self._ruchnoy_update_hints)

        ttk.Label(add_frame, text="Количество:").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.ruchnoy_kol_var = tk.StringVar(value="1")
        ttk.Entry(add_frame, textvariable=self.ruchnoy_kol_var, width=8).grid(row=0, column=3, padx=4)

        btn_f = ttk.Frame(add_frame)
        btn_f.grid(row=0, column=4, padx=(8, 0))
        ttk.Button(btn_f, text="➕ Добавить",  command=self._ruchnoy_dobavit).pack(side="left", padx=2)
        ttk.Button(btn_f, text="✏️ Заменить", command=self._ruchnoy_zamenit).pack(side="left", padx=2)
        ttk.Button(btn_f, text="🗑 Удалить",  command=self._ruchnoy_udalit).pack(side="left", padx=2)

        add_frame.columnconfigure(1, weight=1)

        # Кнопки управления порядком и действий
        bot = ttk.Frame(f)
        bot.pack(fill="x", padx=8, pady=(2, 8))
        ttk.Button(bot, text="⬆ Вверх",       command=self._ruchnoy_up).pack(side="left", padx=2)
        ttk.Button(bot, text="⬇ Вниз",        command=self._ruchnoy_down).pack(side="left", padx=2)
        ttk.Button(bot, text="🗑 Очистить всё", command=self._ruchnoy_clear).pack(side="left", padx=8)
        ttk.Button(bot, text="▶ Перенести в план и выполнить →",
                   command=self._ruchnoy_v_plan).pack(side="right", padx=4)
        ttk.Button(bot, text="💾 Сохранить заказ",
                   command=self._ruchnoy_sohranit).pack(side="right", padx=4)

        # Внутренние данные
        self._ruchnoy_rows: list[tuple[str, str]] = []  # (naim, kol)

    def _ruchnoy_refresh(self):
        self.tree_ruchnoy.delete(*self.tree_ruchnoy.get_children())
        for i, (naim, kol) in enumerate(self._ruchnoy_rows, 1):
            tag = "odd" if i % 2 == 1 else "even"
            fraza, naideny, _sc = naiti_cherez_slovar(naim, self.tovary, self.slovar)
            if naideny and len(naideny) == 1:
                idx_b, t = naideny[0]
                pos_b  = str(idx_b + 1)
                oform  = t.get("oformlenie", "")
                massa  = t.get("massa", "")
                status = "📌 резерв" if t.get("rezerv") else "✓"
            elif naideny:
                pos_b = "?"; oform = ""; massa = ""; status = f"⚠{len(naideny)}"
            else:
                pos_b = "—"; oform = ""; massa = ""; status = "✗"
            self.tree_ruchnoy.insert("", "end", iid=str(i-1), tags=(tag,),
                                     values=(i, pos_b, naim, oform, massa, kol, status))

    def _ruchnoy_update_hints(self, *_):
        kw = _norm(self.ruchnoy_naim_var.get())
        self.ruchnoy_hint_lb.delete(0, "end")
        if len(kw) < 2:
            return
        matches = [t["naimenovanie"] for t in self.tovary
                   if kw in _norm(t["naimenovanie"])][:8]
        for m in matches:
            self.ruchnoy_hint_lb.insert("end", m)

    def _ruchnoy_pick_hint(self, _=None):
        sel = self.ruchnoy_hint_lb.curselection()
        if sel:
            self.ruchnoy_naim_var.set(self.ruchnoy_hint_lb.get(sel[0]))
            self.ruchnoy_hint_lb.delete(0, "end")

    def _ruchnoy_dobavit(self):
        naim = self.ruchnoy_naim_var.get().strip()
        kol  = self.ruchnoy_kol_var.get().strip()
        if not naim:
            messagebox.showwarning("Пусто", "Введите наименование.", parent=self); return
        if not kol.isdigit() or int(kol) <= 0:
            messagebox.showwarning("Ошибка", "Количество должно быть целым числом > 0.", parent=self); return
        self._ruchnoy_rows.append((naim, kol))
        self._ruchnoy_refresh()
        self.ruchnoy_naim_var.set("")
        self.ruchnoy_kol_var.set("1")

    def _ruchnoy_zamenit(self):
        iid = self.tree_ruchnoy.focus()
        if not iid:
            messagebox.showwarning("Не выбрано", "Выберите строку для замены.", parent=self); return
        naim = self.ruchnoy_naim_var.get().strip()
        kol  = self.ruchnoy_kol_var.get().strip()
        if not naim or not kol.isdigit():
            messagebox.showwarning("Ошибка", "Заполните наименование и количество.", parent=self); return
        self._ruchnoy_rows[int(iid)] = (naim, kol)
        self._ruchnoy_refresh()

    def _ruchnoy_edit_row(self, _=None):
        iid = self.tree_ruchnoy.focus()
        if not iid:
            return
        naim, kol = self._ruchnoy_rows[int(iid)]
        self.ruchnoy_naim_var.set(naim)
        self.ruchnoy_kol_var.set(kol)

    def _ruchnoy_udalit(self):
        iid = self.tree_ruchnoy.focus()
        if not iid:
            return
        self._ruchnoy_rows.pop(int(iid))
        self._ruchnoy_refresh()

    def _ruchnoy_up(self):
        iid = self.tree_ruchnoy.focus()
        if not iid:
            return
        i = int(iid)
        if i > 0:
            self._ruchnoy_rows[i], self._ruchnoy_rows[i-1] = (
                self._ruchnoy_rows[i-1], self._ruchnoy_rows[i])
            self._ruchnoy_refresh()
            self.tree_ruchnoy.focus(str(i-1))
            self.tree_ruchnoy.selection_set(str(i-1))

    def _ruchnoy_down(self):
        iid = self.tree_ruchnoy.focus()
        if not iid:
            return
        i = int(iid)
        if i < len(self._ruchnoy_rows) - 1:
            self._ruchnoy_rows[i], self._ruchnoy_rows[i+1] = (
                self._ruchnoy_rows[i+1], self._ruchnoy_rows[i])
            self._ruchnoy_refresh()
            self.tree_ruchnoy.focus(str(i+1))
            self.tree_ruchnoy.selection_set(str(i+1))

    def _ruchnoy_clear(self):
        if messagebox.askyesno("Очистить?", "Удалить все строки заказа?"):
            self._ruchnoy_rows.clear()
            self._ruchnoy_refresh()

    def _ruchnoy_sohranit(self):
        name = self.ruchnoy_name_var.get().strip()
        if not name:
            messagebox.showwarning("Нет названия", "Введите название заказа."); return
        if not self._ruchnoy_rows:
            messagebox.showwarning("Пусто", "Нет строк."); return
        ZAKAZY_DIR.mkdir(parents=True, exist_ok=True)
        safe  = re.sub(r'[<>:"/\\\\|?*]', "_", name)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path  = ZAKAZY_DIR / f"{safe}__{stamp}.txt"
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(f"# Заказ: {name}\n")
            for naim, kol in self._ruchnoy_rows:
                f.write(f"{naim};{kol}\n")
        self._log(f"💾 Заказ '{name}' сохранён: {path.name}")
        messagebox.showinfo("Сохранено", f"{path.name}")

    def _ruchnoy_v_plan(self):
        if not self._ruchnoy_rows:
            messagebox.showwarning("Пусто", "Нет строк для переноса."); return
        name = self.ruchnoy_name_var.get().strip()
        if name:
            self.zakaz_name_var.set(name)
        pozicii = [(naim, "", "", kol) for naim, kol in self._ruchnoy_rows]
        self._postroit_plan(pozicii)
        self.nb.select(self.tab_zakaz)
        self._log(f"Ручной заказ перенесён в план: {len(pozicii)} позиций")

    # ╔══════════════════════════════════════════════════════════════╗
    # ║  ВКЛАДКА: СПИСОК ТОВАРОВ (редактирование базы)              ║
    # ╚══════════════════════════════════════════════════════════════╝

    def _build_tab_tovary(self):
        f = self.tab_tovary

        ttk.Label(f, text="Список товаров — порядок строк = порядок в программе. Изменения сохраняются в tovary.csv.",
                  foreground="#444").pack(anchor="w", padx=8, pady=(6, 0))

        # Фильтр
        ff = ttk.Frame(f)
        ff.pack(fill="x", padx=8, pady=2)
        ttk.Label(ff, text="🔍 Фильтр:").pack(side="left")
        self.tov_filter_var = tk.StringVar()
        self.tov_filter_var.trace_add("write", lambda *_: self._tov_filter())
        ttk.Entry(ff, textvariable=self.tov_filter_var, width=28).pack(side="left", padx=4)
        ttk.Label(ff, text="Показать:").pack(side="left", padx=(12, 2))
        self.tov_show_var = tk.StringVar(value="все")
        cb = ttk.Combobox(ff, textvariable=self.tov_show_var,
                          values=["все", "активные", "резерв", "отключённые"],
                          width=12, state="readonly")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _: self._tov_filter())
        self.lbl_tov_count = ttk.Label(ff, text="", foreground="#555")
        self.lbl_tov_count.pack(side="right")

        # Таблица
        tbl = ttk.LabelFrame(f, text="Товары", padding=4)
        tbl.pack(fill="both", expand=True, padx=8, pady=2)

        cols = ("pos", "naim", "of", "massa", "ost", "rezerv", "active")
        self.tree_tov = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        heads  = {"pos":"№","naim":"Наименование","of":"Оформление",
                  "massa":"Масса","ost":"Остаток","rezerv":"Резерв","active":"Вкл"}
        widths = {"pos":42,"naim":290,"of":130,"massa":65,"ost":65,"rezerv":60,"active":40}
        for c in cols:
            self.tree_tov.heading(c, text=heads[c])
            self.tree_tov.column(c, width=widths[c],
                                 stretch=(c == "naim"),
                                 anchor="center" if c in ("pos","ost","rezerv","active") else "w")
        vsb = ttk.Scrollbar(tbl, orient="vertical",   command=self.tree_tov.yview)
        hsb = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree_tov.xview)
        self.tree_tov.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree_tov.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl.rowconfigure(0, weight=1); tbl.columnconfigure(0, weight=1)

        self.tree_tov.tag_configure("active",   background="#ffffff")
        self.tree_tov.tag_configure("active2",  background="#eef2f8")
        self.tree_tov.tag_configure("rezerv",   background=CLR_REZERV)
        self.tree_tov.tag_configure("inactive", background="#e8e8e8", foreground="#888")

        self.tree_tov.bind("<<TreeviewSelect>>", self._tov_on_select)
        self.tree_tov.bind("<Double-1>",         self._tov_toggle_active)

        # Панель редактирования
        ep = ttk.LabelFrame(f, text="Редактировать / добавить товар", padding=6)
        ep.pack(fill="x", padx=8, pady=2)

        ttk.Label(ep, text="Наименование:").grid(row=0, column=0, sticky="w")
        self.tov_naim_var = tk.StringVar()
        ttk.Entry(ep, textvariable=self.tov_naim_var, width=34).grid(row=0, column=1, sticky="ew", padx=4)

        ttk.Label(ep, text="Оформление:").grid(row=0, column=2, sticky="w", padx=(8,0))
        self.tov_of_var = tk.StringVar()
        ttk.Entry(ep, textvariable=self.tov_of_var, width=16).grid(row=0, column=3, padx=4)

        ttk.Label(ep, text="Масса:").grid(row=0, column=4, sticky="w", padx=(4,0))
        self.tov_massa_var = tk.StringVar()
        ttk.Entry(ep, textvariable=self.tov_massa_var, width=7).grid(row=0, column=5, padx=4)

        ttk.Label(ep, text="Остаток:").grid(row=1, column=0, sticky="w", pady=(4,0))
        self.tov_ost_var = tk.StringVar(value="0")
        ttk.Entry(ep, textvariable=self.tov_ost_var, width=8).grid(row=1, column=1, sticky="w", padx=4, pady=(4,0))

        self.tov_rezerv_var = tk.BooleanVar()
        ttk.Checkbutton(ep, text="Резерв", variable=self.tov_rezerv_var).grid(
            row=1, column=2, sticky="w", padx=(8,0), pady=(4,0))

        self.tov_active_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(ep, text="Активен", variable=self.tov_active_var).grid(
            row=1, column=3, sticky="w", pady=(4,0))

        ep.columnconfigure(1, weight=1)

        # Кнопки
        bp = ttk.Frame(f)
        bp.pack(fill="x", padx=8, pady=(2, 6))
        ttk.Button(bp, text="⬆ Вверх",       command=self._tov_up).pack(side="left", padx=2)
        ttk.Button(bp, text="⬇ Вниз",        command=self._tov_down).pack(side="left", padx=2)
        ttk.Button(bp, text="➕ Добавить после выбранного",
                   command=self._tov_dobavit).pack(side="left", padx=8)
        ttk.Button(bp, text="✏️ Сохранить изменения",
                   command=self._tov_izmenit).pack(side="left", padx=2)
        ttk.Button(bp, text="🗑 Удалить",     command=self._tov_udalit).pack(side="left", padx=2)
        ttk.Button(bp, text="↺⏺ Вкл/Выкл (двойной клик)",
                   command=self._tov_toggle_active).pack(side="left", padx=8)
        ttk.Button(bp, text="💾 Сохранить базу",
                   command=self._tov_save).pack(side="right", padx=4)

    def _tov_refresh(self, select_real_idx: int | None = None):
        """Перерисовать таблицу товаров с учётом фильтра."""
        self.tree_tov.delete(*self.tree_tov.get_children())
        kw    = _norm(self.tov_filter_var.get())
        show  = self.tov_show_var.get()
        n_vis = 0
        for real_idx, t in enumerate(self.tovary):
            # Фильтр по режиму показа
            is_rez  = t.get("rezerv", False)
            is_act  = t.get("active", True)
            if show == "активные"   and (not is_act or is_rez): continue
            if show == "резерв"     and not is_rez: continue
            if show == "отключённые" and is_act: continue
            # Текстовый фильтр
            if kw and kw not in _norm(t["naimenovanie"]): continue

            if not is_act:
                tag = "inactive"
            elif is_rez:
                tag = "rezerv"
            else:
                tag = "active" if n_vis % 2 == 0 else "active2"

            rez_mark = "✓" if is_rez else ""
            act_mark = "✓" if is_act else "✗"
            self.tree_tov.insert("", "end", iid=str(real_idx), tags=(tag,),
                                 values=(real_idx + 1, t["naimenovanie"],
                                         t.get("oformlenie",""), t.get("massa",""),
                                         t.get("ostatok", 0), rez_mark, act_mark))
            n_vis += 1

        self.lbl_tov_count.config(text=f"Показано: {n_vis} / {len(self.tovary)}")
        if select_real_idx is not None:
            iid = str(select_real_idx)
            if self.tree_tov.exists(iid):
                self.tree_tov.focus(iid)
                self.tree_tov.selection_set(iid)
                self.tree_tov.see(iid)

    def _tov_filter(self):
        self._tov_refresh()

    def _tov_on_select(self, _=None):
        iid = self.tree_tov.focus()
        if not iid:
            return
        t = self.tovary[int(iid)]
        self.tov_naim_var.set(t["naimenovanie"])
        self.tov_of_var.set(t.get("oformlenie", ""))
        self.tov_massa_var.set(t.get("massa", ""))
        self.tov_ost_var.set(str(t.get("ostatok", 0)))
        self.tov_rezerv_var.set(t.get("rezerv", False))
        self.tov_active_var.set(t.get("active", True))

    def _tov_toggle_active(self, _=None):
        iid = self.tree_tov.focus()
        if not iid:
            return
        t = self.tovary[int(iid)]
        t["active"] = not t.get("active", True)
        self._tov_refresh(int(iid))

    def _tov_dobavit(self):
        naim = self.tov_naim_var.get().strip()
        if not naim:
            messagebox.showwarning("Пусто", "Введите наименование."); return
        try:
            ost = int(self.tov_ost_var.get() or "0")
        except ValueError:
            ost = 0
        new_t = {
            "naimenovanie": naim,
            "oformlenie":   self.tov_of_var.get().strip(),
            "massa":        self.tov_massa_var.get().strip(),
            "ostatok":      ost,
            "rezerv":       self.tov_rezerv_var.get(),
            "active":       self.tov_active_var.get(),
        }
        iid = self.tree_tov.focus()
        if iid:
            insert_after = int(iid) + 1
            self.tovary.insert(insert_after, new_t)
            self._tov_refresh(insert_after)
        else:
            self.tovary.append(new_t)
            self._tov_refresh(len(self.tovary) - 1)
        self._log(f"Добавлен товар: {naim!r}")

    def _tov_izmenit(self):
        iid = self.tree_tov.focus()
        if not iid:
            messagebox.showwarning("Не выбрано", "Выберите строку."); return
        try:
            ost = int(self.tov_ost_var.get() or "0")
        except ValueError:
            ost = 0
        i = int(iid)
        self.tovary[i].update({
            "naimenovanie": self.tov_naim_var.get().strip(),
            "oformlenie":   self.tov_of_var.get().strip(),
            "massa":        self.tov_massa_var.get().strip(),
            "ostatok":      ost,
            "rezerv":       self.tov_rezerv_var.get(),
            "active":       self.tov_active_var.get(),
        })
        self._tov_refresh(i)
        self._log(f"Изменён товар #{i+1}: {self.tovary[i]['naimenovanie']!r}")

    def _tov_udalit(self):
        iid = self.tree_tov.focus()
        if not iid:
            return
        i = int(iid)
        naim = self.tovary[i]["naimenovanie"]
        if not messagebox.askyesno("Удалить?", f"Удалить товар:\n{naim}\n\n"
                                   "Внимание: это сдвинет номера ВСЕХ последующих позиций!"):
            return
        self.tovary.pop(i)
        self._tov_refresh(min(i, len(self.tovary)-1) if self.tovary else None)
        self._log(f"Удалён товар: {naim!r}")

    def _tov_up(self):
        iid = self.tree_tov.focus()
        if not iid:
            return
        i = int(iid)
        if i > 0:
            self.tovary[i], self.tovary[i-1] = self.tovary[i-1], self.tovary[i]
            self._tov_refresh(i-1)

    def _tov_down(self):
        iid = self.tree_tov.focus()
        if not iid:
            return
        i = int(iid)
        if i < len(self.tovary) - 1:
            self.tovary[i], self.tovary[i+1] = self.tovary[i+1], self.tovary[i]
            self._tov_refresh(i+1)

    def _tov_save(self):
        sohranit_bazu(self.tovary)
        self._reload_baza()
        self._tov_refresh()
        self._log(f"✓ База товаров сохранена: {len(self.tovary)} позиций")
        messagebox.showinfo("Сохранено", f"База обновлена: {len(self.tovary)} товаров.")


    # ╔══════════════════════════════════════════════════════════════╗
    # ║  НАВИГАЦИЯ КЛАВИАТУРОЙ                                       ║
    # ╚══════════════════════════════════════════════════════════════╝

    def _setup_keyboard_nav(self):
        self.bind_all("<Control-c>", self._clipboard_copy)
        self.bind_all("<Control-v>", self._clipboard_paste)
        self.bind_all("<Control-a>", self._select_all)
        self.bind_class("TButton", "<Return>", lambda e: e.widget.invoke())
        for attr in ("tree_plan","tree_ruchnoy","tree_tov","tree_sl","tree_zh"):
            tree = getattr(self, attr, None)
            if tree:
                tree.bind("<Return>", lambda e, t=tree: self._tree_enter(t))
                tree.bind("<space>",  lambda e, t=tree: self._tree_enter(t))

    def _tree_enter(self, tree):
        iid = tree.focus()
        if iid:
            tree.event_generate("<Double-1>")

    def _clipboard_copy(self, event=None):
        w = self.focus_get()
        if isinstance(w, (tk.Entry, ttk.Entry, tk.Text)):
            try: w.event_generate("<<Copy>>")
            except Exception: pass

    def _clipboard_paste(self, event=None):
        w = self.focus_get()
        if isinstance(w, (tk.Entry, ttk.Entry, tk.Text)):
            try: w.event_generate("<<Paste>>")
            except Exception: pass

    def _select_all(self, event=None):
        w = self.focus_get()
        if isinstance(w, (tk.Entry, ttk.Entry)):
            w.select_range(0, "end")
        elif isinstance(w, tk.Text):
            w.tag_add("sel", "1.0", "end")

    def _sohranit_zakaz(self):
        nomer = self.zakaz_nomer_var.get().strip()
        name  = self.zakaz_name_var.get().strip()
        if not name and not nomer:
            messagebox.showwarning("Нет названия", "Введите номер или название заказа."); return
        pozicii = [f"{p['name']};{p['kol']}" for p in self.plan_rows
                   if p["status"] in ("ok","rezerv","multi")]
        if not pozicii:
            messagebox.showwarning("Пусто", "Нет позиций."); return
        ZAKAZY_DIR.mkdir(parents=True, exist_ok=True)
        prefix = f"[{nomer}]_" if nomer else ""
        safe   = re.sub(r'[<>:"/\\\\|?*]', "_", name or "zakaz")
        stamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
        path   = ZAKAZY_DIR / f"{prefix}{safe}__{stamp}.txt"
        try:
            with open(path, "w", encoding="utf-8-sig") as f:
                f.write(f"# [{nomer}] {name}\n" if nomer else f"# {name}\n")
                f.write("\n".join(pozicii))
            self._log(f"💾 {path.name}")
            messagebox.showinfo("Сохранено", path.name)
        except OSError as e:
            messagebox.showerror("Ошибка", str(e))

    def _otkryt_sohranenny(self):
        ZAKAZY_DIR.mkdir(parents=True, exist_ok=True)
        files = sorted(ZAKAZY_DIR.glob("*.txt"),
                       key=lambda p: p.stat().st_mtime, reverse=True)
        if not files:
            messagebox.showinfo("Нет заказов", str(ZAKAZY_DIR)); return
        dlg = tk.Toplevel(self); dlg.title("Сохранённые заказы")
        dlg.geometry("640x400"); dlg.grab_set()
        ttk.Label(dlg, text="Enter или двойной клик — открыть:").pack(anchor="w", padx=8, pady=(6,2))
        tree = ttk.Treeview(dlg, columns=("nom","name","dt","n"), show="headings")
        tree.heading("nom",  text="№");    tree.column("nom",  width=55, anchor="center")
        tree.heading("name", text="Название"); tree.column("name", width=310)
        tree.heading("dt",   text="Дата"); tree.column("dt",   width=130, anchor="center")
        tree.heading("n",    text="Строк"); tree.column("n",   width=55,  anchor="center")
        vsb = ttk.Scrollbar(dlg, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8,0), pady=(0,8))
        vsb.pack(side="right", fill="y", pady=(0,8), padx=(0,8))
        for fp in files:
            txt = prochitat_tekst(str(fp))
            lines = [l for l in txt.splitlines() if l.strip() and not l.startswith("#")]
            m = re.match(r"\[(\d+)\]_(.*?)__", fp.stem)
            zn = m.group(1) if m else ""
            zname = (m.group(2) if m else fp.stem.rsplit("__",1)[0]).replace("_"," ")
            mtime = datetime.fromtimestamp(fp.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
            tree.insert("", "end", iid=str(fp), values=(zn, zname, mtime, len(lines)))
        if tree.get_children():
            tree.focus(tree.get_children()[0]); tree.selection_set(tree.get_children()[0])
        tree.focus_set()
        def open_sel(event=None):
            iid = tree.focus()
            if not iid: return
            fp = Path(iid)
            m = re.match(r"\[(\d+)\]_(.*?)__", fp.stem)
            self.zakaz_nomer_var.set(m.group(1) if m else "")
            self.zakaz_name_var.set((m.group(2) if m else fp.stem.rsplit("__",1)[0]).replace("_"," "))
            dlg.destroy()
            self._zagruzit_zadanie(str(fp), "txt")
        tree.bind("<Double-1>", open_sel)
        tree.bind("<Return>", open_sel)
        ttk.Button(dlg, text="Открыть", command=open_sel).pack(pady=4)


# ── ТОЧКА ВХОДА ──────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
