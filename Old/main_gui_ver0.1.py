# -*- coding: utf-8 -*-
"""
main_gui_ver0.1.py  —  единый GUI для авто-ввода продукции  (v1.0)

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

# ── Настройки ввода ───────────────────────────────────────────────
REZHIM_VVODA     = "vboxmanage"
IMYA_VM          = "Xp"
VBOXMANAGE_PATH  = r"C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
SBROS_PAGEUP     = 20
RAZMER_STRANICY  = 17
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
        tovary.append({
            "naimenovanie": name,
            "oformlenie":   (row.get("oformlenie") or "").strip(),
            "massa":        (row.get("massa") or "").strip(),
            "ostatok":      ost,
            "rezerv":       rezerv_raw in ("1", "yes", "да", "true"),
        })
    return tovary


def sohranit_bazu(tovary: list[dict]) -> None:
    BAZA_DIR.mkdir(parents=True, exist_ok=True)
    with open(TOVARY_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["pozitsiya", "naimenovanie", "oformlenie", "massa", "ostatok", "rezerv"])
        for i, t in enumerate(tovary, 1):
            w.writerow([i, t["naimenovanie"], t["oformlenie"], t["massa"],
                        t["ostatok"], "1" if t["rezerv"] else "0"])


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


def _fuzzy_score(query: str, candidate: str) -> float:
    """
    Оценка нечёткого сходства query и candidate (0.0 … 1.0).
    Комбинирует:
      - долю совпавших токенов (слов)
      - долю совпавших биграмм символов
    Используется для подсказок при вводе в словарь.
    """
    qn = _norm(query)
    cn = _norm(candidate)
    if not qn:
        return 0.0
    if qn == cn:
        return 1.0

    # Токены
    q_tok = set(qn.split())
    c_tok = set(cn.split())
    if q_tok and c_tok:
        tok_score = len(q_tok & c_tok) / max(len(q_tok), len(c_tok))
    else:
        tok_score = 0.0

    # Биграммы символов
    def bigrams(s):
        return [s[i:i+2] for i in range(len(s)-1)]
    qb = bigrams(qn)
    cb = bigrams(cn)
    if qb and cb:
        qb_set = set(qb)
        cb_set = set(cb)
        bi_score = len(qb_set & cb_set) / max(len(qb_set), len(cb_set))
    else:
        bi_score = 0.0

    # Бонус: query целиком входит в candidate или наоборот
    sub_bonus = 0.2 if (qn in cn or cn in qn) else 0.0

    return min(1.0, 0.45 * tok_score + 0.35 * bi_score + sub_bonus)


def fuzzy_podbor(query: str, tovary: list[dict], top_n: int = 8) -> list[tuple[float, int, dict]]:
    """
    Вернуть top_n товаров из базы, наиболее похожих на query.
    Возвращает список (score, idx, tovar), отсортированный по убыванию score.
    Порог: score >= 0.10.
    """
    if not query.strip():
        return []
    results = []
    for idx, t in enumerate(tovary):
        sc = _fuzzy_score(query, t["naimenovanie"])
        if sc >= 0.10:
            results.append((sc, idx, t))
    results.sort(key=lambda x: -x[0])
    return results[:top_n]


def naiti_cherez_slovar(
    zapros: str,
    tovary: list[dict],
    slovar: list[tuple[str, str]],
) -> tuple[str | None, list[tuple[int, dict]]]:
    zn = _norm(zapros)
    baza_podstroka = None
    sovp_fraza = None
    for fraza, baza in slovar:
        if _norm(fraza) in zn:
            sovp_fraza = fraza
            baza_podstroka = baza
            break
    if baza_podstroka is None:
        return None, []
    bn = _norm(baza_podstroka)
    naideny = [(i, t) for i, t in enumerate(tovary) if bn in _norm(t["naimenovanie"])]
    return sovp_fraza, naideny


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
    delta = cel - tek
    if delta == 0:
        return
    pg  = "pagedown" if delta > 0 else "pageup"
    arr = "down"     if delta > 0 else "up"
    for _ in range(abs(delta) // RAZMER_STRANICY):
        nazhat_klavishu(pg)
    for _ in range(abs(delta) % RAZMER_STRANICY):
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

        self._build_ui()
        self._reload_baza()

    # ── BUILD UI ──────────────────────────────────────────────────

    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb

        self.tab_zakaz   = ttk.Frame(nb)
        self.tab_slovar  = ttk.Frame(nb)
        self.tab_zhurnal = ttk.Frame(nb)

        nb.add(self.tab_zakaz,   text="  📋 Заказ  ")
        nb.add(self.tab_slovar,  text="  📖 Словарь  ")
        nb.add(self.tab_zhurnal, text="  📜 Журнал  ")

        self._build_tab_zakaz()
        self._build_tab_slovar()
        self._build_tab_zhurnal()

    # ── ВКЛАДКА ЗАКАЗ ─────────────────────────────────────────────

    def _build_tab_zakaz(self):
        f = self.tab_zakaz

        # Панель загрузки файла
        top = ttk.LabelFrame(f, text="Файл задания", padding=6)
        top.pack(fill="x", padx=6, pady=(6, 2))

        ttk.Button(top, text="📂 Открыть Excel…", command=self._otkryt_excel).grid(row=0, column=0, padx=(0,4))
        ttk.Button(top, text="📂 Открыть TXT…",   command=self._otkryt_txt).grid(row=0, column=1, padx=(0,4))
        ttk.Button(top, text="📂 Из папки zadaniya…", command=self._otkryt_iz_papki).grid(row=0, column=2, padx=(0,8))

        self.lbl_file = ttk.Label(top, text="Файл не выбран", foreground="#777", width=55)
        self.lbl_file.grid(row=0, column=3, sticky="ew")
        top.columnconfigure(3, weight=1)

        # Статус базы и VM
        info = ttk.Frame(f)
        info.pack(fill="x", padx=6, pady=2)
        self.lbl_baza_stat = ttk.Label(info, text="База: не загружена", foreground=CLR_ERR)
        self.lbl_baza_stat.pack(side="left", padx=(0, 20))
        self.lbl_vm_stat = ttk.Label(info, text="VM: неизвестно", foreground=CLR_WARN)
        self.lbl_vm_stat.pack(side="left")
        ttk.Button(info, text="↺ Проверить VM", command=self._check_vm).pack(side="left", padx=8)
        ttk.Button(info, text="↺ Перезагрузить базу", command=self._reload_baza).pack(side="left")

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
        self.tree_plan.tag_configure("ok",    background=CLR_STRIPE1)
        self.tree_plan.tag_configure("ok2",   background=CLR_STRIPE2)
        self.tree_plan.tag_configure("multi", background=CLR_MULTI)
        self.tree_plan.tag_configure("miss",  background=CLR_MISS)

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
            fraza, naideny = naiti_cherez_slovar(name, self.tovary, self.slovar)

            if fraza is None:
                # Не в словаре
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "miss", "idx": None, "t": None, "varianty": [],
                })
            elif not naideny:
                # В словаре, но в базе нет
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "miss", "idx": None, "t": None, "varianty": [],
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
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "ok", "idx": idx, "t": t, "varianty": [],
                })
            else:
                # Несколько вариантов — нужен выбор
                self.plan_rows.append({
                    "name": name, "oform": oform, "massa": massa, "kol": kol,
                    "status": "multi", "idx": None, "t": None, "varianty": naideny,
                })

        if SORTIROVAT:
            ok_rows   = [p for p in self.plan_rows if p["status"] == "ok"]
            other_rows = [p for p in self.plan_rows if p["status"] != "ok"]
            ok_rows.sort(key=lambda p: p["idx"])
            self.plan_rows = ok_rows + other_rows

        self._refresh_plan_tree()
        self._update_plan_stat()
        self._update_start_btn()

    def _refresh_plan_tree(self):
        self.tree_plan.delete(*self.tree_plan.get_children())
        ok_n = 0
        for i, p in enumerate(self.plan_rows, 1):
            st = p["status"]
            if st == "ok":
                ok_n += 1
                tag = "ok" if ok_n % 2 == 1 else "ok2"
                baza_label = format_tovar(p["t"])
                stroka     = str(p["idx"] + 1)
                ost        = str(p["t"]["ostatok"])
                stat_label = "✓ найдено"
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
        total = len(self.plan_rows)
        self.lbl_plan_stat.config(
            text=f"Всего: {total}  |  ✓ {ok}  |  ⚠ выбрать: {multi}  |  ✗ не найдено: {miss}",
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

        def apply():
            sel = tree2.focus()
            if not sel:
                messagebox.showwarning("Выберите товар", "Кликните на строку в таблице.", parent=dlg)
                return
            ci   = int(sel)
            ct   = self.tovary[ci]
            baza = ct["naimenovanie"]
            frz  = fraza_var.get().strip()
            if not frz:
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
            self._refresh_plan_tree()
            self._update_plan_stat()
            self._update_start_btn()
            self._log(f"Добавлено в словарь: {frz!r} → {baza!r}")
            dlg.destroy()

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
        ok_rows = [p for p in self.plan_rows if p["status"] == "ok"]
        if not ok_rows:
            messagebox.showwarning("Нет данных", "Нет готовых строк для ввода."); return

        multi = sum(1 for p in self.plan_rows if p["status"] == "multi")
        miss  = sum(1 for p in self.plan_rows if p["status"] == "miss")
        if multi > 0 or miss > 0:
            msg = []
            if multi: msg.append(f"⚠ {multi} строк требуют выбора (двойной клик)")
            if miss:  msg.append(f"✗ {miss} строк не найдено в базе")
            if not messagebox.askyesno(
                "Неполный план",
                "\n".join(msg) + f"\n\nВсё равно запустить ввод {len(ok_rows)} найденных строк?"):
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

        # Сброс курсора наверх
        self._log(f"Сброс курсора: {SBROS_PAGEUP}× PageUp...")
        for _ in range(SBROS_PAGEUP):
            if self._stop_vvod.is_set(): return
            nazhat_klavishu("pageup")
        self._log("Курсор на строке 1. Начинаю ввод...")

        tek_poz  = 0
        istoriya = []
        teper    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        vvedeno  = 0

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


# ── ТОЧКА ВХОДА ──────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
