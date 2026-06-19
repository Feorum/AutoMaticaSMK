# -*- coding: utf-8 -*-
"""
zakaz_parser.py  (v1.0)

Утилита разбора Excel-заказов и формирования файлов-заданий для auto_vvod.

Поддерживаемые типы документов (автоопределение):
  - Тип A "Евроопт":      столбцы Наименование + Требуемое количество
  - Тип B "Заказ-накладная": столбцы Наименование товара + Количество
  - Тип C "Отвес":        столбцы Наименование продукции + План заявки
  - Универсальный:        если тип не определён — пользователь выбирает столбцы вручную

Результат: файл задания в папке zadaniya/ (или выбранной папке).
Формат строки: Наименование;количество
"""

import os
import sys
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    messagebox.showerror("Ошибка", "Не установлен pandas.\nВыполните: pip install pandas openpyxl")
    sys.exit(1)

# ──────────────────────────────────────────────────────────────────
#  НАСТРОЙКИ
# ──────────────────────────────────────────────────────────────────

# Базовая папка: рядом с exe или рядом со скриптом
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

DEFAULT_ZADANIYA_DIR = BASE_DIR / "zadaniya"

# Ключевые слова для автоопределения столбцов (нижний регистр)
NAME_KEYWORDS    = ["наименование", "название", "продукция", "товар", "номенклатура"]
QTY_KEYWORDS     = ["требуемое", "план", "заявк", "количество", "кол-во", "кол во", "кол_во"]
SKIP_ROWS_MAX    = 20   # сколько строк сверху просматривать в поиске заголовков
MIN_QTY          = 0    # строки с количеством <= MIN_QTY пропускаем

# Строки-заглушки, которые не являются товарами (нижний регистр, вхождение)
SKIP_NAME_PATTERNS = ["итого", "всего", "total", "nan", "наименование", "продукция", "товар"]


# ──────────────────────────────────────────────────────────────────
#  ПАРСЕР EXCEL
# ──────────────────────────────────────────────────────────────────

def normalize_col(s: str) -> str:
    return str(s).lower().strip()


def find_header_row(df_raw: pd.DataFrame) -> int:
    """Найти строку заголовков: ту, где больше всего ключевых слов столбцов."""
    best_row, best_score = 0, 0
    all_kw = NAME_KEYWORDS + QTY_KEYWORDS
    for i in range(min(SKIP_ROWS_MAX, len(df_raw))):
        row_vals = [normalize_col(v) for v in df_raw.iloc[i].values]
        score = sum(1 for kw in all_kw if any(kw in rv for rv in row_vals))
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def find_column(columns: list[str], keywords: list[str]) -> str | None:
    """Найти имя столбца по ключевым словам."""
    for kw in keywords:
        for col in columns:
            if kw in normalize_col(col):
                return col
    return None


def detect_doc_type(columns: list[str]) -> str:
    """Определить тип документа по заголовкам."""
    cols_low = [normalize_col(c) for c in columns]
    joined = " ".join(cols_low)
    if "требуемое" in joined:
        return "evroopt"
    if "план" in joined and ("заявк" in joined or "наименование продукции" in joined):
        return "otves"
    if "наименование товара" in joined or "штрихкод" in joined:
        return "nakl"
    return "unknown"


def load_excel_smart(path: str) -> tuple[pd.DataFrame | None, str, str | None, str | None, str]:
    """
    Загрузить Excel, найти заголовки, вернуть:
      (df, doc_type, name_col, qty_col, sheet_name)
    """
    xl = pd.ExcelFile(path)
    # Берём первый лист (обычно там данные)
    sheet = xl.sheet_names[0]

    df_raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
    header_row = find_header_row(df_raw)

    df = pd.read_excel(path, sheet_name=sheet, header=header_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    doc_type = detect_doc_type(list(df.columns))
    name_col = find_column(list(df.columns), NAME_KEYWORDS)
    qty_col  = find_column(list(df.columns), QTY_KEYWORDS)

    return df, doc_type, name_col, qty_col, sheet


def parse_qty(val) -> float:
    """Извлечь число из ячейки количества."""
    if val is None or str(val).strip().lower() in ("nan", "", "-"):
        return 0.0
    s = re.sub(r"[^\d.,]", "", str(val)).replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def is_skip_name(name: str) -> bool:
    n = name.lower().strip()
    return not n or any(p in n for p in SKIP_NAME_PATTERNS)


def extract_rows(df: pd.DataFrame, name_col: str, qty_col: str) -> list[tuple[str, int]]:
    """Извлечь строки (название, количество) из датафрейма."""
    result = []
    for _, row in df.iterrows():
        name = str(row.get(name_col, "")).strip()
        qty  = parse_qty(row.get(qty_col))
        if is_skip_name(name):
            continue
        if qty <= MIN_QTY:
            continue
        result.append((name, int(qty) if qty == int(qty) else qty))
    return result


# ──────────────────────────────────────────────────────────────────
#  GUI
# ──────────────────────────────────────────────────────────────────

DOC_TYPE_LABELS = {
    "evroopt": "Евроопт",
    "otves":   "Отвес накладная",
    "nakl":    "Заказ-накладная",
    "unknown": "Неизвестный тип",
}

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Парсер заказов → zadaniya  v1.0")
        self.geometry("980x680")
        self.resizable(True, True)

        self.df: pd.DataFrame | None = None
        self.all_columns: list[str] = []
        self.rows: list[tuple[str, int]] = []
        self.zadaniya_dir = tk.StringVar(value=str(DEFAULT_ZADANIYA_DIR))
        self.name_col_var  = tk.StringVar()
        self.qty_col_var   = tk.StringVar()
        self.doc_type_var  = tk.StringVar(value="—")
        self.sheet_var     = tk.StringVar(value="—")
        self.file_path_var = tk.StringVar(value="")
        self.filter_var    = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self._apply_filter())

        self._build_ui()

    # ── BUILD UI ──────────────────────────────────────────────────

    def _build_ui(self):
        # ── Верхняя панель: файл ──
        top = ttk.LabelFrame(self, text="Файл заказа", padding=6)
        top.pack(fill="x", padx=8, pady=(8, 2))

        ttk.Button(top, text="Открыть Excel...", command=self._open_file).grid(row=0, column=0, padx=(0,6))
        ttk.Entry(top, textvariable=self.file_path_var, state="readonly", width=55).grid(row=0, column=1, sticky="ew")
        ttk.Label(top, text="Тип:").grid(row=0, column=2, padx=(10,2))
        ttk.Label(top, textvariable=self.doc_type_var, foreground="#0055aa", width=18).grid(row=0, column=3)
        ttk.Label(top, text="Лист:").grid(row=0, column=4, padx=(6,2))
        ttk.Label(top, textvariable=self.sheet_var, foreground="#555", width=14).grid(row=0, column=5)
        top.columnconfigure(1, weight=1)

        # ── Столбцы ──
        cols_frame = ttk.LabelFrame(self, text="Столбцы (автоопределение; можно поменять)", padding=6)
        cols_frame.pack(fill="x", padx=8, pady=2)

        ttk.Label(cols_frame, text="Наименование:").grid(row=0, column=0, sticky="w")
        self.cb_name = ttk.Combobox(cols_frame, textvariable=self.name_col_var, width=40, state="readonly")
        self.cb_name.grid(row=0, column=1, padx=(4,16), sticky="ew")
        self.cb_name.bind("<<ComboboxSelected>>", lambda _: self._reload_rows())

        ttk.Label(cols_frame, text="Количество:").grid(row=0, column=2, sticky="w")
        self.cb_qty = ttk.Combobox(cols_frame, textvariable=self.qty_col_var, width=28, state="readonly")
        self.cb_qty.grid(row=0, column=3, padx=(4,0), sticky="ew")
        self.cb_qty.bind("<<ComboboxSelected>>", lambda _: self._reload_rows())

        ttk.Button(cols_frame, text="↺ Применить", command=self._reload_rows).grid(row=0, column=4, padx=(10,0))
        cols_frame.columnconfigure(1, weight=1)

        # ── Фильтр ──
        flt_frame = ttk.Frame(self)
        flt_frame.pack(fill="x", padx=8, pady=2)
        ttk.Label(flt_frame, text="🔍 Фильтр строк:").pack(side="left")
        ttk.Entry(flt_frame, textvariable=self.filter_var, width=35).pack(side="left", padx=6)
        ttk.Label(flt_frame, text="(поиск по названию)").pack(side="left")
        self.lbl_count = ttk.Label(flt_frame, text="Строк: 0", foreground="#555")
        self.lbl_count.pack(side="right", padx=8)

        # ── Таблица ──
        tbl_frame = ttk.LabelFrame(self, text="Строки заказа  (галочкой — включить в задание)", padding=4)
        tbl_frame.pack(fill="both", expand=True, padx=8, pady=2)

        cols = ("check", "num", "name", "qty")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("check", text="✓")
        self.tree.heading("num",   text="№")
        self.tree.heading("name",  text="Наименование")
        self.tree.heading("qty",   text="Кол-во")
        self.tree.column("check", width=32,  stretch=False, anchor="center")
        self.tree.column("num",   width=42,  stretch=False, anchor="center")
        self.tree.column("name",  width=580, stretch=True)
        self.tree.column("qty",   width=80,  stretch=False, anchor="center")

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self._on_click)

        # ── Кнопки выделения ──
        sel_frame = ttk.Frame(self)
        sel_frame.pack(fill="x", padx=8, pady=(0, 2))
        ttk.Button(sel_frame, text="✓ Все",          command=self._check_all).pack(side="left", padx=2)
        ttk.Button(sel_frame, text="✗ Снять все",    command=self._uncheck_all).pack(side="left", padx=2)
        ttk.Button(sel_frame, text="↔ Инвертировать",command=self._invert_check).pack(side="left", padx=2)
        ttk.Button(sel_frame, text="✓ Только видимые (фильтр)", command=self._check_visible).pack(side="left", padx=8)

        # ── Нижняя панель: выход ──
        bot = ttk.LabelFrame(self, text="Сохранить задание", padding=6)
        bot.pack(fill="x", padx=8, pady=(2, 8))

        ttk.Label(bot, text="Папка zadaniya:").grid(row=0, column=0, sticky="w")
        ttk.Entry(bot, textvariable=self.zadaniya_dir, width=50).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(bot, text="...", width=3, command=self._pick_dir).grid(row=0, column=2, padx=2)

        ttk.Label(bot, text="Имя файла:").grid(row=1, column=0, sticky="w", pady=(4,0))
        self.fname_var = tk.StringVar(value="zadanie.txt")
        ttk.Entry(bot, textvariable=self.fname_var, width=34).grid(row=1, column=1, sticky="w", padx=4, pady=(4,0))

        ttk.Button(bot, text="💾  Сохранить задание →  zadaniya/",
                   command=self._save_zadanie,
                   style="Accent.TButton").grid(row=0, column=3, rowspan=2, padx=(12,0), sticky="ns")

        self.lbl_status = ttk.Label(bot, text="", foreground="#006600")
        self.lbl_status.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4,0))

        bot.columnconfigure(1, weight=1)

        # Стиль кнопки сохранения
        style = ttk.Style(self)
        style.configure("Accent.TButton", font=("Arial", 10, "bold"))

    # ── ЛОГИКА ───────────────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Выберите файл заказа",
            filetypes=[("Excel файлы", "*.xlsx *.xls *.xlsm"), ("Все файлы", "*.*")]
        )
        if not path:
            return
        self.file_path_var.set(path)
        self._load_file(path)

    def _load_file(self, path: str):
        try:
            df, doc_type, name_col, qty_col, sheet = load_excel_smart(path)
        except Exception as e:
            messagebox.showerror("Ошибка чтения файла", str(e))
            return

        self.df           = df
        self.all_columns  = list(df.columns)
        self.doc_type_var.set(DOC_TYPE_LABELS.get(doc_type, doc_type))
        self.sheet_var.set(sheet)

        self.cb_name["values"] = self.all_columns
        self.cb_qty["values"]  = self.all_columns

        self.name_col_var.set(name_col or (self.all_columns[0] if self.all_columns else ""))
        self.qty_col_var.set(qty_col  or (self.all_columns[1] if len(self.all_columns) > 1 else ""))

        # Предлагаем имя файла задания по имени исходника
        src_stem = Path(path).stem
        stamp    = datetime.now().strftime("%Y%m%d_%H%M")
        self.fname_var.set(f"{src_stem}_{stamp}.txt")

        self._reload_rows()
        self._status(f"Загружено: {Path(path).name}  |  тип: {DOC_TYPE_LABELS.get(doc_type, doc_type)}"
                     f"  |  столбцов: {len(self.all_columns)}")

    def _reload_rows(self):
        name_col = self.name_col_var.get()
        qty_col  = self.qty_col_var.get()
        if self.df is None or not name_col or not qty_col:
            return
        try:
            self.rows = extract_rows(self.df, name_col, qty_col)
        except Exception as e:
            messagebox.showerror("Ошибка разбора", str(e))
            return
        self._fill_tree(self.rows, check_all=True)

    def _fill_tree(self, rows: list[tuple], check_all: bool = True):
        # Запомним состояние галочек по имени (если уже было что-то)
        checked = {self.tree.set(iid, "name"): self.tree.set(iid, "check") == "☑"
                   for iid in self.tree.get_children()}

        self.tree.delete(*self.tree.get_children())
        for i, (name, qty) in enumerate(rows, 1):
            was_checked = checked.get(name, check_all)
            mark = "☑" if was_checked else "☐"
            self.tree.insert("", "end", iid=str(i),
                             values=(mark, i, name, qty))
        self._update_count()

    def _apply_filter(self):
        kw = self.filter_var.get().lower().strip()
        for iid in self.tree.get_children():
            name = self.tree.set(iid, "name").lower()
            if kw and kw not in name:
                self.tree.detach(iid)
            else:
                # Reattach если была скрыта
                try:
                    self.tree.reattach(iid, "", "end")
                except Exception:
                    pass
        self._update_count()

    def _update_count(self):
        visible  = len(self.tree.get_children())
        checked  = sum(1 for iid in self.tree.get_children()
                       if self.tree.set(iid, "check") == "☑")
        self.lbl_count.config(text=f"Видимых: {visible}  |  Отмечено: {checked}")

    def _on_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        col    = self.tree.identify_column(event.x)
        iid    = self.tree.identify_row(event.y)
        if not iid:
            return
        if region == "cell" and col == "#1":
            # Переключить галочку
            cur = self.tree.set(iid, "check")
            self.tree.set(iid, "check", "☑" if cur == "☐" else "☐")
            self._update_count()

    def _check_all(self):
        for iid in self.tree.get_children():
            self.tree.set(iid, "check", "☑")
        self._update_count()

    def _uncheck_all(self):
        for iid in self.tree.get_children():
            self.tree.set(iid, "check", "☐")
        self._update_count()

    def _invert_check(self):
        for iid in self.tree.get_children():
            cur = self.tree.set(iid, "check")
            self.tree.set(iid, "check", "☑" if cur == "☐" else "☐")
        self._update_count()

    def _check_visible(self):
        """Отметить только видимые (после фильтра) строки."""
        all_iids     = set(str(i+1) for i in range(len(self.rows)))
        visible_iids = set(self.tree.get_children())
        hidden_iids  = all_iids - visible_iids
        for iid in visible_iids:
            self.tree.set(iid, "check", "☑")
        for iid in hidden_iids:
            try:
                self.tree.set(iid, "check", "☐")
            except Exception:
                pass
        self._update_count()

    def _pick_dir(self):
        d = filedialog.askdirectory(title="Папка zadaniya")
        if d:
            self.zadaniya_dir.set(d)

    def _save_zadanie(self):
        # Собираем все строки с галочкой (включая скрытые фильтром)
        lines = []
        for iid in self.tree.get_children(""):
            if self.tree.set(iid, "check") == "☑":
                name = self.tree.set(iid, "name").strip()
                qty  = self.tree.set(iid, "qty").strip()
                if name and qty:
                    lines.append(f"{name};{qty}")

        # Также добавим скрытые фильтром, если они отмечены
        # (они detach'ены, но данные сохранены в self.rows + tree iid)
        visible_iids = set(self.tree.get_children())
        for i, (name, qty) in enumerate(self.rows, 1):
            iid = str(i)
            if iid not in visible_iids:
                try:
                    if self.tree.set(iid, "check") == "☑":
                        lines.append(f"{name};{qty}")
                except Exception:
                    pass

        if not lines:
            messagebox.showwarning("Нет данных", "Не отмечено ни одной строки.")
            return

        out_dir = Path(self.zadaniya_dir.get())
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать папку:\n{e}")
            return

        fname = self.fname_var.get().strip() or "zadanie.txt"
        if not fname.endswith(".txt"):
            fname += ".txt"
        out_path = out_dir / fname

        try:
            with open(out_path, "w", encoding="utf-8-sig") as f:
                f.write("\n".join(lines))
        except Exception as e:
            messagebox.showerror("Ошибка записи", str(e))
            return

        self._status(f"✓ Сохранено: {out_path}  ({len(lines)} позиций)")
        messagebox.showinfo("Готово",
                            f"Файл задания сохранён:\n{out_path}\n\nПозиций: {len(lines)}")

    def _status(self, msg: str):
        self.lbl_status.config(text=msg)


# ──────────────────────────────────────────────────────────────────
#  ТОЧКА ВХОДА
# ──────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
