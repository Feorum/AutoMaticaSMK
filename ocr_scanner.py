# -*- coding: utf-8 -*-
"""
ocr_scanner.py  —  Модуль машинного зрения для синхронизации ассортимента
==========================================================================

Принцип работы:
  1. VBoxManage screenshotpng  →  PNG файл экрана DOS-VM
  2. pytesseract OCR           →  текст экрана (кириллица CP866/DOS-шрифт)
  3. Парсинг строк             →  список наименований
  4. Прокрутка PageDown        →  следующая страница (17 строк)
  5. Повтор до конца списка    →  полный ассортимент собран
  6. Diff с tovary.csv         →  новые / удалённые позиции
  7. Обновление tovary.csv     →  добавить новые, active=0 для удалённых

Зависимости:
  pip install pytesseract pillow
  + Tesseract OCR должен быть установлен:
    https://github.com/UB-Mannheim/tesseract/wiki
    При установке выбрать: Russian language pack

Настройки (см. константы ниже или передать через класс OCRScanner):
  VM_NAME         — имя виртуальной машины VirtualBox
  VBOXMANAGE_PATH — путь к VBoxManage.exe
  PAGE_SIZE       — строк в одной странице DOS-программы (17)
  OCR_REGION      — область экрана для OCR (x, y, w, h) в пикселях
                    None = весь экран
  NAME_COL_X1/X2 — горизонтальные границы столбца наименований (пиксели)
                    None = вся ширина

Интеграция с main_gui_v3.py:
  Вызывать из кнопки "🔍 Сканировать ассортимент" в TabTovary.
  Передать callback'и для прогресса и результата.
"""

from __future__ import annotations

import os
import re
import csv
import time
import subprocess
import tempfile
import threading
from pathlib import Path
from typing import Callable, Optional

# ── Попытка импорта pytesseract / PIL ────────────────────────────────────────
try:
    import pytesseract
    from PIL import Image, ImageFilter, ImageEnhance
    TESSERACT_OK = True
except ImportError:
    TESSERACT_OK = False


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  КОНСТАНТЫ (можно переопределить через OCRScanner.__init__)              ║
# ╚══════════════════════════════════════════════════════════════════════════╝

DEFAULT_VBOXMANAGE = r"C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
DEFAULT_VM_NAME    = "Xp"
DEFAULT_PAGE_SIZE  = 17          # строк на страницу (PageDown в DOS)
DEFAULT_MAX_PAGES  = 30          # защита от бесконечного цикла (30×17=510 строк макс)
DEFAULT_PAUSE_KEY  = 0.3         # пауза после нажатия клавиши (сек)
DEFAULT_PAUSE_SCR  = 0.6         # пауза перед скриншотом (дать DOS отрисоваться)

# Путь к Tesseract (если не прописан в PATH)
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# OCR: регион (x, y, width, height) — вырезаем только нужную область.
# None = весь экран. Подберите под вашу DOS-программу после первого скриншота.
DEFAULT_REGION: Optional[tuple[int, int, int, int]] = None

# Столбец наименований: левая и правая граница по X (пиксели).
# None = не обрезать по горизонтали.
DEFAULT_NAME_X1: Optional[int] = None
DEFAULT_NAME_X2: Optional[int] = None

# Минимальная длина строки после OCR, чтобы считать её товаром (фильтр мусора)
MIN_NAME_LEN = 4

# Маркер конца списка — строка OCR, при появлении которой прокрутка
# прекращается. Подберите под вашу DOS-программу.
# Примеры: "ИТОГО", "---", пустая строка после N повторов.
END_MARKERS: list[str] = []    # [] = автоопределение по повтору страниц


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  НОРМАЛИЗАЦИЯ СТРОК                                                      ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def _norm_ocr(s: str) -> str:
    """Нормализация строки OCR для сравнения: убрать лишние символы/пробелы."""
    # DOS-программы часто дают псевдографику и артефакты
    s = re.sub(r"[│├┤┼─┬┴┌┐└┘║╠╣╬═╔╗╚╝╞╡╟╢╤╧╦╩╪╫▓▒░■□▪▫•◦]", " ", s)
    # Убрать незначащие символы, оставить буквы, цифры, /, -, пробелы, точки
    s = re.sub(r"[^\w\s\-/.,`'\"«»()%#№]", " ", s)
    # Схлопнуть пробелы
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _norm_key(s: str) -> str:
    """Ключ для сравнения: нижний регистр, только буквы/цифры."""
    return re.sub(r"[^а-яёa-z0-9]", "", s.lower())


def _names_match(ocr_name: str, db_name: str, threshold: float = 0.75) -> bool:
    """
    Нечёткое сравнение: считаем строки одинаковыми, если >threshold символов совпадают.
    Используем простую метрику пересечения биграмм.
    """
    def bigrams(t: str) -> set:
        t = _norm_key(t)
        return {t[i:i+2] for i in range(len(t)-1)} if len(t) >= 2 else set()

    a, b = bigrams(ocr_name), bigrams(db_name)
    if not a and not b:
        return True
    if not a or not b:
        return False
    score = len(a & b) / max(len(a), len(b))
    return score >= threshold


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  ПАРСИНГ OCR-ТЕКСТА → СПИСОК НАИМЕНОВАНИЙ                               ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def parse_ocr_text(
    raw_text: str,
    name_x1: Optional[int] = None,
    name_x2: Optional[int] = None,
) -> list[str]:
    """
    Из сырого текста OCR извлечь строки-наименования.

    DOS-программа обычно выводит таблицу вида:
      │ 1 │ Чесночная м/р Груп м/атм  │ ... │
    Нас интересует только колонка наименования.

    Стратегия:
      - Разбить на строки
      - Убрать строки с псевдографикой/шапками таблицы
      - Взять первый «смысловой» фрагмент каждой строки
      - Фильтр по MIN_NAME_LEN
    """
    names = []
    for line in raw_text.splitlines():
        line = _norm_ocr(line)
        if not line:
            continue

        # Пропустить строки-разделители таблицы (много тире, знаков рамки)
        if len(re.sub(r"[\-=\s]", "", line)) < 3:
            continue

        # Пропустить строки с ключевыми словами шапки
        low = line.lower()
        if any(kw in low for kw in ("наименование", "итого", "всего", "остаток",
                                     "количество", "цена", "сумма", "накладная",
                                     "ввод", "меню", "выход", "помощь", "f1",
                                     "esc", "enter", "стрел")):
            continue

        # Убрать ведущие номера строк: "123 " или "123. " или "| 123 |"
        line = re.sub(r"^\|?\s*\d{1,4}[\s.|)│]+", "", line).strip()
        if not line:
            continue

        # Если строка содержит разделители столбцов (│ или  2+ пробела),
        # берём только первый столбец (до первого разделителя)
        parts = re.split(r"│|\s{3,}", line)
        name = parts[0].strip() if parts else line.strip()

        # Убрать хвостовые цифры (остатки, цены)
        name = re.sub(r"\s+[\d\-+]+\s*$", "", name).strip()

        if len(name) >= MIN_NAME_LEN:
            names.append(name)

    return names


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  ОСНОВНОЙ КЛАСС                                                          ║
# ╚══════════════════════════════════════════════════════════════════════════╝

class OCRScanner:
    """
    Сканирует ассортимент DOS-программы через VBoxManage screenshot + Tesseract OCR.

    Параметры:
        vm_name        — имя VM в VirtualBox
        vboxmanage     — путь к VBoxManage.exe
        tovary_csv     — Path к файлу tovary.csv
        page_size      — строк на страницу DOS-программы
        region         — (x, y, w, h) область скриншота или None
        name_x1/x2     — горизонтальные границы столбца наименований
        on_progress    — callback(msg: str, step: int, total: int)
        on_done        — callback(result: ScanResult)
        on_error       — callback(err: str)
        pause_key      — пауза после нажатия клавиши (сек)
        pause_scr      — пауза перед скриншотом (сек)
        max_pages      — максимум страниц (защита от зависания)
        end_markers    — строки-маркеры конца списка
        fuzzy_match    — использовать нечёткое сравнение при поиске в базе
        match_threshold — порог нечёткого сравнения (0.0–1.0)
    """

    def __init__(
        self,
        vm_name: str = DEFAULT_VM_NAME,
        vboxmanage: str = DEFAULT_VBOXMANAGE,
        tovary_csv: Optional[Path] = None,
        page_size: int = DEFAULT_PAGE_SIZE,
        region: Optional[tuple] = DEFAULT_REGION,
        name_x1: Optional[int] = DEFAULT_NAME_X1,
        name_x2: Optional[int] = DEFAULT_NAME_X2,
        on_progress: Optional[Callable] = None,
        on_done: Optional[Callable] = None,
        on_error: Optional[Callable] = None,
        pause_key: float = DEFAULT_PAUSE_KEY,
        pause_scr: float = DEFAULT_PAUSE_SCR,
        max_pages: int = DEFAULT_MAX_PAGES,
        end_markers: Optional[list[str]] = None,
        fuzzy_match: bool = True,
        match_threshold: float = 0.75,
    ):
        self.vm_name         = vm_name
        self.vboxmanage      = vboxmanage
        self.tovary_csv      = tovary_csv
        self.page_size       = page_size
        self.region          = region
        self.name_x1         = name_x1
        self.name_x2         = name_x2
        self.on_progress     = on_progress or (lambda msg, step, total: None)
        self.on_done         = on_done or (lambda r: None)
        self.on_error        = on_error or (lambda e: None)
        self.pause_key       = pause_key
        self.pause_scr       = pause_scr
        self.max_pages       = max_pages
        self.end_markers     = end_markers if end_markers is not None else END_MARKERS
        self.fuzzy_match     = fuzzy_match
        self.match_threshold = match_threshold
        self._stop           = threading.Event()

    # ── ПРОВЕРКИ ──────────────────────────────────────────────────────────

    def check_deps(self) -> tuple[bool, str]:
        """Проверить все зависимости перед сканированием."""
        if not TESSERACT_OK:
            return False, ("pytesseract или Pillow не установлены.\n"
                           "pip install pytesseract pillow")

        # Путь к Tesseract
        if os.path.isfile(TESSERACT_CMD):
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
        try:
            ver = pytesseract.get_tesseract_version()
        except Exception as e:
            return False, (f"Tesseract не найден: {e}\n"
                           "Скачать: https://github.com/UB-Mannheim/tesseract/wiki\n"
                           "При установке выбрать: Russian language pack")

        # Русский язык в Tesseract
        try:
            langs = pytesseract.get_languages()
            if "rus" not in langs:
                return False, ("Русский язык не установлен в Tesseract.\n"
                               "Переустановите с опцией 'Russian language pack'.")
        except Exception:
            pass  # get_languages иногда недоступен — не критично

        # VBoxManage
        if not os.path.isfile(self.vboxmanage):
            return False, f"VBoxManage не найден: {self.vboxmanage}"

        # VM запущена?
        try:
            out = subprocess.run(
                [self.vboxmanage, "list", "runningvms"],
                capture_output=True, text=True, check=True,
            )
            if f'"{self.vm_name}"' not in out.stdout:
                return False, f'VM "{self.vm_name}" не запущена'
        except Exception as e:
            return False, f"Ошибка VBoxManage: {e}"

        return True, f"Tesseract {ver} · VM готова"

    # ── VBOXMANAGE: СКРИНШОТ ──────────────────────────────────────────────

    def _screenshot(self, path: str) -> None:
        """Сделать скриншот VM и сохранить в path (PNG)."""
        subprocess.run(
            [self.vboxmanage, "controlvm", self.vm_name, "screenshotpng", path],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
        )

    # ── VBOXMANAGE: НАЖАТИЕ КЛАВИШИ ──────────────────────────────────────

    _SC_EXT  = {"pagedown": (0x51, True), "pageup": (0x49, True), "home": (0x47, True)}
    _SC_MAKE = {"esc": 0x01, "enter": 0x1C}

    def _send_key(self, key: str) -> None:
        """Отправить скан-код клавиши в VM через VBoxManage."""
        key = key.lower()
        if key in self._SC_EXT:
            m, _ = self._SC_EXT[key]
            codes = ["e0", f"{m:02x}", "e0", f"{(m|0x80):02x}"]
        elif key in self._SC_MAKE:
            m = self._SC_MAKE[key]
            codes = [f"{m:02x}", f"{(m|0x80):02x}"]
        else:
            return
        subprocess.run(
            [self.vboxmanage, "controlvm", self.vm_name, "keyboardputscancode"] + codes,
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
        )
        time.sleep(self.pause_key)

    # ── OCR ОДНОГО СКРИНШОТА ──────────────────────────────────────────────

    def _ocr_screenshot(self, png_path: str) -> tuple[str, list[str]]:
        """
        Применить OCR к PNG, вернуть (сырой текст, список наименований).
        """
        img = Image.open(png_path)

        # Вырезать нужную область экрана
        if self.region:
            x, y, w, h = self.region
            img = img.crop((x, y, x + w, y + h))

        # Вырезать столбец наименований по X
        if self.name_x1 is not None or self.name_x2 is not None:
            x1 = self.name_x1 or 0
            x2 = self.name_x2 or img.width
            img = img.crop((x1, 0, x2, img.height))

        # Предобработка для DOS-шрифта:
        #   - перевести в оттенки серого
        #   - увеличить контрастность
        #   - масштабировать ×2 (улучшает OCR мелкого DOS-шрифта)
        img = img.convert("L")
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)

        # Tesseract: русский + английский, режим PSM 6 (равномерный блок текста)
        config = "--psm 6 -l rus+eng"
        raw = pytesseract.image_to_string(img, config=config)

        names = parse_ocr_text(raw, self.name_x1, self.name_x2)
        return raw, names

    # ── ОСНОВНОЙ СКАНЕР (поток) ───────────────────────────────────────────

    def scan_async(self) -> threading.Thread:
        """Запустить сканирование в фоновом потоке. Вернуть поток."""
        self._stop.clear()
        t = threading.Thread(target=self._scan_worker, daemon=True)
        t.start()
        return t

    def stop(self) -> None:
        """Попросить остановить сканирование."""
        self._stop.set()

    def _scan_worker(self) -> None:
        """Воркер: обходит все страницы DOS-программы, собирает наименования."""
        try:
            result = self._do_scan()
            self.on_done(result)
        except Exception as e:
            import traceback
            self.on_error(f"{e}\n\n{traceback.format_exc()}")

    def _do_scan(self) -> "ScanResult":
        all_names: list[str] = []
        prev_page_key: Optional[str] = None   # для детектирования конца списка

        total_est = self.max_pages
        self.on_progress("Подготовка к сканированию…", 0, total_est)

        with tempfile.TemporaryDirectory() as tmpdir:
            for page_idx in range(self.max_pages):
                if self._stop.is_set():
                    break

                step = page_idx + 1
                self.on_progress(
                    f"Страница {step} (найдено строк: {len(all_names)})…",
                    step, total_est,
                )

                # Сделать паузу и скриншот
                time.sleep(self.pause_scr)
                png = os.path.join(tmpdir, f"page_{page_idx:03d}.png")
                self._screenshot(png)

                raw, names = self._ocr_screenshot(png)

                if not names:
                    # Пустая страница — возможно конец
                    self.on_progress(
                        f"Страница {step}: строк не найдено — конец списка.",
                        step, total_est,
                    )
                    break

                # Маркер конца
                if self.end_markers:
                    if any(m.lower() in raw.lower() for m in self.end_markers):
                        # Добавим строки до маркера
                        all_names.extend(names)
                        break

                # Автоопределение конца: если текущая страница совпадает с предыдущей
                page_key = "|".join(names)
                if page_key == prev_page_key:
                    self.on_progress(
                        f"Страница {step}: повтор предыдущей — конец списка.",
                        step, total_est,
                    )
                    break
                prev_page_key = page_key

                all_names.extend(names)

                # Нажать PageDown для следующей страницы
                if not self._stop.is_set():
                    self._send_key("pagedown")

        self.on_progress(
            f"OCR завершён. Считано строк: {len(all_names)}. Сравниваем с базой…",
            total_est, total_est,
        )

        # Вычислить diff с tovary.csv
        return self._compute_diff(all_names)

    # ── DIFF С БАЗОЙ ──────────────────────────────────────────────────────

    def _load_tovary(self) -> list[dict]:
        if not self.tovary_csv or not Path(self.tovary_csv).exists():
            return []
        tovary = []
        for enc in ("utf-8-sig", "cp1251", "utf-8"):
            try:
                with open(self.tovary_csv, encoding=enc, newline="") as f:
                    reader = csv.DictReader(f, delimiter=";")
                    for row in reader:
                        tovary.append(dict(row))
                break
            except UnicodeDecodeError:
                continue
        return tovary

    def _compute_diff(self, ocr_names: list[str]) -> "ScanResult":
        """
        Сравнить OCR-список с tovary.csv.

        Возвращает ScanResult с:
          - matched:  строки OCR, найденные в базе
          - new:      строки OCR, которых нет в базе → надо добавить
          - removed:  строки базы (active=1), которых нет в OCR → active=0
          - unchanged: строки базы, которые совпали
        """
        tovary = self._load_tovary()
        db_names = [t.get("naimenovanie", "").strip() for t in tovary]

        matched_ocr: list[str]   = []
        new_ocr: list[str]       = []
        matched_db_idx: set[int] = set()

        for ocr_name in ocr_names:
            best_idx: Optional[int] = None
            if self.fuzzy_match:
                # Нечёткий поиск: найти наилучшее совпадение в базе
                best_score = 0.0
                for i, db_name in enumerate(db_names):
                    s = _bigram_score(ocr_name, db_name)
                    if s > best_score:
                        best_score = s
                        best_idx = i
                if best_score < self.match_threshold:
                    best_idx = None
            else:
                # Точное нормализованное совпадение
                ocr_key = _norm_key(ocr_name)
                for i, db_name in enumerate(db_names):
                    if _norm_key(db_name) == ocr_key:
                        best_idx = i
                        break

            if best_idx is not None:
                matched_ocr.append(ocr_name)
                matched_db_idx.add(best_idx)
            else:
                new_ocr.append(ocr_name)

        # Строки базы (active=1), не найденные в OCR → кандидаты на удаление
        removed: list[dict] = []
        unchanged: list[dict] = []
        for i, t in enumerate(tovary):
            if t.get("active", "1") in ("0", ""):
                continue  # уже неактивный — не трогаем
            if i in matched_db_idx:
                unchanged.append(t)
            else:
                removed.append(t)

        return ScanResult(
            ocr_names=ocr_names,
            matched=matched_ocr,
            new=new_ocr,
            removed=removed,
            unchanged=unchanged,
            tovary=tovary,
        )

    # ── ПРИМЕНЕНИЕ ИЗМЕНЕНИЙ ──────────────────────────────────────────────

    def apply_changes(
        self,
        result: "ScanResult",
        add_new: bool = True,
        deactivate_removed: bool = True,
    ) -> tuple[int, int]:
        """
        Применить результаты сканирования к tovary.csv.

        Параметры:
            add_new             — добавить новые строки (в конец файла)
            deactivate_removed  — пометить удалённые как active=0

        Возвращает (добавлено, деактивировано).
        """
        tovary = list(result.tovary)  # копия

        # Деактивировать удалённые
        deactivated = 0
        if deactivate_removed:
            removed_names = {t["naimenovanie"] for t in result.removed}
            for t in tovary:
                if t.get("naimenovanie", "") in removed_names:
                    t["active"] = "0"
                    deactivated += 1

        # Добавить новые
        added = 0
        if add_new and result.new:
            max_pos = max(
                (int(t.get("pozitsiya", 0) or 0) for t in tovary),
                default=0,
            )
            for name in result.new:
                max_pos += 1
                tovary.append({
                    "pozitsiya":   str(max_pos),
                    "naimenovanie": name,
                    "oformlenie":  "",
                    "massa":       "",
                    "ostatok":     "0",
                    "rezerv":      "0",
                    "active":      "1",
                })
                added += 1

        # Сохранить
        self._save_tovary(tovary)
        return added, deactivated

    def _save_tovary(self, tovary: list[dict]) -> None:
        fieldnames = ["pozitsiya", "naimenovanie", "oformlenie", "massa",
                      "ostatok", "rezerv", "active"]
        with open(self.tovary_csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";",
                               extrasaction="ignore")
            w.writeheader()
            for t in tovary:
                w.writerow(t)

    # ── УТИЛИТА: один скриншот для отладки ───────────────────────────────

    def debug_screenshot(self, save_path: str = "debug_ocr.png") -> tuple[str, list[str]]:
        """
        Сделать один скриншот и запустить OCR — без прокрутки.
        Удобно для подбора region / name_x1/x2.

        Возвращает (raw_ocr_text, parsed_names).
        """
        ok, msg = self.check_deps()
        if not ok:
            raise RuntimeError(msg)
        self._screenshot(save_path)
        raw, names = self._ocr_screenshot(save_path)
        print(f"=== RAW OCR ===\n{raw}\n")
        print(f"=== PARSED ({len(names)}) ===")
        for i, n in enumerate(names, 1):
            print(f"  {i:3d}. {n!r}")
        return raw, names


# ── Вспомогательная функция биграммного сходства ─────────────────────────────

def _bigram_score(a: str, b: str) -> float:
    def bg(s):
        s = _norm_key(s)
        return {s[i:i+2] for i in range(len(s)-1)} if len(s) >= 2 else set()
    ba, bb = bg(a), bg(b)
    if not ba and not bb:
        return 1.0
    if not ba or not bb:
        return 0.0
    return len(ba & bb) / max(len(ba), len(bb))


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  РЕЗУЛЬТАТ СКАНИРОВАНИЯ                                                  ║
# ╚══════════════════════════════════════════════════════════════════════════╝

class ScanResult:
    """
    Результат одного сканирования.

    Атрибуты:
        ocr_names  — все имена, считанные OCR (порядок с экрана)
        matched    — OCR-строки, найденные в базе
        new        — OCR-строки, которых НЕТ в базе (новые товары)
        removed    — записи базы (active=1), которых нет в OCR (возможно удалены)
        unchanged  — записи базы, подтверждённые OCR
        tovary     — исходный список базы (для apply_changes)
    """
    def __init__(
        self,
        ocr_names: list[str],
        matched: list[str],
        new: list[str],
        removed: list[dict],
        unchanged: list[dict],
        tovary: list[dict],
    ):
        self.ocr_names  = ocr_names
        self.matched    = matched
        self.new        = new
        self.removed    = removed
        self.unchanged  = unchanged
        self.tovary     = tovary

    @property
    def has_changes(self) -> bool:
        return bool(self.new or self.removed)

    def summary(self) -> str:
        lines = [
            f"Всего OCR строк:    {len(self.ocr_names)}",
            f"Совпало с базой:    {len(self.matched)}",
            f"✨ Новых товаров:   {len(self.new)}",
            f"🗑 Исчезло из DOS: {len(self.removed)}",
            f"✓ Без изменений:   {len(self.unchanged)}",
        ]
        return "\n".join(lines)

    def new_names_str(self) -> str:
        return "\n".join(f"  + {n}" for n in self.new)

    def removed_names_str(self) -> str:
        return "\n".join(f"  - {t.get('naimenovanie','?')}" for t in self.removed)


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  ИНТЕГРАЦИЯ С GUI (фрагмент кода для main_gui_v3.py)                    ║
# ╚══════════════════════════════════════════════════════════════════════════╝
#
# Добавить в класс App (вкладка "Список товаров"):
#
#   from ocr_scanner import OCRScanner, ScanResult
#
#   def _build_tab_tovary(self):
#       ...
#       # Кнопка сканирования (добавить в панель кнопок)
#       ttk.Button(btn_frame, text="🔍 Сканировать ассортимент",
#                  command=self._ocr_scan).pack(side="left", padx=4)
#       self.lbl_scan_stat = ttk.Label(btn_frame, text="", foreground=CLR_INFO)
#       self.lbl_scan_stat.pack(side="left", padx=8)
#
#   def _ocr_scan(self):
#       """Запустить OCR-сканирование ассортимента DOS-программы."""
#       from ocr_scanner import OCRScanner
#       scanner = OCRScanner(
#           vm_name=IMYA_VM,
#           vboxmanage=VBOXMANAGE_PATH,
#           tovary_csv=TOVARY_CSV,
#           page_size=RAZMER_STRANICY,
#           on_progress=self._ocr_progress,
#           on_done=self._ocr_done,
#           on_error=self._ocr_error,
#       )
#       ok, msg = scanner.check_deps()
#       if not ok:
#           messagebox.showerror("OCR: ошибка", msg); return
#       self.lbl_scan_stat.config(text="⏳ Сканирование…", foreground=CLR_WARN)
#       self._ocr_scanner = scanner
#       scanner.scan_async()
#
#   def _ocr_progress(self, msg: str, step: int, total: int):
#       self.after(0, lambda: self.lbl_scan_stat.config(
#           text=f"⏳ {msg}", foreground=CLR_WARN))
#
#   def _ocr_done(self, result: ScanResult):
#       self.after(0, lambda: self._ocr_show_result(result))
#
#   def _ocr_error(self, err: str):
#       self.after(0, lambda: (
#           messagebox.showerror("OCR: ошибка", err),
#           self.lbl_scan_stat.config(text="❌ Ошибка", foreground=CLR_ERR),
#       ))
#
#   def _ocr_show_result(self, result: ScanResult):
#       from ocr_scanner import ScanResult
#       if not result.has_changes:
#           self.lbl_scan_stat.config(
#               text=f"✓ Без изменений ({len(result.unchanged)} поз.)",
#               foreground=CLR_OK)
#           messagebox.showinfo("OCR: синхронизация", result.summary())
#           return
#
#       msg = (f"{result.summary()}\n\n"
#              f"Новые:\n{result.new_names_str() or '  —'}\n\n"
#              f"Исчезли из DOS:\n{result.removed_names_str() or '  —'}\n\n"
#              "Применить изменения?")
#       if not messagebox.askyesno("OCR: найдены изменения", msg):
#           self.lbl_scan_stat.config(text="Отменено", foreground=CLR_WARN)
#           return
#
#       added, deact = self._ocr_scanner.apply_changes(result)
#       self._reload_baza()
#       self._tov_refresh()
#       self.lbl_scan_stat.config(
#           text=f"✓ +{added} новых, деакт. {deact}", foreground=CLR_OK)
#       self._log(f"OCR-синхронизация: +{added} новых, деактивировано {deact}")
#       messagebox.showinfo("OCR: готово",
#           f"Добавлено новых позиций:  {added}\n"
#           f"Деактивировано (исчезли): {deact}")


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  ЗАПУСК ИЗ КОМАНДНОЙ СТРОКИ (отладка)                                   ║
# ╚══════════════════════════════════════════════════════════════════════════╝

if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="OCR Scanner — синхронизация ассортимента DOS")
    ap.add_argument("--vm",    default=DEFAULT_VM_NAME,    help="Имя VM")
    ap.add_argument("--vbox",  default=DEFAULT_VBOXMANAGE, help="Путь к VBoxManage.exe")
    ap.add_argument("--csv",   default="baza/tovary.csv",  help="Путь к tovary.csv")
    ap.add_argument("--debug", action="store_true",        help="Один скриншот (без прокрутки)")
    ap.add_argument("--apply", action="store_true",        help="Сразу применить без вопросов")
    ap.add_argument("--region", nargs=4, type=int, metavar=("X","Y","W","H"),
                    help="Область скриншота (пиксели)")
    ap.add_argument("--x1", type=int, help="Левая граница столбца наименований")
    ap.add_argument("--x2", type=int, help="Правая граница столбца наименований")
    args = ap.parse_args()

    scanner = OCRScanner(
        vm_name=args.vm,
        vboxmanage=args.vbox,
        tovary_csv=Path(args.csv),
        region=tuple(args.region) if args.region else None,
        name_x1=args.x1,
        name_x2=args.x2,
    )

    # Проверка зависимостей
    ok, msg = scanner.check_deps()
    print(f"Зависимости: {'OK' if ok else 'ОШИБКА'} — {msg}")
    if not ok:
        raise SystemExit(1)

    if args.debug:
        # Режим отладки: один скриншот
        scanner.debug_screenshot("debug_ocr.png")
        raise SystemExit(0)

    # Полное сканирование
    done_result: list[ScanResult] = []
    errors: list[str] = []

    def _prog(msg, step, total):
        print(f"  [{step:2d}/{total}] {msg}")

    def _done(r):
        done_result.append(r)

    def _err(e):
        errors.append(e)

    t = scanner.scan_async()
    scanner.on_progress = _prog
    scanner.on_done     = _done
    scanner.on_error    = _err
    t.join()

    if errors:
        print(f"\nОШИБКА:\n{errors[0]}")
        raise SystemExit(1)

    result = done_result[0]
    print(f"\n{result.summary()}")

    if result.new:
        print(f"\nНовые товары ({len(result.new)}):")
        print(result.new_names_str())

    if result.removed:
        print(f"\nИсчезли из DOS ({len(result.removed)}):")
        print(result.removed_names_str())

    if result.has_changes:
        if args.apply:
            added, deact = scanner.apply_changes(result)
            print(f"\n✓ Применено: +{added} новых, деактивировано {deact}")
        else:
            ans = input("\nПрименить изменения? [y/N]: ").strip().lower()
            if ans == "y":
                added, deact = scanner.apply_changes(result)
                print(f"✓ Применено: +{added} новых, деактивировано {deact}")
            else:
                print("Отменено.")
    else:
        print("\n✓ Изменений нет.")
