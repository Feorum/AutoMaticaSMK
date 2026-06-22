# -*- coding: utf-8 -*-
"""
auto_vvod.0.56.py  (v0.5)

НОВОЕ В v0.5:
  - Сброс курсора: 20× PageUp в начале каждого задания → tek_poz = 0.
    Настраивается через SBROS_PAGEUP.
  - Резервы: столбец 'rezerv' в tovary.csv (1=резерв, 0=обычный).
    Для резервных товаров остаток отслеживается строго и блокирует ввод
    при нехватке (с возможностью override). Для остальных — фоновый учёт.
  - Дубли: столбцы 'oformlenie' и 'massa' в tovary.csv.
    Поиск и отображение учитывают их. В файле задания можно уточнять:
      Название;оформление;масса;количество
      Название;;;количество   (уточнения пропущены)
  - Выбор задания: нумерованный список в консоли → вводим номер или "все".
"""
import re
import sys
import os
import csv
import time
import difflib
from thefuzz import fuzz
import subprocess
from datetime import datetime

pyautogui = None
try:
    import pyautogui as _pyautogui
    pyautogui = _pyautogui
except Exception:
    pass

# ╔══════════════════════════════════════════════════════════════════╗
# ║                       НАСТРОЙКИ                                  ║
# ╚══════════════════════════════════════════════════════════════════╝

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BAZA_DIR      = os.path.join(BASE_DIR, "../baza")
ZADANIYA_DIR  = os.path.join(BASE_DIR, "../zadaniya")
ARHIV_DIR     = os.path.join(ZADANIYA_DIR, "_arhiv")
TOVARY_CSV    = os.path.join(BAZA_DIR, "tovary.csv")
ISTORIYA_CSV  = os.path.join(BAZA_DIR, "istoriya.csv")

REZHIM_VVODA         = "pyautogui"
IMYA_VM              = "Xp"
VBOXMANAGE_PATH      = r"C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"

OBRATNYI_OTSCHET     = 5
PAUZA_KLAVISHA       = 0.05
PAUZA_MEZHDU_POZICIYAMI = 0.15
RAZMER_STRANICY      = 17
SORTIROVAT_PO_STROKE = True
KLAVISHA_PODTVERZHDENIYA = ""
PORICHE_POISKA       = 0.5
PROVERYAT_OSTATOK    = True
OPROS_PAPKI_SEK      = 2

# Сброс курсора: сколько PageUp слать в начале задания.
# 20 × 17 строк = 340 строк вверх — заведомо больше любого списка.
SBROS_PAGEUP = 20

if pyautogui is not None:
    pyautogui.PAUSE    = PAUZA_KLAVISHA
    pyautogui.FAILSAFE = True


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       БАЗА ТОВАРОВ (CSV)                         ║
# ╚══════════════════════════════════════════════════════════════════╝

def normalize(s: str) -> str:
    return " ".join(s.lower().replace("/", " ").split())


def prochitat_tekst(path: str) -> str:
    for enc in ("utf-8-sig", "cp1251"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read()


def zagruzit_bazu() -> list[dict] | None:
    """Загрузить базу товаров.
    Возвращает список dict: {naimenovanie, oformlenie, massa, ostatok, rezerv}.
    Порядок строк = порядок на экране программы."""
    if not os.path.isfile(TOVARY_CSV):
        return None
    tovary = []
    tekst  = prochitat_tekst(TOVARY_CSV)
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
        rezerv = rezerv_raw in ("1", "yes", "да", "true")
        tovary.append({
            "naimenovanie": name,
            "oformlenie":   (row.get("oformlenie") or "").strip(),
            "massa":        (row.get("massa") or "").strip(),
            "ostatok":      ost,
            "rezerv":       rezerv,
        })
    return tovary


def sohranit_bazu(tovary: list[dict]) -> None:
    os.makedirs(BAZA_DIR, exist_ok=True)
    with open(TOVARY_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["pozitsiya", "naimenovanie", "oformlenie",
                         "massa", "ostatok", "rezerv"])
        for i, t in enumerate(tovary, 1):
            writer.writerow([
                i,
                t["naimenovanie"],
                t["oformlenie"],
                t["massa"],
                t["ostatok"],
                "1" if t["rezerv"] else "0",
            ])


def zapisat_istoriyu(zapisi: list) -> None:
    os.makedirs(BAZA_DIR, exist_ok=True)
    novyi = not os.path.isfile(ISTORIYA_CSV)
    with open(ISTORIYA_CSV, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        if novyi:
            writer.writerow(["data_vremya", "zadanie", "stroka",
                             "naimenovanie", "oformlenie", "massa",
                             "kolichestvo", "ostatok_do", "ostatok_posle",
                             "rezerv"])
        for z in zapisi:
            writer.writerow(z)


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ПОИСК ПО БАЗЕ                              ║
# ╚══════════════════════════════════════════════════════════════════╝

CATEGORIES = {
    'колбаса': ['колбаса', 'к-са', 'к-ка', 'изделие колбасное'],
    'сосиски': ['сосиски', 'сосис'],
    'сардельки': ['сардельки', 'сард'],
    'сало': ['сало'],
    'крылышки': ['крылышки', 'крылья'],
    'паштет': ['паштет', 'паштетный', 'паштетная']
}

def get_product_category(text: str) -> str:
    """Определяет основную категорию продукта из строки."""
    text_lower = text.lower()
    for cat_name, keywords in CATEGORIES.items():
        if any(kw in text_lower for kw in keywords):
            return cat_name
    return ""

def clean_and_stem(text: str) -> str:
    """Удаляет мусор, знаки препинания и срезает окончания для точного сравнения."""
    if not text:
        return ""
    # Удаляем знаки препинания, приводим к нижнему регистру
    text = re.sub(r'[^\w\s]', ' ', text.lower())
    words = text.split()

    # Простейший стемминг: срезаем падежные окончания (ая, ый, ые, ое, ую, а, ы)
    # Чтобы "неженка паштетный" и "неженка паштетная" стали одинаковыми
    stemmed_words = []
    for w in words:
        if len(w) > 4:
            for end in ['ая', 'ый', 'ые', 'ое', 'ую', 'ая', 'их', 'ых']:
                if w.endswith(end):
                    w = w[:-len(end)]
                    break
        stemmed_words.append(w)

    return " ".join(stemmed_words)

def _score_field(zapros: str, znachenie: str) -> float:
    """Оценка совпадения одного поля (оформление или масса)."""
    if not zapros:
        return 0.0
    if not znachenie:
        return -0.3
    nz, nv = clean_and_stem(zapros), clean_and_stem(znachenie)
    if nz == nv:
        return 0.2
    if nz in nv or nv in nz:
        return 0.1
    return (fuzz.ratio(nz, nv) / 100.0) * 0.15 - 0.05


def naiti_kandidatov(
        zapros: str,
        tovary: list[dict],
        zapros_oformlenie: str = "",
        zapros_massa: str = "",
) -> list[tuple]:
    """Найти позиции базы. Возвращает список (idx, товар, оценка)."""

    # Определяем категорию запроса клиента (например, "сардельки")
    zapros_cat = get_product_category(zapros)

    nz = clean_and_stem(zapros)
    if not nz:
        return []

    rez = []
    for idx, t in enumerate(tovary):
        naim = t["naimenovanie"]
        nn = clean_and_stem(naim)
        if not nn:
            continue

        # --- ЖЕСТКАЯ ПРОВЕРКА КАТЕГОРИИ ---
        tovar_cat = get_product_category(naim)
        # Если категории определились и они РАЗНЫЕ (Сардельки vs Крылышки) — пропускаем товар
        if zapros_cat and tovar_cat and zapros_cat != tovar_cat:
            continue

        # --- ОЦЕНКА СХОДСТВА НАЗВАНИЙ ЧЕРЕЗ THEFUZZ ---
        ts_ratio = fuzz.token_set_ratio(nz, nn) / 100.0
        sort_ratio = fuzz.token_sort_ratio(nz, nn) / 100.0

        base = max(ts_ratio, sort_ratio)

        # Бонусы за совпадения
        if nz == nn:
            base = max(base, 1.0)
        elif nz in nn:
            base = max(base, 0.9 + 0.1 * (len(nz) / max(len(nn), 1)))

        # Если в запросе было "премиум плотные", а в базе просто "сосиски сливочные",
        # то fuzz.token_sort_ratio будет очень низким (~0.4).
        # Благодаря этому "Сосиски плотные" не пройдут этот порог.
        if base < PORICHE_POISKA:
            continue

        # Уточнение по оформлению и массе
        score = (base
                 + _score_field(zapros_oformlenie, t["oformlenie"])
                 + _score_field(zapros_massa,       t["massa"]))
        rez.append((idx, t, score))

    # Сортировка по score (индекс 2 в кортеже) в порядке убывания
    rez.sort(key=lambda x: x[2], reverse=True)
    return rez


def format_tovar(t: dict) -> str:
    """Краткое отображение товара со всеми параметрами."""
    parts = [t["naimenovanie"]]
    if t["oformlenie"]:
        parts.append(t["oformlenie"])
    if t["massa"]:
        parts.append(t["massa"])
    return " | ".join(parts)


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ФАЙЛЫ-ЗАДАНИЯ                              ║
# ╚══════════════════════════════════════════════════════════════════╝

def spisok_zadaniy() -> list[str]:
    if not os.path.isdir(ZADANIYA_DIR):
        return []
    return sorted(
        os.path.join(ZADANIYA_DIR, f)
        for f in os.listdir(ZADANIYA_DIR)
        if f.lower().endswith(".txt")
        and os.path.isfile(os.path.join(ZADANIYA_DIR, f))
    )


def vybrat_zadaniya(zadaniya: list[str]) -> list[str]:
    """Показать нумерованный список файлов, вернуть выбранные пользователем."""
    print(f"\nНайдено заданий: {len(zadaniya)}")
    print("-" * 48)
    for i, path in enumerate(zadaniya, 1):
        print(f"  {i:>3}. {os.path.basename(path)}")
    print("-" * 48)
    print("  Введите номер, несколько номеров через запятую,")
    print("  диапазон (например 2-5), или 'все' для всех.")
    print("  Enter без ввода = отмена.")

    raw = input("\nВаш выбор: ").strip().lower()
    if not raw:
        return []
    if raw in ("все", "all", "v"):
        return zadaniya[:]

    vybrano: list[str] = []
    for chunk in raw.split(","):
        chunk = chunk.strip()
        if "-" in chunk:
            parts = chunk.split("-", 1)
            try:
                a, b = int(parts[0]), int(parts[1])
                for n in range(a, b + 1):
                    if 1 <= n <= len(zadaniya):
                        p = zadaniya[n - 1]
                        if p not in vybrano:
                            vybrano.append(p)
            except ValueError:
                print(f"  Не понял диапазон: {chunk!r}, пропускаю.")
        elif chunk.isdigit():
            n = int(chunk)
            if 1 <= n <= len(zadaniya):
                p = zadaniya[n - 1]
                if p not in vybrano:
                    vybrano.append(p)
            else:
                print(f"  Нет задания с номером {n}, пропускаю.")
        else:
            print(f"  Не понял: {chunk!r}, пропускаю.")
    return vybrano


def razobrat_zadanie(path: str) -> tuple[list, list]:
    """Разобрать файл-задание.
    Форматы строки (разделитель ';'):
      Название;кол
      Название;оформление;кол
      Название;оформление;масса;кол
    Пустое поле = не уточняется: 'Молоко;;0.9л;10'."""
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

        # Последнее поле всегда количество
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


def v_arhiv(path: str) -> str | None:
    os.makedirs(ARHIV_DIR, exist_ok=True)
    bn    = os.path.basename(path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst   = os.path.join(ARHIV_DIR, f"{stamp}__{bn}")
    try:
        os.replace(path, dst)
        return dst
    except OSError:
        return None


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ВВОД В ПРОГРАММУ                           ║
# ╚══════════════════════════════════════════════════════════════════╝

SC_MAKE = {
    "0": 0x0B, "1": 0x02, "2": 0x03, "3": 0x04, "4": 0x05,
    "5": 0x06, "6": 0x07, "7": 0x08, "8": 0x09, "9": 0x0A,
    "enter": 0x1C, "tab": 0x0F, "esc": 0x01, "f2": 0x3C,
}
SC_EXT = {
    "up": 0x48, "down": 0x50, "left": 0x4B, "right": 0x4D,
    "pageup": 0x49, "pagedown": 0x51, "home": 0x47, "end": 0x4F,
}


def _vbox_codes_for(klavisha: str) -> list[str]:
    if klavisha in SC_EXT:
        m = SC_EXT[klavisha]
        return ["e0", f"{m:02x}", "e0", f"{(m | 0x80):02x}"]
    if klavisha in SC_MAKE:
        m = SC_MAKE[klavisha]
        return [f"{m:02x}", f"{(m | 0x80):02x}"]
    raise KeyError(f"Нет скан-кода для клавиши: {klavisha!r}")


def _vbox_send(codes: list[str]) -> None:
    cmd = [VBOXMANAGE_PATH, "controlvm", IMYA_VM,
           "keyboardputscancode"] + codes
    subprocess.run(cmd, check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)


def nazhat_klavishu(klavisha: str) -> None:
    if REZHIM_VVODA == "vboxmanage":
        _vbox_send(_vbox_codes_for(klavisha))
        if PAUZA_KLAVISHA > 0:
            time.sleep(PAUZA_KLAVISHA)
    else:
        pyautogui.press(klavisha)


def sbrosit_kursor() -> None:
    """Послать SBROS_PAGEUP нажатий PageUp, чтобы гарантированно
    оказаться на первой строке списка."""
    print(f"  Сброс курсора: {SBROS_PAGEUP}× PageUp...", end=" ", flush=True)
    for _ in range(SBROS_PAGEUP):
        nazhat_klavishu("pageup")
    print("готово. Курсор на строке 1.")


def peremestit_kursor(tek_poz: int, cel_poz: int) -> None:
    delta = cel_poz - tek_poz
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
    if KLAVISHA_PODTVERZHDENIYA:
        nazhat_klavishu(KLAVISHA_PODTVERZHDENIYA)


def proverit_rezhim() -> tuple[bool, str]:
    if REZHIM_VVODA == "vboxmanage":
        if not os.path.isfile(VBOXMANAGE_PATH):
            return False, (f"Не найден VBoxManage.exe:\n  {VBOXMANAGE_PATH}")
        try:
            out = subprocess.run(
                [VBOXMANAGE_PATH, "list", "runningvms"],
                capture_output=True, text=True, check=True)
            if f'"{IMYA_VM}"' not in out.stdout:
                return False, (f"Виртуалка \"{IMYA_VM}\" не запущена.")
        except Exception as e:
            return False, f"Не удалось опросить VirtualBox: {e}"
        return True, f"Режим vboxmanage. VM \"{IMYA_VM}\" запущена."
    else:
        if pyautogui is None:
            return False, "Нужна библиотека pyautogui: pip install pyautogui"
        return True, f"Режим {REZHIM_VVODA}."


# ╔══════════════════════════════════════════════════════════════════╗
# ║                  ОБРАБОТКА ОДНОГО ЗАДАНИЯ                        ║
# ╚══════════════════════════════════════════════════════════════════╝

def obrabotat_zadanie(path: str, tovary: list[dict]) -> bool:
    z_name = os.path.basename(path)
    print("\n" + "=" * 64)
    print(f"  ЗАДАНИЕ: {z_name}")
    print("=" * 64)

    pozicii, oshibki = razobrat_zadanie(path)
    if oshibki:
        print("Строки, которые не удалось разобрать:")
        for nomer, raw, prich in oshibki:
            print(f"  строка {nomer}: {raw!r} — {prich}")
    if not pozicii:
        print("Нет корректных позиций. Задание пропущено (в архив).")
        v_arhiv(path)
        return False

    # ── Сопоставление со базой ──
    plan: list[dict] = []
    ne_naideno: list  = []
    for name, oform, massa, kol in pozicii:
        kand = naiti_kandidatov(name, tovary, oform, massa)
        if not kand:
            ne_naideno.append((name, oform, massa, kol))
            continue
        idx, t, score = kand[0]
        ambig = (len(kand) > 1 and kand[1][2] > score - 0.05)
        plan.append({
            "idx":   idx,
            "t":     t,
            "kol":   int(kol),
            "ambig": ambig,
        })

    if SORTIROVAT_PO_STROKE:
        plan.sort(key=lambda x: x["idx"])

    # ── Показ плана ──
    print(f"\nК вводу: {len(plan)} позиц.")
    print("-" * 64)
    blokirovka = False
    for p in plan:
        t    = p["t"]
        label = format_tovar(t)
        rez_mark = " [R]" if t["rezerv"] else ""
        pometka   = "  [похожие есть!]" if p["ambig"] else ""
        ost       = t["ostatok"]
        posle     = ost - p["kol"]

        if PROVERYAT_OSTATOK:
            ost_info = f"  (ост {ost} → {posle})"
            if posle < 0:
                if t["rezerv"]:
                    ost_info += "  !!! НЕ ХВАТАЕТ (РЕЗЕРВ — БЛОК)"
                    blokirovka = True
                else:
                    ost_info += "  ! не хватает (фон)"
        else:
            ost_info = ""

        print(f"  стр {p['idx']+1:>3}: {label:<32}"
              f"= {p['kol']}{ost_info}{rez_mark}{pometka}")

    if ne_naideno:
        print("\n  НЕ НАЙДЕНЫ в базе (пропущены):")
        for name, oform, massa, kol in ne_naideno:
            uточн = " | ".join(x for x in [oform, massa] if x)
            print(f"    {name}{' | ' + uточн if uточн else ''} = {kol}")

    if blokirovka:
        print("\n  СТОП: по резервным товарам остатка не хватает.")
        override = input("  Всё равно продолжить? (да/Enter=нет): ").strip().lower()
        if override not in ("да", "yes", "y", "д"):
            print("Задание отменено. Файл остаётся в папке.")
            return False

    if not plan:
        print("Нечего вводить. Задание в архив.")
        v_arhiv(path)
        return False

    # ── Стартовая позиция и подтверждение ──
    go = input(f"\nЗапустить ввод {len(plan)} позиц.? (Enter=да, н=отмена): ").strip().lower()
    if go in ("н", "n", "нет", "no"):
        print("Отменено. Задание остаётся в папке.")
        return False

    if REZHIM_VVODA != "vboxmanage":
        print(f"\nПереключитесь в окно старой программы! Старт через:")
        for i in range(OBRATNYI_OTSCHET, 0, -1):
            print(f"  {i}...", flush=True)
            time.sleep(1)
        print("  ПОЕХАЛИ\n")
    else:
        print("\nОтправляю клавиши прямо в виртуалку.\n")

    # ── Сброс курсора наверх ──
    sbrosit_kursor()
    tek_poz = 0

    # ── Ввод ──
    istoriya: list = []
    teper    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    vvedeno  = 0
    failsafe_exc = getattr(pyautogui, "FailSafeException", None) if pyautogui else None

    try:
        for n, p in enumerate(plan, 1):
            peremestit_kursor(tek_poz, p["idx"])
            vvesti_kolichestvo(str(p["kol"]))
            tek_poz = p["idx"]

            t       = p["t"]
            ost_do  = t["ostatok"]
            ost_posle = ost_do - p["kol"]
            t["ostatok"] = ost_posle

            istoriya.append([
                teper, z_name, p["idx"] + 1,
                t["naimenovanie"], t["oformlenie"], t["massa"],
                p["kol"], ost_do, ost_posle,
                "1" if t["rezerv"] else "0",
                               ])
            vvedeno += 1

            label = format_tovar(t)
            print(f"  [{n}/{len(plan)}] стр {p['idx']+1}: "
                  f"{label} = {p['kol']}  (ост {ost_do}→{ost_posle})  ✓")

            if PAUZA_MEZHDU_POZICIYAMI > 0:
                time.sleep(PAUZA_MEZHDU_POZICIYAMI)

        print("\nЗадание введено.")

    except subprocess.CalledProcessError as e:
        err = (e.stderr or b"").decode("utf-8", "ignore") \
            if isinstance(e.stderr, bytes) else (e.stderr or "")
        print(f"\n  !!! ОШИБКА VBoxManage: {err.strip()}")
        print("  Сохраняю введённое.")
    except Exception as e:
        if failsafe_exc is not None and isinstance(e, failsafe_exc):
            print("\n  !!! АВАРИЙНЫЙ СТОП. Ввод прерван. Сохраняю.")
        else:
            raise

    if istoriya:
        sohranit_bazu(tovary)
        zapisat_istoriyu(istoriya)
        print(f"  Учёт обновлён, история: {len(istoriya)} строк.")

    if vvedeno == len(plan):
        dst = v_arhiv(path)
        if dst:
            print(f"  Задание → архив: {os.path.basename(dst)}")
    else:
        print("  Задание выполнено НЕ полностью — файл остаётся в папке.")

    return vvedeno > 0


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ГЛАВНЫЙ ЦИКЛ                               ║
# ╚══════════════════════════════════════════════════════════════════╝

def main() -> None:
    print("=" * 64)
    print("  ПОМОЩНИК ВВОДА ПРОДУКЦИИ  v0.5")
    print("=" * 64)
    print(f"Режим ввода: {REZHIM_VVODA}")
    print(f"База товаров: {TOVARY_CSV}")
    print(f"Задания:      {ZADANIYA_DIR}")

    ok, soobshchenie = proverit_rezhim()
    print(soobshchenie)
    if not ok:
        print("Ввод невозможен — исправьте ошибку выше.")

    tovary = zagruzit_bazu()
    if tovary is None:
        print(f"\nНе найдена база: {TOVARY_CSV}")
        print("Создайте baza/tovary.csv со столбцами:")
        print("  pozitsiya;naimenovanie;oformlenie;massa;ostatok;rezerv")
        return
    print(f"Загружена база: {len(tovary)} товаров.")

    os.makedirs(ZADANIYA_DIR, exist_ok=True)
    print("\nРежим ЦИКЛА. Кладите задания (.txt) в папку zadaniya.")
    print("Команды: Enter — проверить папку;  'выход' — завершить.\n")

    while True:
        zadaniya = spisok_zadaniy()
        if not zadaniya:
            try:
                cmd = input("Ожидание задания... (Enter=проверить, 'выход'): ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                break
            if cmd in ("выход", "exit", "quit", "q"):
                break
            time.sleep(OPROS_PAPKI_SEK)
            continue

        # ── Выбор заданий из списка ──
        vybrano = vybrat_zadaniya(zadaniya)
        if not vybrano:
            print("Ничего не выбрано.\n")
            time.sleep(OPROS_PAPKI_SEK)
            continue

        # Перезагружаем базу — вдруг пользователь правил CSV вручную
        tovary = zagruzit_bazu() or tovary

        for path in vybrano:
            if not os.path.isfile(path):
                print(f"  Файл пропал: {os.path.basename(path)}, пропускаю.")
                continue
            obrabotat_zadanie(path, tovary)

        print("\nВсе выбранные задания обработаны. Жду новые.\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано.")
    except Exception:
        import traceback
        print("\n!!! ОШИБКА !!!")
        traceback.print_exc()
    finally:
        try:
            input("\n--- Нажмите Enter для выхода. ---")
        except EOFError:
            pass