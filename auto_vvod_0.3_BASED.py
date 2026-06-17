# -*- coding: utf-8 -*-
"""
auto_vvod.py  (v0.3)  — помощник ввода продукции в старую DOS/Oracle-программу.

НОВОЕ В v0.3:
  - ЦИКЛ: скрипт не выключается после задания. Выполнил → перенёс файл задания
    в архив (zadaniya/_arhiv) → ждёт новый файл в папке zadaniya → повторяет.
  - ПОСТОЯННАЯ БАЗА ТОВАРОВ: baza/tovary.csv (позиция;наименование;остаток).
    Это и "карта экрана" (порядок строк), и учёт остатков. Скрипт читает базу,
    при отгрузке вычитает остаток и сохраняет обратно — переживает перезапуск.
  - ИСТОРИЯ: baza/istoriya.csv — все операции дописываются (дата, задание,
    товар, количество, остаток до/после).

ВАЖНО: остатки вы ведёте САМИ (старая программа их пока не отдаёт без зрения).
  Скрипт честно предупреждает, если по ЕГО учёту остатка не хватает, но ввод
  в программу всё равно разрешает — решение за вами.

ПРИНЦИП ВВОДА ("вслепую по списку"): скрипт НЕ видит экран. Он доверяет порядку
  строк в базе и счётчику позиции курсора. Базу держите в порядке экрана,
  а курсор перед стартом — на известной строке (по умолчанию первой).
"""

import sys
import os
import csv
import time
import difflib
from datetime import datetime

try:
    import pyautogui
except ImportError:
    print("Не установлена библиотека pyautogui.  Установите:  pip install pyautogui")
    input("Нажмите Enter для выхода.")
    sys.exit(1)

# ╔══════════════════════════════════════════════════════════════════╗
# ║                       НАСТРОЙКИ                                     ║
# ╚══════════════════════════════════════════════════════════════════╝

# Базовая папка: рядом с exe (если собрано PyInstaller) или рядом с .py.
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BAZA_DIR = os.path.join(BASE_DIR, "baza")               # база товаров + история
ZADANIYA_DIR = os.path.join(BASE_DIR, "zadaniya")       # входящие задания
ARHIV_DIR = os.path.join(ZADANIYA_DIR, "_arhiv")        # выполненные задания

TOVARY_CSV = os.path.join(BAZA_DIR, "tovary.csv")
ISTORIYA_CSV = os.path.join(BAZA_DIR, "istoriya.csv")

OBRATNYI_OTSCHET = 5           # секунд на переключение в окно программы
PAUZA_KLAVISHA = 0.05          # пауза между нажатиями
PAUZA_MEZHDU_POZICIYAMI = 0.15 # пауза между товарами
RAZMER_STRANICY = 17           # на сколько строк прыгает PageDown/PageUp
SORTIROVAT_PO_STROKE = True    # минимум перемещений курсора
KLAVISHA_PODTVERZHDENIYA = ""  # что нажать после цифр: "", "enter", "f2", "tab", "down"
PORICHE_POISKA = 0.55          # строгость нечёткого поиска (0..1)
PROVERYAT_OSTATOK = True       # предупреждать, если остатка не хватает

OPROS_PAPKI_SEK = 2            # как часто проверять появление нового задания

# ── Защита pyautogui ──
pyautogui.PAUSE = PAUZA_KLAVISHA
pyautogui.FAILSAFE = True   # мышь в левый верхний угол = аварийный СТОП


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       БАЗА ТОВАРОВ (CSV)                            ║
# ╚══════════════════════════════════════════════════════════════════╝

def normalize(s: str) -> str:
    return " ".join(s.lower().replace("/", " ").split())


def zagruzit_bazu():
    """Загрузить базу товаров. Возвращает список dict: {pozitsiya, naimenovanie, ostatok}.
    Порядок строк в файле = порядок на экране программы."""
    if not os.path.isfile(TOVARY_CSV):
        return None
    tovary = []
    with open(TOVARY_CSV, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=",")
        for row in reader:
            name = (row.get("naimenovanie") or "").strip()
            if not name:
                continue
            try:
                ost = int((row.get("ostatok") or "0").strip() or 0)
            except ValueError:
                ost = 0
            tovary.append({"naimenovanie": name, "ostatok": ost})
    return tovary


def sohranit_bazu(tovary):
    """Сохранить базу товаров обратно в CSV (с BOM для Excel)."""
    os.makedirs(BAZA_DIR, exist_ok=True)
    with open(TOVARY_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["pozitsiya", "naimenovanie", "ostatok"])
        for i, t in enumerate(tovary, 1):
            writer.writerow([i, t["naimenovanie"], t["ostatok"]])


def zapisat_istoriyu(zapisi):
    """Дописать строки в историю операций (создаёт файл с заголовком при первом запуске)."""
    os.makedirs(BAZA_DIR, exist_ok=True)
    novyi = not os.path.isfile(ISTORIYA_CSV)
    with open(ISTORIYA_CSV, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        if novyi:
            writer.writerow(["data_vremya", "zadanie", "stroka",
                             "naimenovanie", "kolichestvo",
                             "ostatok_do", "ostatok_posle"])
        for z in zapisi:
            writer.writerow(z)


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ПОИСК ПО БАЗЕ                                 ║
# ╚══════════════════════════════════════════════════════════════════╝

def naiti_kandidatov(zapros, tovary):
    """Найти позиции базы под запрос. Возвращает (индекс, наимен., остаток, оценка)."""
    nz = normalize(zapros)
    rez = []
    for idx, t in enumerate(tovary):
        nn = normalize(t["naimenovanie"])
        if not nn:
            continue
        if nz in nn:
            score = 0.9 + 0.1 * (len(nz) / max(len(nn), 1))
        elif len(nz) >= 4 and any(w.startswith(nz) for w in nn.split()):
            score = 0.88
        else:
            score = difflib.SequenceMatcher(None, nz, nn).ratio()
        if score >= PORICHE_POISKA:
            rez.append((idx, t["naimenovanie"], t["ostatok"], score))
    rez.sort(key=lambda x: x[3], reverse=True)
    return rez


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ФАЙЛЫ-ЗАДАНИЯ                                 ║
# ╚══════════════════════════════════════════════════════════════════╝

def spisok_zadaniy():
    """Список .txt-файлов заданий в папке (без архива)."""
    if not os.path.isdir(ZADANIYA_DIR):
        return []
    return sorted(
        os.path.join(ZADANIYA_DIR, f)
        for f in os.listdir(ZADANIYA_DIR)
        if f.lower().endswith(".txt") and os.path.isfile(os.path.join(ZADANIYA_DIR, f))
    )


def razobrat_zadanie(path):
    """Разобрать файл-задание. Возвращает (позиции, ошибки).
    Позиция = (название, количество). Формат строки: Название;Кол  или  Название Кол."""
    pozicii, oshibki = [], []
    with open(path, encoding="utf-8-sig") as f:
        for nomer, raw in enumerate(f, 1):
            s = raw.strip()
            if not s or s.startswith("#"):
                continue
            if ";" in s:
                name, _, kol = s.partition(";")
            else:
                toks = s.rsplit(" ", 1)
                if len(toks) == 2 and toks[1].strip().isdigit():
                    name, kol = toks[0], toks[1]
                else:
                    oshibki.append((nomer, raw.strip(), "нет количества"))
                    continue
            name, kol = name.strip(), kol.strip()
            if not name:
                oshibki.append((nomer, raw.strip(), "пустое название"))
            elif not kol.isdigit():
                oshibki.append((nomer, raw.strip(), "количество не число"))
            else:
                pozicii.append((name, kol))
    return pozicii, oshibki


def v_arhiv(path):
    """Перенести выполненное задание в архив с отметкой времени."""
    os.makedirs(ARHIV_DIR, exist_ok=True)
    bn = os.path.basename(path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(ARHIV_DIR, f"{stamp}__{bn}")
    try:
        os.replace(path, dst)
        return dst
    except OSError:
        return None


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ВВОД В ПРОГРАММУ                              ║
# ╚══════════════════════════════════════════════════════════════════╝

def peremestit_kursor(tek_poz, cel_poz):
    delta = cel_poz - tek_poz
    if delta == 0:
        return
    pg = "pagedown" if delta > 0 else "pageup"
    arr = "down" if delta > 0 else "up"
    shagov = abs(delta)
    for _ in range(shagov // RAZMER_STRANICY):
        pyautogui.press(pg)
    for _ in range(shagov % RAZMER_STRANICY):
        pyautogui.press(arr)


def vvesti_kolichestvo(kol):
    for ch in kol:
        pyautogui.press(ch)
    if KLAVISHA_PODTVERZHDENIYA:
        pyautogui.press(KLAVISHA_PODTVERZHDENIYA)


# ╔══════════════════════════════════════════════════════════════════╗
# ║                  ОБРАБОТКА ОДНОГО ЗАДАНИЯ                           ║
# ╚══════════════════════════════════════════════════════════════════╝

def obrabotat_zadanie(path, tovary):
    """Обработать один файл-задание. Возвращает True, если ввод был выполнен."""
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

    # Сопоставление со базой
    plan, ne_naideno = [], []
    for zapros, kol in pozicii:
        kand = naiti_kandidatov(zapros, tovary)
        if not kand:
            ne_naideno.append((zapros, kol))
            continue
        idx, name, ost, score = kand[0]
        plan.append({"idx": idx, "name": name, "kol": int(kol),
                     "ost": ost, "ambig": len(kand) > 1})

    if SORTIROVAT_PO_STROKE:
        plan.sort(key=lambda x: x["idx"])

    # Показ плана с проверкой остатков
    print(f"\nК вводу: {len(plan)} позиц.")
    print("-" * 64)
    nehvatka = []
    for p in plan:
        pometka = "  [дубль->первый]" if p["ambig"] else ""
        ostatok_info = ""
        if PROVERYAT_OSTATOK:
            posle = p["ost"] - p["kol"]
            ostatok_info = f"  (остаток {p['ost']} -> {posle})"
            if posle < 0:
                ostatok_info += "  !!! НЕ ХВАТАЕТ"
                nehvatka.append(p)
        print(f"  стр {p['idx']+1:>3}: {p['name']:<26} = {p['kol']}{ostatok_info}{pometka}")
    if ne_naideno:
        print("\n  НЕ НАЙДЕНЫ в базе (пропущены):")
        for zapros, kol in ne_naideno:
            print(f"    {zapros} = {kol}")
    if nehvatka:
        print("\n  ВНИМАНИЕ: по учёту базы остатка не хватает на позиции выше.")
        print("  (Ввод разрешён — решение за вами. Остаток может уйти в минус.)")

    if not plan:
        print("Нечего вводить. Задание в архив.")
        v_arhiv(path)
        return False

    # ── ПРОВЕРКА/ПОДТВЕРЖДЕНИЕ КАЖДОЙ ПОЗИЦИИ (пока ОТКЛЮЧЕНО по просьбе) ──
    # plan_ok = []
    # for p in plan:
    #     c = input(f"  Ввести '{p['name']}' = {p['kol']}? (Enter=да, н=пропуск): ").strip().lower()
    #     if c not in ("н","n","нет","no"):
    #         plan_ok.append(p)
    # plan = plan_ok
    # ──────────────────────────────────────────────────────────────────────

    # Стартовая позиция курсора и общее подтверждение
    start = input(f"\nНа какой строке стоит курсор? (1..{len(tovary)}, Enter=1): ").strip()
    tek_poz = (int(start) - 1) if start.isdigit() else 0
    if not (0 <= tek_poz < len(tovary)):
        tek_poz = 0

    go = input(f"Запустить ввод {len(plan)} позиц.? (Enter=да, н=отмена): ").strip().lower()
    if go in ("н", "n", "нет", "no"):
        print("Отменено. Задание остаётся в папке (не в архиве).")
        return False

    print(f"\nПереключитесь в окно старой программы! Старт через:")
    for i in range(OBRATNYI_OTSCHET, 0, -1):
        print(f"  {i}...", flush=True)
        time.sleep(1)
    print("  ПОЕХАЛИ\n")

    # Ввод + учёт остатков + история
    istoriya = []
    teper = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    vvedeno = 0
    try:
        for n, p in enumerate(plan, 1):
            peremestit_kursor(tek_poz, p["idx"])
            vvesti_kolichestvo(str(p["kol"]))
            tek_poz = p["idx"]
            ost_do = tovary[p["idx"]]["ostatok"]
            ost_posle = ost_do - p["kol"]
            tovary[p["idx"]]["ostatok"] = ost_posle   # вычитаем из учёта
            istoriya.append([teper, z_name, p["idx"]+1, p["name"],
                             p["kol"], ost_do, ost_posle])
            vvedeno += 1
            print(f"  [{n}/{len(plan)}] стр {p['idx']+1}: {p['name']} = {p['kol']}  (ост {ost_do}->{ost_posle})  ✓")
            if PAUZA_MEZHDU_POZICIYAMI > 0:
                time.sleep(PAUZA_MEZHDU_POZICIYAMI)
        print("\nЗадание введено.")
    except pyautogui.FailSafeException:
        print("\n  !!! АВАРИЙНЫЙ СТОП (мышь в углу). Ввод прерван. !!!")
        print("  Сохраняю уже введённое в учёт и историю.")

    # Сохраняем учёт и историю даже при частичном вводе
    if istoriya:
        sohranit_bazu(tovary)
        zapisat_istoriyu(istoriya)
        print(f"  Учёт обновлён, в историю записано: {len(istoriya)} строк.")

    # В архив только если задание полностью прошло
    if vvedeno == len(plan):
        dst = v_arhiv(path)
        if dst:
            print(f"  Задание перенесено в архив: {os.path.basename(dst)}")
    else:
        print("  Задание выполнено НЕ полностью — оставляю файл в папке.")
    return vvedeno > 0


# ╔══════════════════════════════════════════════════════════════════╗
# ║                       ГЛАВНЫЙ ЦИКЛ                                  ║
# ╚══════════════════════════════════════════════════════════════════╝

def main():
    print("=" * 64)
    print("  ПОМОЩНИК ВВОДА ПРОДУКЦИИ  v0.3  (цикл + база + история)")
    print("=" * 64)
    print("Аварийный СТОП: резко уведите мышь в ЛЕВЫЙ ВЕРХНИЙ угол экрана.")
    print(f"База товаров: {TOVARY_CSV}")
    print(f"Задания:      {ZADANIYA_DIR}")
    print()

    tovary = zagruzit_bazu()
    if tovary is None:
        print(f"Не найдена база товаров: {TOVARY_CSV}")
        print("Создайте папку 'baza' рядом с программой и положите туда tovary.csv")
        print("со столбцами: pozitsiya,naimenovanie,ostatok")
        return
    print(f"Загружена база: {len(tovary)} товаров.")

    os.makedirs(ZADANIYA_DIR, exist_ok=True)
    print("\nРежим ЦИКЛА. Кладите файлы-задания (.txt) в папку zadaniya.")
    print("Команды: Enter — проверить папку сейчас;  'выход' — завершить.\n")

    while True:
        zadaniya = spisok_zadaniy()
        if not zadaniya:
            try:
                cmd = input("Ожидание задания... (Enter=проверить, 'выход'): ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                break
            if cmd in ("выход", "exit", "quit", "q"):
                break
            # короткая пауза перед повторной проверкой
            time.sleep(OPROS_PAPKI_SEK)
            continue

        # Есть задания — обрабатываем по очереди (перезагружаем базу с диска,
        # чтобы подхватить ручные правки остатков между заданиями).
        tovary = zagruzit_bazu() or tovary
        for path in zadaniya:
            if not os.path.isfile(path):
                continue
            obrabotat_zadanie(path, tovary)

        print("\nВсе текущие задания обработаны. Жду новые.\n")

    print("\nЗавершение работы.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nПрервано пользователем.")
    except Exception:
        import traceback
        print("\n!!! ОШИБКА !!!")
        traceback.print_exc()
    finally:
        try:
            input("\n--- Окно не закроется. Нажмите Enter для выхода. ---")
        except EOFError:
            pass
