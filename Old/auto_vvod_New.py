# -*- coding: utf-8 -*-
"""
auto_vvod.py — помощник для ввода количества продукции в старую DOS/Oracle-программу.

РЕЖИМ: ПАКЕТНЫЙ (по заданию из файла).
  Вы заранее готовите файл-задание (например zadanie.txt) со строками вида
      Название;Количество
  Скрипт берёт справочник продукции (порядок строк = порядок на экране),
  идёт по заданию сверху вниз, для каждой позиции:
    - находит товар в справочнике,
    - перемещает курсор стрелками вверх/вниз от текущей позиции,
    - набирает цифры количества.
  Курсор после ввода остаётся на той же строке (так работает ваша программа),
  поэтому скрипт считает следующий шаг от только что введённой строки.

ВАЖНО: скрипт НЕ видит экран. Он доверяет справочнику и счётчику позиции.
  Перед запуском:
    1) Справочник (.txt в папке spiski) совпадает по порядку с экраном.
    2) Курсор в программе стоит на известной строке (по умолчанию — первой).

Автор: подготовлено для оператора ЭВМ (Слонимский мясокомбинат).
"""

import sys
import os
import time
import difflib

# ── Зависимость: pyautogui (эмуляция клавиатуры). Установка: pip install pyautogui
try:
    import pyautogui
except ImportError:
    print("Не установлена библиотека pyautogui.")
    print("Установите её командой:  pip install pyautogui")
    sys.exit(1)

# ╔══════════════════════════════════════════════════════════════════╗
# ║                       НАСТРОЙКИ (меняйте под себя)                  ║
# ╚══════════════════════════════════════════════════════════════════╝

# Папка со справочниками продукции (.txt). По умолчанию — рядом со скриптом.
SPISKI_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../spiski")

# Папка с файлами-заданиями (.txt). По умолчанию — рядом со скриптом.
ZADANIYA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../zadaniya")

# Обратный отсчёт ПЕРЕД началом ввода (секунды). За это время вы должны
# успеть кликнуть в окно старой программы. Увеличьте, если не успеваете.
OBRATNYI_OTSCHET = 5

# Пауза между нажатиями клавиш (секунды). Старые программы тормозят —
# если стрелки "пролетают" мимо, увеличьте значение (например 0.08 или 0.12).
PAUZA_KLAVISHA = 0.05

# Пауза между позициями задания (секунды). Даёт программе "перевести дух".
PAUZA_MEZHDU_POZICIYAMI = 0.15

# На сколько строк прыгают PageDown/PageUp в вашей программе.
RAZMER_STRANICY = 17

# Что нажимать ПОСЛЕ набора цифр количества, чтобы зафиксировать ввод.
# Варианты: "" (ничего), "enter", "f2", "tab", "down".
# Вы говорили "просто нажатием цифр" — оставлено пустым.
KLAVISHA_PODTVERZHDENIYA = ""

# Порог нечёткого поиска (0..1). Чем выше — тем строже совпадение.
PORICHE_POISKA = 0.55

# ── Защита pyautogui ──
pyautogui.PAUSE = PAUZA_KLAVISHA
# FAILSAFE: резко увести мышь в ЛЕВЫЙ ВЕРХНИЙ угол экрана — аварийный СТОП.
pyautogui.FAILSAFE = True


# ╔══════════════════════════════════════════════════════════════════╗
# ║                     СПРАВОЧНИК И ПОИСК                              ║
# ╚══════════════════════════════════════════════════════════════════╝

def normalize(s: str) -> str:
    """Привести строку к виду для сравнения: нижний регистр, убрать лишние пробелы."""
    return " ".join(s.lower().replace("/", " ").split())


def zagruzit_txt_spiski(directory: str) -> dict:
    """Загрузить все .txt-файлы из папки. Возвращает {имя_файла: [строки]}.
    Строки-комментарии (начинаются с #) пропускаются."""
    rezultat = {}
    if not os.path.isdir(directory):
        return rezultat
    for fname in sorted(os.listdir(directory)):
        if fname.lower().endswith(".txt"):
            path = os.path.join(directory, fname)
            with open(path, encoding="utf-8") as f:
                lines = [ln.rstrip("\n") for ln in f]
            lines = [ln for ln in lines if not ln.lstrip().startswith("#")]
            rezultat[fname[:-4]] = lines
    return rezultat


def vybrat_iz(slovar: dict, chto: str):
    """Дать пользователю выбрать один элемент из словаря {имя: данные}."""
    names = list(slovar.keys())
    if len(names) == 1:
        print(f"{chto}: {names[0]} ({len(slovar[names[0]])} строк)")
        return names[0]
    print(f"\nДоступные {chto}:")
    for i, n in enumerate(names, 1):
        print(f"  {i}. {n} ({len(slovar[n])} строк)")
    while True:
        ch = input(f"Выберите номер ({chto}): ").strip()
        if ch.isdigit() and 1 <= int(ch) <= len(names):
            return names[int(ch) - 1]
        print("Неверный номер, попробуйте ещё раз.")


def naiti_kandidatov(zapros: str, spravochnik: list):
    """Найти позиции справочника, подходящие под запрос.
    Возвращает список (индекс, наименование, оценка) по убыванию оценки.
    Все дубли возвращаются."""
    nz = normalize(zapros)
    rezultaty = []
    for idx, name in enumerate(spravochnik):
        nn = normalize(name)
        if not nn:
            continue
        if nz in nn:
            score = 0.9 + 0.1 * (len(nz) / max(len(nn), 1))
        elif len(nz) >= 4 and any(w.startswith(nz) for w in nn.split()):
            score = 0.88
        else:
            score = difflib.SequenceMatcher(None, nz, nn).ratio()
        if score >= PORICHE_POISKA:
            rezultaty.append((idx, name, score))
    rezultaty.sort(key=lambda x: x[2], reverse=True)
    return rezultaty


# ╔══════════════════════════════════════════════════════════════════╗
# ║                     РАЗБОР ФАЙЛА-ЗАДАНИЯ                            ║
# ╚══════════════════════════════════════════════════════════════════╝

def razobrat_zadanie(stroki: list):
    """Разобрать строки задания в список (название, количество).
    Поддерживает 'Название;Количество' и 'Название Количество'.
    Возвращает (позиции, ошибки)."""
    pozicii = []
    oshibki = []
    for nomer, raw in enumerate(stroki, 1):
        s = raw.strip()
        if not s:
            continue
        if ";" in s:
            name, _, kol = s.partition(";")
        else:
            toks = s.rsplit(" ", 1)
            if len(toks) == 2 and toks[1].strip().isdigit():
                name, kol = toks[0], toks[1]
            else:
                oshibki.append((nomer, raw, "нет количества"))
                continue
        name = name.strip()
        kol = kol.strip()
        if not name:
            oshibki.append((nomer, raw, "пустое название"))
        elif not kol.isdigit():
            oshibki.append((nomer, raw, "количество не число"))
        else:
            pozicii.append((name, kol))
    return pozicii, oshibki


# ╔══════════════════════════════════════════════════════════════════╗
# ║                     ВВОД В ПРОГРАММУ (КЛАВИАТУРА)                   ║
# ╚══════════════════════════════════════════════════════════════════╝

def peremestit_kursor(tek_poz: int, cel_poz: int):
    """Нажать стрелки вверх/вниз, чтобы дойти от tek_poz до cel_poz."""
    delta = cel_poz - tek_poz
    if delta == 0:
        return
    pg = "pagedown" if delta > 0 else "pageup"
    klavisha = "down" if delta > 0 else "up"
    shagov = abs(delta)
# Крупные прыжки страницами, затем остаток стрелками.
    for _ in range(shagov // RAZMER_STRANICY):
        pyautogui.press(pg)
    for _ in range(shagov % RAZMER_STRANICY):
        pyautogui.press(klavisha)


def vvesti_kolichestvo(kol: str):
    """Набрать цифры количества и (опционально) клавишу подтверждения."""
    for ch in kol:
        pyautogui.press(ch)
    if KLAVISHA_PODTVERZHDENIYA:
        pyautogui.press(KLAVISHA_PODTVERZHDENIYA)


# ╔══════════════════════════════════════════════════════════════════╗
# ║                            ГЛАВНАЯ ЛОГИКА                           ║
# ╚══════════════════════════════════════════════════════════════════╝

def main():
    print("=" * 64)
    print("  ПОМОЩНИК ВВОДА ПРОДУКЦИИ  (ПАКЕТНЫЙ режим, по заданию)")
    print("=" * 64)
    print("Аварийный СТОП в любой момент: резко уведите мышь в ЛЕВЫЙ")
    print("ВЕРХНИЙ угол экрана.")
    print()

    # 1) Справочник продукции
    spravochniki = zagruzit_txt_spiski(SPISKI_DIR)
    if not spravochniki:
        print(f"В папке '{SPISKI_DIR}' нет ни одного справочника (.txt).")
        sys.exit(1)
    s_name = vybrat_iz(spravochniki, "справочники")
    spravochnik = spravochniki[s_name]

    # 2) Файл-задание
    zadaniya = zagruzit_txt_spiski(ZADANIYA_DIR)
    if not zadaniya:
        print(f"\nВ папке '{ZADANIYA_DIR}' нет ни одного задания (.txt).")
        print("Создайте файл, например 'zadanie.txt', и впишите строки вида:")
        print("    Название;Количество")
        print("по одной позиции в строке.")
        sys.exit(1)
    z_name = vybrat_iz(zadaniya, "задания")
    pozicii, oshibki = razobrat_zadanie(zadaniya[z_name])

    if oshibki:
        print("\nВ задании есть строки, которые не удалось разобрать:")
        for nomer, raw, prichina in oshibki:
            print(f"  строка {nomer}: {raw!r}  — {prichina}")

    if not pozicii:
        print("\nВ задании нет ни одной корректной позиции. Завершение.")
        sys.exit(1)

    # 3) Сопоставляем каждую позицию задания со справочником.
    #    Для дублей берём первое (лучшее) совпадение автоматически.
    #    Несопоставленные и неоднозначные — показываем.
    plan = []          # список (cel_idx, cel_name, kol, zapros)
    ne_naideno = []
    neodnoznachno = []
    for zapros, kol in pozicii:
        kand = naiti_kandidatov(zapros, spravochnik)
        if not kand:
            ne_naideno.append((zapros, kol))
            continue
        if len(kand) > 1:
            neodnoznachno.append((zapros, kol, kand))
        plan.append((kand[0][0], kand[0][1], kol, zapros))

    # 4) Показываем план целиком для предварительной проверки.
    print("\n" + "-" * 64)
    print(f"ЗАДАНИЕ '{z_name}' по справочнику '{s_name}'.")
    print(f"Готово к вводу: {len(plan)} позиц.")
    print("-" * 64)
    for cel_idx, cel_name, kol, zapros in plan:
        pometka = ""
        if any(z == zapros and k == kol for z, k, _ in neodnoznachno):
            pometka = "  [было несколько совпадений — взято первое]"
        print(f"  строка {cel_idx + 1:>3}: {cel_name:<28} = {kol}{pometka}")
    if ne_naideno:
        print("\n  НЕ НАЙДЕНЫ в справочнике (будут пропущены):")
        for zapros, kol in ne_naideno:
            print(f"    {zapros} = {kol}")

    # ── ПРОВЕРКА/ПОДТВЕРЖДЕНИЕ КАЖДОЙ ПОЗИЦИИ (пока ОТКЛЮЧЕНО по просьбе) ──
    # Чтобы включить пошаговое подтверждение каждой строки, раскомментируйте
    # блок ниже и используйте переменную plan_podtverzhdennyi вместо plan.
    #
    # plan_podtverzhdennyi = []
    # for cel_idx, cel_name, kol, zapros in plan:
    #     c = input(f"  Ввести '{cel_name}' = {kol}? (Enter=да, н=пропустить): ").strip().lower()
    #     if c in ("н", "n", "нет", "no"):
    #         print("    пропущено")
    #         continue
    #     plan_podtverzhdennyi.append((cel_idx, cel_name, kol, zapros))
    # plan = plan_podtverzhdennyi
    # ──────────────────────────────────────────────────────────────────────

    if not plan:
        print("\nНечего вводить. Завершение.")
        sys.exit(0)

    # 5) Одно общее подтверждение запуска всего пакета.
    print("\n" + "-" * 64)
    start = input(f"На какой строке сейчас стоит курсор? (1..{len(spravochnik)}, Enter=1): ").strip()
    tek_poz = (int(start) - 1) if start.isdigit() else 0
    if tek_poz < 0 or tek_poz >= len(spravochnik):
        tek_poz = 0
    print(f"Стартовая позиция курсора: строка {tek_poz + 1} — '{spravochnik[tek_poz]}'")

    go = input(f"\nЗапустить ввод {len(plan)} позиц.? (Enter=да, н=отмена): ").strip().lower()
    if go in ("н", "n", "нет", "no"):
        print("Отменено.")
        sys.exit(0)

    # 6) Обратный отсчёт — успеть переключиться в окно программы.
    print(f"\nПереключитесь в окно старой программы! Старт через:")
    for i in range(OBRATNYI_OTSCHET, 0, -1):
        print(f"  {i}...", flush=True)
        time.sleep(1)
    print("  ПОЕХАЛИ\n")

    # 7) Ввод по плану.
    try:
        for n, (cel_idx, cel_name, kol, zapros) in enumerate(plan, 1):
            peremestit_kursor(tek_poz, cel_idx)
            vvesti_kolichestvo(kol)
            tek_poz = cel_idx  # курсор остаётся на введённой строке
            print(f"  [{n}/{len(plan)}] строка {cel_idx + 1}: {cel_name} = {kol}  ✓")
            if PAUZA_MEZHDU_POZICIYAMI > 0:
                time.sleep(PAUZA_MEZHDU_POZICIYAMI)
        print("\nГотово. Всё задание введено.")
    except pyautogui.FailSafeException:
        print("\n  !!! АВАРИЙНЫЙ СТОП (мышь в углу экрана). Ввод прерван. !!!")
        print("  ВНИМАНИЕ: позиция курсора могла сбиться — проверьте экран.")


if __name__ == "__main__":
    main()
