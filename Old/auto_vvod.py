# -*- coding: utf-8 -*-
"""
auto_vvod.py — помощник для ввода количества продукции в старую DOS/Oracle-программу.

ПРИНЦИП РАБОТЫ ("вслепую по списку"):
  Скрипт хранит у себя список наименований В ТОМ ЖЕ ПОРЯДКЕ, что на экране.
  Он помнит, на какой строке сейчас стоит курсор (подсветка).
  Вы вводите название товара и количество.
  Скрипт находит товар в списке, считает на сколько строк нажать вверх/вниз,
  жмёт стрелки, затем набирает цифры количества.

ВАЖНО: скрипт НЕ видит экран. Он доверяет своему списку и счётчику позиции.
  Поэтому перед запуском убедитесь, что:
    1) Список в файле .txt совпадает по порядку с тем, что на экране.
    2) Курсор в программе стоит на ПЕРВОЙ строке списка (или укажите стартовую строку).

Автор: подготовлено для оператора ЭВМ (Слонимский мясокомбинат).
"""

import sys
import os
import time
import difflib
import pyautogui
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

# Папка, где лежат файлы списков (.txt). По умолчанию — рядом со скриптом.
SPISKI_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../spiski")

# Пауза между нажатиями клавиш (секунды). Старые программы тормозят —
# если стрелки "пролетают" мимо, увеличьте значение (например 0.08 или 0.12).
PAUZA_KLAVISHA = 0.05

# Пауза перед стартом нажатий после подтверждения (чтобы вы успели переключиться
# в окно старой программы, если потребуется). Обычно фокус уже там.
PAUZA_PERED_VVODOM = 0.0

# Что нажимать ПОСЛЕ набора цифр количества, чтобы зафиксировать ввод.
# Варианты: "" (ничего), "enter", "f2", "tab", "down".
# Вы говорили "просто нажатием цифр" — оставлено пустым. Если нужно
# подтверждение — впишите, например, "enter".
KLAVISHA_PODTVERZHDENIYA = ""

# Спрашивать подтверждение (Enter) перед каждым реальным вводом в программу.
# True — безопаснее (рекомендуется на старте). False — быстрее.
PODTVERZHDAT_PERED_VVODOM = True

# Порог нечёткого поиска (0..1). Чем выше — тем строже совпадение.
PORICHE_POISKA = 0.55

# ── Защита pyautogui ──
pyautogui.PAUSE = PAUZA_KLAVISHA
# FAILSAFE: резко увести мышь в ЛЕВЫЙ ВЕРХНИЙ угол экрана — аварийный СТОП.
pyautogui.FAILSAFE = True


# ╔══════════════════════════════════════════════════════════════════╗
# ║                         РАБОТА СО СПИСКАМИ                          ║
# ╚══════════════════════════════════════════════════════════════════╝

def normalize(s: str) -> str:
    """Привести строку к виду для сравнения: нижний регистр, убрать лишние пробелы."""
    return " ".join(s.lower().replace("/", " ").split())


def zagruzit_spiski(directory: str) -> dict:
    """Загрузить все .txt-списки из папки. Возвращает {имя_файла: [строки]}."""
    spiski = {}
    if not os.path.isdir(directory):
        return spiski
    for fname in sorted(os.listdir(directory)):
        if fname.lower().endswith(".txt"):
            path = os.path.join(directory, fname)
            with open(path, encoding="utf-8") as f:
                lines = [ln.rstrip("\n") for ln in f]
            # Пустые строки сохраняем как позиции (они тоже занимают строку на экране!),
            # но строки-комментарии (начинаются с #) пропускаем.
            lines = [ln for ln in lines if not ln.lstrip().startswith("#")]
            spiski[fname[:-4]] = lines
    return spiski


def vybrat_spisok(spiski: dict):
    """Дать пользователю выбрать активный список."""
    if not spiski:
        print(f"В папке '{SPISKI_DIR}' нет ни одного файла-списка (.txt).")
        print("Создайте файл, например 'kolbasnyi.txt', и впишите наименования по одному в строке,")
        print("в ТОМ ЖЕ порядке, что на экране программы.")
        sys.exit(1)
    names = list(spiski.keys())
    if len(names) == 1:
        print(f"Загружен список: {names[0]} ({len(spiski[names[0]])} позиций)")
        return names[0]
    print("\nДоступные списки:")
    for i, n in enumerate(names, 1):
        print(f"  {i}. {n} ({len(spiski[n])} позиций)")
    while True:
        ch = input("Выберите номер списка: ").strip()
        if ch.isdigit() and 1 <= int(ch) <= len(names):
            return names[int(ch) - 1]
        print("Неверный номер, попробуйте ещё раз.")


def naiti_kandidatov(zapros: str, spisok: list):
    """
    Найти позиции в списке, подходящие под запрос.
    Возвращает список кортежей (индекс, наименование, оценка) по убыванию оценки.
    Учитывает дубли — все совпадения возвращаются.
    """
    nz = normalize(zapros)
    rezultaty = []
    for idx, name in enumerate(spisok):
        nn = normalize(name)
        if not nn:
            continue
        # Точное вхождение подстроки даёт высокий приоритет.
        if nz in nn:
            score = 0.9 + 0.1 * (len(nz) / max(len(nn), 1))
        # Совпадение по началу какого-либо слова (напр. 'чеснок' -> 'чесночная').
        elif len(nz) >= 4 and any(w.startswith(nz) for w in nn.split()):
            score = 0.88
        else:
            score = difflib.SequenceMatcher(None, nz, nn).ratio()
        if score >= PORICHE_POISKA:
            rezultaty.append((idx, name, score))
    rezultaty.sort(key=lambda x: x[2], reverse=True)
    return rezultaty


# ╔══════════════════════════════════════════════════════════════════╗
# ║                     ВВОД В ПРОГРАММУ (КЛАВИАТУРА)                   ║
# ╚══════════════════════════════════════════════════════════════════╝

def peremestit_kursor(tek_poz: int, cel_poz: int):
    """Нажать стрелки вверх/вниз, чтобы дойти от tek_poz до cel_poz."""
    delta = cel_poz - tek_poz
    if delta == 0:
        print("  Курсор уже на нужной строке.")
        return
    klavisha = "down" if delta > 0 else "up"
    shagov = abs(delta)
    napr = "вниз" if delta > 0 else "вверх"
    print(f"  Перемещение: {shagov} раз(а) {napr}...")
    for _ in range(shagov):
        pyautogui.press(klavisha)


def vvesti_kolichestvo(kol: str):
    """Набрать цифры количества и (опционально) клавишу подтверждения."""
    print(f"  Ввод количества: {kol}")
    for ch in kol:
        pyautogui.press(ch)
    if KLAVISHA_PODTVERZHDENIYA:
        print(f"  Подтверждение: {KLAVISHA_PODTVERZHDENIYA}")
        pyautogui.press(KLAVISHA_PODTVERZHDENIYA)


# ╔══════════════════════════════════════════════════════════════════╗
# ║                            ГЛАВНЫЙ ЦИКЛ                             ║
# ╚══════════════════════════════════════════════════════════════════╝

def main():
    print("=" * 64)
    print("  ПОМОЩНИК ВВОДА ПРОДУКЦИИ  (режим: вслепую по списку)")
    print("=" * 64)
    print("Аварийный СТОП в любой момент: резко уведите мышь в ЛЕВЫЙ")
    print("ВЕРХНИЙ угол экрана (сработает защита pyautogui).")
    print()

    spiski = zagruzit_spiski(SPISKI_DIR)
    aktiv = vybrat_spisok(spiski)
    spisok = spiski[aktiv]

    # Текущая позиция курсора в программе (индекс строки, 0 = первая).
    start = input(f"\nНа какой строке сейчас стоит курсор? (1..{len(spisok)}, Enter = 1): ").strip()
    tek_poz = (int(start) - 1) if start.isdigit() else 0
    if tek_poz < 0 or tek_poz >= len(spisok):
        tek_poz = 0
    print(f"Стартовая позиция курсора: строка {tek_poz + 1} — '{spisok[tek_poz]}'")

    print("\nГотово. Вводите запросы. Команды:")
    print("  - название;количество   (например: Докторская;15)")
    print("  - можно через пробел:    Докторская 15")
    print("  - 'список'  — сменить активный список")
    print("  - 'поз N'   — вручную задать текущую позицию курсора")
    print("  - 'выход'   — завершить")
    print("-" * 64)

    while True:
        try:
            raw = input("\nЗапрос > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nЗавершение.")
            break

        if not raw:
            continue
        low = raw.lower()
        if low in ("выход", "exit", "quit", "q"):
            print("Завершение.")
            break
        if low == "список":
            aktiv = vybrat_spisok(spiski)
            spisok = spiski[aktiv]
            tek_poz = 0
            print(f"Активный список: {aktiv}. Курсор сброшен на строку 1.")
            continue
        if low.startswith("поз"):
            parts = low.split()
            if len(parts) == 2 and parts[1].isdigit():
                n = int(parts[1])
                if 1 <= n <= len(spisok):
                    tek_poz = n - 1
                    print(f"Текущая позиция установлена: строка {n} — '{spisok[tek_poz]}'")
                else:
                    print("Номер вне диапазона.")
            else:
                print("Формат: поз N  (например: поз 5)")
            continue

        # Разбор "название;количество" или "название количество"
        if ";" in raw:
            name_part, _, kol_part = raw.partition(";")
        else:
            toks = raw.rsplit(" ", 1)
            if len(toks) == 2 and toks[1].strip().isdigit():
                name_part, kol_part = toks[0], toks[1]
            else:
                print("Не вижу количества. Формат: Название;Количество  или  Название Количество")
                continue

        name_part = name_part.strip()
        kol_part = kol_part.strip()
        if not kol_part.isdigit():
            print("Количество должно быть числом.")
            continue

        kandidaty = naiti_kandidatov(name_part, spisok)
        if not kandidaty:
            print(f"  Ничего похожего на '{name_part}' не найдено в списке '{aktiv}'.")
            continue

        # Обработка нескольких совпадений (дубли или похожие названия)
        if len(kandidaty) > 1:
            print(f"  Найдено несколько совпадений для '{name_part}':")
            for i, (idx, name, score) in enumerate(kandidaty, 1):
                print(f"    {i}. строка {idx + 1}: {name}  (схожесть {score:.0%})")
            ch = input("  Какой нужен? (номер, Enter = первый, 'отмена'): ").strip().lower()
            if ch == "отмена":
                continue
            if ch == "":
                vybor = kandidaty[0]
            elif ch.isdigit() and 1 <= int(ch) <= len(kandidaty):
                vybor = kandidaty[int(ch) - 1]
            else:
                print("  Неверный выбор, пропускаю.")
                continue
        else:
            vybor = kandidaty[0]

        cel_idx, cel_name, score = vybor
        print(f"  Цель: строка {cel_idx + 1} — '{cel_name}', количество {kol_part}")

        if PODTVERZHDAT_PERED_VVODOM:
            c = input("  Подтвердить ввод? (Enter = да, 'н' = нет): ").strip().lower()
            if c in ("н", "n", "нет", "no"):
                print("  Отменено.")
                continue

        if PAUZA_PERED_VVODOM > 0:
            print(f"  Старт через {PAUZA_PERED_VVODOM} c... (переключитесь в окно программы)")
            time.sleep(PAUZA_PERED_VVODOM)

        # Выполняем ввод
        try:
            peremestit_kursor(tek_poz, cel_idx)
            vvesti_kolichestvo(kol_part)
            tek_poz = cel_idx  # обновляем текущую позицию
            print(f"  Готово. Курсор теперь на строке {tek_poz + 1}.")
        except pyautogui.FailSafeException:
            print("\n  !!! АВАРИЙНЫЙ СТОП (мышь в углу экрана). Ввод прерван. !!!")
            print("  ВНИМАНИЕ: позиция курсора могла сбиться — проверьте экран и")
            print("  при необходимости задайте её заново командой 'поз N'.")
            break


if __name__ == "__main__":
    main()
