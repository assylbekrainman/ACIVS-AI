#!/usr/bin/env python3
"""
Черновая классификация файлов проектной подпапки по чек-листу
sz-fed-perechislenie / perechislenie-tracker.

Это НЕ финальный вердикт — только быстрая первая раскладка по ключевым
словам в имени файла, чтобы не читать вручную десятки имён в каждой
подпапке. Всегда проверяй "unclassified" и спорные случаи сам: живые
названия файлов пишут по-разному (опечатки, "счет"/"счёт", кальки с
казахского и т.д.), regex такое пропускает.

Использование:
    python3 classify_files.py <путь_к_папке_проекта> [<путь_к_другой_папке> ...]

Печатает JSON построчно (один объект на папку) со полями:
    folder, checklist (какие из 8 пунктов найдены -> список файлов),
    unclassified (файлы, которые не удалось уверенно отнести ни к одному пункту)
"""
import json
import os
import re
import sys

# Порядок и ключевые слова совпадают с чек-листом ШАГ 0 sz-fed-perechislenie /
# ШАГ 2 perechislenie-tracker. Ключевые слова — по-русски и по-казахски,
# без учёта регистра; ё нормализуется в е перед сравнением.
CHECKLIST = [
    (1, "Счёт на оплату", [r"счет.*оплат", r"счёт.*оплат", r"\bинвойс\b"]),
    (2, "Договор эскроу-счёта", [r"эскроу.*догов", r"догов.*эскроу"]),
    (3, "Справка о реквизитах/наличии эскроу-счёта", [r"справк.*эскроу", r"справк.*счет", r"справк.*счёт"]),
    (4, "Выписка по счёту софинансирования", [r"выписк.*софин", r"софин.*выписк"]),
    (5, "Свидетельство НДС", [r"ндс"]),
    (6, "Смета расходов / ДС со сметой", [r"смет", r"\bдс\s*№?\s*\d"]),
    (7, "Уведомление о смене реквизитов", [r"уведомлен"]),
    (8, "СЗ в ФЭД", [r"сз.*фэд", r"фэд.*сз", r"служебн.*запи?ск"]),
]


def normalize(name: str) -> str:
    return name.lower().replace("ё", "е")


def classify_folder(path: str) -> dict:
    result = {"folder": path, "checklist": {str(n): [] for n, _, _ in CHECKLIST}, "unclassified": []}
    try:
        entries = sorted(os.listdir(path))
    except OSError as e:
        result["error"] = str(e)
        return result

    for entry in entries:
        full = os.path.join(path, entry)
        if os.path.isdir(full):
            continue  # подпапки (например, "N этап") обрабатывай отдельным вызовом
        norm = normalize(entry)
        matched = False
        for num, _label, patterns in CHECKLIST:
            if any(re.search(p, norm) for p in patterns):
                result["checklist"][str(num)].append(entry)
                matched = True
                break  # первый совпавший пункт — большинство имён однозначны
        if not matched:
            result["unclassified"].append(entry)
    return result


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 1
    for folder in argv[1:]:
        print(json.dumps(classify_folder(folder), ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
