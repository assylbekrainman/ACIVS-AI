#!/usr/bin/env python3
"""
Проверка аффилированности поставщика с грантополучателем через внутренний
реестр проектных команд Фонда науки (references/team_roster.csv,
2022-2025 гг., без ИИН — только ФИО/ИРН проекта/роль/год).

Логика: если учредитель, руководитель или подписант поставщика — то же лицо,
что и член ПРОЕКТНОЙ ГРУППЫ ГП (руководитель проекта, специалист по
коммерциализации и т.д.), это self-dealing — Фонд науки платит поставщику,
которым фактически владеет/управляет тот же человек, что заявлен в проекте.

НЕ заменяет проверку по открытым реестрам (adata.kz, kompra.kz, egov.kz) —
это ДОПОЛНИТЕЛЬНЫЙ внутренний источник, который открытые реестры не видят
(проектная группа — это не всегда учредитель/директор юрлица в устав. Данных).

Использование:
    python3 check_affiliation.py "Иванов Иван Иванович"
    python3 check_affiliation.py "Иванов Иван Иванович" --irn DP12345678
    python3 check_affiliation.py "Иванов И.И." "Петров П.П." --irn AP98765432

Сравнение — по нормализованным токенам (без учёта регистра/порядка слов),
плюс мягкое совпадение (difflib) на случай опечаток/сокращений инициалов.
Точное поспадение по всем 3 токенам ФИО — HIGH; частичное (2 из 3, например
из-за отчества/инициалов) — MEDIUM, показывается с пометкой "проверить
вручную". Скрипт ничего не решает сам — он только показывает совпадения,
финальную оценку аффилированности делает человек.
"""

import argparse
import csv
import difflib
import os
import sys

ROSTER_PATH = os.path.join(os.path.dirname(__file__), "..", "references", "team_roster.csv")


def normalize(name):
    return [tok.strip(".,").lower() for tok in name.replace("ё", "е").split() if tok.strip(".,")]


def load_roster(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def score_match(query_tokens, roster_tokens):
    """Возвращает (совпавшие_токены, оценка) — простая эвристика, не ML."""
    matched = 0
    for qt in query_tokens:
        for rt in roster_tokens:
            if qt == rt:
                matched += 1
                break
            if difflib.SequenceMatcher(None, qt, rt).ratio() > 0.85:
                matched += 1
                break
    total = max(len(query_tokens), len(roster_tokens))
    return matched, total


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("names", nargs="+", help="ФИО для проверки (одно или несколько, каждое в кавычках)")
    ap.add_argument("--irn", help="Ограничить поиск конкретным проектом (ИРН/DP-код), иначе ищет по всей базе 2022-2025")
    ap.add_argument("--roster", default=ROSTER_PATH, help="Путь к team_roster.csv (по умолчанию — references/ рядом со скриптом)")
    args = ap.parse_args()

    if not os.path.exists(args.roster):
        print(f"ОШИБКА: не найден реестр {args.roster}", file=sys.stderr)
        sys.exit(1)

    roster = load_roster(args.roster)
    if args.irn:
        roster = [r for r in roster if r["ирн_проекта"].strip().upper() == args.irn.strip().upper()]
        if not roster:
            print(f"В реестре нет ни одной записи по проекту {args.irn} — "
                  f"либо опечатка в коде, либо проект не входит в 2022-2025 гг. "
                  f"по этому файлу. Сверка по этому источнику невозможна, скажи об этом прямо.")
            return

    any_match = False
    for name in args.names:
        q_tokens = normalize(name)
        hits = []
        for row in roster:
            r_tokens = normalize(row["фио"])
            matched, total = score_match(q_tokens, r_tokens)
            if matched >= 2:
                level = "HIGH" if matched == total and matched >= 3 else "MEDIUM (проверить вручную)"
                hits.append((level, row))

        print(f"\n=== Запрос: «{name}» ===")
        if not hits:
            print("  Совпадений в реестре проектных команд не найдено.")
            continue
        any_match = True
        for level, row in hits:
            grantee = f", ГП: {row['грантополучатель']}" if row.get("грантополучатель") else ""
            print(f"  [{level}] {row['фио']} — проект {row['ирн_проекта']} ({row['год']}), "
                  f"роль: {row['роль_в_проекте']}, статус: {row['статус']}{grantee}")

    if any_match:
        print("\nЕсть совпадения — это ПОВОД проверить вручную (учредитель/директор/"
              "подписант поставщика = член проектной группы?), а не автоматический "
              "красный флаг. Имена в казахстанских документах часто совпадают у "
              "однофамильцев — не делай вывод об аффилированности по одному "
              "совпадению ФИО без второго подтверждения (ИИН, дата рождения, адрес).")


if __name__ == "__main__":
    main()
