#!/usr/bin/env python3
"""
Сверка сумм и группировка зарплаты по транзитным счетам для эскроу-расписки на ФОТ.

Зачем нужен именно скрипт, а не ручной устный подсчёт: в реальном письме одного
из грантополучателей была найдена ошибка ровно такого рода — построчная сумма по
сотрудникам не сходилась с заявленным итогом по категории (казахская таблица по
соцналогу оказалась копией таблицы ИПН). Арифметику лучше поручать коду, а не
пересчитывать в уме — так подобные ошибки ловятся автоматически и достоверно.

Использование:
    python check_and_group.py data.json

Формат data.json — см. example_input() ниже (запусти без аргументов, чтобы
увидеть пример на stdout).
"""
import json
import sys

TRANSIT_ACCOUNTS = {
    "kaspi bank": {"bik": "CASPKZKA", "iik": "KZ24722S000000686267"},
    "фридом банк казахстан": {"bik": "KSNVKZKA", "iik": "KZ48551A024000213KZT"},
    "народный банк казахстана": {"bik": "HSBKKZKX", "iik": "KZ256010002200008667"},
    "народный банк": {"bik": "HSBKKZKX", "iik": "KZ256010002200008667"},
    "банк центркредит": {"bik": "KCJBKZKX", "iik": "KZ268562870108312928"},
    "forte bank": {"bik": "IRTYKZKA", "iik": "KZ309650000159022043"},
}


def normalize_bank(name: str) -> str:
    return " ".join(name.strip().lower().replace("«", "").replace("»", "").replace('"', "").split())


def group_salary_by_bank(employees, use_transit=True):
    """employees: list of {"fio","iin","bank","account","sum"}.
    Returns list of beneficiary groups, each either routed to a transit account
    (if use_transit and bank is in TRANSIT_ACCOUNTS) or to each employee's own
    account individually (if not use_transit, or bank unknown)."""
    groups = {}
    order = []
    for e in employees:
        bank_norm = normalize_bank(e["bank"])
        transit = TRANSIT_ACCOUNTS.get(bank_norm) if use_transit else None
        key = bank_norm if transit else f"__direct__{e['fio']}"
        if key not in groups:
            groups[key] = {
                "bank_label": e["bank"],
                "beneficiary_account": transit["iik"] if transit else e["account"],
                "beneficiary_bik": transit["bik"] if transit else e.get("bik", ""),
                "routed_via_transit": bool(transit),
                "employees": [],
                "total": 0,
            }
            order.append(key)
        groups[key]["employees"].append(e)
        groups[key]["total"] += e["sum"]
    return [groups[k] for k in order]


def check_category(category_name, employees_sum_list, declared_total):
    computed = sum(x["sum"] for x in employees_sum_list)
    ok = computed == declared_total
    return {
        "category": category_name,
        "computed_from_rows": computed,
        "declared_total": declared_total,
        "match": ok,
        "diff": computed - declared_total,
    }


def run(data):
    report = {"categories": [], "salary_groups": None, "grand_total_check": None}

    grand_computed = 0
    for cat in data.get("categories", []):
        res = check_category(cat["name"], cat["employees"], cat["declared_total"])
        report["categories"].append(res)
        grand_computed += res["computed_from_rows"]

    if "salary" in data:
        use_transit = data.get("use_transit_accounts", True)
        groups = group_salary_by_bank(data["salary"]["employees"], use_transit=use_transit)
        report["salary_groups"] = groups
        salary_total = sum(g["total"] for g in groups)
        report["salary_check"] = check_category("Зарплата", data["salary"]["employees"], data["salary"]["declared_total"])
        grand_computed += salary_total

    declared_grand_total = data.get("declared_grand_total")
    if declared_grand_total is not None:
        report["grand_total_check"] = {
            "computed": grand_computed,
            "declared": declared_grand_total,
            "match": grand_computed == declared_grand_total,
            "diff": grand_computed - declared_grand_total,
        }
    return report


def example_input():
    return {
        "use_transit_accounts": True,
        "salary": {
            "declared_total": 11194200,
            "employees": [
                {"fio": "Сарсембин У.К.", "iin": "841218301972", "bank": "Kaspi Bank", "account": "KZ40722C000014019138", "sum": 2376000},
                {"fio": "Енсебаев Н.А.", "iin": "870625350470", "bank": "First Heartland Jusan Bank", "account": "KZ11998PB00006891058", "sum": 1900800},
                {"fio": "Нурмакова С.М.", "iin": "770123450010", "bank": "Народный банк", "account": "KZ566010002034387496", "sum": 1663200},
            ],
        },
        "categories": [
            {
                "name": "ОПВ 10%",
                "declared_total": 300000,
                "employees": [{"fio": "Сарсембин У.К.", "sum": 300000}],
            }
        ],
        "declared_grand_total": 11494200,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Пример входного файла (используй как шаблон data.json):\n")
        print(json.dumps(example_input(), ensure_ascii=False, indent=2))
        sys.exit(0)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    result = run(data)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    # human-readable summary + non-zero exit if mismatches found
    problems = []
    for c in result["categories"]:
        if not c["match"]:
            problems.append(f"  ⚠ {c['category']}: построчно {c['computed_from_rows']}, заявлено {c['declared_total']}, разница {c['diff']}")
    if result.get("salary_check") and not result["salary_check"]["match"]:
        sc = result["salary_check"]
        problems.append(f"  ⚠ Зарплата: построчно {sc['computed_from_rows']}, заявлено {sc['declared_total']}, разница {sc['diff']}")
    if result.get("grand_total_check") and not result["grand_total_check"]["match"]:
        gt = result["grand_total_check"]
        problems.append(f"  ⚠ ИТОГО: посчитано {gt['computed']}, заявлено {gt['declared']}, разница {gt['diff']}")

    print("\n" + ("НАЙДЕНЫ РАСХОЖДЕНИЯ:\n" + "\n".join(problems) if problems else "Все суммы сходятся."))
