"""Тесты check_and_group.py на синтетических данных (вымышленные ФИО/ИИН/счета).
Запуск: python3 -m unittest escrow-fot/scripts/test_check_and_group.py -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(__file__))
import check_and_group as c  # noqa: E402


def emp(fio, bank, s, account="KZ00000000000000TEST"):
    return {"fio": fio, "iin": "000000000000", "bank": bank, "account": account, "sum": s}


class GroupingTests(unittest.TestCase):
    def test_same_bank_grouped_into_one_transit_beneficiary(self):
        groups = c.group_salary_by_bank([
            emp("Тест А.", "Kaspi Bank", 100), emp("Тест Б.", "kaspi bank", 250),
        ])
        self.assertEqual(len(groups), 1)
        self.assertTrue(groups[0]["routed_via_transit"])
        self.assertEqual(groups[0]["total"], 350)
        self.assertEqual(groups[0]["beneficiary_account"], c.TRANSIT_ACCOUNTS["kaspi bank"]["iik"])

    def test_unknown_bank_goes_direct_to_employee_account(self):
        groups = c.group_salary_by_bank([emp("Тест В.", "Неизвестный банк", 500, account="KZ11TESTACC")])
        self.assertFalse(groups[0]["routed_via_transit"])
        self.assertEqual(groups[0]["beneficiary_account"], "KZ11TESTACC")

    def test_pre_2026_no_transit(self):
        groups = c.group_salary_by_bank([emp("Тест А.", "Kaspi Bank", 100), emp("Тест Б.", "Kaspi Bank", 200)],
                                        use_transit=False)
        self.assertEqual(len(groups), 2)
        self.assertTrue(all(not g["routed_via_transit"] for g in groups))

    def test_bank_name_quotes_and_case_normalized(self):
        self.assertEqual(c.normalize_bank('  АО «Народный  Банк» '), "ао народный банк")
        groups = c.group_salary_by_bank([emp("Тест Г.", "«Народный банк»", 10)])
        self.assertTrue(groups[0]["routed_via_transit"])


class ReconciliationTests(unittest.TestCase):
    def test_category_match_and_mismatch(self):
        rows = [{"fio": "Тест А.", "sum": 100}, {"fio": "Тест Б.", "sum": 50}]
        self.assertTrue(c.check_category("ОПВ", rows, 150)["match"])
        bad = c.check_category("ОПВ", rows, 140)
        self.assertFalse(bad["match"])
        self.assertEqual(bad["diff"], 10)

    def test_grand_total_includes_salary_and_categories(self):
        data = {
            "salary": {"declared_total": 300, "employees": [emp("Тест А.", "Kaspi Bank", 100), emp("Тест Б.", "Народный банк", 200)]},
            "categories": [{"name": "ОПВ 10%", "declared_total": 30, "employees": [{"fio": "Тест А.", "sum": 30}]}],
            "declared_grand_total": 330,
        }
        r = c.run(data)
        self.assertTrue(r["salary_check"]["match"])
        self.assertTrue(r["grand_total_check"]["match"])

    def test_grand_total_mismatch_detected(self):
        data = {
            "salary": {"declared_total": 300, "employees": [emp("Тест А.", "Kaspi Bank", 300)]},
            "categories": [],
            "declared_grand_total": 301,
        }
        self.assertEqual(c.run(data)["grand_total_check"]["diff"], -1)


if __name__ == "__main__":
    unittest.main()
