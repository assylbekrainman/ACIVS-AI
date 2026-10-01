"""Тесты ugd_lookup.py. Запуск: python3 -m unittest skills/ugd-bin-lookup/scripts/test_ugd_lookup.py -v"""
import collections
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(__file__))
import ugd_lookup as u  # noqa: E402


class RegistryTests(unittest.TestCase):
    def test_size_and_checksums(self):
        rows = u.load()
        self.assertEqual(len(rows), 246)
        self.assertEqual([r["bin"] for r in rows if not u.bin_checksum_ok(r["bin"])], [])

    def test_unique_bins_and_codes_format(self):
        rows = u.load()
        dup = [b for b, c in collections.Counter(r["bin"] for r in rows).items() if c > 1]
        self.assertEqual(dup, [])
        self.assertTrue(all(len(r["code"]) == 4 and r["code"].isdigit() for r in rows))

    def test_checksum_rejects_garbage(self):
        self.assertFalse(u.bin_checksum_ok("123456789012"))
        self.assertFalse(u.bin_checksum_ok("91074000012"))
        self.assertFalse(u.bin_checksum_ok("91074000012X"))


class SearchTests(unittest.TestCase):
    def one(self, q):
        res = u.search(q)
        self.assertEqual(len(res), 1, (q, res))
        return res[0]

    def test_known_examples_from_real_letters(self):
        self.assertEqual(self.one("Медеуский район г. Алматы")["bin"], "910740000123")
        self.assertEqual(self.one("Шардаринский район")["bin"], "021140001969")
        self.assertEqual(self.one("Бостандыкский район г. Алматы")["bin"], "910740000044")
        self.assertEqual(self.one("Айыртауский район Северо-Казахстанской области")["bin"], "980740001302")
        self.assertEqual(self.one("Карасайский район")["bin"], "900340000058")

    def test_similar_names_not_confused(self):
        self.assertNotEqual(self.one("Карасайский район")["bin"], self.one("Карасуский район")["bin"])

    def test_short_query_prefix(self):
        self.assertEqual(self.one("Медеу Алматы")["name"], "УГД по Медеускому району")

    def test_ambiguous_returns_several(self):
        self.assertGreater(len(u.search("Алматинская область")), 1)

    def test_not_found(self):
        self.assertEqual(u.search("Несуществующий район"), [])
        self.assertEqual(u.search("район по"), [])

    def test_cli_exit_codes(self):
        self.assertEqual(u.main(["x", "Медеуский район г. Алматы"]), 0)
        self.assertEqual(u.main(["x", "Алматинская область"]), 3)
        self.assertEqual(u.main(["x", "Несуществующий"]), 2)
        self.assertEqual(u.main(["x", "--bin", "910740000123"]), 0)
        self.assertEqual(u.main(["x", "--bin", "000000000000"]), 2)
        self.assertEqual(u.main(["x", "--check"]), 0)


if __name__ == "__main__":
    unittest.main()
