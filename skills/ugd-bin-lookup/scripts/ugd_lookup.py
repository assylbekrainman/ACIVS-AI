#!/usr/bin/env python3
"""Поиск БИН и наименования УГД/ДГД по району/городу/области (справочник КГД МФ РК, 2025).

Использование:
    python3 ugd_lookup.py "Медеуский район г. Алматы"
    python3 ugd_lookup.py --bin 910740000123      # обратный поиск по БИН
    python3 ugd_lookup.py --check                 # контрольная сумма всех БИН справочника

Код возврата: 0 — найдена ровно одна запись; 2 — не найдено; 3 — несколько кандидатов
(нужно уточнить у пользователя, ничего не выбирать наугад).
"""
import csv
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REGISTRY = os.path.join(HERE, "..", "references", "ugd-registry.csv")

# Постоянные реквизиты получателя-банка (из шапки кодификатора), одинаковы для всех УГД.
TREASURY = {"bik": "KKMFKZ2A", "iik": "KZ24070105KSN0000000",
            "bank": "ГУ «Комитет казначейства Министерства финансов РК»"}

STOP = {"р", "н", "г", "по", "обл", "угд", "дгд", "и", "в", "им", "рк", "казахстан"}
# служебные слова в любых падежах: район(а/у/е/ы), область/области/..., город(а/у/е), республики
STOP_STEMS = ("район", "облас", "город", "респу")


def load():
    with open(REGISTRY, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def bin_checksum_ok(b: str) -> bool:
    """Контрольная цифра БИН/ИИН: веса 1..11, при остатке 10 — веса 3..11,1,2; повторно 10 — невалидный."""
    if not re.fullmatch(r"\d{12}", b):
        return False
    d = [int(c) for c in b[:11]]
    s = sum((i + 1) * d[i] for i in range(11)) % 11
    if s == 10:
        w = [3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2]
        s = sum(w[i] * d[i] for i in range(11)) % 11
        if s == 10:
            return False
    return s == int(b[11])


def tokens(text):
    text = text.lower().replace("ё", "е")
    toks = re.findall(r"[a-zа-яәіңғүұқөһ]+", text)
    return [t for t in toks if t not in STOP and not t.startswith(STOP_STEMS)]


def tok_match(q, n):
    """Токен запроса q совпадает с токеном наименования n.
    Короткий запрос (<=6 букв) — как префикс («Медеу» → «Медеускому»);
    длинный — по первым 6 буквам (Карасайский ≠ Карасуский, Алматинская ≠ Алматы)."""
    if len(q) <= 6:
        return n.startswith(q)
    return n[:6] == q[:6]


def search(query, rows=None):
    rows = rows if rows is not None else load()
    q = tokens(query)
    if not q:
        return []
    hits, by_name = [], []
    for r in rows:
        nt, rt = tokens(r["name"]), tokens(r["region"])
        if all(any(tok_match(t, n) for n in nt + rt) for t in q):
            hits.append(r)
            if all(any(tok_match(t, n) for n in nt) for t in q):
                by_name.append(r)
    # запрос, полностью совпавший с наименованием органа, точнее, чем совпадение по разделу/области
    return by_name or hits


def fmt(r):
    return f'{r["bin"]}  {r["name"]}  [{r["region"]}, код {r["code"]}]'


def main(argv):
    if len(argv) >= 2 and argv[1] == "--check":
        bad = [r for r in load() if not bin_checksum_ok(r["bin"])]
        print("Записей:", len(load()), "| с неверной контрольной цифрой:", len(bad))
        for r in bad:
            print("  ", fmt(r))
        return 1 if bad else 0
    if len(argv) >= 3 and argv[1] == "--bin":
        res = [r for r in load() if r["bin"] == argv[2]]
        print("\n".join(fmt(r) for r in res) or "БИН в справочнике УГД/ДГД не найден")
        return 0 if res else 2
    if len(argv) < 2:
        print(__doc__)
        return 2
    res = search(" ".join(argv[1:]))
    if not res:
        print("Не найдено. Уточните район/город и область регистрации ГП.")
        return 2
    for r in res:
        print(fmt(r))
    if len(res) > 1:
        print(f"Кандидатов: {len(res)} — уточните у пользователя, не выбирайте наугад.")
        return 3
    print(f'Получатель-банк: {TREASURY["bank"]}, БИК {TREASURY["bik"]}, ИИК {TREASURY["iik"]}')
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
