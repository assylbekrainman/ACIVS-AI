#!/usr/bin/env python3
"""Выгрузка атрибутов проектов из реестра РННТД (xlsx) в tools/rnntd_registry.csv.

    python3 tools/rnntd_registry.py <реестр.xlsx> tools/rnntd_registry.csv

Персональные данные руководителей (ФИО, пол, дата рождения, возраст) намеренно не выгружаются.
"""
import sys, csv, re, datetime, openpyxl

def norm(v): return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''
def num(v):
    try: return round(float(v)) if v not in (None, '') else ''
    except (TypeError, ValueError): return ''
def date(v): return v.date().isoformat() if isinstance(v, datetime.datetime) else norm(v)

# колонка реестра -> (поле, тип); индексы проверяются по заголовкам ниже
COLS = {2: ('year', norm), 3: ('source', norm), 5: ('name', norm), 8: ('applicant', norm), 14: ('grantee', norm),
        19: ('model', lambda v: norm(v).lower().replace('стартап-компания', 'стартап').replace('стартап компания', 'стартап')),
        20: ('mechanism', norm), 21: ('period', norm), 22: ('months', num), 23: ('contract_no', norm), 24: ('contract_date', date),
        25: ('grant', num), 26: ('grant_s1', num), 27: ('grant_s2', num), 28: ('grant_s3', num), 29: ('grant_s4', num),
        30: ('cofin_plan', num), 37: ('product_service', norm), 38: ('product_goods', norm), 40: ('region', norm), 46: ('industry', norm),
        84: ('status', norm), 85: ('closure_reason', norm), 86: ('nns_decision', norm), 87: ('direction', norm),
        90: ('paid', num), 96: ('r_sales', num), 102: ('r_cofin', num), 108: ('r_export', num), 115: ('r_jobs', num),
        128: ('r_taxes', num), 152: ('r_assets', num), 153: ('r_ntr', num)}
EXPECT = {1: 'ИРН', 2: 'конкурс', 5: 'Название проекта', 25: 'Общая сумма гранта', 46: 'Отрасль', 84: 'Текущий статус', 87: 'Проектная', 153: 'НТР'}
IP = [53, 54, 55, 56, 57, 58, 59, 60, 61, 63, 64]   # охранные документы и лицензии (шт.)

def main(src, dst):
    rows = list(openpyxl.load_workbook(src, data_only=True, read_only=True).worksheets[0].iter_rows(values_only=True))
    head = [norm(c) for c in rows[0]]
    for i, frag in EXPECT.items():
        if frag not in head[i]: raise SystemExit(f'Структура реестра изменилась: колонка {i} «{head[i]}», ожидалось «{frag}»')
    out = []
    for r in rows[1:]:
        if not norm(r[1]): continue
        rec = {'irn': norm(r[1])}
        for i, (k, fn) in COLS.items(): rec[k] = fn(r[i])
        rec['industry'] = rec['industry'].strip(); rec['status'] = rec['status'].capitalize() if rec['status'] else ''
        d = rec['direction']; rec['direction'] = re.sub(r'ДКТ\s*', 'ДКТ ', d) if d.startswith('ДКТ') else ''
        if not rec['direction'] and d not in ('', '0') and rec['nns_decision'].lower() in ('', 'нет'): rec['nns_decision'] = d  # у части старых строк решение ННС записано в графе дирекции
        m = re.search(r'(?:№|механизм)\s*(\d)|^(\d)$', rec['mechanism']); rec['mechanism'] = (m.group(1) or m.group(2)) if m else ''
        rec['source'] = {'Инициативная': 'Инициативный'}.get(rec['source'], rec['source'])
        rec['ip_total'] = sum(num(r[i]) or 0 for i in IP)
        out.append(rec)
    with open(dst, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, list(out[0])); w.writeheader(); w.writerows(out)
    print(len(out), 'проектов ->', dst)

if __name__ == '__main__': main(*sys.argv[1:3])
