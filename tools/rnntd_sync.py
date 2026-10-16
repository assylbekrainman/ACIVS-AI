#!/usr/bin/env python3
"""Разбор файла «Показатели проектов на ДД.ММ.ГГГГ.xlsx» (Google Drive) в плоские CSV для дашборда РННТД.

Использование:
    python3 tools/rnntd_sync.py <показатели.xlsx> <реестр_атрибутов.csv> <папка_вывода>

Колонки последнего квартала определяются по заголовкам (регулярные выражения),
поэтому скрипт работает и после появления новых кварталов без правок.
Результат: projects.csv (проекты), quarterly.csv (лист «Свод»), meta.json (даты, предупреждения).
"""
import sys, re, csv, json, datetime
import openpyxl

IRN_RE = re.compile(r'^(?:\d{3,4}-\d{2}-ГК|[A-Z]{2}\d{6,})$')
DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')

# поле -> регулярное выражение заголовка; берётся ПОСЛЕДНЯЯ подходящая колонка листа
FIELDS = {
    'grant':  r'^Сумма гранта',
    'paid':   r'^Фактически перечислено',
    'sales':  r'^Общие продажи всего|^Продажи, тенге$',
    'cofin':  r'^Общее фактическое софинансирование|^Фактическое софинансирование, тенге$',
    'export': r'^Общий фактический экспорт|^Фактический экспорт, тенге$',
    'jobs':   r'^Количество созданных рабочих мест',
    'taxes':  r'^Налоги суммарно',
    'assets': r'^Общая по закупу основных средств',
    'ntr':    r'^НТР$',
    'direction': r'^Проектная дирекция',
}

def norm(v):
    return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''

def num(v):
    if v is None or v == '': return ''
    if isinstance(v, (int, float)): return round(v)
    s = str(v).replace(' ', '').replace('\xa0', '').replace(',', '')
    try: return round(float(s))
    except ValueError: return ''   # напр. #VALUE!

def parse_date(text):
    m = DATE_RE.search(text or '')
    return datetime.date(int(m[3]), int(m[2]), int(m[1])).isoformat() if m else ''

def read_projects(wb, attrs, warnings):
    out = []
    for ws in wb:
        if ws.title == 'Свод': continue
        rows = list(ws.iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows) if r and r[0] and '№' in str(r[0])), None)
        if hi is None: warnings.append(f'{ws.title}: не найден заголовок'); continue
        head = [norm(c) for c in rows[hi]]
        col = {}
        for key, rx in FIELDS.items():
            idx = [i for i, h in enumerate(head) if re.search(rx, h)]
            if idx: col[key] = idx[-1]
            else: warnings.append(f'{ws.title}: нет колонки «{key}»')
        asof = parse_date(head[col['sales']]) if 'sales' in col else ''
        if not asof:
            asof = parse_date(norm(rows[0][0]))  # у старых листов дата в заголовке листа
        n = 0
        for r in rows[hi + 1:]:
            irn = norm(r[1])
            if not IRN_RE.match(irn): continue
            rec = {'irn': irn, 'sheet': ws.title, 'asof': asof}
            for key, i in col.items():
                rec[key] = norm(r[i]) if key == 'direction' else num(r[i])
            a = attrs.get(irn, {})
            for k in ('year', 'industry', 'region', 'model', 'status'):
                rec[k] = a.get(k, '')
            if not rec['year']: rec['year'] = ws.title
            out.append(rec); n += 1
        if n == 0: warnings.append(f'{ws.title}: 0 проектов')
    return out

def read_quarterly(wb, warnings):
    ws = wb['Свод']
    rows = [r for r in ws.iter_rows(values_only=True)]
    head = [norm(c) for c in rows[0]]
    keys = {'Продажи': 'sales', 'Налоги': 'taxes', 'Роялти': 'royalty', 'Софинансирование': 'cofin',
            'рабочих мест': 'jobs', 'лицензионных': 'licenses', 'Экспорт': 'export',
            'зар.плата': 'salary', 'Основные средства': 'assets', 'Вклад в Р': 'ntr'}
    cols = {}
    for i, h in enumerate(head):
        for frag, k in keys.items():
            if frag in h and k not in cols.values(): cols[i] = k
    out = []
    for r in rows[1:]:
        d = parse_date(norm(r[0]))
        if not d or 'прирост' in norm(r[0]): continue
        rec = {'date': d}
        for i, k in cols.items():
            v = num(r[i])
            rec[k] = v
            if v == '' and r[i] is not None and r[i] != '':
                warnings.append(f'Свод {d}: «{k}» содержит ошибку формулы ({r[i]}); оставлено пустым')
        out.append(rec)
    out.sort(key=lambda x: x['date'])
    return out

def main(src, attrs_csv, outdir):
    warnings = []
    wb = openpyxl.load_workbook(src, data_only=True)
    attrs = {r['irn']: r for r in csv.DictReader(open(attrs_csv, encoding='utf-8'))}
    projects = read_projects(wb, attrs, warnings)
    quarterly = read_quarterly(wb, warnings)
    pf = ['irn','sheet','year','industry','region','model','status','direction','asof','grant','paid','sales','cofin','export','taxes','jobs','assets','ntr']
    with open(f'{outdir}/projects.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, pf, extrasaction='ignore'); w.writeheader(); w.writerows(projects)
    qf = ['date','sales','taxes','royalty','cofin','jobs','licenses','export','salary','assets','ntr']
    with open(f'{outdir}/quarterly.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, qf, extrasaction='ignore'); w.writeheader(); w.writerows(quarterly)
    matched = sum(1 for p in projects if p['status'])
    meta = {'projects': len(projects), 'with_status': matched, 'latest_quarter': quarterly[-1]['date'] if quarterly else '',
            'asof_by_sheet': {p['sheet']: p['asof'] for p in projects}, 'warnings': sorted(set(warnings))}
    json.dump(meta, open(f'{outdir}/meta.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(meta, ensure_ascii=False, indent=1))

if __name__ == '__main__':
    main(*sys.argv[1:4])
