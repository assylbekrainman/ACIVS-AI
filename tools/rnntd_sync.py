#!/usr/bin/env python3
"""Сведение файла «Показатели проектов на ДД.ММ.ГГГГ.xlsx» (Google Drive) и реестра РННТД в CSV для дашборда.

    python3 tools/rnntd_sync.py <показатели.xlsx> tools/rnntd_registry.csv <папка_вывода>

Результат:
  projects.csv   — все проекты (объединение реестра и листов показателей), показатели на последнюю доступную дату,
                   признаки полноты данных (src, asof, miss, quality)
  history.csv    — накопленные продажи проекта по датам (irn, date, sales)
  quarterly.csv  — лист «Свод» (портфель по кварталам)
  quality.json   — охват, расхождения и предупреждения
Колонки последнего квартала определяются по заголовкам, поэтому скрипт работает и в следующих кварталах.
"""
import sys, re, csv, json, datetime
import openpyxl

IRN_RE = re.compile(r'^(?:\d{3,4}-\d{2}-ГК|[A-ZА-Я]{2}\d{6,})$')
DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')
LAT = str.maketrans('АРВСЕКМНОТХ', 'APBCEKMHOTX')   # кириллица -> латиница в ИРН

# поле -> регулярное выражение заголовка; берётся ПОСЛЕДНЯЯ подходящая колонка листа
FIELDS = {
    'sales':  r'^Общие продажи всего|^Продажи, тенге$',
    'cofin':  r'^Общее фактическое софинансирование|^Фактическое софинансирование, тенге$',
    'export': r'^Общий фактический экспорт|^Фактический экспорт, тенге$',
    'author_fee': r'^Общее авторское вознаграждение|^Фактическое авторское вознаграждение',
    'jobs':   r'^Количество созданных рабочих мест',
    'taxes':  r'^Налоги суммарно',
    'licenses': r'^Общее количество заключенных лицензионных|^Количество заключенных лицензионных соглашений, ед',
    'salary': r'Общая сумма зар\. ?платы, выплаченной команде',
    'salary_sci': r'^(Общая сумма зар\.платы ученым|Сумма зар\.платы ученым)',
    'assets': r'^Общая по закупу основных средств',
    'ntr':    r'^НТР$',
}
ESSENTIAL = {'industry': 'отрасль', 'region': 'регион', 'status': 'статус', 'contract_no': 'договор', 'grant': 'сумма гранта', 'sales': 'продажи'}

def norm(v): return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''
def key(irn): return irn.translate(LAT)
def num(v):
    if v is None or v == '': return ''
    if isinstance(v, (int, float)): return round(v)
    s = str(v).replace(' ', '').replace('\xa0', '').replace(',', '')
    try: return round(float(s))
    except ValueError: return ''
def parse_date(t):
    m = DATE_RE.search(t or '')
    return datetime.date(int(m[3]), int(m[2]), int(m[1])).isoformat() if m else ''

def read_drive(wb, warnings):
    out, hist = {}, []
    for ws in wb:
        if ws.title == 'Свод': continue
        rows = list(ws.iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows) if r and r[0] and '№' in str(r[0])), None)
        if hi is None: warnings.append(f'Лист {ws.title}: не найден заголовок'); continue
        head = [norm(c) for c in rows[hi]]
        col = {}
        for k, rx in FIELDS.items():
            idx = [i for i, h in enumerate(head) if re.search(rx, h)]
            if idx: col[k] = idx[-1]
            else: warnings.append(f'Лист {ws.title}: нет колонки «{k}»')
        pcol = lambda rx: next((i for i, h in enumerate(head) if re.search(rx, h)), None)
        c_name, c_grantee, c_grant, c_paid = pcol(r'^Название проекта'), pcol(r'^Грантополучатель'), pcol(r'^Сумма гранта'), pcol(r'^Фактически перечислено')
        asof = parse_date(head[col['sales']]) if 'sales' in col else ''
        asof = asof or parse_date(norm(rows[0][0]))
        ytd = [(i, parse_date(h)) for i, h in enumerate(head) if re.match(r'^Продажи на \d{2}\.\d{2}\.\d{4}', h)]
        for r in rows[hi + 1:]:
            irn = norm(r[1])
            if not IRN_RE.match(irn): continue
            rec = {'asof': asof, 'sheet': ws.title, 'irn': irn,
                   'd_name': norm(r[c_name]) if c_name is not None else '', 'd_grantee': norm(r[c_grantee]) if c_grantee is not None else '',
                   'd_grant': num(r[c_grant]) if c_grant is not None else '', 'd_paid': num(r[c_paid]) if c_paid is not None else ''}
            for k, i in col.items(): rec[k] = num(r[i])
            out[key(irn)] = rec
            if len(ytd) >= 2:   # колонки «на 31.12.ГГГГ» (база) и накопленные итоги с начала года
                base = num(r[ytd[0][0]]) or 0
                hist.append((irn, ytd[0][1], base))
                for i, d in ytd[1:]:
                    v = num(r[i])
                    if v != '' and d: hist.append((irn, d, base + v))
    return out, hist

def read_quarterly(wb, warnings):
    ws = wb['Свод']; rows = list(ws.iter_rows(values_only=True)); head = [norm(c) for c in rows[0]]
    keys = {'Продажи': 'sales', 'Налоги': 'taxes', 'Роялти': 'royalty', 'Софинансирование': 'cofin', 'рабочих мест': 'jobs',
            'лицензионных': 'licenses', 'Экспорт': 'export', 'зар.плата': 'salary', 'Основные средства': 'assets', 'Вклад в Р': 'ntr'}
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
            v = num(r[i]); rec[k] = v
            if v == '' and r[i] not in (None, ''): warnings.append(f'Лист «Свод» {d}: «{k}» — ошибка формулы ({r[i]}), значение пропущено')
        out.append(rec)
    return sorted(out, key=lambda x: x['date'])

def main(src, registry_csv, outdir):
    warnings = []
    wb = openpyxl.load_workbook(src, data_only=True)
    drive, hist = read_drive(wb, warnings)
    quarterly = read_quarterly(wb, warnings)
    reg, dups = {}, {}
    for r in csv.DictReader(open(registry_csv, encoding='utf-8')):
        k = key(r['irn'])
        if k in reg: dups.setdefault(k, [reg[k]['status']]).append(r['status'])
        reg[k] = r   # при дубле остаётся последняя строка реестра
    METRICS = ['sales', 'cofin', 'export', 'taxes', 'jobs', 'assets', 'ntr']
    projects = []
    for k in sorted(set(reg) | set(drive)):
        r, d = reg.get(k, {}), drive.get(k)
        rec = {f: r.get(f, '') for f in ['irn', 'year', 'source', 'name', 'applicant', 'grantee', 'model', 'mechanism', 'period', 'months', 'contract_no',
               'contract_date', 'grant', 'grant_s1', 'grant_s2', 'grant_s3', 'grant_s4', 'cofin_plan', 'product_service', 'product_goods', 'region',
               'industry', 'status', 'closure_reason', 'nns_decision', 'direction', 'paid', 'ip_total']}
        rec['irn'] = rec['irn'] or d['irn']
        rec['name'] = rec['name'] or (d or {}).get('d_name', ''); rec['grantee'] = rec['grantee'] or (d or {}).get('d_grantee', '')
        rec['grant'] = rec['grant'] if rec['grant'] != '' else (d or {}).get('d_grant', '')
        rec['year'] = rec['year'] or (d or {}).get('sheet', '')
        # показатели: берём источник с более поздней датой; при равенстве — файл Drive
        rasof = '2026-03-31' if any(r.get('r_' + m, '') != '' for m in METRICS) else ''
        use_drive = d is not None and any(d.get(m, '') != '' for m in METRICS) and (not rasof or d['asof'] >= rasof)
        if use_drive:
            rec['src'], rec['asof'] = 'drive', d['asof']
            for m in METRICS + ['author_fee', 'licenses', 'salary', 'salary_sci']: rec[m] = d.get(m, '')
            if rec['paid'] == '': rec['paid'] = d.get('d_paid', '')
        elif rasof:
            rec['src'], rec['asof'] = 'registry', rasof
            for m in METRICS: rec[m] = r.get('r_' + m, '')
        else:
            rec['src'], rec['asof'] = 'none', ''
        for m in METRICS + ['author_fee', 'licenses', 'salary', 'salary_sci']: rec.setdefault(m, '')
        miss = [lab for f, lab in ESSENTIAL.items() if rec.get(f, '') in ('', None)]
        rec['miss'] = '; '.join(miss)
        rec['quality'] = 'Нет показателей' if rec['src'] == 'none' else ('Полные' if not miss else 'Частичные')
        if d and not r: rec['in_registry'] = 'нет'
        rec['in_registry'] = 'да' if r else 'нет'; rec['in_drive'] = 'да' if d else 'нет'
        projects.append(rec)
    pf = ['irn', 'year', 'name', 'grantee', 'applicant', 'source', 'model', 'mechanism', 'direction', 'industry', 'region', 'status', 'closure_reason', 'nns_decision',
          'contract_no', 'contract_date', 'period', 'months', 'grant', 'grant_s1', 'grant_s2', 'grant_s3', 'grant_s4', 'cofin_plan', 'paid',
          'src', 'asof', 'sales', 'cofin', 'export', 'taxes', 'jobs', 'assets', 'ntr', 'author_fee', 'licenses', 'salary', 'salary_sci', 'ip_total',
          'product_service', 'product_goods', 'in_registry', 'in_drive', 'miss', 'quality']
    def w(path, fields, rows):
        with open(path, 'w', newline='', encoding='utf-8') as f:
            wr = csv.DictWriter(f, fields, extrasaction='ignore'); wr.writeheader(); wr.writerows(rows)
    w(f'{outdir}/projects.csv', pf, projects)
    w(f'{outdir}/history.csv', ['irn', 'date', 'sales'], [{'irn': i, 'date': d, 'sales': s} for i, d, s in hist])
    w(f'{outdir}/quarterly.csv', ['date', 'sales', 'taxes', 'royalty', 'cofin', 'jobs', 'licenses', 'export', 'salary', 'assets', 'ntr'], quarterly)
    from collections import Counter
    qc = Counter(p['quality'] for p in projects); sc = Counter(p['src'] for p in projects)
    meta = {'projects': len(projects), 'registry': len(reg), 'drive': len(drive), 'both': len(set(reg) & set(drive)),
            'only_registry': len(set(reg) - set(drive)), 'only_drive': len(set(drive) - set(reg)), 'quality': dict(qc), 'source': dict(sc),
            'latest_quarter': quarterly[-1]['date'] if quarterly else '', 'warnings': sorted(set(warnings))}
    # --- замечания по качеству данных (показываются на дашборде) ---
    F = []
    def add(sev, finding, count, rec): F.append({'severity': sev, 'finding': finding, 'count': count, 'recommendation': rec})
    err = [x for x in warnings if 'ошибка формулы' in x]
    if err: add('Высокая', 'Лист «Свод»: ошибка формулы (#VALUE!) в итогах последнего квартала — ' + '; '.join(sorted({{'sales': 'продажи', 'ntr': 'вклад в НТР'}.get(re.findall('«(\\w+)»', x)[-1], re.findall('«(\\w+)»', x)[-1]) for x in err})), len(err), 'Исправить формулы в листе «Свод»; до исправления значения на дашборде пропущены.')
    if dups: add('Высокая', 'Дубли ИРН в реестре с разными статусами: ' + ', '.join(f"{k} ({' / '.join(v)})" for k, v in dups.items()), len(dups), 'Оставить одну запись на ИРН; смена грантополучателя оформлять отдельным полем, а не новой строкой.')
    stale = {}
    for p in projects:
        if p['src'] == 'drive' and p['asof'] and p['asof'] < meta['latest_quarter']: stale[p['asof']] = stale.get(p['asof'], 0) + 1
    old_sheets = sorted({(w['sheet'], w['asof']) for w in drive.values() if w['asof'] and w['asof'] < meta['latest_quarter']})
    if old_sheets:
        n_old = sum(1 for w in drive.values() if w['asof'] < meta['latest_quarter'])
        add('Средняя', 'Листы с устаревшей отчётной датой: ' + ', '.join(f"{s} — на {a[8:]}.{a[5:7]}.{a[:4]}" for s, a in old_sheets), n_old, 'Обновить листы до последнего квартала; пока показатели этих проектов берутся из реестра (31.03.2026), если он свежее.')
    only_r = [p for p in projects if p['in_drive'] == 'нет']
    if only_r: add('Средняя', 'Проекты из реестра отсутствуют в листах показателей Drive (закрытые, завершённые или не внесённые)', len(only_r), 'Подтвердить статус; для завершённых и закрытых зафиксировать итоговые показатели на дату закрытия.')
    only_d = [p for p in projects if p['in_registry'] == 'нет']
    if only_d: add('Средняя', 'Проекты из листов Drive отсутствуют в реестре — нет отрасли, региона, статуса, условий договора', len(only_d), 'Внести проекты в реестр; пока на дашборде они отображаются как «Не указано».')
    for f_, lab in [('industry', 'отрасль'), ('region', 'регион'), ('contract_no', 'номер договора'), ('grant', 'сумма гранта')]:
        n = sum(1 for p in projects if p[f_] in ('', None))
        if n: add('Низкая', f'Не заполнено поле «{lab}»', n, 'Дозаполнить в реестре.')
    for m_, lab in [('taxes', 'налоги'), ('cofin', 'софинансирование'), ('export', 'экспорт'), ('jobs', 'рабочие места'), ('assets', 'основные средства')]:
        tot = sum(w[m_] for w in drive.values() if w.get(m_, '') != ''); sv = quarterly[-1].get(m_, '')
        if sv != '' and sv and abs(tot - sv) / sv > 0.03:
            add('Средняя', f'Расхождение «Свод» и суммы по листам проектов — {lab}: {sv:,} против {tot:,}'.replace(',', ' '), '', 'Сверить методику свода (какие проекты и периоды входят в итог).')
    w(f'{outdir}/quality.csv', ['severity', 'finding', 'count', 'recommendation'], F)
    RECON = []
    for m_, lab in [('sales', 'Продажи'), ('taxes', 'Налоги'), ('cofin', 'Софинансирование'), ('export', 'Экспорт'), ('jobs', 'Рабочие места'), ('assets', 'Основные средства'), ('ntr', 'Вклад в НТР')]:
        vals = [w[m_] for w in drive.values() if w.get(m_, '') != '']
        RECON.append({'metric': m_, 'label': lab, 'svod': quarterly[-1].get(m_, ''), 'drive_sum': sum(vals) if vals else ''})
    w(f'{outdir}/recon.csv', ['metric', 'label', 'svod', 'drive_sum'], RECON)
    json.dump(meta, open(f'{outdir}/quality.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(meta, ensure_ascii=False, indent=1))

if __name__ == '__main__': main(*sys.argv[1:4])
