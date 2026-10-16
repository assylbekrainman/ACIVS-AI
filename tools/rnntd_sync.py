#!/usr/bin/env python3
"""Сведение квартальных файлов «Показатели проектов на ДД.ММ.ГГГГ» (Google Drive) в CSV для дашборда РННТД.
Источник — только Google Drive. Лист «Свод» не используется: итоги считаются по проектам.

    python3 tools/rnntd_sync.py <папка_со_снимками> <проекты_2022-2026.xlsx> <папка_вывода>

В папке со снимками лежат файлы с именами ГГГГ-ММ-ДД.xlsx (по одному на отчётную дату).
Результат:
  projects.csv   — все проекты, показатели из последнего снимка, где проект есть; признаки полноты и отчётности
  history.csv    — показатели проекта на каждую отчётную дату
  quarterly.csv  — итоги портфеля по датам (сумма по проектам)
  quality.csv    — замечания по качеству данных
Колонки находятся по заголовкам, поэтому новые кварталы работают без правок.
"""
import sys, os, re, csv, json, datetime, glob
import openpyxl

IRN_RE = re.compile(r'^(?:\d{3,4}-\d{2}-ГК|[A-ZА-Я]{2}\d{6,})$')
DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')
LAT = str.maketrans('АРВСЕКМНОТХ', 'APBCEKMHOTX')
FIELDS = {   # поле -> заголовок; берётся последняя подходящая колонка листа
    'sales': r'^Общие продажи всего|^Продажи, тенге$', 'cofin': r'^Общее фактическое софинансирование|^Фактическое софинансирование, тенге$',
    'export': r'^Общий фактический экспорт|^Фактический экспорт, тенге$', 'author_fee': r'^Общее авторское вознаграждение|^Фактическое авторское вознаграждение',
    'jobs': r'^Количество созданных рабочих мест', 'taxes': r'^Налоги суммарно', 'licenses': r'^Общее количество заключенных лицензионных|^Количество заключенных лицензионных соглашений, ед',
    'salary': r'Общая сумма зар\. ?платы, выплаченной команде', 'assets': r'^Общая по закупу основных средств', 'ntr': r'^НТР$'}
METRICS = ['sales', 'cofin', 'export', 'taxes', 'jobs', 'assets', 'ntr']
ESSENTIAL = {'industry': 'отрасль', 'region': 'регион', 'grant': 'сумма гранта', 'sales': 'продажи'}

def norm(v): return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''
def key(i): return i.translate(LAT)
def num(v, errs=None):
    if v is None or v == '': return ''
    if isinstance(v, (int, float)): return round(v)
    s = str(v).replace(' ', '').replace('\xa0', '').replace(',', '')
    try: return round(float(s))
    except ValueError:
        if errs is not None and s.startswith('#'): errs.append(s)
        return ''
def pdate(t):
    m = DATE_RE.search(t or ''); return datetime.date(int(m[3]), int(m[2]), int(m[1])).isoformat() if m else ''
def dmy(d): return f'{d[8:]}.{d[5:7]}.{d[:4]}'

def read_snapshot(path, errs):
    wb = openpyxl.load_workbook(path, data_only=True); out = {}; sheets = {}; dup = []
    for ws in wb:
        if ws.title == 'Свод': continue
        rows = list(ws.iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows) if r and r[0] and '№' in str(r[0])), None)
        if hi is None: continue
        head = [norm(c) for c in rows[hi]]
        col = {k: [i for i, h in enumerate(head) if re.search(rx, h)][-1] for k, rx in FIELDS.items() if any(re.search(rx, h) for h in head)}
        p = lambda rx: next((i for i, h in enumerate(head) if re.search(rx, h)), None)
        cn, cg, cgr, cp, cd = p(r'^Название проекта'), p(r'^Грантополучатель'), p(r'^Сумма гранта'), p(r'^Фактически перечислено'), p(r'^Проектная дирекция')
        g = lambda r, i: r[i] if i is not None and i < len(r) else None
        asof = pdate(head[col['sales']]) if 'sales' in col else ''
        sheets[ws.title] = asof or pdate(norm(rows[0][0]))
        for r in rows[hi + 1:]:
            irn = norm(r[1])
            if not IRN_RE.match(irn): continue
            rec = {'irn': irn, 'sheet': ws.title, 'name': norm(g(r, cn)), 'grantee': norm(g(r, cg)), 'grant': num(g(r, cgr), errs), 'paid': num(g(r, cp), errs), 'direction': norm(g(r, cd))}
            for k, i in col.items(): rec[k] = num(g(r, i), errs)
            if key(irn) in out: dup.append(irn)
            out[key(irn)] = rec
    return out, sheets, dup

def region_of(oblast, city):
    M = {'тараз': 'Жамбылская область', 'караганда': 'Карагандинская область', 'туркестан': 'Туркестанская область', 'усть-каменогорск': 'Восточно-Казахстанская область', 'степногорск': 'Акмолинская область', 'кокшетау': 'Акмолинская область', 'костанай': 'Костанайская область', 'павлодар': 'Павлодарская область', 'актобе': 'Актюбинская область', 'атырау': 'Атырауская область', 'актау': 'Мангистауская область', 'семей': 'Абайская область', 'петропавловск': 'Северо-Казахстанская область', 'уральск': 'Западно-Казахстанская область', 'кызылорда': 'Кызылординская область', 'талдыкорган': 'Жетысуская область', 'конаев': 'Алматинская область'}
    def clean(s):
        s = norm(s).strip(' ,;')
        return '' if s.lower() in ('', '-', 'none', 'нет', '0', '0.0') else s
    for s in (clean(oblast), clean(city)):
        if not s: continue
        l = s.lower().replace('г.', '').replace('город', '').replace('гор.', '').strip()
        for c in ('алматы', 'астана', 'шымкент'):
            if l == c or (c in l and 'област' not in l and 'обл' not in l): return c.capitalize()
        for c, rg in M.items():
            if c in l and 'обл' not in l: return rg
        m = re.match(r'^([а-яё\- ]+?)\s*(област|обл\.?)', l)
        if m:
            b = m.group(1).strip().capitalize()
            return {'Вко': 'Восточно-Казахстанская область', 'Ско': 'Северо-Казахстанская область', 'Зко': 'Западно-Казахстанская область'}.get(b, b + ' область')
        if l in ('вко',): return 'Восточно-Казахстанская область'
        if l in ('ско',): return 'Северо-Казахстанская область'
        if l in ('зко',): return 'Западно-Казахстанская область'
        if oblast == s: continue
        return s.capitalize()
    return ''

def industry_of(t):
    l = norm(t).lower()
    for k, v in [('агропромышл', 'АПК'), ('обрабатывающ', 'Обрабатывающая промышленность'), ('цифров', 'Цифровизация'), ('экологи', 'Экология'), ('здравоохран', 'Здравоохранение'), ('геолог', 'Геология и недропользование'), ('машиностро', 'Машиностроение'), ('хими', 'Химическая промышленность'), ('образован', 'Качественное образование'), ('энергет', 'Энергетика'), ('металлург', 'Металлургия'), ('безопасност', 'Национальная безопасность')]:
        if k in l: return v
    return ''

def model_of(t):
    l = norm(t).lower()
    if l in ('', 'none', '0.0', '3.0', '0', '-'): return ''
    if 'вуз' in l: return 'ВУЗ'
    if l in ('нпц',): return 'НПЦ'
    if 'нии' in l: return 'НИИ'
    if 'физ' in l: return 'Физическое лицо'
    return 'Бизнес-организация' if any(x in l for x in ('тоо', 'ао', 'нао', 'бизнес', 'нпп', 'ип', 'стартап')) else t

def read_attrs(path):
    ws = openpyxl.load_workbook(path, data_only=True).active; out = {}
    for r in list(ws.iter_rows(values_only=True))[1:]:
        irn = norm(r[1])
        if not IRN_RE.match(irn): continue
        out[key(irn)] = {'industry': industry_of(r[10]), 'region': region_of(r[4], r[5]), 'model': model_of(r[6]), 'period': norm(r[8]), 'applicant': norm(r[3])}
    return out

def main(snapdir, attrs_xlsx, outdir):
    files = sorted(glob.glob(f'{snapdir}/????-??-??.xlsx')); dates = [os.path.basename(f)[:10] for f in files]
    snaps, sheets, dups, errs = {}, {}, {}, {}
    for f, d in zip(files, dates):
        e = []; snaps[d], sheets[d], dups[d] = read_snapshot(f, e); errs[d] = e
    latest = dates[-1]; attrs = read_attrs(attrs_xlsx)
    keys = sorted({k for d in dates for k in snaps[d]} | set())
    projects, hist, F = [], [], []
    for k in keys:
        last = next(d for d in reversed(dates) if k in snaps[d]); s = snaps[last][k]; a = attrs.get(k, {})
        rec = {'irn': s['irn'], 'year': s['sheet'], 'name': s['name'], 'grantee': s['grantee'], 'direction': s['direction'], 'industry': a.get('industry', ''), 'region': a.get('region', ''), 'model': a.get('model', ''),
               'period': a.get('period', ''), 'status': 'В отчётности' if last == latest else 'Вне отчётности', 'last_date': last, 'in_attrs': 'да' if a else 'нет', 'n_snap': sum(1 for d in dates if k in snaps[d])}
        for f_ in ['grant', 'paid'] + METRICS + ['author_fee', 'licenses', 'salary']: rec[f_] = s.get(f_, '')
        if not rec['name'] and a.get('applicant'): rec['grantee'] = rec['grantee'] or a['applicant']
        miss = [lab for f_, lab in ESSENTIAL.items() if rec[f_] in ('', None)]
        rec['miss'] = '; '.join(miss); rec['quality'] = 'Нет показателей' if all(rec[m] == '' for m in METRICS) else ('Полные' if not miss else 'Частичные')
        projects.append(rec)
        for d in dates:
            if k in snaps[d]:
                h = {'irn': s['irn'], 'date': d}; h.update({m: snaps[d][k].get(m, '') for m in METRICS}); hist.append(h)
    quarterly = []
    for d in dates:
        q = {'date': d, 'n_projects': len(snaps[d])}
        for m in METRICS: q[m] = sum(v.get(m, 0) or 0 for v in snaps[d].values() if v.get(m, '') != '')
        quarterly.append(q)
    # --- замечания ---
    def add(sev, finding, count, rec): F.append({'severity': sev, 'finding': finding, 'count': count, 'recommendation': rec})
    old = {sh: dt for sh, dt in sheets[latest].items() if dt and dt < latest}
    if old: add('Средняя', 'Листы последнего файла с устаревшей отчётной датой: ' + ', '.join(f'{s} — на {dmy(dt)}' for s, dt in sorted(old.items())), sum(1 for v in snaps[latest].values() if v['sheet'] in old), 'Обновить листы до отчётной даты файла.')
    drop = [p for p in projects if p['status'] == 'Вне отчётности']
    if drop: add('Средняя', f'Проекты были в прежних файлах, но отсутствуют в последнем ({dmy(latest)}); их показатели взяты на последнюю дату наличия', len(drop), 'Подтвердить причину (завершение, закрытие, пропуск) и зафиксировать итоговые показатели.')
    for d, e in errs.items():
        if e and d == latest: add('Высокая', f'В файле на {dmy(d)} есть ячейки с ошибками формул: {", ".join(sorted(set(e)))}', len(e), 'Исправить формулы; такие ячейки на дашборде считаются пустыми.')
    dd = [i for i in dups[latest]]
    if dd: add('Высокая', 'ИРН встречается на нескольких листах последнего файла: ' + ', '.join(dd[:6]), len(dd), 'Оставить одну запись на проект.')
    # уменьшение накопленных продаж между снимками
    dec = []
    for k in keys:
        prev = None
        for d in dates:
            if k in snaps[d]:
                v = snaps[d][k].get('sales', '')
                if v != '' and prev not in (None, '') and v < prev - 1: dec.append((k, prev - v))
                prev = v if v != '' else prev
    decn = len({k for k, _ in dec})
    if decn: add('Средняя', f'Накопленные продажи уменьшились между отчётными датами (сумма снижений {sum(x for _, x in dec)/1e6:,.0f} млн ₸)'.replace(',', ' '), decn, 'Проверить корректировки; накопленные продажи не должны убывать.')
    no_attr = [p for p in projects if p['in_attrs'] == 'нет']
    if no_attr: add('Средняя', 'Для проектов нет описательных данных (отрасль, регион, тип организации): файл с атрибутами на Drive охватывает только конкурсы 2022–2026', len(no_attr), 'Разместить на Drive реестр проектов 2016–2021 или добавить атрибуты в файлы показателей.')
    for f_, lab in [('industry', 'отрасль'), ('region', 'регион'), ('model', 'тип организации'), ('grant', 'сумма гранта')]:
        n = sum(1 for p in projects if p[f_] in ('', None))
        if n: add('Низкая', f'Не заполнено поле «{lab}»', n, 'Дозаполнить.')
    nn = sum(1 for p in projects if p['quality'] == 'Нет показателей')
    if nn: add('Средняя', 'Нет ни одного финансового показателя', nn, 'Запросить отчётность у проектных менеджеров.')
    def w(path, fields, rows):
        with open(path, 'w', newline='', encoding='utf-8') as fh:
            wr = csv.DictWriter(fh, fields, extrasaction='ignore'); wr.writeheader(); wr.writerows(rows)
    w(f'{outdir}/projects.csv', ['irn', 'year', 'name', 'grantee', 'direction', 'industry', 'region', 'model', 'period', 'status', 'last_date', 'in_attrs', 'n_snap', 'grant', 'paid'] + METRICS + ['author_fee', 'licenses', 'salary', 'miss', 'quality'], projects)
    w(f'{outdir}/history.csv', ['irn', 'date'] + METRICS, hist)
    w(f'{outdir}/quarterly.csv', ['date', 'n_projects'] + METRICS, quarterly)
    w(f'{outdir}/quality.csv', ['severity', 'finding', 'count', 'recommendation'], F)
    from collections import Counter
    print(json.dumps({'snapshots': dates, 'projects': len(projects), 'quality': Counter(p['quality'] for p in projects), 'status': Counter(p['status'] for p in projects), 'findings': len(F)}, ensure_ascii=False, default=dict))

if __name__ == '__main__': main(*sys.argv[1:4])
