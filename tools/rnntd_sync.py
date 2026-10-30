#!/usr/bin/env python3
"""Сведение квартальных файлов «Показатели проектов на ДД.ММ.ГГГГ» (Google Drive) в CSV для дашборда РННТД.
Источник — только Google Drive. Лист «Свод» не используется: итоги считаются по проектам.

    python3 tools/rnntd_sync.py <папка_со_снимками> <проекты_2022-2026.xlsx> <перечень_проектов.xlsx> <папка_вывода>

В папке со снимками лежат файлы с именами ГГГГ-ММ-ДД.xlsx (по одному на отчётную дату).
Результат:
  projects.csv   — все проекты, показатели из последнего снимка, где проект есть; признаки полноты и отчётности
  history.csv    — показатели проекта на каждую отчётную дату
  quarterly.csv  — итоги портфеля по датам (сумма по проектам)
  quality.csv    — замечания по качеству данных
Перечень проектов 2022–2026 (с датами завершения) даёт договор, дату завершения, регион, место реализации,
частного партнёра, ДКТ, менеджера и решения ННС о закрытии; по нему определяется статус проекта.
Колонки находятся по заголовкам, поэтому новые кварталы работают без правок.
"""
import sys, os, re, csv, json, datetime, glob
import openpyxl
from collections import Counter

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
def mln(v): return f'{v/1e6:,.1f}'.replace(',', ' ').replace('.', ',')
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

MODEL_TYPES = ['ВУЗ', 'НИИ', 'НПЦ', 'Бизнес-организация', 'Государственная организация', 'Общественное объединение', 'Физическое лицо']

def model_of(t):
    """Нормализует значение поля «Тип организации» к единым категориям; пусто — если значение не распознано."""
    l = norm(t).lower()
    if l in ('', 'none', '0.0', '3.0', '0', '-', 'договор не заключен'): return ''
    if 'вуз' in l or 'университет' in l or 'учебн' in l: return 'ВУЗ'
    if 'нии' in l: return 'НИИ'
    if l in ('нпц', 'нкц', 'нпо'): return 'НПЦ'
    if 'физ' in l: return 'Физическое лицо'
    if 'общественн' in l: return 'Общественное объединение'
    if 'государствен' in l or l in ('учреждение', 'частное учреждение'): return 'Государственная организация' if 'государствен' in l else 'НПЦ'
    if any(x in l for x in ('тоо', 'ао', 'нао', 'бизнес', 'нпп', 'ип', 'стартап', 'start', 'частного партнера')): return 'Бизнес-организация'
    return ''

def model_by_name(name):
    """Тип организации по её названию (для проектов без поля «Тип организации»)."""
    l = norm(name).lower()
    if not l or l in ('-',): return ''
    if re.search(r'общественн', l): return 'Общественное объединение'
    if re.search(r'университет|university|академи|высш(ая|ей) школ|колледж|казниту|казну|satbayev', l): return 'ВУЗ'
    if re.search(r'научно-исследовательск|\bнии\b|институт', l): return 'НИИ'
    if re.search(r'научно-производствен|научно-технологическ|научно-техническ|научный центр|национальный центр|\bнпц\b|\bнпп\b|\bнптц\b|лаборатори|технопарк', l): return 'НПЦ'
    if re.search(r'общественн|объединени', l): return 'Общественное объединение'
    if re.search(r'\bрг[кп]п\b|\bгкп\b|государственн|учреждени|филиал рг', l): return 'Государственная организация'
    if re.search(r'\b(тоо|ао|нао|ип|ооо|ltd|llp|кх|пк)\b|товарищество|акционерн|компани|корпораци|холдинг|фирма|завод|«|"', l): return 'Бизнес-организация'
    if re.fullmatch(r'[а-яёәіңғүұқөһ\-]+( [а-яёәіңғүұқөһ\-]+){1,3}', l): return 'Физическое лицо'
    return 'Бизнес-организация'

def read_attrs(path):
    ws = openpyxl.load_workbook(path, data_only=True).active; out = {}
    for r in list(ws.iter_rows(values_only=True))[1:]:
        irn = norm(r[1])
        if not IRN_RE.match(irn): continue
        out[key(irn)] = {'industry': industry_of(r[10]), 'region': region_of(r[4], r[5]), 'model': model_of(r[6]), 'period': norm(r[8]), 'applicant': norm(r[3])}
    return out

CLOSE_RE = re.compile(r'отказ|закрыт|прекращ|неисполн|расторг|пр?о?токол|выписк|непредостав|не заключ|возврат', re.I)
SKIP_RE = re.compile(r'Заявитель|Грантополучатель|Частный|Софинансирующая|Наименование проекта|Адрес|Место реализации|Менеджер|ДКТ|БИН|эскроу|банка', re.I)

def read_perechen(path, report_date):
    """Перечень проектов по конкурсам: статус выводится из решений ННС и даты завершения."""
    wb = openpyxl.load_workbook(path, data_only=True); out, dup, bad, sheets = {}, [], [], []
    for ws in wb:
        rows = list(ws.iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows) if r and any('ИРН' in str(c) for c in r if c)), None)
        if hi is None: continue
        head = [norm(c) for c in rows[hi]]
        p = lambda rx: next((i for i, h in enumerate(head) if re.search(rx, h)), None)
        c = {'contract': p(r'договор'), 'end': p(r'^Дата завершения'), 'region': p(r'^Регион|^Область/Город'), 'place': p(r'^Место реализации'),
             'partner': p(r'^Частный партнер|^Софинансирующая'), 'bin': p(r'^БИН'), 'name': p(r'^Наименование проекта'), 'grantee': p(r'^Грантополучатель'),
             'applicant': p(r'^Заявитель'), 'grant': p(r'^Одобренная сумма гранта'), 'grant_2026': p(r'^Сумма гранта на 2026'),
             'cofin_2026': p(r'^Сумма софинансирования'), 'manager': p(r'^Менеджер'), 'dkt': p(r'^ДКТ')}
        m = re.search(r'в\s*реал\.?\s*(\d+)', ws.title); year = re.search(r'20\d\d', ws.title)
        g = lambda r, i: r[i] if i is not None and i < len(r) else None
        cnt = 0
        for r in rows[hi + 1:]:
            irn = norm(r[1])
            if not IRN_RE.match(irn): continue
            notes = [norm(x) for i, x in enumerate(r) if x and i > 1 and i < len(head) and not SKIP_RE.search(head[i]) and CLOSE_RE.search(str(x))]
            notes += [norm(x) for i, x in enumerate(r) if x and i >= len(head) and CLOSE_RE.search(str(x))]
            end = g(r, c['end'])
            if notes: st = 'Закрыт/прекращён'
            elif isinstance(end, datetime.datetime): st = 'Завершён' if end.date().isoformat() <= report_date else 'В реализации'
            else:
                st = 'Не определён'; bad.append(f'{irn}: «{norm(end) or "пусто"}»')
            con = norm(g(r, c['contract']))
            b = g(r, c['bin']); b = (str(int(b)) if isinstance(b, float) else norm(b)).zfill(12) if b not in (None, '', '-') and re.match(r'^\d+(\.0)?$', str(b).strip()) else ''
            dk = norm(g(r, c['dkt'])); dk = re.sub(r'ДКТ\s*', 'ДКТ ', dk) if dk.startswith('ДКТ') else ''
            rec = {'irn': irn, 'p_year': year.group(0) if year else '', 'status': st, 'closure': '; '.join(dict.fromkeys(notes))[:300],
                   'contract': '' if 'не заключ' in con.lower() else con, 'end_date': end.date().isoformat() if isinstance(end, datetime.datetime) else '',
                   'p_region': region_of(g(r, c['region']), ''), 'place': norm(g(r, c['place'])), 'partner': norm(g(r, c['partner'])), 'bin': b,
                   'p_name': norm(g(r, c['name'])), 'p_grantee': norm(g(r, c['grantee'])), 'p_applicant': norm(g(r, c['applicant'])),
                   'p_grant': num(g(r, c['grant'])), 'grant_2026': num(g(r, c['grant_2026'])), 'cofin_2026': num(g(r, c['cofin_2026'])),
                   'manager': norm(g(r, c['manager'])), 'p_dkt': dk}
            if rec['manager'].startswith('ДКТ'): rec['manager'] = ''   # сдвиг колонок в строках закрытых проектов
            if not rec['p_name'] or CLOSE_RE.search(rec['p_name']):
                # строка закрытого проекта сдвинута влево: название — самая длинная ячейка, не похожая на организацию или договор
                ORG = re.compile(r'^(ТОО|НАО|АО|РГП|РГКП|ЧУ|ИП|ДТОО|Дочернее|Товарищество|Некоммерческое|Акционерное|Филиал|Частн|Республиканск|Общественн|ОО|«|")', re.I)
                cells = [(i, norm(r[i])) for i in range(2, min(len(r), 11)) if isinstance(r[i], str)]
                cand = [(i, t) for i, t in cells if len(t) > 25 and not ORG.match(t) and not CLOSE_RE.search(t) and not re.match(r'^№?\s*\d+\s*от', t)]
                if cand:
                    ni, rec['p_name'] = max(cand, key=lambda x: len(x[1]))
                    orgs = [t for i, t in cells if i < ni and i > 2 and ORG.match(t)]
                    if orgs: rec['p_grantee'] = orgs[-1]
                    nums = [num(r[i]) for i in range(ni + 1, min(len(r), ni + 4)) if num(r[i]) != '' and num(r[i]) > 1e6]
                    if nums: rec['p_grant'] = nums[0]
            if key(irn) in out: dup.append(irn)
            out[key(irn)] = rec; cnt += 1
        sheets.append({'title': ws.title, 'year': year.group(0) if year else '', 'declared': int(m.group(1)) if m else None})
    return out, dup, bad, sheets

ORG_FORMS = r'\b(тоо|нао|ао|рго|ргп|на пхв|гкп|ип|чу|ку|ооо|llp|ltd|жшс|частное учреждение|частная компания|некоммерческое акционерное общество|товарищество с ограниченной ответственностью|акционерное общество|республиканское государственное предприятие на праве хозяйственного ведения|республиканское государственное предприятие)\b'
ORG_ALIAS = {'казниту': 'казахский национальный исследовательский технический университет имени к и сатпаева', 'казну': 'казахский национальный университет имени аль фараби'}

def org_key(name):
    """Ключ организации для группировки: без кавычек, организационно-правовой формы и регистра."""
    s = name.lower().replace('ё', 'е')
    s = re.sub(r'[«»"“”„\'`]', ' ', s)
    s = re.sub(ORG_FORMS, ' ', s)
    s = re.sub(r'[^\w ]', ' ', s)
    s = re.sub(r'\bим\b', 'имени', s).replace('ауэзов', 'ауезов')
    s = re.sub(r'\s+', ' ', s).strip()
    for a, full in ORG_ALIAS.items():
        if s.startswith(a + ' ') or s == a: s = full
    return s

def valid_bin(b):
    return bool(re.fullmatch(r'\d{12}', b or '')) and 1 <= int(b[2:4]) <= 12 and not re.fullmatch(r'0+[1-9]0+', b)

def assign_orgs(projects, field, use_bin):
    """Группирует проекты по организации: одинаковый ключ названия или (для грантополучателя) одинаковый БИН."""
    from collections import Counter
    parent = {}
    def find(x):
        while parent.setdefault(x, x) != x: parent[x] = parent[parent[x]]; x = parent[x]
        return x
    def union(a, b): parent[find(a)] = find(b)
    def similar(a, b):   # варианты написания одной организации
        ca, cb = a.replace(' ', '').replace('i', 'і'), b.replace(' ', '').replace('i', 'і')
        return ca[:8] == cb[:8] or a.split()[-1] == b.split()[-1] or ca in cb or cb in ca
    by_bin = {}
    for p in projects:
        k = org_key(p[field]) if p[field] else ''
        if not k: continue
        union('n:' + k, 'n:' + k)
        if use_bin and valid_bin(p.get('bin', '')): by_bin.setdefault(p['bin'], set()).add(k)
    conflicts = []
    for b, ks in by_bin.items():
        ks = sorted(ks)
        for a in ks:
            for c in ks:
                if a < c:
                    if similar(a, c): union('n:' + a, 'n:' + c)
                    else: conflicts.append((b, a, c))
    labels = {}
    for p in projects:
        k = org_key(p[field]) if p[field] else ''
        if not k: p[field + '_id'] = ''; continue
        root = find('n:' + k); p[field + '_id'] = root; labels.setdefault(root, Counter())[re.sub(r'\s+', ' ', p[field]).strip()] += 1
    ids = {r: f'{field[0].upper()}{i + 1:04d}' for i, r in enumerate(sorted(labels))}
    for p in projects:
        r = p[field + '_id']
        p[field + '_id'] = ids.get(r, ''); p[field + '_org'] = labels[r].most_common(1)[0][0] if r else ''
    return conflicts

def main(snapdir, attrs_xlsx, perechen_xlsx, outdir):
    files = sorted(glob.glob(f'{snapdir}/????-??-??.xlsx')); dates = [os.path.basename(f)[:10] for f in files]
    snaps, sheets, dups, errs = {}, {}, {}, {}
    for f, d in zip(files, dates):
        e = []; snaps[d], sheets[d], dups[d] = read_snapshot(f, e); errs[d] = e
    latest = dates[-1]; attrs = read_attrs(attrs_xlsx)
    per, per_dup, per_bad, per_sheets = read_perechen(perechen_xlsx, latest)
    keys = sorted({k for d in dates for k in snaps[d]} | set(per))
    projects, hist, F = [], [], []
    for k in keys:
        a = attrs.get(k, {}); pr = per.get(k, {})
        last = next((d for d in reversed(dates) if k in snaps[d]), None); s_ = snaps[last][k] if last else {}
        rec = {'irn': s_.get('irn') or next(iter([x for x in [pr.get('irn')] if x]), '') or k, 'year': s_.get('sheet') or pr.get('p_year', ''),
               'name': s_.get('name') or pr.get('p_name', ''), 'grantee': s_.get('grantee') or pr.get('p_grantee', ''),
               'direction': pr.get('p_dkt') or s_.get('direction', ''), 'industry': a.get('industry', ''), 'region': pr.get('p_region') or a.get('region', ''),
               'model': a.get('model', ''), 'period': a.get('period', ''), 'status': pr.get('status', 'Нет данных'),
               'report': ('В отчётности' if last == latest else 'Вне отчётности') if last else 'Нет в файлах показателей',
               'last_date': last or '', 'in_attrs': 'да' if a else 'нет', 'in_list': 'да' if pr else 'нет', 'n_snap': sum(1 for d in dates if k in snaps[d])}
        for f_ in ['contract', 'end_date', 'place', 'partner', 'bin', 'manager', 'closure', 'grant_2026', 'cofin_2026']: rec[f_] = pr.get(f_, '')
        rec['applicant'] = pr.get('p_applicant') or a.get('applicant', '')
        for f_ in ['grant', 'paid'] + METRICS + ['author_fee', 'licenses', 'salary']: rec[f_] = s_.get(f_, '')
        if rec['grant'] == '' and pr.get('p_grant', '') != '': rec['grant'] = pr['p_grant']
        miss = [lab for f_, lab in ESSENTIAL.items() if rec[f_] in ('', None)]
        rec['miss'] = '; '.join(miss); rec['quality'] = 'Нет показателей' if all(rec[m] == '' for m in METRICS) else ('Полные' if not miss else 'Частичные')
        projects.append(rec)
        for d in dates:
            if k in snaps[d]:
                h = {'irn': rec['irn'], 'date': d}; h.update({m: snaps[d][k].get(m, '') for m in METRICS}); hist.append(h)
    # отрасль по названию проекта (ручная классификация) — только где в источниках отрасли нет
    byname = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'rnntd_industry_by_name.csv')
    cls = {key(r['irn']): r['industry'] for r in csv.DictReader(open(byname, encoding='utf-8'))} if os.path.exists(byname) else {}
    for p in projects:
        if p['industry']: p['industry_src'] = 'источник'
        elif key(p['irn']) in cls: p['industry'], p['industry_src'] = cls[key(p['irn'])], 'по названию проекта'
        else: p['industry_src'] = ''
        miss = [lab for f_, lab in ESSENTIAL.items() if p[f_] in ('', None)]
        p['miss'] = '; '.join(miss); p['quality'] = 'Нет показателей' if all(p[m] == '' for m in METRICS) else ('Полные' if not miss else 'Частичные')
    for p in projects:   # тип организации: из источника, иначе по названию заявителя, иначе грантополучателя
        if p['model']: p['model_src'] = 'источник'
        elif p['applicant'] and model_by_name(p['applicant']): p['model'], p['model_src'] = model_by_name(p['applicant']), 'по названию заявителя'
        elif p['grantee'] and model_by_name(p['grantee']): p['model'], p['model_src'] = model_by_name(p['grantee']), 'по названию грантополучателя'
        else: p['model_src'] = ''
    assign_orgs(projects, 'applicant', False); bin_conf = assign_orgs(projects, 'grantee', True)
    quarterly = []
    for d in dates:
        q = {'date': d, 'n_projects': len(snaps[d])}
        for m in METRICS: q[m] = sum(v.get(m, 0) or 0 for v in snaps[d].values() if v.get(m, '') != '')
        quarterly.append(q)
    # --- замечания ---
    def add(sev, finding, count, rec): F.append({'severity': sev, 'finding': finding, 'count': count, 'recommendation': rec})
    old = {sh: dt for sh, dt in sheets[latest].items() if dt and dt < latest}
    if old: add('Средняя', 'Листы последнего файла с устаревшей отчётной датой: ' + ', '.join(f'{s} — на {dmy(dt)}' for s, dt in sorted(old.items())), sum(1 for v in snaps[latest].values() if v['sheet'] in old), 'Обновить листы до отчётной даты файла.')
    drop = [p for p in projects if p['report'] == 'Вне отчётности']
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
    # резкий рост: прирост за квартал больше 1 млрд ₸ и больше чем в 20 раз выше прежнего накопленного итога
    jump = []
    for k in keys:
        prev = None
        for d in dates:
            if k in snaps[d]:
                v = snaps[d][k].get('sales', '')
                if v != '' and prev not in (None, '') and v - prev > 1e9 and v > 20 * max(prev, 1): jump.append((snaps[d][k]['irn'], d, prev, v))
                prev = v if v != '' else prev
    if jump: add('Высокая', 'Подозрительный скачок накопленных продаж за квартал (возможна ошибка ввода): ' + '; '.join(f'{i} на {dmy(d)}: {mln(p)} → {mln(v)} млн ₸' for i, d, p, v in jump), len(jump), 'Проверить первичные документы; при ошибке исправить в файле показателей. До проверки значение входит в итоги панели.')
    if decn: add('Средняя', f'Накопленные продажи уменьшились между отчётными датами (сумма снижений {sum(x for _, x in dec)/1e6:,.0f} млн ₸)'.replace(',', ' '), decn, 'Проверить корректировки; накопленные продажи не должны убывать.')
    no_attr = [p for p in projects if p['in_attrs'] == 'нет']
    if no_attr: add('Средняя', 'Нет описательных данных (тип организации; отрасль в источнике): файл «Проекты 2022–2026 г. данные» охватывает только конкурсы 2022–2026', len(no_attr), 'Разместить на Drive реестр проектов 2016–2021 или добавить атрибуты в файлы показателей.')
    for f_, lab in [('industry', 'отрасль'), ('region', 'регион'), ('model', 'тип организации'), ('grant', 'сумма гранта')]:
        n = sum(1 for p in projects if p[f_] in ('', None))
        if n: add('Низкая', f'Не заполнено поле «{lab}»', n, 'Дозаполнить.')
    for sh in per_sheets:
        if sh['declared'] is None: continue
        act = sum(1 for v in per.values() if v['p_year'] == sh['year'] and v['status'] != 'Закрыт/прекращён')
        if act != sh['declared']:
            add('Средняя', f"Перечень, лист «{sh['title']}»: в названии указано «в реал. {sh['declared']}», по строкам незакрытых проектов — {act}", abs(act - sh['declared']), 'Сверить статусы строк с названием листа.')
    if per_bad: add('Средняя', 'Перечень: дата завершения не распознана (статус «Не определён»): ' + '; '.join(per_bad[:8]), len(per_bad), 'Указать дату в формате ДД.ММ.ГГГГ или отметить решение ННС.')
    if per_dup: add('Высокая', 'Перечень: ИРН повторяется: ' + ', '.join(sorted(set(per_dup))), len(per_dup), 'Оставить одну строку на проект.')
    cl_rep = [p for p in projects if p['status'] == 'Закрыт/прекращён' and p['report'] == 'В отчётности']
    if cl_rep: add('Высокая', 'Проекты закрыты/прекращены по перечню, но остаются в последнем файле показателей: ' + ', '.join(p['irn'] for p in cl_rep[:10]), len(cl_rep), 'Подтвердить статус; исключить из отчётности или зафиксировать итоговые показатели.')
    nc = [p for p in projects if p['in_list'] == 'да' and not p['contract'] and p['status'] != 'Закрыт/прекращён']
    if nc: add('Средняя', 'Перечень: у незакрытых проектов нет номера и даты договора', len(nc), 'Внести реквизиты договора.')
    nf = [p for p in projects if p['in_list'] == 'да' and p['report'] == 'Нет в файлах показателей' and p['status'] != 'Закрыт/прекращён']
    if nf: add('Средняя', 'Незакрытые проекты из перечня отсутствуют во всех файлах показателей', len(nf), 'Включить проекты в квартальную отчётность.')
    if bin_conf: add('Средняя', 'Перечень: один БИН указан у разных грантополучателей (вероятно, в графе БИН указан БИН заявителя): ' + '; '.join(f'{b} — {a} / {c}' for b, a, c in bin_conf[:6]), len(bin_conf), 'Проверить БИН грантополучателя; такие проекты на панели не объединены.')
    nbn = sum(1 for p in projects if p.get('industry_src') == 'по названию проекта')
    if nbn: add('Низкая', 'Отрасль не указана в источниках Drive и определена по названию проекта (tools/rnntd_industry_by_name.csv)', nbn, 'Внести отрасль в файлы Drive; классификацию по названию проверить проектным менеджерам.')
    nm = Counter(p['model_src'] for p in projects)
    if nm.get('по названию заявителя') or nm.get('по названию грантополучателя'): add('Низкая', f"Тип организации не указан в источниках и определён по названию: заявителя — {nm.get('по названию заявителя', 0)}, грантополучателя — {nm.get('по названию грантополучателя', 0)} (конкурсы 2016–2018, где заявитель неизвестен)", nm.get('по названию заявителя', 0) + nm.get('по названию грантополучателя', 0), 'Указать тип организации в файлах Drive.')
    nn = sum(1 for p in projects if p['quality'] == 'Нет показателей')
    if nn: add('Средняя', 'Нет ни одного финансового показателя', nn, 'Запросить отчётность у проектных менеджеров.')
    def w(path, fields, rows):
        with open(path, 'w', newline='', encoding='utf-8') as fh:
            wr = csv.DictWriter(fh, fields, extrasaction='ignore'); wr.writeheader(); wr.writerows(rows)
    w(f'{outdir}/projects.csv', ['irn', 'year', 'name', 'grantee', 'applicant', 'applicant_id', 'applicant_org', 'grantee_id', 'grantee_org', 'bin', 'direction', 'manager', 'industry', 'industry_src', 'region', 'place', 'partner', 'model', 'model_src', 'period', 'contract', 'end_date', 'status', 'closure', 'report', 'last_date', 'in_attrs', 'in_list', 'n_snap', 'grant', 'grant_2026', 'cofin_2026', 'paid'] + METRICS + ['author_fee', 'licenses', 'salary', 'miss', 'quality'], projects)
    w(f'{outdir}/history.csv', ['irn', 'date'] + METRICS, hist)
    w(f'{outdir}/quarterly.csv', ['date', 'n_projects'] + METRICS, quarterly)
    w(f'{outdir}/quality.csv', ['severity', 'finding', 'count', 'recommendation'], F)
    print(json.dumps({'snapshots': dates, 'projects': len(projects), 'quality': Counter(p['quality'] for p in projects), 'status': Counter(p['status'] for p in projects), 'report': Counter(p['report'] for p in projects), 'findings': len(F)}, ensure_ascii=False, default=dict))

if __name__ == '__main__': main(*sys.argv[1:5])
