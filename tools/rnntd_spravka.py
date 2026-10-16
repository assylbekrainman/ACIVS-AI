#!/usr/bin/env python3
"""Справки по программе коммерциализации РННТД в Word / Excel / PDF.

    python3 tools/rnntd_spravka.py <папка_с_projects.csv_и_quarterly.csv> <папка_вывода> \
        [--variant summary|industry|region|quarter|all] [--format docx|xlsx|pdf|all] [--year 2024]

Варианты:
  summary  — общая справка: показатели на последнюю дату, динамика за год, топ отраслей и регионов
  industry — разрез по отраслям экономики
  region   — разрез по регионам
  quarter  — квартальная динамика портфеля (лист «Свод») и прирост
Данные берутся из CSV, которые формирует tools/rnntd_sync.py, поэтому справка всегда «на последний квартал».
"""
import csv, sys, argparse, datetime, os
from collections import defaultdict

def fnum(v): return int(v) if v not in ('', None) else 0
def bn(v): return f'{v/1e9:,.1f}'.replace(',', ' ')
def dmy(iso): y, m, d = iso.split('-'); return f'{d}.{m}.{y}'

def load(folder, year=None):
    P = list(csv.DictReader(open(f'{folder}/projects.csv', encoding='utf-8')))
    if year: P = [p for p in P if p['year'] == str(year)]
    Q = list(csv.DictReader(open(f'{folder}/quarterly.csv', encoding='utf-8')))
    return P, Q

def group(P, key):
    g = defaultdict(lambda: defaultdict(int)); 
    for p in P:
        k = p[key] or 'Не указано'; g[k]['n'] += 1
        for f in ('grant', 'paid', 'sales', 'export', 'taxes', 'jobs'): g[k][f] += fnum(p[f])
    rows = [[k, v['n'], bn(v['grant']), bn(v['sales']), bn(v['export']), bn(v['taxes']), v['jobs']] for k, v in g.items()]
    return sorted(rows, key=lambda r: -r[1])

GH = ['', 'Проектов', 'Грант, млрд ₸', 'Продажи, млрд ₸', 'Экспорт, млрд ₸', 'Налоги, млрд ₸', 'Рабочих мест']

def quarter_rows(Q):
    rows = []
    for i, q in enumerate(Q):
        rows.append([dmy(q['date']), bn(fnum(q['sales'])) if q['sales'] else '—', bn(fnum(q['taxes'])), bn(fnum(q['cofin'])),
                     bn(fnum(q['export'])) if q['export'] else '—', fnum(q['jobs']), bn(fnum(q['ntr'])) if q['ntr'] else '—'])
    return rows
QH = ['Дата', 'Продажи, млрд ₸', 'Налоги, млрд ₸', 'Софинансирование, млрд ₸', 'Экспорт, млрд ₸', 'Рабочих мест', 'Вклад в НТР, млрд ₸']

def build(variant, P, Q, year=None):
    last = Q[-1]; prev = Q[-2]; ly = next((q for q in reversed(Q) if q['date'][5:] == last['date'][5:] and q['date'] < last['date']), None)
    asof = dmy(last['date']); scope = f' (конкурс {year} г.)' if year else ''
    doc = {'title': f'Справка по грантовой программе коммерциализации РННТД{scope}', 'subtitle': f'Данные на {asof}', 'blocks': []}
    B = doc['blocks']
    def delta(k):
        a, b = fnum(last[k]), fnum(prev[k]); return a - b if last[k] and prev[k] else None
    if variant in ('summary', 'quarter'):
        s = [['Налоги', bn(fnum(last['taxes'])) + ' млрд ₸'], ['Софинансирование', bn(fnum(last['cofin'])) + ' млрд ₸'],
             ['Экспорт', bn(fnum(last['export'])) + ' млрд ₸'], ['Основные средства', bn(fnum(last['assets'])) + ' млрд ₸'],
             ['Рабочих мест', f"{fnum(last['jobs']):,}".replace(',', ' ')], ['Лицензионных соглашений', last['licenses']]]
        if last['sales']: s.insert(0, ['Продажи', bn(fnum(last['sales'])) + ' млрд ₸'])
        else:
            sm = sum(fnum(p['sales']) for p in P)
            s.insert(0, ['Продажи (сумма по листам проектов)', bn(sm) + ' млрд ₸'])
        B.append(('h', 'Ключевые показатели на ' + asof)); B.append(('t', ['Показатель', 'Значение'], s))
        d = delta('jobs')
        txt = f"За квартал (с {dmy(prev['date'])}) рабочих мест {'+' if d>=0 else ''}{d}; налоги {'+' if delta('taxes')>=0 else ''}{bn(delta('taxes'))} млрд ₸."
        B.append(('p', txt))
    if variant == 'quarter':
        B.append(('h', 'Квартальная динамика портфеля')); B.append(('t', QH, quarter_rows(Q[-12:])))
        if any(q['sales'] == '' for q in Q[-3:]): B.append(('p', 'Примечание: в исходной таблице лист «Свод» содержит ошибку формулы (#VALUE!) в графах «Продажи» или «Вклад в НТР» за последний квартал — значение не показано.'))
    if variant in ('summary', 'industry'):
        r = group(P, 'industry'); B.append(('h', 'Разрез по отраслям экономики')); B.append(('t', ['Отрасль'] + GH[1:], r if variant == 'industry' else r[:8]))
    if variant in ('summary', 'region'):
        r = group(P, 'region'); B.append(('h', 'Разрез по регионам')); B.append(('t', ['Регион'] + GH[1:], r if variant == 'region' else r[:8]))
    B.append(('p', f'Источник: «Показатели проектов на {asof}», проектов в выборке — {len(P)}. Сформировано {datetime.date.today().strftime("%d.%m.%Y")}.'))
    return doc

def to_docx(doc, path):
    import docx
    from docx.shared import Pt
    d = docx.Document(); st = d.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(10.5)
    d.add_heading(doc['title'], 0); d.add_paragraph(doc['subtitle'])
    for b in doc['blocks']:
        if b[0] == 'h': d.add_heading(b[1], 2)
        elif b[0] == 'p': d.add_paragraph(b[1])
        else:
            t = d.add_table(rows=1, cols=len(b[1])); t.style = 'Light Grid Accent 1'
            for i, h in enumerate(b[1]): t.rows[0].cells[i].text = str(h)
            for r in b[2]:
                cs = t.add_row().cells
                for i, v in enumerate(r): cs[i].text = str(v)
    d.save(path)

def to_xlsx(doc, path):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Справка'; ws.append([doc['title']]); ws['A1'].font = Font(bold=True, size=14)
    ws.append([doc['subtitle']]); ws.append([])
    for b in doc['blocks']:
        if b[0] == 'h': ws.append([b[1]]); ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
        elif b[0] == 'p': ws.append([b[1]])
        else:
            ws.append(b[1])
            for c in ws[ws.max_row]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='2F5597'); c.alignment = Alignment(wrap_text=True)
            for r in b[2]: ws.append(list(r))
            ws.append([])
    ws.column_dimensions['A'].width = 42
    for col in 'BCDEFG': ws.column_dimensions[col].width = 20
    wb.save(path)

def to_pdf(doc, path):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pdfmetrics.registerFont(TTFont('DV', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
    pdfmetrics.registerFont(TTFont('DVB', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
    ss = getSampleStyleSheet(); n = ParagraphStyle('n', parent=ss['Normal'], fontName='DV', fontSize=9, leading=12)
    h = ParagraphStyle('h', parent=n, fontName='DVB', fontSize=11, spaceBefore=10, spaceAfter=4); t0 = ParagraphStyle('t', parent=n, fontName='DVB', fontSize=15, leading=19)
    cell = ParagraphStyle('c', parent=n, fontSize=7.5, leading=9)
    S = [Paragraph(doc['title'], t0), Paragraph(doc['subtitle'], n), Spacer(1, 6)]
    for b in doc['blocks']:
        if b[0] == 'h': S.append(Paragraph(b[1], h))
        elif b[0] == 'p': S.append(Paragraph(b[1], n))
        else:
            data = [[Paragraph(str(x), cell) for x in b[1]]] + [[Paragraph(str(x), cell) for x in r] for r in b[2]]
            t = Table(data, repeatRows=1); t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#DCE6F2')), ('GRID', (0, 0), (-1, -1), .3, colors.grey), ('VALIGN', (0, 0), (-1, -1), 'TOP')])); S.append(t)
    SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28).build(S)

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('data'); ap.add_argument('out'); ap.add_argument('--variant', default='all'); ap.add_argument('--format', default='all'); ap.add_argument('--year')
    a = ap.parse_args(); os.makedirs(a.out, exist_ok=True); P, Q = load(a.data, a.year)
    variants = ['summary', 'industry', 'region', 'quarter'] if a.variant == 'all' else [a.variant]
    fmts = ['docx', 'xlsx', 'pdf'] if a.format == 'all' else [a.format]
    for v in variants:
        doc = build(v, P, Q, a.year)
        for f in fmts:
            p = f"{a.out}/spravka_rnntd_{v}_{Q[-1]['date']}{'_' + a.year if a.year else ''}.{f}"
            {'docx': to_docx, 'xlsx': to_xlsx, 'pdf': to_pdf}[f](doc, p); print(p)

if __name__ == '__main__': main()
