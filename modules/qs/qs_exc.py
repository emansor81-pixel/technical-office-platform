# -*- coding: utf-8 -*-
"""QS - Excavation takeoff from WIR requests (Contractor / Water Networks).
Reads: QS_INPUT.xlsx (SETTINGS, BOQ, LEVELS) + WIR LOG (SW sheet)
Writes: QS_EXC_OUTPUT.xlsx  (detail rows split at segment boundaries, BOQ summary, notes)
Usage:  python qs_exc.py [QS_INPUT.xlsx]
"""
import sys, os, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import qs_core as core
INPUT = sys.argv[1] if len(sys.argv) > 1 else core.DEFAULT_INPUT
EPS = 1e-6

def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def todate(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    return None

def sta(x):
    km = int(x // 1000 + EPS); m = x - km * 1000
    s = f"{m:06.2f}" if abs(m - round(m)) > 0.004 else f"{int(round(m)):03d}"
    return f"{km}+{s}"

# ---------------- read input
wi = openpyxl.load_workbook(INPUT, data_only=True)
S = {r[0]: r[1] for r in wi['SETTINGS'].iter_rows(min_row=2, values_only=True) if r[0]}
S.setdefault('WIR_LOG_PATH', core.DEFAULTS['WIR_LOG_PATH']); S.setdefault('OUTPUT_PATH', core.DEFAULTS['OUTPUT_PATH'])
for _k in core.PATH_KEYS:
    if _k in S: S[_k] = core.project_path(S[_k])
BOQ = [r for r in wi['BOQ'].iter_rows(min_row=2, values_only=True) if r[0] is not None]
LEV = {}
for r in wi['LEVELS'].iter_rows(min_row=2, values_only=True):
    if r[0] is None or num(r[3]) is None or num(r[4]) is None: continue
    key = (str(r[0]).strip(), str(r[1]).strip())
    LEV.setdefault(key, []).append(dict(dia=num(r[2]), a=num(r[3]), b=num(r[4]), da=num(r[6]), db=num(r[7]), w=num(r[8])))
for k in LEV: LEV[k].sort(key=lambda s: s['a'])
cutoff = todate(S.get('CUTOFF_DATE')) or dt.date.today()
appcol = CI(str(S.get('APPROVAL_COL', 'M')).strip())
exc_code = str(S.get('EXC_CODE', 'EXC')).strip()
widths = {300: num(S.get('WIDTH_300')) or 0.85, 400: num(S.get('WIDTH_400')) or 0.95, 600: num(S.get('WIDTH_600')) or 1.25}

# ---------------- read requests
wl = openpyxl.load_workbook(S['WIR_LOG_PATH'], data_only=True)
sw = wl[S.get('WIR_SHEET', 'SW')]
reqs = []
_hdr = [str(sw.cell(2, c).value or '').strip() for c in range(1, 41)]
CODE_COL = _hdr.index('NAM') + 1 if 'NAM' in _hdr else CI('N')
for r in range(4, sw.max_row + 1):
    typ = sw.cell(r, CODE_COL).value
    if typ is None or str(typ).strip() != exc_code: continue
    a, b = num(sw.cell(r, CI('G')).value), num(sw.cell(r, CI('H')).value)
    if a is None or b is None: continue
    if a > b: a, b = b, a
    reqs.append(dict(row=r, no=sw.cell(r, CI('A')).value, dist=str(sw.cell(r, CI('C')).value or '').strip(),
                     street=str(sw.cell(r, CI('D')).value or '').strip(), line=str(sw.cell(r, CI('E')).value or '').strip(),
                     a=a, b=b, dia=num(sw.cell(r, CI('J')).value), reqdate=todate(sw.cell(r, CI('L')).value),
                     appdate=todate(sw.cell(r, appcol).value)))
for q in reqs:
    q['approved'] = q['appdate'] is not None and q['appdate'] <= cutoff
# approved first (by approval date) so overlaps are charged to the earliest approved request
reqs.sort(key=lambda q: (not q['approved'], q['appdate'] or dt.date.max, q['no'] if isinstance(q['no'], (int, float)) else 9e9))

def subtract(iv, covered):
    """iv=(a,b) minus list of covered intervals -> list of free sub-intervals"""
    out = [iv]
    for c0, c1 in covered:
        nxt = []
        for x0, x1 in out:
            if c1 <= x0 + EPS or c0 >= x1 - EPS: nxt.append((x0, x1)); continue
            if c0 > x0 + EPS: nxt.append((x0, c0))
            if c1 < x1 - EPS: nxt.append((c1, x1))
        out = nxt
    return out

rows, notes = [], []
covered = {}
for q in reqs:
    key = (q['line'], q['dist'])
    segs = LEV.get(key)
    if not segs:
        notes.append(f"ريكويست {q['no']}: مفيش مناسيب للخط {q['line']} / {q['dist']} في شيت LEVELS"); continue
    if q['a'] < segs[0]['a'] - EPS or q['b'] > segs[-1]['b'] + EPS:
        notes.append(f"ريكويست {q['no']}: المحطات {sta(q['a'])} - {sta(q['b'])} طالعة بره مدى المناسيب")
    for s in segs:
        a, b = max(q['a'], s['a']), min(q['b'], s['b'])
        if b - a <= EPS: continue
        pieces = [(a, b)]
        if q['approved']:
            cov = covered.setdefault(key, [])
            pieces = subtract((a, b), cov)
            if len(pieces) != 1 or abs(pieces[0][0] - a) > EPS or abs(pieces[0][1] - b) > EPS:
                notes.append(f"ريكويست {q['no']}: جزء من {sta(a)} - {sta(b)} متحسب قبل كده في ريكويست تاني (اتخصم)")
            cov.extend(pieces)
        for x0, x1 in pieces:
            def dep(x):
                if s['da'] is None or s['db'] is None: return None
                return s['da'] + (s['db'] - s['da']) * (x - s['a']) / (s['b'] - s['a'])
            dia = q['dia'] or s['dia']
            w = s['w'] or widths.get(int(dia) if dia else 0)
            rows.append(dict(q=q, x0=x0, x1=x1, d0=dep(x0), d1=dep(x1), dia=dia, w=w))
            if dep(x0) is None:
                notes.append(f"ريكويست {q['no']}: مناسيب ناقصة للسجمنت {sta(s['a'])} - {sta(s['b'])}")
rows.sort(key=lambda r: (r['q']['line'], r['q']['dist'], r['x0'], not r['q']['approved']))

# ---------------- write output
wo = openpyxl.Workbook()
H = PatternFill('solid', fgColor='1F4E78'); SUB = PatternFill('solid', fgColor='DDEBF7')
GRAY = PatternFill('solid', fgColor='EDEDED'); TOT = PatternFill('solid', fgColor='FFF2CC')
thin = Side(style='thin', color='999999'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
ws = wo.active; ws.title = 'حصر الحفر'; ws.sheet_view.rightToLeft = True
ws['A1'] = 'حصر أعمال الحفر من طلبات الاستلام المعتمدة'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = f"تاريخ القطع: {cutoff:%d-%m-%Y}    |    تاريخ التحديث: {dt.datetime.now():%d-%m-%Y %H:%M}"
heads = ['م', 'رقم الطلب', 'الحي', 'الشارع', 'الخط', 'القطر', 'من', 'إلى', 'الطول (م)', 'عمق البداية', 'عمق النهاية',
         'متوسط العمق', 'العرض', 'حتى 1.5 ارتفاع', 'حتى 1.5 كمية', '1.5-2.5 ارتفاع', '1.5-2.5 كمية',
         '2.5-3.5 ارتفاع', '2.5-3.5 كمية', '>3.5 ارتفاع', '>3.5 كمية', 'إجمالي الحفر م3', 'تاريخ الطلب', 'تاريخ الاعتماد', 'الحالة']
wid = [5, 9, 11, 14, 6, 7, 9, 9, 9, 9, 9, 9, 7, 9, 10, 9, 10, 9, 10, 9, 10, 12, 12, 12, 12]
HR = 4
for i, (h, w) in enumerate(zip(heads, wid), 1):
    c = ws.cell(HR, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H; c.border = B
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); ws.column_dimensions[CL(i)].width = w
ws.row_dimensions[HR].height = 32
r = HR + 1
for n, it in enumerate(rows, 1):
    q = it['q']; ok = q['approved'] and it['d0'] is not None
    vals = [n, q['no'], q['dist'], q['street'], q['line'], it['dia'], sta(it['x0']), sta(it['x1']),
            round(it['x1'] - it['x0'], 2), None if it['d0'] is None else round(it['d0'], 3), None if it['d1'] is None else round(it['d1'], 3)]
    for j, v in enumerate(vals, 1): ws.cell(r, j, v)
    if ok:
        ws.cell(r, 12, f'=(J{r}+K{r})/2'); ws.cell(r, 13, it['w'])
        ws.cell(r, 14, f'=IF(L{r}>1.5,1.5,L{r})');                ws.cell(r, 15, f'=I{r}*M{r}*N{r}')
        ws.cell(r, 16, f'=IF(L{r}>=2.5,1,MAX(L{r}-N{r},0))');      ws.cell(r, 17, f'=I{r}*M{r}*P{r}')
        ws.cell(r, 18, f'=IF(L{r}>=3.5,1,MAX(L{r}-N{r}-P{r},0))'); ws.cell(r, 19, f'=I{r}*M{r}*R{r}')
        ws.cell(r, 20, f'=MAX(L{r}-N{r}-P{r}-R{r},0)');            ws.cell(r, 21, f'=I{r}*M{r}*T{r}')
        ws.cell(r, 22, f'=O{r}+Q{r}+S{r}+U{r}')
        status = 'معتمد'
    else:
        ws.cell(r, 13, it['w'])
        status = 'مناسيب ناقصة' if it['d0'] is None else ('غير معتمد' if q['appdate'] is None else 'بعد تاريخ القطع')
    ws.cell(r, 23, q['reqdate']); ws.cell(r, 24, q['appdate']); ws.cell(r, 25, status)
    for j in range(1, 26):
        c = ws.cell(r, j); c.border = B; c.alignment = Alignment(horizontal='center')
        if j in (10, 11, 12, 14, 16, 18, 20): c.number_format = '0.00'
        if j in (9, 15, 17, 19, 21, 22): c.number_format = '#,##0.00'
        if j in (23, 24): c.number_format = 'dd-mm-yyyy'
        if not ok: c.fill = GRAY
    r += 1
last = r - 1
ws.cell(r, 1, 'الإجمالي (المعتمد)').font = Font(bold=True); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
for col in ('I', 'O', 'Q', 'S', 'U', 'V'):
    if col == 'I':
        ws[f'I{r}'] = f'=SUMIF(Y{HR+1}:Y{last},"معتمد",I{HR+1}:I{last})'
    else:
        ws[f'{col}{r}'] = f'=SUM({col}{HR+1}:{col}{last})'
    ws[f'{col}{r}'].number_format = '#,##0.00'
for j in range(1, 26):
    ws.cell(r, j).fill = TOT; ws.cell(r, j).border = B; ws.cell(r, j).font = Font(bold=True)
ws.freeze_panes = ws.cell(HR + 1, 1)
TOTROW = r
# BOQ summary
sm = wo.create_sheet('ملخص البنود'); sm.sheet_view.rightToLeft = True
sh = ['رقم البند', 'كود البند', 'الوصف', 'الوحدة', 'الكمية المعتمدة', 'سعر الوحدة', 'القيمة']
for i, (h, w) in enumerate(zip(sh, [12, 12, 48, 8, 16, 12, 16]), 1):
    c = sm.cell(1, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H; c.border = B
    c.alignment = Alignment(horizontal='center'); sm.column_dimensions[CL(i)].width = w
qcol = {1: 'O', 2: 'Q', 3: 'S', 4: 'U'}
for i, bq in enumerate(BOQ, 2):
    band, d0, d1, item, code, unit, rate, desc = bq[:8]
    vals = [item, code, desc, unit, f"='حصر الحفر'!{qcol[int(band)]}{TOTROW}", rate, f'=E{i}*F{i}']
    for j, v in enumerate(vals, 1):
        c = sm.cell(i, j, v); c.border = B
        if j in (5, 6, 7): c.number_format = '#,##0.00'
t = len(BOQ) + 2
sm.cell(t, 3, 'الإجمالي').font = Font(bold=True)
sm.cell(t, 5, f'=SUM(E2:E{t-1})').number_format = '#,##0.00'; sm.cell(t, 7, f'=SUM(G2:G{t-1})').number_format = '#,##0.00'
for j in range(1, 8): sm.cell(t, j).fill = TOT; sm.cell(t, j).border = B; sm.cell(t, j).font = Font(bold=True)
# notes
nt = wo.create_sheet('ملاحظات'); nt.sheet_view.rightToLeft = True; nt.column_dimensions['A'].width = 110
nt['A1'] = 'ملاحظات المراجعة'; nt['A1'].font = Font(bold=True)
for i, s_ in enumerate(dict.fromkeys(notes) or ['لا توجد ملاحظات'], 2): nt.cell(i, 1, s_)
out = S.get('OUTPUT_PATH') or 'QS_EXC_OUTPUT.xlsx'
os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
wo.save(out)
print(f"OK: {len(reqs)} excavation requests, {len(rows)} rows -> {out}")
for s_ in dict.fromkeys(notes): print(' -', s_)
