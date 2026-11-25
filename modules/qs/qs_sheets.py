# -*- coding: utf-8 -*-
"""Extra sheets in the QS export, laid out like the consultant's template so they can be copied straight in:
   UNIFIER (WIR Unifier no. per activity per segment), CHAMBERS, CONNECTIONS, BORING."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

THIN = Side(style='thin', color='000000'); MED = Side(style='medium', color='000000')
BRD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR = PatternFill('solid', fgColor='D9DDF0')
TITLE = PatternFill('solid', fgColor='F2F2F2')
ID = PatternFill('solid', fgColor='EDEDED')
RED = PatternFill('solid', fgColor='F8CBAD')
ORG = PatternFill('solid', fgColor='FFE699')
GRN = PatternFill('solid', fgColor='C6EFCE')
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
F_B = Font(bold=True, size=10); F_N = Font(size=10)

# activity groups (template BC..CH, Subbase groups removed). Codes can be overridden in SETTINGS: UNF_<n> = "CODE1, CODE2"
GROUPS = [
    ('الحفر\nExcavation', ['EXC']),
    ('الفرشة أسفل الماسورة\nSand Bedding Under Pipe', ['BED-D-P']),
    ('التمديد\nPipe Installation', ['PIP']),
    ('اختبار الأنابيب\nPipe Test', ['PIP -T-1', 'PIP -T-2']),
    ('الفرشة حول الماسورة\nBedding And Pipe Surround', ['COV-P']),
    ('الردم طبقة\nSubgrade', 'BF'),
    ('قاعدة الطريق طبقة 1\nBase Course 1', ['RB-1']),
    ('قاعدة الطريق طبقة 2\nBase Course 2', ['RB-2']),
    ('الأسفلت\nAsphalt Type 1', 'ASP1_CODE'),                 # كود الطلب من الإعدادات (ASP-1)
    ('الكشط\nMilling', 'ASP2_CODE'),                          # ASP-2-ML
    ('الطبقة السطحية\nWearing Course', 'ASP_WEAR_CODE'),      # نفس طلب الكشط (ASP-2-ML)
]


def _codes(p, i, default):
    v = p.S.get(f'UNF_{i}')
    if v is None or str(v).strip() == '':
        if isinstance(default, str) and default.endswith('_CODE'):
            c = p.code(default)
            return [c] if c else []
        return default
    return [c.strip() for c in str(v).split(',') if c.strip()]


def _ov(a0, a1, b0, b1): return max(0.0, min(a1, b1) - max(a0, b0))


def _approved(p, q):
    return q['appdate'] is not None and q['appdate'] <= p.cutoff


def _hits(p, codes, key, x0, x1):
    """approved requests of these codes that overlap the segment -> list of labels (Unifier no., or (WIR n) if missing)"""
    out, missing = [], False
    for q in p.reqs_all:
        if q['type'] not in codes or q['key'] != key or q['a'] is None or q['b'] is None: continue
        if not _approved(p, q): continue
        a, b = sorted((q['a'], q['b']))
        if _ov(a, b, x0, x1) <= 1e-6: continue
        u = str(q.get('unf') or '').strip()
        if not u:
            n = q.get('no') if str(q.get('no') or '').strip() else q.get('seq')
            u = f"(WIR {n})"; missing = True
        if u not in out: out.append(u)
    return out, missing


def _cell(ws, r, c, v=None, fill=None, font=F_N, brd=BRD):
    x = ws.cell(r, c, v); x.border = brd; x.alignment = CEN; x.font = font
    if fill: x.fill = fill
    return x


def add_unifier(p, wb, rows, items, notes):
    if 'UNIFIER' in wb.sheetnames: del wb['UNIFIER']
    ws = wb.create_sheet('UNIFIER'); ws.sheet_view.rightToLeft = True
    pre = str(p.S.get('BF_PREFIX', 'BF-')).strip()
    nmax = int(float(p.S.get('BF_MAX_LAYERS') or 12))
    # identity columns (not to copy) + gap + block
    idh = ['م', 'المنطقة', 'الخط', 'من', 'إلى', 'الطبقات المطلوبة']
    for j, h in enumerate(idh, 1):
        _cell(ws, 2, j, h, ID, F_B); ws.merge_cells(start_row=2, start_column=j, end_row=3, end_column=j)
        _cell(ws, 3, j, None, ID)
        ws.column_dimensions[CL(j)].width = 9 if j != 2 else 12
    ws.column_dimensions['G'].width = 3
    c0 = 8
    spans = []      # (group index, first col, kind)
    c = c0
    for i, (title, codes) in enumerate(GROUPS, 1):
        if codes == 'BF':
            _cell(ws, 2, c, title, HDR, F_B); [_cell(ws, 2, c + k, None, HDR) for k in (1, 2)]
            ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c + 2)
            for k, h in enumerate(['رقم الطلب\nUN NO', 'عدد\nالطبقات', 'رقم المستخلص\nINV NO']):
                _cell(ws, 3, c + k, h, HDR, F_B)
            ws.column_dimensions[CL(c)].width = 30; ws.column_dimensions[CL(c + 1)].width = 8; ws.column_dimensions[CL(c + 2)].width = 9
            spans.append((i, c, 'BF')); c += 3
        else:
            _cell(ws, 2, c, title, HDR, F_B); _cell(ws, 2, c + 1, None, HDR)
            ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c + 1)
            _cell(ws, 3, c, 'رقم الطلب\nUN NO', HDR, F_B); _cell(ws, 3, c + 1, 'رقم المستخلص\nINV NO', HDR, F_B)
            ws.column_dimensions[CL(c)].width = 11; ws.column_dimensions[CL(c + 1)].width = 9
            spans.append((i, c, _codes(p, i, codes))); c += 2
    last_c = c - 1
    _cell(ws, 1, c0, 'طلبات الاستلام WIR UNIFIER', TITLE, Font(bold=True, size=11))
    for k in range(c0 + 1, last_c + 1): _cell(ws, 1, k, None, TITLE)
    ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=last_c)
    ws.row_dimensions[2].height = 48; ws.row_dimensions[3].height = 40

    bf_items = items.get('bf', [])
    r0 = 4
    vals = {}   # (col) -> list of values per row for vertical merge
    miss = set()
    for i, R in enumerate(rows):
        r = r0 + i
        key = (str(R['C']).strip(), str(R['B']).strip())
        x0, x1 = float(R['E']), float(R['F'])
        # required layers
        need = [f['n'] for f in bf_items if f['q']['key'] == key and f.get('n') and _ov(f['x0'], f['x1'], x0, x1) > 1e-6]
        for j, v in enumerate([R['A'], R['B'], R['C'], x0, x1, max(need) if need else None], 1):
            _cell(ws, r, j, v, ID)
        for gi, col, kind in spans:
            if kind == 'BF':
                parts, got, miss_unf = [], [], False
                for n in range(1, nmax + 1):
                    h, m = _hits(p, [f'{pre}{n}'], key, x0, x1)
                    if h: parts.append(f'{n}) ' + ' + '.join(h)); got.append(n); miss_unf |= m
                    if m: miss.update(x for x in h if x.startswith('(WIR'))
                cnt = 0
                while cnt + 1 in got: cnt += 1
                gap = bool(got) and max(got) > cnt
                a = _cell(ws, r, col, '   '.join(parts) or None, ORG if miss_unf else None)
                a.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
                _cell(ws, r, col + 1, cnt or None, RED if gap else None, F_B)
                _cell(ws, r, col + 2, inv_of(p, key[0], key[1], x0, x1, gi) or None)
                if gap: notes.append(f"UNIFIER: {R['C']} {x0:g}-{x1:g}: طبقات ردم ناقصة في النص (المعتمد {got})")
            else:
                h, m = _hits(p, kind, key, x0, x1) if kind else ([], False)
                v = ' + '.join(h) or None
                _cell(ws, r, col, v, ORG if m else None); _cell(ws, r, col + 1, inv_of(p, key[0], key[1], x0, x1, gi) or None)
                vals.setdefault(col, []).append((v, key))
                if m: miss.update(x for x in h if x.startswith('(WIR'))
    for x in sorted(miss): notes.append(f"UNIFIER: طلب معتمد من غير رقم يونيفاير {x} - اكتب رقمه في سجل الطلبات")
    # vertical merge of identical request numbers on consecutive rows of the same line
    for col, lst in vals.items():
        s = 0
        while s < len(lst):
            e = s
            while e + 1 < len(lst) and lst[e + 1] == lst[s] and lst[s][0]: e += 1
            if e > s: ws.merge_cells(start_row=r0 + s, start_column=col, end_row=r0 + e, end_column=col)
            s = e + 1
    for k in range(1, last_c + 1):
        for rr in (2, 3): ws.cell(rr, k).border = BRD
    ws.freeze_panes = ws.cell(r0, c0)
    ws['A1'] = 'انسخ من العمود H لحد ' + CL(last_c) + ' والصق في BC في شيت حصر مجمع'
    ws['A1'].font = Font(italic=True, size=9, color='808080')
    return ws


# ---------------- chambers (same layout as the consultant's "حصر الغرف")
STAGES = [('أعمال المساحة (توقيع محاور الغرفة)', 0.025), ('الحفر وتسوية قاع الغرفة', 0.025), ('صب الفرشة (خرسانة عادية)', 0.05),
          ('تسليح ونجارة القواعد', 0.05), ('صب القواعد (خرسانة مسلحة)', 0.05), ('تسليح ونجارة الحوائط', 0.05),
          ('صب الحوائط (خرسانة مسلحة)', 0.05), ('تسليح ونجارة السقف', 0.025), ('صب السقف (خرسانة مسلحة)', 0.025),
          ('الأعمال الميكانيكية (Sleeves - الأنابيب - المحابس)', 0.35), ('تركيب غطاء غرفة المحابس', 0.10),
          ('العزل الداخلي (الإيبوكسي)', 0.10), ('العزل الخارجي (الميمبرين)', 0.05), ('الدفان (كل طبقة 30 سم)', 0.05)]
CH_TYPES = [('S2', '33.2.16.2', 1000001), ('S3', '33.2.16.3', 1000002), ('D2', '33.2.17.2', 1000003), ('D3', '33.2.17.3', 1000004),
            ('AV-1', '33.2.18.1', 1000005), ('WO', '33.2.19', 1000006), ('PRV', '33.2.20', 1000007)]


def add_chambers(p, wb, n=20):
    if 'CHAMBERS' in wb.sheetnames: del wb['CHAMBERS']
    ws = wb.create_sheet('CHAMBERS'); ws.sheet_view.rightToLeft = True
    ws['A1'] = 'حصر الغرف - نسب الإنجاز'; ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = 'نوع الغرفة: ' + '   '.join(f'{t} = {it} ({code})' for t, it, code in CH_TYPES); ws['A2'].font = Font(size=9, color='595959')
    fixed = ['م', 'المنطقة ZONE', 'اسم الخط', 'نوع الغرفة', 'القطر', 'المحطة', 'تركيب المحابس', 'البيان']
    for j, h in enumerate(fixed, 1):
        _cell(ws, 4, j, h, HDR, F_B); ws.merge_cells(start_row=4, start_column=j, end_row=5, end_column=j); _cell(ws, 5, j, None, HDR)
    s0 = len(fixed) + 1
    for k, (h, _) in enumerate(STAGES): _cell(ws, 5, s0 + k, h, HDR, F_B); ws.column_dimensions[CL(s0 + k)].width = 11
    groups = [('الأعمال الإنشائية', 0, 9), ('الأعمال الميكانيكية', 9, 11), ('التشطيبات', 11, 14)]
    for t, a, b in groups:
        _cell(ws, 4, s0 + a, t, HDR, F_B)
        for k in range(a + 1, b): _cell(ws, 4, s0 + k, None, HDR)
        ws.merge_cells(start_row=4, start_column=s0 + a, end_row=4, end_column=s0 + b - 1)
    tc = s0 + len(STAGES)
    _cell(ws, 4, tc, 'نسبة إنجاز الغرفة %', HDR, F_B); ws.merge_cells(start_row=4, start_column=tc, end_row=5, end_column=tc); _cell(ws, 5, tc, None, HDR)
    _cell(ws, 6, 1, 'النسب المتفق عليها %', ID, F_B)
    for k in range(2, s0): _cell(ws, 6, k, None, ID)
    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=s0 - 1)
    for k, (_, pc) in enumerate(STAGES): x = _cell(ws, 6, s0 + k, pc, ID, F_B); x.number_format = '0.0%'
    x = _cell(ws, 6, tc, f'=SUM({CL(s0)}6:{CL(tc - 1)}6)', ID, F_B); x.number_format = '0%'
    ws.row_dimensions[5].height = 60
    for col, w in zip('ABCDEFGH', [5, 12, 12, 10, 7, 10, 10, 16]): ws.column_dimensions[col].width = w
    ws.column_dimensions[CL(tc)].width = 12
    r = 7
    data = getattr(p, 'extra', {}).get('CHAMBERS', [])
    n = max(n, len(data))
    for i in range(1, n + 1):
        d_ = data[i - 1] if i <= len(data) else {}
        fixv = [i, d_.get('zone'), d_.get('line'), d_.get('type'), d_.get('dia'), d_.get('station'), d_.get('valves')]
        for j in range(1, s0 - 1):
            _cell(ws, r, j, (fixv[j - 1] or None) if j - 1 < len(fixv) else None)
            for d in (1, 2): _cell(ws, r + d, j, None)
            ws.merge_cells(start_row=r, start_column=j, end_row=r + 2, end_column=j)
        for d, lbl in enumerate(['رقم الطلب UN NO', 'رقم المستخلص INV NO', 'النسبة']):
            _cell(ws, r + d, s0 - 1, lbl, ID, F_B)
        for k in range(len(STAGES)):
            cc = s0 + k; L = CL(cc)
            _cell(ws, r, cc, d_.get(f'un{k + 1}') or None); _cell(ws, r + 1, cc, d_.get(f'inv{k + 1}') or None)
            x = _cell(ws, r + 2, cc, f'=IF({L}{r}<>"",{L}$6,0)'); x.number_format = '0.0%'
        x = _cell(ws, r, tc, f'=SUM({CL(s0)}{r + 2}:{CL(tc - 1)}{r + 2})', None, F_B); x.number_format = '0%'
        for d in (1, 2): _cell(ws, r + d, tc, None)
        ws.merge_cells(start_row=r, start_column=tc, end_row=r + 2, end_column=tc)
        r += 3
    ws.freeze_panes = ws.cell(7, s0)
    return ws


CONN_TYPES = [(300, 110), (300, 150), (300, 160), (300, 300), (300, 600), (300, 1200), (400, 400), (400, 1200), (600, 600), (600, 1200), (300, 1500)]
ZONES = ['ZONE A', 'ZONE B', 'ZONE C', 'ZONE D']   # أسماء المناطق/الأحياء بتاعة مشروعك


def add_connections(p, wb, n=10):
    if 'CONNECTIONS' in wb.sheetnames: del wb['CONNECTIONS']
    ws = wb.create_sheet('CONNECTIONS'); ws.sheet_view.rightToLeft = True
    ws['A1'] = 'حصر أعمال الربط على الخطوط القائمة'; ws['A1'].font = Font(bold=True, size=13)
    for j, h in enumerate(['الرقم\nNO', 'المنطقة\nZone', 'رقم الخط\nLine No'], 1):
        _cell(ws, 3, j, h, HDR, F_B)
        for rr in (4, 5): _cell(ws, rr, j, None, HDR)
        ws.merge_cells(start_row=3, start_column=j, end_row=5, end_column=j)
    _cell(ws, 4, 4, 'خط مقترح', HDR, F_B); _cell(ws, 5, 4, 'خط قائم', HDR, F_B); _cell(ws, 3, 4, None, HDR)
    for k, (a, b) in enumerate(CONN_TYPES):
        cc = 5 + k
        _cell(ws, 3, cc, 'Connection', HDR, F_B); _cell(ws, 4, cc, a, HDR, F_B); _cell(ws, 5, cc, b, HDR, F_B)
        ws.column_dimensions[CL(cc)].width = 11
    tc = 5 + len(CONN_TYPES)
    _cell(ws, 3, tc, 'TOTAL /ZONE', HDR, F_B)
    for rr in (4, 5): _cell(ws, rr, tc, None, HDR)
    ws.merge_cells(start_row=3, start_column=tc, end_row=5, end_column=tc)
    data = getattr(p, 'extra', {}).get('CONNECTIONS', [])
    agg = {}
    for d_ in data:
        k = (d_.get('zone', ''), d_.get('line', ''))
        try: t = (int(float(d_.get('prop'))), int(float(d_.get('exist'))))
        except (TypeError, ValueError): continue
        agg.setdefault(k, {}); agg[k][t] = agg[k].get(t, 0) + 1
    keys = sorted(agg); n = max(n, len(keys))
    for i in range(1, n + 1):
        r = 5 + i
        k = keys[i - 1] if i <= len(keys) else ('', '')
        for j in range(1, tc + 1):
            v = i if j == 1 else (k[0] or None) if j == 2 else (k[1] or None) if j == 3 else None
            if 5 <= j < tc and i <= len(keys): v = agg[k].get(CONN_TYPES[j - 5]) or None
            _cell(ws, r, j, v)
        ws.cell(r, tc).value = f'=SUM(E{r}:{CL(tc - 1)}{r})'
    t = 6 + n
    _cell(ws, t, 1, 'TOTAL', ID, F_B)
    for j in range(2, tc + 1):
        _cell(ws, t, j, f'=SUM({CL(j)}6:{CL(j)}{t - 1})' if j >= 5 else None, ID, F_B)
    for col, w in zip('ABCD', [6, 14, 12, 10]): ws.column_dimensions[col].width = w
    return ws


def add_boring(p, wb):
    if 'BORING' in wb.sheetnames: del wb['BORING']
    ws = wb.create_sheet('BORING'); ws.sheet_view.rightToLeft = True
    ws['A1'] = 'حصر أعمال الثقب الأفقي (33.1.2 - 1000008)'; ws['A1'].font = Font(bold=True, size=13)
    hdr = ['الرقم\nNO', 'المنطقة\nZone', 'رقم الخط\nLine No', 'قطر الخط\nDiameter', 'المحطة\nStation', None, 'Boring Works']
    for j, h in enumerate(hdr, 1):
        _cell(ws, 3, j, h, HDR, F_B)
    ws.merge_cells('E3:F3')
    for j, h in enumerate([None, None, None, None, 'From', 'To', 'LM'], 1):
        _cell(ws, 4, j, h, HDR, F_B)
    for j in (1, 2, 3, 4): ws.merge_cells(start_row=3, start_column=j, end_row=4, end_column=j)
    r = 5
    data = getattr(p, 'extra', {}).get('BORING', [])
    src = [[d_.get('zone'), d_.get('line'), d_.get('dia'), d_.get('x0'), d_.get('x1')] for d_ in data] or [[z, None, None, None, None] for z in ZONES]
    for i, row in enumerate(src, 1):
        vals = [i] + [(float(v) if isinstance(v, str) and v.replace('.', '', 1).isdigit() else v) or None for v in row] + [None]
        for j, v in enumerate(vals, 1): _cell(ws, r, j, v)
        ws.cell(r, 7).value = f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),ABS(F{r}-E{r}),"")'
        r += 1
    _cell(ws, r, 1, 'TOTAL', ID, F_B)
    for j in range(2, 8): _cell(ws, r, j, f'=SUM(G5:G{r - 1})' if j == 7 else None, ID, F_B)
    for col, w in zip('ABCDEFG', [6, 16, 12, 10, 10, 10, 12]): ws.column_dimensions[col].width = w
    return ws


def add_all(p, wb, rows, items, notes):
    add_unifier(p, wb, rows, items, notes)
    add_chambers(p, wb)
    add_connections(p, wb)
    add_boring(p, wb)


# ================= manual data kept in QS_INPUT.xlsx (sheets INV, CHAMBERS, CONNECTIONS, BORING)
CH_RATES = {'S2': 1000, 'S3': 1000, 'D2': 1000, 'D3': 1000, 'AV-1': 1000, 'WO': 1000, 'PRV': 1000}   # sample chamber rates
CH_COLS = ['zone', 'line', 'type', 'dia', 'station', 'valves'] + [f'{k}{i}' for i in range(1, 15) for k in ('un', 'inv')]
CON_COLS = ['zone', 'line', 'station', 'prop', 'exist', 'un', 'inv']
BOR_COLS = ['zone', 'line', 'dia', 'x0', 'x1', 'un', 'inv']
CON_PROP = [100, 150, 300, 400, 600]
CON_EXIST = [110, 150, 160, 300, 400, 600, 1200, 1500]


def _s(v): return '' if v is None else (f'{v:g}' if isinstance(v, float) else str(v)).strip()


def load_extra(p, wb):
    ex = {'INV': {}, 'CHAMBERS': [], 'CONNECTIONS': [], 'BORING': []}
    if 'INV' in wb.sheetnames:
        for r in wb['INV'].iter_rows(min_row=2, values_only=True):
            if not r or r[0] is None or r[5] in (None, ''): continue
            try: ex['INV'][(_s(r[0]), _s(r[1]), round(float(r[2]), 3), round(float(r[3]), 3), int(r[4]))] = _s(r[5])
            except (TypeError, ValueError): pass
    for name, cols in (('CHAMBERS', CH_COLS), ('CONNECTIONS', CON_COLS), ('BORING', BOR_COLS)):
        if name in wb.sheetnames:
            for r in wb[name].iter_rows(min_row=2, values_only=True):
                if not r or all(v in (None, '') for v in r): continue
                ex[name].append({k: _s(r[i]) if i < len(r) else '' for i, k in enumerate(cols)})
    p.extra = ex


def save_extra(p, wb, sheet):
    ex = getattr(p, 'extra', None)
    if not ex: return
    ws = sheet('INV', ['Line', 'District', 'From (m)', 'To (m)', 'Activity #', 'INV No.'], [8, 16, 11, 11, 10, 12])
    r = 2
    for (line, dist, a, b, g), v in sorted(ex['INV'].items()):
        if not v: continue
        for j, x in enumerate((line, dist, a, b, g, v), 1): ws.cell(r, j, x)
        r += 1
    for name, cols in (('CHAMBERS', CH_COLS), ('CONNECTIONS', CON_COLS), ('BORING', BOR_COLS)):
        ws = sheet(name, cols, [10] * len(cols))
        for i, row in enumerate(ex[name], 2):
            for j, k in enumerate(cols, 1):
                v = row.get(k, '')
                try: v = float(v) if k in ('dia', 'station', 'x0', 'x1', 'prop', 'exist') and v != '' else v
                except ValueError: pass
                ws.cell(i, j, v)


def ch_pct(row):
    return sum(pc for i, (_, pc) in enumerate(STAGES, 1) if _s(row.get(f'un{i}')))


def unifier_rows(p, rows, items):
    """per approved segment: dict(line, dist, x0, x1, need, groups=[(text, missing, cnt, gap) ...])"""
    pre = str(p.S.get('BF_PREFIX', 'BF-')).strip()
    nmax = int(float(p.S.get('BF_MAX_LAYERS') or 12))
    bf_items = items.get('bf', [])
    out = []
    for R in rows:
        key = (str(R['C']).strip(), str(R['B']).strip())
        x0, x1 = float(R['E']), float(R['F'])
        need = [f['n'] for f in bf_items if f['q']['key'] == key and f.get('n') and _ov(f['x0'], f['x1'], x0, x1) > 1e-6]
        gs = []
        for i, (title, codes) in enumerate(GROUPS, 1):
            if codes == 'BF':
                parts, got, miss = [], [], False
                for n in range(1, nmax + 1):
                    h, m = _hits(p, [f'{pre}{n}'], key, x0, x1)
                    if h: parts.append(f'{n}) ' + ' + '.join(h)); got.append(n); miss |= m
                cnt = 0
                while cnt + 1 in got: cnt += 1
                gs.append(('   '.join(parts), miss, cnt, bool(got) and max(got) > cnt))
            else:
                cs = _codes(p, i, codes)
                h, m = _hits(p, cs, key, x0, x1) if cs else ([], False)
                gs.append((' + '.join(h), m, None, False))
        out.append(dict(n=R['A'], dist=key[1], line=key[0], x0=x0, x1=x1, need=max(need) if need else None, groups=gs))
    return out


def inv_of(p, line, dist, x0, x1, g):
    return getattr(p, 'extra', {}).get('INV', {}).get((line, dist, round(x0, 3), round(x1, 3), g), '')
