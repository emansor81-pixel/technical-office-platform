# -*- coding: utf-8 -*-
"""QS core (Contractor / CLIENT Sample Water Networks)
Items: Excavation (EXC) - Bedding below & around pipe (COV-P) - Pipe Laying (PIP + test + flushing %)
Reads the WIR log (SW sheet) as-is, splits requests every 25 m, exports to the approved QS-TOTAL template.
"""
import os, re, math, shutil, tempfile, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL
from openpyxl.worksheet.page import PageMargins

EPS = 1e-6
HERE = os.path.dirname(os.path.abspath(__file__))
# ---------------- مكان فولدر المشروع: من ملف SETTINGS.json (بيدوّر عليه في الفولدرات اللي فوق)
def _find_settings():
    d = HERE
    for _ in range(4):
        f = os.path.join(d, 'SETTINGS.json')
        if os.path.exists(f): return f
        d = os.path.dirname(d)
    return os.path.join(os.path.dirname(HERE), 'SETTINGS.json')


SETTINGS_JSON = _find_settings()


def _read_project_root():
    try:
        import json
        with open(SETTINGS_JSON, encoding='utf-8-sig') as f:
            return str(json.load(f).get('PROJECT_ROOT', '')).strip()
    except Exception:
        return ''


PROJECT_ROOT = _read_project_root()


def project_path(p):
    """مسار نسبي (زي 08_Cost_Payments\\QS\\QS_INPUT.xlsx) يتحسب من فولدر المشروع. المسار الكامل يفضل زي ما هو."""
    if p is None: return p
    p = str(p).strip()
    if not p or os.path.isabs(p) or re.match(r'^[A-Za-z]:[\\/]', p): return p
    p = p.replace('\\', os.sep).replace('/', os.sep)
    return os.path.join(PROJECT_ROOT or HERE, p)


QS_DIR_REL = r'08_Cost_Payments\QS'
DEFAULT_INPUT = project_path(QS_DIR_REL + r'\QS_INPUT.xlsx')
QS_DIR = project_path(QS_DIR_REL)
PATH_KEYS = ('WIR_LOG_PATH', 'OUTPUT_PATH', 'TEMPLATE_PATH', 'HISTORY_PATH')
ROCK, OTHER = 'Rock', 'Other'
SOIL_AR = {ROCK: 'صخرية', OTHER: 'أخرى'}

DEFAULTS = {
    'WIR_LOG_PATH': r'00_Registers\01-LOG WIR.xlsm',
    'WIR_SHEET': 'SW',
    'APPROVAL_COL': 'M',
    'CUTOFF_DATE': None,
    'OUTPUT_PATH': r'08_Cost_Payments\QS\QS_EXC_OUTPUT.xlsx',
    'TEMPLATE_PATH': os.path.join(HERE, 'QS-TOTAL_TEMPLATE.xlsx'),
    'HISTORY_PATH': r'08_Cost_Payments\QS\QS_HISTORY.xlsx',
    'EXC_CODE': 'EXC',
    'BED_CODE': 'COV-P',
    'BED_BELOW_CODE': 'BED-D-P',   # bedding below the pipe (progress only - the item is paid when COV-P is approved)          # bedding item is counted 100% when the pipe-cover request is approved
    'PIP_CODE': 'PIP',
    'TEST_CODE': 'PIP -T-1',
    'FLUSH_CODE': 'FLU',
    'SEGMENT_LENGTH': 25,
    'WIDTH_300': 0.85, 'WIDTH_400': 0.95, 'WIDTH_600': 1.25,
    'OD_300': 0.30, 'OD_400': 0.40, 'OD_600': 0.60,      # nominal diameter like the consultant template (DI outside: 0.326 / 0.429 / 0.635)
    'BED_ROCK': 0.20, 'BED_OTHER': 0.15, 'COVER_ABOVE': 0.30,
    'PIP_PCT_LAID': 85, 'PIP_PCT_TEST': 90, 'PIP_PCT_FLUSH': 100,
    # Sample rates only – put your own contract rates in the SETTINGS sheet
    'RATE_BED': 100, 'RATE_BF': 100, 'RATE_RB': 100, 'RATE_PIP_300': 100, 'RATE_PIP_400': 100, 'RATE_PIP_600': 100,
    'BF_PREFIX': 'BF-', 'BF_MAX_LAYERS': 12, 'BF_LAYER': 0.30, 'ROAD_BASE': 0.30, 'ROAD_BASE_LAYERS': 0.15,
    'RB_PREFIX': 'RB-', 'RB_LAYERS': 2, 'RB_LAYER': 0.15, 'ASPH_DEFAULT': 10,
    # asphalt (consultant «حصر مجمع»): 32.4.1 = L × (trench width + 0.30) × thickness  ← ASP-1
    #                                  32.4.2 = L × (3.60 − asphalt width) × 0.04       ← milling request (ASP-2-ML)
    #                                  wearing course = same quantity, not paid           ← same request as milling (ASP-2-ML)
    'ASP1_CODE': 'ASP-1', 'ASP2_CODE': 'ASP-2-ML', 'ASP_WEAR_CODE': 'ASP-2-ML', 'MC_CODE': 'MC', 'RC_CODE': 'RC',
    'ASP_EXTRA': 0.30, 'LANE_WIDTH': 3.60, 'WEAR_TH': 0.04, 'RATE_ASP1': 100, 'RATE_ASP2': 100,
}
DEFAULT_BOQ = [(1, '31.2.3.1', 1000009, 100.0), (2, '31.2.3.2', 1000010, 100.0),
               (3, '31.2.3.3', 1000011, 100.0), (4, '31.2.3.4', 1000012, 100.0)]   # sample rates
DIAS = (300, 400, 600)
PI_C = 3.1416          # زي معادلة الاستشاري (3.1416 × نق²)
# PO items for the other works: key -> (item no, PO code, description, unit, rate setting)
PO_ITEMS = {
    'bed': ('31.2.4.3', 1000013, 'Bedding and Initial Backfill for pipes – أعمال الردم الأولية والفرشة الحبيبية للأنابيب', 'm3', 'RATE_BED'),
    'bf': ('31.2.4.4', 1000014, 'Satisfactory fill material for backfilling utility trenches – مواد الردم المناسبة لإعادة ردم خنادق المرافق', 'm3', 'RATE_BF'),
    'rb': ('31.2.4.5', 1000015, 'Selected road subgrade layer – طبقة قاعدة الطريق من المواد المختارة', 'm3', 'RATE_RB'),
    300: ('33.2.3.5', 1000016, 'DI C-Class pipe dia 300 mm – أنابيب حديد دكتايل قطر 300 مم', 'L.M', 'RATE_PIP_300'),
    400: ('33.2.3.6', 1000017, 'DI C-Class pipe dia 400 mm – أنابيب حديد دكتايل قطر 400 مم', 'L.M', 'RATE_PIP_400'),
    600: ('33.2.3.8', 1000018, 'DI C-Class pipe dia 600 mm – أنابيب حديد دكتايل قطر 600 مم', 'L.M', 'RATE_PIP_600'),
    'asp1': ('32.4.1', 1000019, 'Asphalt paving – سفلتة عادية', 'm3', 'RATE_ASP1'),
    'asp2': ('32.4.2', 1000020, 'Asphalt paving with paver and milling machine – سفلتة باستخدام الفرادة والكشاطة', 'm3', 'RATE_ASP2'),
}
# أرقام الطلبات جنب كل بند في الملخص (برّه أعمدة نموذج الاستشاري: AZ..BH)
WIR_COLS = [('AZ', 'الحفر'), ('BA', 'الفرشة أسفل الماسورة'), ('BB', 'الفرشة حول الماسورة'), ('BC', 'التمديد'), ('BD', 'الردم'),
            ('BE', 'قاعدة الطريق'), ('BF', 'الأسفلت'), ('BG', 'الكشط'), ('BH', 'الطبقة السطحية')]
WIR_HDR = 'أرقام الطلبات (WIR)'


def bi_lines(desc):
    """الوصف «English – عربي (…)» على سطرين للطباعة: الإنجليزي (ومعاه الملاحظة اللي بين قوسين) فوق والعربي تحت،
    عشان الاتجاهين ميتلخبطوش في نفس السطر في إكسل والـ PDF."""
    s = str(desc or '').strip()
    if ' – ' not in s: return s
    en, ar = s.split(' – ', 1)
    m = re.search(r'\s*(\([^()]*[A-Za-z%][^()]*\))\s*$', ar)
    if m: en, ar = f'{en} {m.group(1)}', ar[:m.start()].rstrip()
    return f'{en}\n{ar}'


def qround(v):
    """تقريب كميات البنود زي ملخص حصر المقاول: ROUNDDOWN(ROUND(x,4),0) – رقم صحيح لتحت"""
    return float(math.floor(round(v or 0.0, 4) + 1e-9))


def wno(v):
    if v is None or v == '': return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return str(v).strip()


def wjoin(vals):
    out = []
    for v in vals:
        for x in str(v or '').replace('،', ',').split(','):
            x = wno(x.strip())
            if x and x not in out: out.append(x)
    return ', '.join(out)


EXC_DESC = {1: 'Trench excavation to a depth of 1.5m – حفر حتى عمق 1.5 م', 2: 'Trench excavation 1.5m to 2.5m – حفر من 1.5 إلى 2.5 م',
            3: 'Trench excavation 2.5m to 3.5m – حفر من 2.5 إلى 3.5 م', 4: 'Trench excavation deeper than 3.5m – حفر أعمق من 3.5 م'}


def num(v):
    try:
        if v is None or (isinstance(v, str) and not v.strip()): return None
        return float(v)
    except (TypeError, ValueError):
        return None


def todate(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, str):
        for f in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
            try: return dt.datetime.strptime(v.strip(), f).date()
            except ValueError: pass
    return None


def sta(x):
    if x is None: return ''
    km = int(x // 1000 + EPS); m = x - km * 1000
    s = f"{m:06.2f}" if abs(m - round(m)) > 0.004 else f"{int(round(m)):03d}"
    return f"{km}+{s}"


def raw(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): return int(v)
    return v


def kstr(v):
    return str(raw(v)).strip()


def bands(avg):
    h1 = min(avg, 1.5)
    h2 = 1.0 if avg >= 2.5 else max(avg - h1, 0.0)
    h3 = 1.0 if avg >= 3.5 else max(avg - h1 - h2, 0.0)
    h4 = max(avg - h1 - h2 - h3, 0.0)
    return [h1, h2, h3, h4]


def subtract(iv, covered):
    out = [iv]
    for c0, c1 in covered:
        nxt = []
        for x0, x1 in out:
            if c1 <= x0 + EPS or c0 >= x1 - EPS: nxt.append((x0, x1)); continue
            if c0 > x0 + EPS: nxt.append((x0, c0))
            if c1 < x1 - EPS: nxt.append((c1, x1))
        out = nxt
    return out


def split_cover(x0, x1, ivs):
    """split [x0,x1] into parts (a, b, covered?) against a list of intervals"""
    cuts = {x0, x1}
    for a, b in ivs:
        if x0 + EPS < a < x1 - EPS: cuts.add(a)
        if x0 + EPS < b < x1 - EPS: cuts.add(b)
    cuts = sorted(cuts); out = []
    for a, b in zip(cuts[:-1], cuts[1:]):
        m = (a + b) / 2
        cov = any(c0 - EPS <= m <= c1 + EPS for c0, c1 in ivs)
        if out and out[-1][2] == cov and abs(out[-1][1] - a) < EPS: out[-1] = (out[-1][0], b, cov)
        else: out.append((a, b, cov))
    return out


def overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


class Project:
    def __init__(self, input_path):
        self.input_path = input_path
        self.S = dict(DEFAULTS)
        self.boq = list(DEFAULT_BOQ)
        self.points = {}      # (line, district) -> {station: depth}
        self.soil = {}        # (line, district) -> {(from, to): 'Other'}   (default = Rock)
        self.asph = {}        # (line, district) -> {(from, to): thickness (cm)}
        self.reqs_all = []
        self.reqs = []
        self.dirty = False
        self.cutoff = dt.date.today()
        self.mode = 'approved'      # 'approved' = QS-TOTAL ; 'requested' = every WIR in the log counted as if approved
        self.load_input()

    # ---------------- input file
    def load_input(self):
        wb = openpyxl.load_workbook(self.input_path, data_only=True)
        if 'SETTINGS' in wb.sheetnames:
            for r in wb['SETTINGS'].iter_rows(min_row=2, values_only=True):
                if r and r[0] and r[1] not in (None, ''):
                    self.S[str(r[0]).strip()] = r[1]
        for k in PATH_KEYS:
            self.S[k] = project_path(self.S.get(k))
        if 'BOQ' in wb.sheetnames:
            b = []
            for r in wb['BOQ'].iter_rows(min_row=2, values_only=True):
                if r and r[0] is not None and num(r[0]):
                    b.append((int(num(r[0])), r[3], r[4], num(r[6]) or 0.0))
            if b: self.boq = b
        self.points, self.soil, self.asph = {}, {}, {}
        if 'POINTS' in wb.sheetnames:
            for r in wb['POINTS'].iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None or num(r[2]) is None or num(r[3]) is None: continue
                self.points.setdefault((kstr(r[0]), kstr(r[1])), {})[round(num(r[2]), 3)] = num(r[3])
        elif 'LEVELS' in wb.sheetnames:
            for r in wb['LEVELS'].iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None: continue
                key = (kstr(r[0]), kstr(r[1])); p = self.points.setdefault(key, {})
                a, b, da, db = num(r[3]), num(r[4]), num(r[6]), num(r[7])
                if a is not None and da is not None: p[round(a, 3)] = da
                if b is not None and db is not None: p[round(b, 3)] = db
        if 'SOIL' in wb.sheetnames:
            for r in wb['SOIL'].iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None or num(r[2]) is None or num(r[3]) is None: continue
                if str(r[4]).strip().lower().startswith('o') or str(r[4]).strip() == SOIL_AR[OTHER]:
                    self.soil.setdefault((kstr(r[0]), kstr(r[1])), {})[(round(num(r[2]), 3), round(num(r[3]), 3))] = OTHER
        try:
            import qs_sheets; qs_sheets.load_extra(self, wb)
        except Exception:
            self.extra = {'INV': {}, 'CHAMBERS': [], 'CONNECTIONS': [], 'BORING': []}
        if 'ASPHALT' in wb.sheetnames:
            for r in wb['ASPHALT'].iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None or num(r[2]) is None or num(r[3]) is None or num(r[4]) is None: continue
                self.asph.setdefault((kstr(r[0]), kstr(r[1])), {})[(round(num(r[2]), 3), round(num(r[3]), 3))] = num(r[4])
        self.seg_len = num(self.S.get('SEGMENT_LENGTH')) or 25.0
        self.widths, self.ods = {}, {}
        for k, v in self.S.items():
            if k.startswith('WIDTH_') and k[6:].isdigit() and num(v): self.widths[int(k[6:])] = num(v)
            if k.startswith('OD_') and k[3:].isdigit() and num(v): self.ods[int(k[3:])] = num(v)
        self.dirty = False

    def save_points(self):
        wb = openpyxl.load_workbook(self.input_path)
        H = PatternFill('solid', fgColor='1F4E78'); thin = Side(style='thin', color='999999')
        def sheet(name, hdr, widths):
            if name in wb.sheetnames: del wb[name]
            ws = wb.create_sheet(name)
            for i, h in enumerate(hdr, 1):
                c = ws.cell(1, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H
                c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                ws.column_dimensions[CL(i)].width = widths[i - 1]
            ws.freeze_panes = 'A2'
            return ws
        ws = sheet('POINTS', ['Line', 'District', 'Station (m)', 'Excavation Depth (m)', 'Station'], [8, 16, 12, 20, 12])
        r = 2
        for (line, dist), pts in sorted(self.points.items()):
            for x in sorted(pts):
                for j, v in enumerate((line, dist, x, pts[x], sta(x)), 1): ws.cell(r, j, v)
                r += 1
        ws = sheet('SOIL', ['Line', 'District', 'From (m)', 'To (m)', 'Soil (Rock = default)'], [8, 16, 11, 11, 20])
        r = 2
        for (line, dist), d in sorted(self.soil.items()):
            for (a, b), v in sorted(d.items()):
                for j, x in enumerate((line, dist, a, b, v), 1): ws.cell(r, j, x)
                r += 1
        ws = sheet('ASPHALT', ['Line', 'District', 'From (m)', 'To (m)', 'Asphalt Thickness (cm)'], [8, 16, 11, 11, 22])
        r = 2
        for (line, dist), d in sorted(self.asph.items()):
            for (a, b), v in sorted(d.items()):
                for j, x in enumerate((line, dist, a, b, v), 1): ws.cell(r, j, x)
                r += 1
        import qs_sheets
        qs_sheets.save_extra(self, wb, sheet)
        wb.save(self.input_path)
        self.dirty = False

    # ---------------- requests (read as-is)
    def load_requests(self):
        src = str(self.S['WIR_LOG_PATH'])
        tmp = os.path.join(tempfile.gettempdir(), 'qs_wir_copy' + os.path.splitext(src)[1])
        shutil.copy2(src, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
        ws = wb[str(self.S.get('WIR_SHEET', 'SW'))]
        ac = CI(str(self.S.get('APPROVAL_COL', 'M')).strip()) - 1
        reqs = []
        # columns found by the markers row (NAM = request code, UNF = Unifier no., S NAM = title) – searched in the
        # first rows, so rows added above it in the log don't break anything (fallback: row 2, columns N / none)
        hdr, hrow = [], 2
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, max_col=40, values_only=True), 1):
            h = [str(v).strip() if v is not None else '' for v in row]
            if 'NAM' in h and ('S NAM' in h or 'UNF' in h): hdr, hrow = h, i; break
        cc = hdr.index('NAM') if 'NAM' in hdr else 13
        uc = hdr.index('UNF') if 'UNF' in hdr else None
        fc = hdr.index('S NAM') if 'S NAM' in hdr else cc + 1   # full request title (file name)
        self.code_col, self.unf_col = cc, uc
        width = max(16, ac + 1, cc + 1, fc + 1, (uc or 0) + 1)
        for i, r in enumerate(ws.iter_rows(min_row=hrow + 2, max_col=width, values_only=True), hrow + 2):
            r = list(r) + [None] * width
            if r[cc] is None or not str(r[cc]).strip(): continue
            reqs.append(dict(row=i, seq=raw(r[0]), no=raw(r[1]), dist=raw(r[2]), street=raw(r[3]), line=raw(r[4]),
                             desc=raw(r[5]), a_raw=raw(r[6]), b_raw=raw(r[7]), a=num(r[6]), b=num(r[7]),
                             dia_raw=raw(r[9]), dia=num(r[9]), reqdate=todate(r[11]), appdate=todate(r[ac]),
                             type=str(r[cc]).strip(), unf=raw(r[uc]) if uc is not None else '', title=raw(r[fc]), L_raw=raw(r[8]), rev=raw(r[10]),
                             key=(kstr(r[4]), kstr(r[2]))))
        wb.close()
        try: os.remove(tmp)
        except OSError: pass
        self.reqs_all = reqs
        self.reqs = self.of('EXC_CODE')
        return reqs

    def code(self, k): return str(self.S.get(k, DEFAULTS.get(k, ''))).strip()

    def of(self, codekey):
        c = self.code(codekey)
        return [q for q in self.reqs_all if q['type'] == c]

    # ---------------- depths / soil
    def depth_at(self, key, x):
        p = self.points.get(key, {})
        for s, d in p.items():
            if abs(s - x) < 0.005: return d, False
        g0 = (x // self.seg_len) * self.seg_len; g1 = g0 + self.seg_len
        d0 = d1 = None
        for s, d in p.items():
            if abs(s - g0) < 0.005: d0 = d
            if abs(s - g1) < 0.005: d1 = d
        if d0 is not None and d1 is not None:
            return d0 + (d1 - d0) * (x - g0) / self.seg_len, True
        return None, False

    def set_depth(self, key, x, val):
        p = self.points.setdefault(key, {})
        for s in list(p):
            if abs(s - x) < 0.005: del p[s]
        if val is not None: p[round(x, 3)] = val
        self.dirty = True

    def soil_at(self, key, x):
        for (a, b), v in self.soil.get(key, {}).items():
            if a - EPS <= x <= b + EPS: return v
        return ROCK

    def set_soil(self, key, x0, x1, val):
        d = self.soil.setdefault(key, {})
        for (a, b) in list(d):
            if overlap(a, b, x0, x1) > EPS: del d[(a, b)]
        if val == OTHER: d[(round(x0, 3), round(x1, 3))] = OTHER
        self.dirty = True

    def asph_at(self, key, x):
        for (a, b), v in self.asph.get(key, {}).items():
            if a - EPS <= x <= b + EPS: return v
        return None

    def set_asph(self, key, x0, x1, val):
        d = self.asph.setdefault(key, {})
        for (a, b) in list(d):
            if overlap(a, b, x0, x1) > EPS: del d[(a, b)]
        if val is not None: d[(round(x0, 3), round(x1, 3))] = val
        self.dirty = True

    def width(self, dia): return self.widths.get(int(dia)) if dia else None
    def od(self, dia): return self.ods.get(int(dia)) if dia else None

    # ---------------- generic request -> 25 m pieces (approved overlaps counted once)
    def pieces(self, reqs, label, notes):
        cutoff = self.cutoff
        qs = []
        for q in reqs:
            if q['a'] is None or q['b'] is None:
                notes.append(f"{label} WIR {q['no']}: From/To missing in requests sheet"); continue
            a, b = sorted((q['a'], q['b']))
            if b - a < EPS:
                notes.append(f"{label} WIR {q['no']}: From = To ({q['a_raw']})"); continue
            q2 = dict(q, x0=a, x1=b)
            q2['approved'] = self.ok(q)
            q2['after_cutoff'] = self.mode == 'approved' and q['appdate'] is not None and q['appdate'] > cutoff
            qs.append(q2)
        qs.sort(key=lambda q: (not q['approved'], q['appdate'] or dt.date.max,
                               q['no'] if isinstance(q['no'], (int, float)) else 9e9))
        covered, out = {}, []
        L = self.seg_len
        for q in qs:
            a, b = q['x0'], q['x1']
            cuts = [a] + [k * L for k in range(int(a // L) + 1, int(b // L) + 1) if a + EPS < k * L < b - EPS] + [b]
            for x0, x1 in zip(cuts[:-1], cuts[1:]):
                ps = [(x0, x1)]
                if q['approved']:
                    cov = covered.setdefault(q['key'], [])
                    ps = subtract((x0, x1), cov)
                    if not (len(ps) == 1 and abs(ps[0][0] - x0) < EPS and abs(ps[0][1] - x1) < EPS):
                        notes.append(f"{label} WIR {q['no']}: part of {sta(x0)} - {sta(x1)} already counted in another approved WIR (deducted)")
                    cov.extend(ps)
                out += [(q, p0, p1) for p0, p1 in ps]
        return out

    def ok(self, q):
        if self.mode == 'requested': return q['reqdate'] is None or q['reqdate'] <= self.cutoff
        return q['appdate'] is not None and q['appdate'] <= self.cutoff

    def approved_ivs(self, codekey):
        d = {}
        for q in self.of(codekey):
            if q['a'] is None or q['b'] is None or not self.ok(q): continue
            a, b = sorted((q['a'], q['b']))
            d.setdefault(q['key'], []).append((a, b))
        return d

    def ivs_code(self, code):
        d = {}
        for q in self.reqs_all:
            if q['type'] != code or q['a'] is None or q['b'] is None: continue
            if not self.ok(q): continue
            a, b = sorted((q['a'], q['b']))
            d.setdefault(q['key'], []).append((a, b, q['no']))
        return d

    def wir_at(self, codekey, key, x0, x1):
        out = []
        for q in self.of(codekey):
            if q['key'] != key or q['a'] is None or q['b'] is None: continue
            if not self.ok(q): continue
            a, b = sorted((q['a'], q['b']))
            if overlap(a, b, x0, x1) > EPS: out.append(str(q['no']))
        return ', '.join(dict.fromkeys(out))

    def status_of(self, q):
        if q['approved']: return 'Approved'
        return 'After Cut-off' if q['after_cutoff'] else 'Not Approved'

    # ---------------- build all items
    def build(self, cutoff=None):
        self.cutoff = cutoff or todate(self.S.get('CUTOFF_DATE')) or dt.date.today()
        notes = []
        exc = [self.calc_exc(q, a, b) for q, a, b in self.pieces(self.of('EXC_CODE'), 'EXC', notes)]
        exc.sort(key=lambda r: (kstr(r['q']['dist']), r['q']['key'][0], r['x0'], not r['q']['approved']))
        for r in exc:
            if r['status'] == 'Missing Levels':
                notes.append(f"EXC WIR {r['q']['no']}: missing depth at {sta(r['x0'])} - {sta(r['x1'])}")
            if r['status'] == 'No Width':
                notes.append(f"EXC WIR {r['q']['no']}: no trench width for dia {r['q']['dia_raw']} (add WIDTH_{r['q']['dia_raw']} in SETTINGS)")
        exc_iv = self.approved_ivs('EXC_CODE')
        # bedding: each part is paid when ITS OWN request is approved (2026-11-04), only on approved excavation
        #   below the pipe  -> BED-D-P  (part='below', qty = below)
        #   around + cover  -> COV-P    (part='around', qty = net)
        bed = []
        for code, part in (('BED_BELOW_CODE', 'below'), ('BED_CODE', 'around')):
            for q, a, b in self.pieces(self.of(code), 'BED', notes):
                for p0, p1, cov in (split_cover(a, b, exc_iv.get(q['key'], [])) if q['approved'] else [(a, b, True)]):
                    r = self.calc_bed(q, p0, p1)
                    r['part'] = part
                    if r['total'] is not None:
                        if part == 'below': r['net'] = 0.0; r['gross'] = 0.0; r['pipe'] = 0.0; r['total'] = r['below']
                        else: r['below'] = 0.0; r['total'] = r['net']
                    if q['approved'] and not cov:
                        r['status'] = 'Ahead of EXC'
                        notes.append(f"BED WIR {q['no']}: {sta(p0)} - {sta(p1)} has no approved excavation (not counted)")
                    bed.append(r)
        bed.sort(key=lambda r: (kstr(r['q']['dist']), r['q']['key'][0], r['x0'], not r['q']['approved']))
        # pipe laying – %: laid 85 / tested 90 / flushed 100, only on approved excavation
        test_iv, flush_iv = self.approved_ivs('TEST_CODE'), self.approved_ivs('FLUSH_CODE')
        pip = []
        for q, a, b in self.pieces(self.of('PIP_CODE'), 'PIP', notes):
            parts = split_cover(a, b, exc_iv.get(q['key'], [])) if q['approved'] else [(a, b, True)]
            for p0, p1, cov in parts:
                if q['approved'] and cov:
                    for t0, t1, tc in split_cover(p0, p1, test_iv.get(q['key'], [])):
                        for f0, f1, fc in split_cover(t0, t1, flush_iv.get(q['key'], [])):
                            pip.append(self.calc_pip(q, f0, f1, tc, fc))
                else:
                    r = self.calc_pip(q, p0, p1, False, False)
                    if q['approved']:
                        r['status'] = 'Ahead of EXC'
                        notes.append(f"PIP WIR {q['no']}: {sta(p0)} - {sta(p1)} has no approved excavation (not counted)")
                    pip.append(r)
        pip.sort(key=lambda r: (kstr(r['q']['dist']), r['q']['key'][0], r['x0'], not r['q']['approved']))
        bf = self.build_bf(exc, notes)
        rb = self.build_rb(exc, notes)
        asp = self.build_asp(exc, notes)
        bedp = self.build_bedp(exc)
        self.items = dict(exc=exc, bed=bed, pip=pip, bf=bf, rb=rb, asp=asp, bedp=bedp)
        return self.items, list(dict.fromkeys(notes))

    # ---------------- bedding progress: below (BED-D-P) + around (COV-P) -> % done and value done
    def build_bedp(self, exc):
        biv = self.ivs_code(str(self.S.get('BED_BELOW_CODE', 'BED-D-P')).strip())
        civ = self.ivs_code(str(self.S.get('BED_CODE', 'COV-P')).strip())
        rate = self.rate('bed'); out = []
        for e in exc:
            q = e['q']
            if not q['approved']: continue
            key = q['key']
            cuts = {e['x0'], e['x1']}
            for iv in (biv, civ):
                for a, b, _ in iv.get(key, []):
                    if e['x0'] + EPS < a < e['x1'] - EPS: cuts.add(a)
                    if e['x0'] + EPS < b < e['x1'] - EPS: cuts.add(b)
            cuts = sorted(cuts)
            for a, b in zip(cuts[:-1], cuts[1:]):
                m = (a + b) / 2
                bw = [str(no) for c0, c1, no in biv.get(key, []) if c0 - EPS <= m <= c1 + EPS]
                cw = [str(no) for c0, c1, no in civ.get(key, []) if c0 - EPS <= m <= c1 + EPS]
                r = self.calc_bed(q, a, b)
                r.update(below_ok=bool(bw), cov_ok=bool(cw), wirs=', '.join(dict.fromkeys(bw + cw)))
                if r['total'] is None:
                    r.update(done=0.0, pct=None, value=0.0, inv=0.0, status='No Width'); out.append(r); continue
                if r['cov_ok']: done = r['total']                     # item closed -> below + around
                elif r['below_ok']: done = r['below']
                else: done = 0.0
                r.update(done=done, pct=done / r['total'] if r['total'] else None, value=done * rate,
                         inv=(r['below'] if r['below_ok'] else 0.0) + (r['net'] if r['cov_ok'] else 0.0))   # كل جزء بطلبه
                r['status'] = 'Closed' if r['cov_ok'] else ('Below Done' if r['below_ok'] else 'Not Started')
                out.append(r)
        return out

    # ---------------- road base layers (RB-1, RB-2): each layer = length x trench width x 0.15, own BOQ item
    def build_rb(self, exc, notes):
        pre = str(self.S.get('RB_PREFIX', 'RB-')).strip()
        nl = int(num(self.S.get('RB_LAYERS')) or 2)
        th = num(self.S.get('RB_LAYER')) or 0.15
        ivs = {k: self.ivs_code(f'{pre}{k}') for k in range(1, nl + 1)}
        out = []
        for e in exc:
            q = e['q']
            if not q['approved']: continue
            key = q['key']
            base = dict(q=q, x0=e['x0'], x1=e['x1'], L=e['L'], w=e['w'], th=th, nl=nl, layers=[], q1=0.0, q2=0.0,
                        qty=0.0, wirs='')
            if not e['w']:
                out.append(dict(base, status='No Width')); continue
            cuts = {e['x0'], e['x1']}
            for k in ivs:
                for a, b, _ in ivs[k].get(key, []):
                    if e['x0'] + EPS < a < e['x1'] - EPS: cuts.add(a)
                    if e['x0'] + EPS < b < e['x1'] - EPS: cuts.add(b)
            cuts = sorted(cuts)
            for a, b in zip(cuts[:-1], cuts[1:]):
                m = (a + b) / 2
                got, wirs = [], []
                for k in ivs:
                    hit = [no for c0, c1, no in ivs[k].get(key, []) if c0 - EPS <= m <= c1 + EPS]
                    if hit: got.append(k); wirs += [str(x) for x in hit]
                L = b - a
                lq = {k: (L * e['w'] * th if k in got else 0.0) for k in range(1, nl + 1)}
                r = dict(base, x0=a, x1=b, L=L, layers=got, q1=lq.get(1, 0.0), q2=lq.get(2, 0.0), qty=sum(lq.values()),
                         wirs=', '.join(dict.fromkeys(wirs)))
                r['status'] = 'Complete' if len(got) >= nl else ('Partial' if got else 'No Layers')
                out.append(r)
        return out

    # ---------------- asphalt – like the consultant sheet (AS..AY), on approved excavation, each part by its own approved request
    #   32.4.1 normal asphalt (ASP-1)    = L × AS × AT    AS = trench width + 0.30   AT = thickness (cm ÷ 100)
    #   32.4.2 paver & milling (ASP-2-ML) = L × AV × AW    AV = 3.60 − AS             AW = 0.04
    #   wearing course (ASP-2-ML too)    = L × AV × AW    shown only – not in the consultant summary (not paid)
    #   (the trench strip's wearing is inside AT: the consultant counts both layers there, 10 cm, under 32.4.1)
    #   prime coats MC / RC have no quantity – kept to go with their item's WIRs in the IPC
    def asp_dims(self, w, asph):
        extra = num(self.S.get('ASP_EXTRA'))
        if extra is None: extra = 0.30
        lane = num(self.S.get('LANE_WIDTH')) or 3.60
        AS = (w or 0.0) + extra
        return AS, (asph or 0.0) / 100, max(lane - AS, 0.0), num(self.S.get('WEAR_TH')) or 0.04

    def build_asp(self, exc, notes):
        codes = dict(b=self.code('ASP1_CODE'), m=self.code('ASP2_CODE'), wc=self.code('ASP_WEAR_CODE'),
                     mc=self.code('MC_CODE'), rc=self.code('RC_CODE'))
        ivs = {k: (self.ivs_code(c) if c else {}) for k, c in codes.items()}
        out = []
        for e in exc:
            q = e['q']
            if not q['approved']: continue
            key = q['key']
            base = dict(q=q, x0=e['x0'], x1=e['x1'], L=e['L'], w=e['w'], asph=e['asph'], asph_def=e.get('asph_def', False),
                        AS=None, AT=None, AV=None, AW=None, f1=0.0, f2=0.0, q1=0.0, q2=0.0, q3=0.0,
                        w1='', w2='', w3='', wmc='', wrc='', wirs='')
            if not e['w']:
                out.append(dict(base, status='No Width')); continue
            AS, AT, AV, AW = self.asp_dims(e['w'], e['asph'])
            cuts = {e['x0'], e['x1']}
            for k in ivs:
                for a, b, _ in ivs[k].get(key, []):
                    if e['x0'] + EPS < a < e['x1'] - EPS: cuts.add(a)
                    if e['x0'] + EPS < b < e['x1'] - EPS: cuts.add(b)
            cuts = sorted(cuts)
            for a, b in zip(cuts[:-1], cuts[1:]):
                m = (a + b) / 2
                hit = {k: [wno(no) for c0, c1, no in ivs[k].get(key, []) if c0 - EPS <= m <= c1 + EPS] for k in ivs}
                L = b - a
                f1, f2 = L * AS * AT, L * AV * AW
                r = dict(base, x0=a, x1=b, L=L, AS=AS, AT=AT, AV=AV, AW=AW, f1=f1, f2=f2,
                         q1=f1 if hit['b'] else 0.0, q2=f2 if hit['m'] else 0.0, q3=f2 if hit['wc'] else 0.0,
                         w1=wjoin(hit['b']), w2=wjoin(hit['m']), w3=wjoin(hit['wc']), wmc=wjoin(hit['mc']), wrc=wjoin(hit['rc']),
                         wirs=wjoin(hit['b'] + hit['m'] + hit['wc']))
                done = [bool(hit['b']), bool(hit['m']), bool(hit['wc'])]
                r['status'] = 'Complete' if all(done) else ('Partial' if any(done) else 'Not Started')
                if hit['b'] and e['asph'] is None:
                    notes.append(f"ASP: {sta(a)} - {sta(b)} asphalt thickness not entered (taken 0)")
                out.append(r)
        return out

    # ---------------- backfill by layers (BF-1 .. BF-10) on approved excavation
    def build_bf(self, exc, notes):
        pre = str(self.S.get('BF_PREFIX', 'BF-')).strip()
        nmax = int(num(self.S.get('BF_MAX_LAYERS')) or 12)
        lay = num(self.S.get('BF_LAYER')) or 0.30
        road = num(self.S.get('ROAD_BASE')) or 0.30           # consultant formula -> paid quantity
        road_n = num(self.S.get('ROAD_BASE_LAYERS'))          # real height -> number of layers
        if road_n is None: road_n = 0.15
        cover = num(self.S.get('COVER_ABOVE')) or 0.30
        ivs = {k: self.ivs_code(f'{pre}{k}') for k in range(1, nmax + 1)}
        out = []
        for e in exc:
            q = e['q']
            if not q['approved']: continue
            key = q['key']
            base = dict(q=q, x0=e['x0'], x1=e['x1'], L=e['L'], avg=e['avg'], asph=e['asph'], soil=e['soil'], w=e['w'],
                        t=None, ch=None, H=None, Hn=None, n=None, layers=[], cnt=0, pct=None, qty=None, paid=None, wirs='')
            if e['status'] != 'Approved':
                out.append(dict(base, status='Missing Levels' if e['status'] == 'Missing Levels' else e['status'])); continue
            od = self.od(q['dia'])
            if not od:
                out.append(dict(base, status='No Width')); continue
            t = num(self.S.get('BED_ROCK' if e['soil'] == ROCK else 'BED_OTHER'))
            ch = od + cover
            H = e['avg'] - road - (e['asph'] or 0) / 100 - t - ch        # quantity height (consultant)
            Hn = e['avg'] - road_n - (e['asph'] or 0) / 100 - t - ch     # real backfill height (layers)
            n = 0 if Hn <= 0.005 else int(math.ceil(Hn / lay - 1e-6))
            if e['asph'] is None:
                notes.append(f"BF: {sta(e['x0'])} - {sta(e['x1'])} asphalt thickness not entered (taken 0)")
            # split the excavation piece at the ends of approved layer WIRs
            cuts = {e['x0'], e['x1']}
            for k in ivs:
                for a, b, _ in ivs[k].get(key, []):
                    if e['x0'] + EPS < a < e['x1'] - EPS: cuts.add(a)
                    if e['x0'] + EPS < b < e['x1'] - EPS: cuts.add(b)
            cuts = sorted(cuts)
            for a, b in zip(cuts[:-1], cuts[1:]):
                m = (a + b) / 2
                got, wirs = [], []
                for k in ivs:
                    hit = [no for c0, c1, no in ivs[k].get(key, []) if c0 - EPS <= m <= c1 + EPS]
                    if hit: got.append(k); wirs += [str(x) for x in hit]
                cnt = len([k for k in got if k <= n])
                extra = [k for k in got if k > n]
                if extra:
                    notes.append(f"BF: {sta(a)} - {sta(b)} has approved layer(s) {extra} but needs only {n} layer(s)")
                L = b - a
                pct = (cnt / n) if n else None
                qty = L * e['w'] * H if H > 0 else 0.0
                r = dict(base, x0=a, x1=b, L=L, t=t, ch=ch, H=H, Hn=Hn, n=n, layers=got, cnt=cnt, pct=pct, qty=qty,
                         paid=qty * pct if pct is not None else 0.0, wirs=', '.join(dict.fromkeys(wirs)))
                r['status'] = 'No Backfill' if n == 0 else ('Complete' if cnt >= n else ('Partial' if cnt else 'No Layers'))
                out.append(r)
        return out

    def calc_exc(self, q, x0, x1):
        d0, i0 = self.depth_at(q['key'], x0)
        d1, i1 = self.depth_at(q['key'], x1)
        w = self.width(q['dia'])
        r = dict(q=q, x0=x0, x1=x1, L=x1 - x0, d0=d0, d1=d1, interp=(i0, i1), w=w, avg=None,
                 h=[None] * 4, v=[None] * 4, total=None, soil=self.soil_at(q['key'], (x0 + x1) / 2),
                 asph=self.asph_at(q['key'], (x0 + x1) / 2))
        if d0 is not None and d1 is not None:
            r['avg'] = (d0 + d1) / 2
            r['h'] = bands(r['avg'])
            if w:
                r['v'] = [r['L'] * w * h for h in r['h']]
                r['total'] = sum(r['v'])
        if r['asph'] is None:
            dflt = num(self.S.get('ASPH_DEFAULT'))
            if dflt is not None: r['asph'], r['asph_def'] = dflt, True
        if not q['approved']: r['status'] = self.status_of(q)
        elif d0 is None or d1 is None: r['status'] = 'Missing Levels'
        elif not w: r['status'] = 'No Width'
        else: r['status'] = 'Approved'
        return r

    def calc_bed(self, q, x0, x1):
        L = x1 - x0; w = self.width(q['dia']); od = self.od(q['dia'])
        soil = self.soil_at(q['key'], (x0 + x1) / 2)
        t = num(self.S.get('BED_ROCK' if soil == ROCK else 'BED_OTHER'))
        r = dict(q=q, x0=x0, x1=x1, L=L, w=w, od=od, soil=soil, t=t, below=None, h=None, gross=None, pipe=None,
                 net=None, total=None)
        if w and od:
            r['h'] = od + (num(self.S.get('COVER_ABOVE')) or 0.30)
            r['below'] = L * w * t
            r['gross'] = L * w * r['h']
            r['pipe'] = L * PI_C / 4 * od ** 2
            r['net'] = r['gross'] - r['pipe']
            r['total'] = r['below'] + r['net']
        r['status'] = self.status_of(q) if not q['approved'] else ('Approved' if (w and od) else 'No Width')
        if r['status'] == 'No Width' and q['approved']:
            pass
        return r

    def calc_pip(self, q, x0, x1, tested, flushed):
        pct = num(self.S.get('PIP_PCT_FLUSH' if flushed else ('PIP_PCT_TEST' if tested else 'PIP_PCT_LAID')))
        r = dict(q=q, x0=x0, x1=x1, L=x1 - x0, dia=q['dia'], pct=pct, paid=(x1 - x0) * pct / 100,
                 test=self.wir_at('TEST_CODE', q['key'], x0, x1) if tested else '',
                 flush=self.wir_at('FLUSH_CODE', q['key'], x0, x1) if flushed else '')
        r['status'] = self.status_of(q)
        if r['status'] == 'Approved' and (not q['dia'] or int(q['dia']) not in DIAS): r['status'] = 'No Dia'
        return r

    # ---------------- totals for KPI bar
    def rate(self, key):
        return num(self.S.get(PO_ITEMS[key][4])) or 0.0

    # ---------------- weekly updates (snapshots of quantities & values)
    HIST_HDR = ['Update Date', 'Basis', 'Item No.', 'PO Code', 'Description', 'Unit', 'Qty', 'Rate', 'Value', 'Saved At', 'Update For']

    def boq_at(self, date, mode):
        """BOQ rows calculated live for a basis ('approved' / 'requested') – weekly / monthly update.
        طلب المستخدم 15-11: التحديث بيعدّ كل الريكويستات اللي في الـ LOG لحد لحظة الحساب (النهارده)، مش لحد تاريخ
        التحديث – لأنه ساعات بيتأخر يوم في عمل التحديث. تاريخ التحديث اسم الفترة بس. (Summary والـ Cut-off بتاعه زي ما هم)"""
        cur_mode, cur_cut = self.mode, self.cutoff
        try:
            self.mode = mode
            items, _ = self.build(max(date, dt.date.today()))
            return self.boq_rows(items)
        finally:
            self.mode = cur_mode; self.build(cur_cut)

    def load_history(self, track=None):
        """{date: {mode: {item_no: dict}}} – track = 'Company' / 'Owner' (عمود Update For).
        التحديثات القديمة من غير نوع بتتحسب للاتنين؛ ولو نفس اليوم فيه تحديث بنوعه، هو اللي بيتاخد."""
        p = str(self.S.get('HISTORY_PATH'))
        out = {}
        if not os.path.exists(p): return out
        wb = openpyxl.load_workbook(p, data_only=True)
        if 'HISTORY' not in wb.sheetnames: return out
        raw = {}
        for r in wb['HISTORY'].iter_rows(min_row=2, values_only=True):
            if not r or r[0] is None: continue
            d = todate(r[0]); mode = str(r[1]).strip().lower()
            if d is None: continue
            tag = str(r[10]).strip() if len(r) > 10 and r[10] else ''
            if track and tag and tag != track: continue
            q, rate, v = num(r[6]) or 0.0, num(r[7]) or 0.0, num(r[8]) or 0.0
            if abs(qround(q) - q) > 1e-6:       # تحديث اتسجل قبل تقريب الكميات لرقم صحيح لتحت: نفس القاعدة وقت القراية
                q = qround(q); v = round(q * rate, 2)          # (الملف نفسه ميتغيرش – عشان الفرق ميطلعش كسور بالسالب)
            raw.setdefault((d, mode), {}).setdefault(tag, {})[str(r[2])] = dict(
                code=r[3], desc=r[4], unit=r[5], qty=q, rate=rate, value=v)
        for (d, mode), tags in raw.items():
            pick = tags.get(track) if track and track in tags else tags.get('', next(iter(tags.values())))
            out.setdefault(d, {})[mode] = pick
        return out

    def save_update(self, date, track=None):
        """store quantities & values of both bases at this date (replaces an earlier save of the same date and the same
        update – Company / Owner)"""
        p = str(self.S.get('HISTORY_PATH'))
        tag = track or ''
        rows = []
        if os.path.exists(p):
            wb = openpyxl.load_workbook(p)
            ws = wb['HISTORY'] if 'HISTORY' in wb.sheetnames else wb.create_sheet('HISTORY')
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None: continue
                old = str(r[10]).strip() if len(r) > 10 and r[10] else ''
                if todate(r[0]) != date or old != tag: rows.append(list(r)[:11])
            wb.remove(ws)
        else:
            wb = openpyxl.Workbook(); wb.remove(wb.active)
        now = dt.datetime.now()
        for mode in ('approved', 'requested'):
            for b in self.boq_at(date, mode):
                rows.append([date, mode.capitalize(), str(b[0]), b[1], b[2], b[3], round(b[4], 3), b[5], round(b[6], 2), now, tag or None])
        ws = wb.create_sheet('HISTORY', 0)
        H = PatternFill('solid', fgColor='1F4E78')
        for i, h in enumerate(self.HIST_HDR, 1):
            c = ws.cell(1, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H
        rows.sort(key=lambda r: (todate(r[0]), str(r[10] or '') if len(r) > 10 else '', str(r[1]), str(r[2])))
        for i, r in enumerate(rows, 2):
            for j, v in enumerate(r, 1):
                c = ws.cell(i, j, v)
                if j == 1: c.number_format = 'dd-mm-yyyy'
                if j == 10: c.number_format = 'dd-mm-yyyy hh:mm'
                if j == 7: c.number_format = '#,##0.000'
                if j in (8, 9): c.number_format = '#,##0.00'
        for col, w in zip('ABCDEFGHIJK', [12, 10, 10, 10, 60, 6, 13, 9, 14, 16, 12]): ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A2'
        os.makedirs(os.path.dirname(os.path.abspath(p)), exist_ok=True)
        wb.save(p)
        return p

    def compare(self, date, mode, period='week', track=None, since=None):
        """compare quantities/values at `date` with the last saved update before it – of the same update (track).
        since = first day of the update's period (from the update calendar): the base is the last saved update BEFORE it,
        so the period counts from its first day (P1 01-11 → 07-11 is compared with 31-10, not with a save made on 04-11).
        (period='month': the last saved update before the 1st of that month)"""
        hist = self.load_history(track)
        limit = date.replace(day=1) if period == 'month' else (since if since and since <= date else date)
        prev_dates = sorted(d for d in hist if d < limit and mode in hist[d])
        prev_d = prev_dates[-1] if prev_dates else None
        prev = hist[prev_d][mode] if prev_d else {}
        saved_now = date in hist and mode in hist[date]
        cur = hist[date][mode] if saved_now else {str(b[0]): dict(code=b[1], desc=b[2], unit=b[3], qty=b[4], rate=b[5], value=b[6])
                                                   for b in self.boq_at(date, mode)}
        rows = []
        for k in list(dict.fromkeys(list(cur) + list(prev))):
            c = cur.get(k, {}); pv = prev.get(k, {})
            base = c or pv
            rows.append(dict(item=k, code=base.get('code'), desc=base.get('desc'), unit=base.get('unit'), rate=base.get('rate'),
                             pq=pv.get('qty', 0.0), pv=pv.get('value', 0.0), cq=c.get('qty', 0.0), cv=c.get('value', 0.0)))
        for r in rows:
            r['dq'] = r['cq'] - r['pq']; r['dv'] = r['cv'] - r['pv']
        days = (date - prev_d).days if prev_d else None
        return dict(date=date, prev_date=prev_d, days=days, rows=rows, saved=saved_now, period=period, track=track,
                    prev_total=sum(r['pv'] for r in rows), cur_total=sum(r['cv'] for r in rows),
                    history=sorted(d for d in hist if mode in hist[d]))

    def update_number(self, date, track=None):
        """رقم التحديث المسلسل للجهة دي من أول المشروع = ترتيبه بين التحديثات المتسجلة (القديمة من غير نوع بتتحسب للاتنين)."""
        return len([d for d in self.load_history(track) if d < date]) + 1

    def history_totals(self, mode, track=None):
        """[(date, total value)] of saved updates (of this update – Company / Owner), oldest first"""
        hist = self.load_history(track)
        return [(d, sum(v.get('value', 0.0) for v in hist[d][mode].values())) for d in sorted(hist) if mode in hist[d]]

    def history_series(self, mode, track=None):
        """للرسومات – قراية واحدة للملف: ([(تاريخ, إجمالي)], [(تاريخ, {نوع الشغل: قيمة})]) للتحديثات المتسجلة، الأقدم الأول"""
        import qs_charts
        hist = self.load_history(track)
        ds = [d for d in sorted(hist) if mode in hist[d]]
        ht = [(d, sum(v.get('value', 0.0) for v in hist[d][mode].values())) for d in ds]
        hg = [(d, qs_charts.group_values((k, v.get('value', 0.0)) for k, v in hist[d][mode].items())) for d in ds]
        return ht, hg

    def export_compare(self, cmp, mode, out, ht=None, title=None, hg=None, contract=None, names=None):
        """ملف التحديث: شيت المقارنة + شيت Charts (لو ht = [(تاريخ, إجمالي)] للتحديثات المتسجلة؛
        hg = نفس التواريخ متقسمة على نوع الشغل، contract = قيمة العقد للنسبة، names = {تاريخ: 'P1'} من تقويم التحديث)."""
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Monthly Update' if cmp.get('period') == 'month' else 'Weekly Update'
        thin = Side(style='thin', color='808080'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
        H = PatternFill('solid', fgColor='1F4E78'); T = PatternFill('solid', fgColor='FFF2CC')
        pd_ = cmp['prev_date']
        ws['A1'] = (f"{cmp['track'] + ' – ' if cmp.get('track') else ''}{'Monthly' if cmp.get('period') == 'month' else 'Weekly'} Update "
                    f"({mode.capitalize()}) – {cmp['date']:%d-%m-%Y}")
        ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
        ws['A2'] = ((cmp['label'] + '   |   ' if cmp.get('label') else '') +
                    (f"Previous update: {pd_:%d-%m-%Y}   |   Days: {cmp['days']}" if pd_ else 'No previous update saved'))
        ws['A2'].font = Font(size=11)
        comp = cmp.get('track') == 'Company'                   # تحديث الشركة: كود الشركة بدل رقم العقد، وكود أمر الشراء مستخبي
        hdr = ['Company Code' if comp else 'Item No.', 'PO Code', 'Description', 'Unit', 'Rate (SAR)', 'Total Qty (to date)',
               'Total Value to date (SAR)', 'Previous Qty', 'Previous Value (SAR)', 'Qty This Period', 'Value This Period (SAR)']
        rp = None
        if comp:                                               # عمود L: المرفوع للشركة لحد تاريخ التحديث (تقارير الإنتاج) – 15-11
            try: rp = reported_at(cmp['date'], cmp['rows'])
            except Exception: rp = None
            if rp and rp['n']: hdr += ['Reported to Company (to date)', 'To Report to CTR']     # M: المحسوب − المرفوع (السالب = صفر)
            else: rp = None
        NC = len(hdr)
        for i, h in enumerate(hdr, 1):
            c = ws.cell(4, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H; c.border = B
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        # الطباعة: الوصف على سطرين (إنجليزي فوق وعربي تحت) بدل ما يتقص أو يدخل في بعضه، والصفوف بتطول لوحدها
        A_DESC = Alignment(vertical='center', wrap_text=True)
        A_MID, A_NUM = Alignment(horizontal='center', vertical='center'), Alignment(vertical='center')
        r = 5
        for x in cmp['rows']:
            rq = rp['items'].get(x['item']) if rp else None
            if abs(x['cq'] or 0) < 1e-9 and abs(x['pq'] or 0) < 1e-9 and not rq: continue      # item with no quantity: hidden
            vals = [x.get('cc', '') if comp else x['item'], x['code'], bi_lines(x['desc']), x['unit'], x['rate'], x['cq'], x['cv'], x['pq'], x['pv'],
                    f'=F{r}-H{r}', f'=G{r}-I{r}']
            if rp: vals += [float(rq) if rq else None, round(max((x['cq'] or 0) - (rq or 0), 0.0), 4)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, round(v, 4) if isinstance(v, float) else v); c.border = B
                c.alignment = A_DESC if j == 3 else A_MID if j in (1, 2, 4) else A_NUM
                if j in (6, 8, 10, 12, 13): c.number_format = '#,##0.000'
                if j in (5, 7, 9, 11): c.number_format = '#,##0.00'
            r += 1
        ws.cell(r, 3, 'TOTAL VALUE (SAR)' if comp else 'TOTAL')
        for col in 'GIK': ws[f'{col}{r}'] = f'=SUM({col}5:{col}{r - 1})'
        if comp:                                               # الشركة: الكميات بس، والقيمة إجمالي تحت عمود الكمية بتاعها
            for qc, vc in (('F', 'G'), ('H', 'I'), ('J', 'K')): ws[f'{qc}{r}'] = f'={vc}{r}'
        if rp:
            ws.cell(r, 12, round(rp['value'], 2))                # قيمة المرفوع للشركة إجمالي تحت عموده
            ws.cell(r, 13, f'=SUMPRODUCT(E5:E{r - 1},M5:M{r - 1})')   # قيمة اللي لسه هيترفع
            ws.cell(r + 1, 3, 'NET – after items reported above the calculated (SAR)')   # الصافي = المحسوب − المرفوع
            ws.cell(r + 1, 13, f'=G{r}-L{r}')
            for j in (3, 13):
                c = ws.cell(r + 1, j); c.font = Font(bold=True, color='C55A11'); c.number_format = '#,##0.00'; c.border = B
        for j in range(1, NC + 1):
            c = ws.cell(r, j); c.fill = T; c.font = Font(bold=True); c.border = B; c.number_format = '#,##0.00'; c.alignment = A_NUM
        ws.row_dimensions[r].height = 21
        if cmp['days']:
            vcol = 10 if comp else 11
            g = 3 if rp else 2                                  # تحت سطر الصافي لو موجود
            ws.cell(r + g, 3, f"Value this period ({cmp['days']} days)"); ws.cell(r + g, vcol, f'=K{r}')
            ws.cell(r + g + 1, 3, 'Average value per day'); ws.cell(r + g + 1, vcol, f"=K{r}/{cmp['days']}")
            for rr in (r + g, r + g + 1):
                ws.cell(rr, 3).font = Font(bold=True, size=12); ws.cell(rr, vcol).number_format = '#,##0.00'
                ws.cell(rr, vcol).font = Font(bold=True, size=12)
        # عرض الأعمدة على قد ورقة A4 بالعرض بمقاس 100% (الوصف أضيق ويلف على سطور)
        for col, w in zip('ABCDEFGHIJK', [9, 9.5, 36, 5.5, 9.5, 11, 12.5, 11, 12.5, 11, 12.5]): ws.column_dimensions[col].width = w
        if rp: ws.column_dimensions['L'].width = 13; ws.column_dimensions['M'].width = 11
        if comp:
            for col in 'BGIK': ws.column_dimensions[col].hidden = True
            ws.column_dimensions['C'].width = 60
        ws.freeze_panes = 'A5'
        ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.3, footer=0.3)
        ws.print_options.horizontalCentered = True
        ws.print_title_rows = '4:4'                                    # عنوان الجدول يتكرر لو التحديث طلع أكتر من صفحة
        if ht is not None:
            import qs_charts
            rt = None
            if comp:                                           # Company: المرفوع للشركة في الرسومات كمان (15-11)
                try: rt = [x for x in reported_series() if x[0] <= cmp['date']]
                except Exception: rt = []
            qs_charts.add_sheet(wb, ht, hg or [], contract, title or ws['A1'].value, ref=cmp['date'], names=names, rep=rt)
        wb.save(out)
        return out

    def boq_rows(self, items):
        """[(item no, code, description, unit, qty, rate, value)] for every priced item"""
        out = []
        ok = [r for r in items['exc'] if r['status'] == 'Approved']
        for b in self.boq:
            band, item, code, rate = b
            q = qround(sum(r['v'][band - 1] for r in ok))
            out.append((item, code, EXC_DESC.get(band, ''), 'm3', q, rate, q * rate))
        bq = qround(sum(r['total'] for r in items['bed'] if r['status'] == 'Approved'))
        it = PO_ITEMS['bed']; out.append((it[0], it[1], it[2], it[3], bq, self.rate('bed'), bq * self.rate('bed')))
        for d in DIAS:
            q = qround(sum(r['paid'] for r in items['pip'] if r['status'] == 'Approved' and int(r['dia']) == d))
            it = PO_ITEMS[d]; out.append((it[0], it[1], it[2] + ' (paid length after %)', it[3], q, self.rate(d), q * self.rate(d)))
        fq = qround(sum(r['paid'] for r in items.get('bf', []) if r['paid']))
        it = PO_ITEMS['bf']; out.append((it[0], it[1], it[2] + ' (by approved layers)', it[3], fq, self.rate('bf'), fq * self.rate('bf')))
        rq = qround(sum(r['qty'] for r in items.get('rb', [])))
        it = PO_ITEMS['rb']; out.append((it[0], it[1], it[2] + ' (RB-1 + RB-2)', it[3], rq, self.rate('rb'), rq * self.rate('rb')))
        for k, f in (('asp1', 'q1'), ('asp2', 'q2')):
            aq = qround(sum(r[f] for r in items.get('asp', [])))
            it = PO_ITEMS[k]; out.append((it[0], it[1], it[2], it[3], aq, self.rate(k), aq * self.rate(k)))
        return out

    def totals(self, item, rows):
        ok = [r for r in rows if r['status'] == 'Approved']
        n_ex = len(rows) - len(ok)
        if item == 'exc':
            bandq = [sum(r['v'][i] for r in ok) for i in range(4)]
            rate = {int(b[0]): b[3] for b in self.boq}
            return [('Approved Length', f"{sum(r['L'] for r in ok):,.3f} m"), ('Excavation Qty', f"{sum(bandq):,.3f} m³"),
                    ('Value', f"{sum(qround(bandq[i]) * rate.get(i + 1, 0) for i in range(4)):,.0f} SAR"), ('Not Approved / Missing', f'{n_ex} rows')]
        if item == 'bedp':
            dn = sum(r['done'] for r in rows); tot = sum(r['total'] or 0 for r in rows); inv = sum(r['inv'] for r in rows)
            return [('Below Done', f"{sum(r['below'] or 0 for r in rows if r['below_ok'] or r['cov_ok']):,.3f} m³"),
                    ('Around Done', f"{sum(r['net'] or 0 for r in rows if r['cov_ok']):,.3f} m³"),
                    ('Done / Total', f"{dn:,.3f} / {tot:,.3f} m³  ({(dn / tot * 100) if tot else 0:.0f}%)"),
                    ('Done Value (progress)', f"{qround(dn) * self.rate('bed'):,.0f} SAR"),
                    ('Invoice (by own request)', f"{inv:,.3f} m³ = {qround(inv) * self.rate('bed'):,.0f} SAR")]
        if item == 'bed':
            return [('Approved Length', f"{sum(r['L'] for r in ok):,.3f} m"), ('Below Pipe', f"{sum(r['below'] for r in ok):,.3f} m³"),
                    ('Around Pipe (net)', f"{sum(r['net'] for r in ok):,.3f} m³"), ('Bedding Total', f"{sum(r['total'] for r in ok):,.3f} m³"),
                    ('Value', f"{qround(sum(r['total'] for r in ok)) * self.rate('bed'):,.0f} SAR"), ('Not Counted', f'{n_ex} rows')]
        if item == 'pip':
            out = []
            for d in DIAS:
                rs = [r for r in ok if int(r['dia']) == d]
                if rs: out.append((f'DN{d} Laid / Paid', f"{sum(r['L'] for r in rs):,.3f} / {sum(r['paid'] for r in rs):,.3f} m"))
            val = sum(qround(sum(r['paid'] for r in ok if int(r['dia']) == d)) * self.rate(d) for d in DIAS)
            return out + [('Value', f"{val:,.0f} SAR"), ('Not Counted', f'{n_ex} rows')]
        if item == 'bf':
            rs = [r for r in rows if r['qty']]
            return [('Backfill Qty (full)', f"{sum(r['qty'] for r in rs):,.3f} m³"),
                    ('Backfill Qty (by layers)', f"{sum(r['paid'] for r in rs):,.3f} m³"),
                    ('Value', f"{qround(sum(r['paid'] for r in rs)) * self.rate('bf'):,.0f} SAR"),
                    ('Complete', f"{len([r for r in rows if r['status'] == 'Complete'])} rows"),
                    ('Partial / No Layers', f"{len([r for r in rows if r['status'] in ('Partial', 'No Layers')])} rows")]
        if item == 'rb':
            return [('Layer 1 (RB-1)', f"{sum(r['q1'] for r in rows):,.3f} m³"), ('Layer 2 (RB-2)', f"{sum(r['q2'] for r in rows):,.3f} m³"),
                    ('Road Base Total', f"{sum(r['qty'] for r in rows):,.3f} m³"),
                    ('Value', f"{qround(sum(r['qty'] for r in rows)) * self.rate('rb'):,.0f} SAR"),
                    ('Both Layers / One / None', f"{len([r for r in rows if r['status'] == 'Complete'])} / "
                                                  f"{len([r for r in rows if r['status'] == 'Partial'])} / "
                                                  f"{len([r for r in rows if r['status'] == 'No Layers'])} rows")]
        if item == 'asp':
            a1, a2, a3 = (sum(r[f] for r in rows) for f in ('q1', 'q2', 'q3'))
            return [('32.4.1 Asphalt (ASP-1)', f"{a1:,.3f} m³ = {qround(a1) * self.rate('asp1'):,.0f} SAR"),
                    ('32.4.2 Milling (ASP-2-ML)', f"{a2:,.3f} m³ = {qround(a2) * self.rate('asp2'):,.0f} SAR"),
                    ('Wearing Course (not paid)', f"{a3:,.3f} m³"),
                    ('Complete / Partial / Not Started', f"{len([r for r in rows if r['status'] == 'Complete'])} / "
                                                         f"{len([r for r in rows if r['status'] == 'Partial'])} / "
                                                         f"{len([r for r in rows if r['status'] == 'Not Started'])} rows")]
        if item == 'sum':
            b = self.boq_rows(self.items)
            n = len(self.boq)
            grp = [('Excavation', b[:n]), ('Bedding', b[n:n + 1]), ('Pipe Laying', b[n + 1:n + 4]),
                   ('Backfill', b[n + 4:n + 5]), ('Road Base', b[n + 5:n + 6]), ('Asphalt', b[n + 6:])]
            return [(f'{g} Value', f"{sum(x[6] for x in rs):,.0f} SAR") for g, rs in grp] + \
                   [('TOTAL VALUE', f"{sum(x[6] for x in b):,.0f} SAR")]
        return []

    # ---------------- QS-TOTAL rows (one row per approved excavation piece; other items mapped by overlap)
    def summary_rows(self, items=None):
        it = items or self.items
        exc = [r for r in it['exc'] if r['q']['approved']]
        exc.sort(key=lambda r: (kstr(r['q']['dist']), r['q']['key'][0], r['x0']))
        bed = [r for r in it['bed'] if r['status'] == 'Approved']
        pip = [r for r in it['pip'] if r['status'] == 'Approved']
        rows = []
        for i, e in enumerate(exc, 1):
            q = e['q']; k = q['key']
            R = dict(A=i, B=q['dist'], C=q['line'], E=e['x0'], F=e['x1'], J=e['L'], L=e['w'], M=SOIL_AR[e['soil']],
                     N=e['d0'], O=e['d1'], P=e['avg'], Q=q['dia_raw'], exc_ok=e['status'] == 'Approved')
            for c, h, v in zip('RSTU', e['h'], e['v']):
                R[c] = h
            for c, v in zip('VWXY', e['v']): R[c] = v
            # bedding
            bl = [(b, overlap(b['x0'], b['x1'], e['x0'], e['x1'])) for b in bed if b['q']['key'] == k]
            bl = [(b, ol) for b, ol in bl if ol > EPS]
            F = R['_f'] = {}
            if bl:
                bb = [(b, ol) for b, ol in bl if b.get('part') == 'below']      # BED-D-P
                ba = [(b, ol) for b, ol in bl if b.get('part') != 'below']      # COV-P
                F['bed'] = dict(ol_b=sum(ol for _, ol in bb), ol_a=sum(ol for _, ol in ba), od=bl[0][0].get('od'),
                                same=len({b['t'] for b, _ in bb}) <= 1 and len({b['h'] for b, _ in ba}) <= 1)
                R['Z'] = (bb or bl)[0][0]['t']
                if bb: R['AA'] = sum(b['below'] * ol / b['L'] for b, ol in bb)
                if ba:
                    R['AB'] = ba[0][0]['h']
                    R['AC'] = sum(b['gross'] * ol / b['L'] for b, ol in ba)
                    R['AD'] = sum(b['pipe'] * ol / b['L'] for b, ol in ba)
                    R['AE'] = sum(b['net'] * ol / b['L'] for b, ol in ba)
            # pipe laying
            pl = [(p, overlap(p['x0'], p['x1'], e['x0'], e['x1'])) for p in pip if p['q']['key'] == k]
            pl = [(p, ol) for p, ol in pl if ol > EPS]
            if pl:
                R['K'] = sum(ol for _, ol in pl)
                F['pipdias'] = {int(p['dia']) for p, _ in pl}
                for d, cq, cp in ((300, 'AF', 'AG'), (400, 'AH', 'AI'), (600, 'AJ', 'AK')):
                    ps = [(p, ol) for p, ol in pl if int(p['dia']) == d]
                    if ps:
                        ln = sum(ol for _, ol in ps)
                        R[cq] = ln; R[cp] = sum(p['pct'] * ol for p, ol in ps) / ln / 100
            # backfill (pieces lie inside the excavation piece)
            fl = [(f, overlap(f['x0'], f['x1'], e['x0'], e['x1'])) for f in it.get('bf', []) if f['q']['key'] == k and f['qty']]
            fl = [(f, ol) for f, ol in fl if ol > EPS]
            # road base (pieces lie inside the excavation piece)
            rl = [(x, overlap(x['x0'], x['x1'], e['x0'], e['x1'])) for x in it.get('rb', []) if x['q']['key'] == k and x['qty']]
            rl = [(x, ol) for x, ol in rl if ol > EPS]
            if rl:
                q1 = sum(x['q1'] * ol / x['L'] for x, ol in rl); q2 = sum(x['q2'] * ol / x['L'] for x, ol in rl)
                if q1: R['AO'] = rl[0][0]['th']; R['AP'] = q1
                if q2: R['AQ'] = rl[0][0]['th']; R['AR'] = q2
                F['rb'] = dict(ol1=sum(ol for x, ol in rl if x['q1']), ol2=sum(ol for x, ol in rl if x['q2']))
            # asphalt (pieces lie inside the excavation piece) – AS..AY like the consultant, thickness AT / AW in cm here
            if e['w']:
                AS, AT, AV, AW = self.asp_dims(e['w'], e['asph'])
                R['AS'], R['AV'], R['AW'] = AS, AV, round(AW * 100, 6)
                if e['asph'] is not None: R['AT'] = e['asph']
            al = [(x, overlap(x['x0'], x['x1'], e['x0'], e['x1'])) for x in it.get('asp', []) if x['q']['key'] == k and x['AS'] is not None]
            al = [(x, ol) for x, ol in al if ol > EPS]
            if al:
                for col, fq in (('AU', 'q1'), ('AX', 'q2'), ('AY', 'q3')):
                    v = sum(x[fq] * ol / x['L'] for x, ol in al)
                    if v: R[col] = v
                F['asp'] = dict(ol1=sum(ol for x, ol in al if x['q1']), ol2=sum(ol for x, ol in al if x['q2']),
                                ol3=sum(ol for x, ol in al if x['q3']))
                for col, fw in (('BF', 'w1'), ('BG', 'w2'), ('BH', 'w3')):
                    v = wjoin(x[fw] for x, _ in al if x[fw])
                    if v: R[col] = v
            R['AZ'] = wno(q.get('no'))
            if bl:
                if bb: R['BA'] = wjoin(b['q'].get('no') for b, _ in bb)      # الفرشة أسفل الماسورة
                if ba: R['BB'] = wjoin(b['q'].get('no') for b, _ in ba)      # الفرشة حول الماسورة
            if pl:
                R['BC'] = wjoin(p['q'].get('no') for p, _ in pl)
                tst = wjoin(p.get('test') for p, _ in pl)
                if tst: R['BC'] += f' (T: {tst})'
            if fl: R['BD'] = wjoin(f.get('wirs') for f, _ in fl)
            if rl: R['BE'] = wjoin(x.get('wirs') for x, _ in rl)
            if fl:
                f0 = fl[0][0]
                F['bf'] = dict(ol=sum(ol for _, ol in fl), same=len({round(f['H'], 9) for f, _ in fl}) == 1,
                               road=num(self.S.get('ROAD_BASE')) or 0.30, asph=f0['asph'] or 0, t=f0['t'], ch=f0['ch'])
                ln = sum(ol for _, ol in fl)
                R['AL'] = fl[0][0]['H']
                R['AM'] = sum((f['pct'] or 0) * ol for f, ol in fl) / ln
                R['AN'] = sum(f['qty'] * (f['pct'] or 0) * ol / f['L'] for f, ol in fl)   # الحجم × نسبة الطبقات المعتمدة (زي شيت الاستشاري)
            rows.append(R)
        return rows

    # ---------------- export to QS-TOTAL template
    def export(self, items, notes, out=None):
        out = out or str(self.S.get('OUTPUT_PATH'))
        tp = str(self.S.get('TEMPLATE_PATH'))
        if not os.path.exists(tp): tp = os.path.join(HERE, 'QS-TOTAL_TEMPLATE.xlsx')
        wb = openpyxl.load_workbook(tp)
        thin = Side(style='thin', color='808080'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
        CEN = Alignment(horizontal='center', vertical='center')
        TOT = PatternFill('solid', fgColor='FFF2CC')
        NC = CI(WIR_COLS[-1][0])   # 51 عمود نموذج الاستشاري + أعمدة أرقام الطلبات (AZ..)
        R0 = 5
        # both bases: approved (QS-TOTAL) and requested (QS-REQUESTED)
        cur = self.mode
        self.mode = 'approved'; items, notes = self.build(self.cutoff)
        self.mode = 'requested'; items_req, _ = self.build(self.cutoff)
        self.mode = cur; self.build(self.cutoff)
        self._boq_sheet(wb, self.boq_rows(items), self.boq_rows(items_req))
        if 'QS-REQUESTED' in wb.sheetnames: del wb['QS-REQUESTED']
        wsr = wb.copy_worksheet(wb['QS-TOTAL']); wsr.title = 'QS-REQUESTED'
        wb.move_sheet('QS-REQUESTED', offset=wb.sheetnames.index('QS-TOTAL') + 1 - wb.sheetnames.index('QS-REQUESTED'))
        self._fill_qs(wsr, self.summary_rows(items_req), B, CEN, TOT, NC, R0)
        ws = wb['QS-TOTAL']
        rows = self.summary_rows(items)
        t = self._fill_qs(ws, rows, B, CEN, TOT, NC, R0)
        last = R0 + len(rows) - 1
        if 'EX' in wb.sheetnames:
            ex = wb['EX']
            for i in range(len(rows)):
                n = R0 + i
                for col, src in zip('ABCDEFGHIJK', 'ABCDEFJVWXY'):
                    ex[f'{col}{n}'] = f"='QS-TOTAL'!{src}{n}"
                for c in range(1, 12):
                    cell = ex.cell(n, c); cell.border = B; cell.alignment = CEN
                    if c >= 7: cell.number_format = '#,##0.000'
            ex[f'A{t}'] = 'الإجمالي'
            if rows:
                for col in 'GHIJK': ex[f'{col}{t}'] = f'=SUM({col}{R0}:{col}{last})'
            for c in range(1, 12):
                cell = ex.cell(t, c); cell.border = B; cell.fill = TOT; cell.font = Font(bold=True)
                cell.alignment = CEN; cell.number_format = '#,##0.000'
        try:
            import qs_sheets
            qs_sheets.add_all(self, wb, rows, items, notes)
        except Exception as e:
            notes.append(f'UNIFIER/extra sheets failed: {e}')
        return self._export_rest(wb, items, notes, out, B, CEN)

    def _boq_sheet(self, wb, appr, req):
        if 'BOQ-VALUE' in wb.sheetnames: del wb['BOQ-VALUE']
        ws = wb.create_sheet('BOQ-VALUE', wb.sheetnames.index('QS-TOTAL'))
        thin = Side(style='thin', color='808080'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
        H = PatternFill('solid', fgColor='1F4E78'); T = PatternFill('solid', fgColor='FFF2CC')
        ws['A1'] = f'BOQ Value – Approved vs Requested   |   Cut-off {self.cutoff:%d-%m-%Y}'; ws['A1'].font = Font(bold=True, size=13)
        hdr = ['Item No.', 'PO Code', 'Description', 'Unit', 'Rate (SAR)', 'Approved Qty', 'Approved Value', 'Requested Qty', 'Requested Value']
        for i, h in enumerate(hdr, 1):
            c = ws.cell(3, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H; c.border = B
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for col, w in zip('ABCDEFGHI', [10, 11, 70, 7, 11, 14, 16, 14, 16]): ws.column_dimensions[col].width = w
        r = 4
        for a, q in zip(appr, req):
            for j, v in enumerate([a[0], a[1], a[2], a[3], a[5], a[4], f'=E{r}*F{r}', q[4], f'=E{r}*H{r}'], 1):
                c = ws.cell(r, j, round(v, 4) if isinstance(v, float) else v); c.border = B
                if j >= 5: c.number_format = '#,##0.000' if j in (6, 8) else '#,##0.00'
            r += 1
        ws.cell(r, 3, 'TOTAL')
        ws.cell(r, 7, f'=SUM(G4:G{r - 1})'); ws.cell(r, 9, f'=SUM(I4:I{r - 1})')
        for j in range(1, 10):
            c = ws.cell(r, j); c.fill = T; c.font = Font(bold=True); c.border = B; c.number_format = '#,##0.00'
        ws.freeze_panes = 'A4'

    def _fill_qs(self, ws, rows, B, CEN, TOT, NC, R0):
        # عناوين أعمدة أرقام الطلبات (AZ..BD)
        H = PatternFill('solid', fgColor='D9E1F2')
        c0, c1 = CI(WIR_COLS[0][0]), CI(WIR_COLS[-1][0])
        ws.cell(1, c0, WIR_HDR)
        try: ws.merge_cells(start_row=1, start_column=c0, end_row=2, end_column=c1)
        except Exception: pass
        for col, name in WIR_COLS:
            ws[f'{col}3'] = name
            try: ws.merge_cells(f'{col}3:{col}4')
            except Exception: pass
            ws.column_dimensions[col].width = 14
        for r in range(1, 5):
            for c in range(c0, c1 + 1):
                x = ws.cell(r, c); x.fill = H; x.border = B; x.font = Font(bold=True)
                x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for i, R in enumerate(rows):
            n = R0 + i
            for col in ('A', 'B', 'C', 'E', 'F', 'L', 'M', 'Q', 'K', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE',
                        'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR',
                        'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', *[c for c, _ in WIR_COLS]):
                v = R.get(col)
                if v is None or v == '': continue
                ws[f'{col}{n}'] = round(v, 4) if isinstance(v, float) else v
            ws[f'J{n}'] = f'=F{n}-E{n}-H{n}'
            if R['exc_ok']:
                ws[f'N{n}'] = round(R['N'], 3); ws[f'O{n}'] = round(R['O'], 3)
                ws[f'P{n}'] = f'=(N{n}+O{n})/2'
                ws[f'R{n}'] = f'=MIN(P{n},1.5)'; ws[f'S{n}'] = f'=IF(P{n}>=2.5,1,MAX(P{n}-R{n},0))'
                ws[f'T{n}'] = f'=IF(P{n}>=3.5,1,MAX(P{n}-R{n}-S{n},0))'; ws[f'U{n}'] = f'=MAX(P{n}-R{n}-S{n}-T{n},0)'
                for c, h in zip('VWXY', 'RSTU'): ws[f'{c}{n}'] = f'=J{n}*L{n}*{h}{n}'
            self._formulas(ws, n, R)
            for c in range(1, NC + 1):
                cell = ws.cell(n, c); cell.border = B; cell.alignment = CEN
                L_ = CL(c)
                if L_ in ('N', 'O', 'P', 'R', 'S', 'T', 'U', 'Z', 'AB'): cell.number_format = '0.000'
                elif L_ in ('AG', 'AI', 'AK', 'AM'): cell.number_format = '0%'
                elif L_ in ('AL', 'AO', 'AQ'): cell.number_format = '0.000'
                elif L_ == 'L': cell.number_format = '#,##0.00'
                elif L_ in ('AT', 'AW'): cell.number_format = '0'
                elif c >= 10 and L_ not in ('M', 'Q') and c <= 51: cell.number_format = '#,##0.000'
        for col in ('AT', 'AW'):                                   # السمك بالسنتي
            v = ws[f'{col}4'].value
            if isinstance(v, str): ws[f'{col}4'] = v.replace('(م)', '(سم)')
        last = R0 + len(rows) - 1
        t = last + 1
        ws[f'A{t}'] = 'الإجمالي'
        if rows:
            for col in ('J', 'K', 'V', 'W', 'X', 'Y', 'AA', 'AE', 'AF', 'AH', 'AJ', 'AN', 'AP', 'AR', 'AU', 'AX', 'AY'):     # من غير AC/AD زي الاستشاري
                ws[f'{col}{t}'] = f'=SUM({col}{R0}:{col}{last})'
        for c in range(1, NC + 1):
            cell = ws.cell(t, c); cell.border = B; cell.fill = TOT; cell.font = Font(bold=True); cell.alignment = CEN
            cell.number_format = '#,##0.000'
        ws.merge_cells(start_row=t, start_column=1, end_row=t, end_column=9)
        ws.freeze_panes = 'B5'
        ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.print_title_rows = '1:4'
        ws.print_area = f'A1:{CL(NC)}{t}'
        return t

    def _formulas(self, ws, n, R):
        """المعادلات في كل خلية محسوبة (الفرشة، التمديد، الردم، قاعدة الطريق) – بتتكتب بس لو نتيجتها = الرقم المحسوب"""
        F = R.get('_f') or {}
        J = R.get('J') or 0
        def ok(v, want): return want is not None and abs(v - want) <= 1e-6 * max(1, abs(want))
        def length(ol): return f'J{n}' if abs(ol - J) < 1e-9 else f'{ol:g}'
        w = R.get('L') or 0
        b = F.get('bed')
        if b and b['same'] and b['od'] and w:
            if b['ol_b'] > EPS and R.get('AA') is not None:
                L_ = length(b['ol_b'])
                if ok(R['Z'] * w * b['ol_b'], R.get('AA')): ws[f'AA{n}'] = f'=Z{n}*L{n}*{L_}'
            if b['ol_a'] > EPS and R.get('AE') is not None:
                L_ = length(b['ol_a']); ol = b['ol_a']
                if ok(R['AB'] * w * ol, R.get('AC')): ws[f'AC{n}'] = f'=AB{n}*L{n}*{L_}'
                if ok(PI_C / 4 * b['od'] ** 2 * ol, R.get('AD')): ws[f'AD{n}'] = f"=3.1416/4*{b['od']:g}^2*{L_}"
                ws[f'AE{n}'] = f'=AC{n}-AD{n}'
        dias = F.get('pipdias') or set()
        if len(dias) == 1 and R.get('K') is not None:
            cq = {300: 'AF', 400: 'AH', 600: 'AJ'}.get(next(iter(dias)))
            pc = {'AF': 'AG', 'AH': 'AI', 'AJ': 'AK'}.get(cq)
            if cq and ok(R['K'], R.get(cq)):                  # الكمية = الطول × نسبة التمديد (زي الاستشاري)
                ws[f'{cq}{n}'] = f'=K{n}*{pc}{n}' if R.get(pc) is not None else f'=K{n}'
        f = F.get('bf')
        if f and f['same'] and R.get('P') is not None:
            t = f'Z{n}' if R.get('Z') is not None and abs(R['Z'] - f['t']) < 1e-9 else f"{f['t']:g}"
            ch = f'AB{n}' if R.get('AB') is not None and abs(R['AB'] - f['ch']) < 1e-9 else f"{f['ch']:g}"
            H = R['P'] - f['road'] - f['asph'] / 100 - f['t'] - f['ch']
            if ok(H, R.get('AL')):                            # سمك الأسفلت (سم) في AT – الردم مربوط بيه زي الاستشاري: − AT
                if f['asph']: ws[f'AT{n}'] = f['asph']
                ws[f'AL{n}'] = f"=P{n}-{f['road']:g}-AT{n}/100-{t}-{ch}"
                if ok(max(H, 0) * w * f['ol'] * (R.get('AM') or 0), R.get('AN')): ws[f'AN{n}'] = f"=MAX(AL{n},0)*L{n}*{length(f['ol'])}*AM{n}"
        r = F.get('rb')
        if r and w:
            if R.get('AP') and ok(R['AO'] * w * r['ol1'], R['AP']): ws[f'AP{n}'] = f"=AO{n}*L{n}*{length(r['ol1'])}"
            if R.get('AR') and ok(R['AQ'] * w * r['ol2'], R['AR']): ws[f'AR{n}'] = f"=AQ{n}*L{n}*{length(r['ol2'])}"
        # الأسفلت زي الاستشاري: AS = L + 0.3 | AV = 3.6 − AS | AU = J × AS × AT | AX / AY = J × AV × AW  (AT و AW بالسنتي هنا)
        if R.get('AS') is not None and w:
            extra = num(self.S.get('ASP_EXTRA'))
            if extra is None: extra = 0.30
            lane = num(self.S.get('LANE_WIDTH')) or 3.60
            if ok(w + extra, R['AS']): ws[f'AS{n}'] = f'=L{n}+{extra:g}'
            if R.get('AV') is not None and ok(lane - R['AS'], R['AV']): ws[f'AV{n}'] = f'={lane:g}-AS{n}'
            a = F.get('asp')
            if a:
                AT, AW = R.get('AT') or 0, R.get('AW') or 0
                if R.get('AU') and ok(R['AS'] * AT / 100 * a['ol1'], R['AU']): ws[f'AU{n}'] = f"={length(a['ol1'])}*AS{n}*AT{n}/100"
                for col, ol in (('AX', a['ol2']), ('AY', a['ol3'])):
                    if R.get(col) and R.get('AV') is not None and ok(R['AV'] * AW / 100 * ol, R[col]):
                        ws[f'{col}{n}'] = f"={length(ol)}*AV{n}*AW{n}/100"

    def _export_rest(self, wb, items, notes, out, B, CEN):
        # detail sheets (traceability)
        H = PatternFill('solid', fgColor='1F4E78'); G = PatternFill('solid', fgColor='EDEDED')
        def detail(name, hdr, data, statuses):
            if name in wb.sheetnames: del wb[name]
            d = wb.create_sheet(name)
            for i, h in enumerate(hdr, 1):
                c = d.cell(1, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = H; c.border = B; c.alignment = CEN
                d.column_dimensions[CL(i)].width = max(8, len(h) + 2)
            for i, (vals, st) in enumerate(zip(data, statuses), 2):
                for j, v in enumerate(vals, 1):
                    c = d.cell(i, j, round(v, 4) if isinstance(v, float) else v); c.border = B; c.alignment = CEN
                    if isinstance(v, float): c.number_format = '#,##0.000'
                    if isinstance(v, dt.date): c.number_format = 'dd-mm-yyyy'
                    if st != 'Approved': c.fill = G
            d.freeze_panes = 'A2'
        e_ = items['exc']
        detail('WIR-EXC', ['#', 'WIR No.', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Soil', 'Asphalt (cm)', 'Depth Start',
                           'Depth End', 'Avg Depth', 'Width', '0-1.5', '1.5-2.5', '2.5-3.5', '>3.5', 'Total (m3)',
                           'Request Date', 'Approval Date', 'Status'],
               [[i, r['q']['no'], r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'],
                 r['soil'], r['asph'], r['d0'], r['d1'], r['avg'], r['w'], *r['v'], r['total'], r['q']['reqdate'], r['q']['appdate'], r['status']]
                for i, r in enumerate(e_, 1)], [r['status'] for r in e_])
        b_ = items['bed']
        detail('WIR-BED', ['#', 'WIR No.', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Soil', 'Width',
                           'Below Thk', 'Below Qty', 'Around Height', 'Around Gross', 'Pipe Volume', 'Around Net', 'Total (m3)',
                           'Approval Date', 'Status'],
               [[i, r['q']['no'], r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'],
                 r['soil'], r['w'], r['t'], r['below'], r['h'], r['gross'], r['pipe'], r['net'], r['total'], r['q']['appdate'], r['status']]
                for i, r in enumerate(b_, 1)], [r['status'] for r in b_])
        p_ = items['pip']
        detail('WIR-PIP', ['#', 'WIR No.', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Test WIR', 'Flushing WIR',
                           '%', 'Paid Length (m)', 'Approval Date', 'Status'],
               [[i, r['q']['no'], r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'],
                 r['test'], r['flush'], r['pct'], r['paid'], r['q']['appdate'], r['status']]
                for i, r in enumerate(p_, 1)], [r['status'] for r in p_])
        f_ = items.get('bf', [])
        detail('WIR-BF', ['#', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Avg Depth', 'Asphalt (cm)', 'Soil',
                          'Bedding Below', 'Cover Height', 'Layers Height', 'Backfill Height (Qty)', 'Layers Required', 'Layers Approved', '%',
                          'Width', 'Backfill Qty (m3)', 'Qty by Layers (m3)', 'Status', 'BF WIR No.'],
               [[i, r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'], r['avg'], r['asph'],
                 r['soil'], r['t'], r['ch'], r.get('Hn'), r['H'], r['n'], ', '.join(str(k) for k in r['layers']),
                 None if r['pct'] is None else round(r['pct'] * 100, 1), r['w'], r['qty'], r['paid'], r['status'], r['wirs']]
                for i, r in enumerate(f_, 1)], ['Approved' if r['status'] == 'Complete' else r['status'] for r in f_])
        bp = items.get('bedp', [])
        detail('BED-PROGRESS', ['#', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Soil', 'Width', 'Below Thk',
                                'Below Qty', 'BED-D-P', 'Around Net', 'COV-P', 'Total', 'Done Qty', '% Done', 'Done Value (SAR)',
                                'Invoice Qty', 'Status', 'WIR No.'],
               [[i, r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'], r['soil'], r['w'],
                 r['t'], r['below'], 'Yes' if r['below_ok'] else '', r['net'], 'Yes' if r['cov_ok'] else '', r['total'], r['done'],
                 None if r['pct'] is None else round(r['pct'] * 100, 1), r['value'], r['inv'], r['status'], r['wirs']]
                for i, r in enumerate(bp, 1)], ['Approved' if r['status'] == 'Closed' else r['status'] for r in bp])
        b_ = items.get('rb', [])
        detail('WIR-RB', ['#', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Width', 'Layer (m)',
                          'RB-1 Qty (m3)', 'RB-2 Qty (m3)', 'Road Base Qty (m3)', 'Status', 'RB WIR No.'],
               [[i, r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'], r['w'], r['th'],
                 r['q1'], r['q2'], r['qty'], r['status'], r['wirs']] for i, r in enumerate(b_, 1)],
               ['Approved' if r['status'] == 'Complete' else r['status'] for r in b_])
        a_ = items.get('asp', [])
        detail('WIR-ASP', ['#', 'District', 'Street', 'Line', 'Dia', 'From', 'To', 'Length (m)', 'Trench Width', 'Asphalt Width (AS)',
                           'Thickness (cm)', 'Lane Rest (AV)', 'Milling Thk (m)', '32.4.1 Asphalt (m3)', '32.4.2 Milling (m3)',
                           'Wearing Course (m3)', 'ASP-1 WIR', 'Milling WIR', 'Wearing WIR', 'MC WIR', 'RC WIR', 'Status'],
               [[i, r['q']['dist'], r['q']['street'], r['q']['line'], r['q']['dia_raw'], r['x0'], r['x1'], r['L'], r['w'], r['AS'],
                 r['asph'], r['AV'], r['AW'], r['q1'], r['q2'], r['q3'], r['w1'], r['w2'], r['w3'], r['wmc'], r['wrc'], r['status']]
                for i, r in enumerate(a_, 1)],
               ['Approved' if r['status'] in ('Complete', 'Partial') else r['status'] for r in a_])
        if 'Notes' in wb.sheetnames: del wb['Notes']
        nt = wb.create_sheet('Notes'); nt.column_dimensions['A'].width = 110
        nt['A1'] = f'Cut-off date: {self.cutoff:%d-%m-%Y}   |   Generated: {dt.datetime.now():%d-%m-%Y %H:%M}'
        nt['A1'].font = Font(bold=True)
        for i, s in enumerate(notes or ['No notes'], 3): nt.cell(i, 1, s)
        wb.active = wb.sheetnames.index('QS-TOTAL')
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        wb.save(out)
        return out


def template_header(path, sheet='QS-TOTAL', hdr_rows=4):
    """Header layout of the approved QS sheet: column widths (Excel units) + cells (r0,c0,r1,c1,text,fill)."""
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    ncol = 0
    for r in range(1, hdr_rows + 1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value is not None: ncol = max(ncol, c)
    for m in ws.merged_cells.ranges:
        if m.min_row <= hdr_rows: ncol = max(ncol, m.max_col)
    widths = []
    for c in range(1, ncol + 1):
        cd = ws.column_dimensions.get(CL(c))
        widths.append(cd.width if cd is not None and cd.width else 9.0)
    cells, used = [], set()
    def fill_of(cell):
        try:
            rgb = cell.fill.fgColor.rgb
            if cell.fill.fill_type and isinstance(rgb, str) and len(rgb) == 8 and rgb != '00000000': return '#' + rgb[2:]
        except Exception: pass
        return '#D9E1F2'
    for m in ws.merged_cells.ranges:
        if m.min_row > hdr_rows or m.min_col > ncol: continue
        r1 = min(m.max_row, hdr_rows); c1 = min(m.max_col, ncol)
        top = ws.cell(m.min_row, m.min_col)
        cells.append((m.min_row, m.min_col, r1, c1, '' if top.value is None else str(top.value), fill_of(top)))
        for r in range(m.min_row, r1 + 1):
            for c in range(m.min_col, c1 + 1): used.add((r, c))
    for r in range(1, hdr_rows + 1):
        for c in range(1, ncol + 1):
            if (r, c) in used: continue
            cell = ws.cell(r, c)
            cells.append((r, c, r, c, '' if cell.value is None else str(cell.value), fill_of(cell)))
    if sheet == 'QS-TOTAL' and ncol < CI(WIR_COLS[-1][0]):
        c0 = ncol + 1
        for col, name in WIR_COLS:
            c = CI(col)
            while ncol < c:
                ncol += 1; widths.append(14.0)
            cells.append((3, c, hdr_rows, c, name, '#D9E1F2'))
        cells.append((1, c0, 2, ncol, WIR_HDR, '#D9E1F2'))
    letters = [CL(c) for c in range(1, ncol + 1)]
    return dict(ncol=ncol, letters=letters, widths=widths, cells=cells, rows=hdr_rows)


# ================================================================ المرفوع للشركة (تقارير الإنتاج الأسبوعية) – طلب المستخدم 15-11
# تقرير الإنتاج من سيستم الشركة (PDF) بيتحط في 07_Planning_Progress\REPORTS_CTR\WEEKLY. البرنامج بيقرا تقرير الإجمالي بس
# (تقرير «منفذ الأعمال» بيتساب عشان الكمية ماتتحسبش مرتين)، وبياخد الكمية المنجزة لكل بند ويجمعها من أول تقرير
# لحد تاريخ التحديث. البند بيتربط ببند البرنامج بسعر الوحدة (ولو فيه أكتر من بند بنفس السعر ← بالوحدة كمان).
PROD_REL = os.path.join('07_Planning_Progress', 'REPORTS_CTR')          # WEEKLY + MONTHLY (15-11)
_PROD_CACHE = {}
_NUM = re.compile(r'\(?\d[\d,]*\.\d{2}\)?')


def _unit_key(u):
    u = str(u or '').lower().replace(' ', '')
    if 'م3' in u or 'm3' in u or 'cum' in u: return 'm3'
    if 'م2' in u or 'm2' in u or 'sqm' in u: return 'm2'
    if 'عدد' in u or u in ('no', 'nr', 'nos', 'ea', 'each', 'no.'): return 'no'
    if 'طو' in u or u in ('m', 'lm', 'ml', 'l.m', 'rm') or 'م.ط' in u: return 'm'
    return u or None


def read_production(path):
    """تقرير إنتاج واحد ← dict(d0, d1, items={سعر الوحدة: [(كمية، وحدة)]}) أو None لو مش تقرير إجمالي."""
    try: mt = os.path.getmtime(path)
    except OSError: return None
    c = _PROD_CACHE.get(path)
    if c and c[0] == mt: return c[1]
    res = None
    try:
        import pypdf
        t = '\n'.join((p.extract_text() or '') for p in pypdf.PdfReader(path).pages)
        if 'نتاج' in t and 'المنجزة' in t and 'منفذ' not in t:      # تقرير إنتاج بس (مش «بيان متابعة»)
            ds = sorted(set(re.findall(r'20\d{2}-\d{2}-\d{2}', t)))
            if ds:
                items = {}
                for ln in t.splitlines():
                    unit = _unit_key('م3' if 'م3' in ln else 'م2' if 'م2' in ln else 'م طو' if re.search(r'م\s*طو', ln) else
                                     'عدد' if 'عدد' in ln else '')
                    if not unit: continue
                    n = [(-1 if x.startswith('(') else 1) * float(x.strip('()').replace(',', '')) for x in _NUM.findall(ln)]
                    hit = None
                    for i in range(len(n)):                 # الكمية المنجزة × سعر الوحدة = القيمة المنجزة (بالترتيب ده في السطر)
                        if n[i] <= 0: continue
                        for j in range(i + 1, len(n)):
                            if n[j] <= 0: continue
                            if any(abs(n[i] * n[j] - n[k]) <= max(0.6, n[k] * 0.001) for k in range(j + 1, len(n))):
                                hit = (n[i], n[j]); break
                        if hit: break
                    if hit: items.setdefault(round(hit[1], 2), []).append((hit[0], unit))
                res = dict(d0=dt.date.fromisoformat(ds[0]), d1=dt.date.fromisoformat(ds[-1]), items=items, path=path)
    except Exception:
        res = None
    _PROD_CACHE[path] = (mt, res)
    return res


def production_reports():
    root = project_path(PROD_REL)
    out, seen = [], set()
    for dp, _, fs in os.walk(root) if os.path.isdir(root) else []:
        for f in sorted(fs):
            if not f.lower().endswith('.pdf'): continue
            r = read_production(os.path.join(dp, f))
            if r and (r['d0'], r['d1']) not in seen:      # نفس الفترة مرتين (نسخة) ← مرة واحدة
                seen.add((r['d0'], r['d1'])); out.append(r)
    return sorted(out, key=lambda r: r['d1'])


def reported_at(date, rows):
    """المرفوع للشركة لحد تاريخ التحديث: {item: كمية}، القيمة، عدد التقارير، الفترة، والبنود اللي ماتربطتش."""
    reps = [r for r in production_reports() if r['d1'] <= date]
    # تقرير جوه فترة تقرير أكبر منه (أسبوع جوه الشهري) ← الكبير بس، عشان الكمية ماتتحسبش مرتين (15-11)
    reps = [r for r in reps if not any(s is not r and s['d0'] <= r['d0'] and r['d1'] <= s['d1'] and
                                       (s['d1'] - s['d0']) > (r['d1'] - r['d0']) for s in reps)]
    tot = {}                                               # (سعر، وحدة) ← كمية من أول تقرير
    for r in reps:
        for rate, lst in r['items'].items():
            for q, u in lst: tot[(rate, u)] = tot.get((rate, u), 0.0) + q
    by_item, value, loose = {}, 0.0, []
    for (rate, u), q in tot.items():
        value += q * rate
        cand = [x for x in rows if x.get('rate') and abs(float(x['rate']) - rate) < 0.005]
        if len(cand) > 1: cand = [x for x in cand if _unit_key(x.get('unit')) == u] or cand
        if len(cand) == 1: by_item[cand[0]['item']] = by_item.get(cand[0]['item'], 0.0) + q
        else: loose.append(f'{q:,.2f} {u} @ {rate:,.2f}' + (' (more than one item)' if cand else ''))
    return dict(items=by_item, value=value, n=len(reps), first=reps[0]['d0'] if reps else None,
                last=reps[-1]['d1'] if reps else None, loose=loose)


def reported_series():
    """للرسومات (Company): [(آخر يوم في التقرير، المرفوع للشركة المتراكم لحد التاريخ ده)]، الأقدم الأول – 15-11"""
    ds = sorted({r['d1'] for r in production_reports()})
    return [(d, reported_at(d, [])['value']) for d in ds]


# ================================================================ كود البند عند الشركة (7.09.xxx) – طلب المستخدم 15-11
# من خطة الشركة الربع سنوية في 07_Planning_Progress\PLANS_CTR\QUARTERLY: العمود A = رقم البند في العقد، B = كود الشركة.
PLAN_REL = os.path.join('07_Planning_Progress', 'PLANS_CTR', 'QUARTERLY')
_CC_CACHE = {}


def _item_key(x):
    return re.sub(r'[\s ]+', '', str(x or '')).rstrip('.')


def company_codes():
    """{رقم البند في العقد: كود الشركة} – من أحدث ملف خطة فيه عمود «كود البند»."""
    root = project_path(PLAN_REL)
    files = []
    for dp, _, fs in os.walk(root) if os.path.isdir(root) else []:
        files += [os.path.join(dp, f) for f in fs if f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]
    out = {}
    for p in sorted(files, key=os.path.getmtime):            # الأحدث آخر عشان يكتب فوق القديم
        try: mt = os.path.getmtime(p)
        except OSError: continue
        c = _CC_CACHE.get(p)
        if not c or c[0] != mt:
            m = {}
            try:
                wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                for ws in wb.worksheets:
                    for row in ws.iter_rows(max_col=3, values_only=True):
                        a, b = (row + (None, None))[:2]
                        if a is None or b is None: continue
                        b = str(b).strip()
                        if re.fullmatch(r'\d+(\.\d+){1,4}', b) and re.fullmatch(r'\d+(\.\d+)*', _item_key(a)):
                            m.setdefault(_item_key(a), b)
                wb.close()
            except Exception:
                m = {}
            c = _CC_CACHE[p] = (mt, m)
        out.update(c[1])
    return out
