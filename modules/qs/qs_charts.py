# -*- coding: utf-8 -*-
"""رسومات التحديثات (من التحديثات المتسجلة للجهة وأساس الحساب المختارين):
1) المنفّذ: رسمة كبيرة بزراير This month (كل تحديث في الشهر قصاد اللي قبله) / شهري / ربع سنوي / نص سنوي / سنوي،
   وفوق كل عمود بكام % أكتر أو أقل من اللي قبله. الفترة اللي لسه مخلصتش مكتوب عليها in progress (من غير نسبة).
2) منحنى التقدّم: الإجمالي المتراكم مع كل تحديث متسجل + نسبته من قيمة العقد.
3) نوع الشغل: نفس الاختيار، والعمود متقسم حفر / فرشة / مواسير / ردم / قاعدة طريق / أسفلت.
نفس الأرقام في شاشة البرنامج (Canvas من غير مكتبات) وفي ملف التحديث (شيت Charts في الإكسل → PDF)."""
import calendar
import datetime as dt
import math

MON = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
WD = ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')
KINDS = [('U', 'This month'), ('M', 'Monthly'), ('Q', 'Quarterly'), ('H', 'Half-year'), ('Y', 'Yearly')]
# نوع الشغل من رقم البند: اللي بينتهي بنقطة = كل البنود اللي بتبدأ بيه
GROUPS = [('Excavation', ('31.2.3.',), '#9C6B3C'), ('Bedding', ('31.2.4.3',), '#E2B13C'), ('Pipes', ('33.2.3.',), '#1F4E78'),
          ('Backfill', ('31.2.4.4',), '#5E9E3A'), ('Road base', ('31.2.4.5',), '#9AA3AE'), ('Asphalt', ('32.4.',), '#3A3F45'),
          ('Other', (), '#B4C7E7')]
NAVY, BLUE, PALE, ORANGE, GREEN, RED, GRID, TXT = '#1F4E78', '#5B9BD5', '#A9C8EA', '#C55A11', '#2E7D4F', '#C00000', '#E3E8EF', '#556070'
FONT = 'Segoe UI'
EMPTY = 'No saved updates yet – save an update from «Weekly & Monthly Update»'


# ---------------------------------------------------------------- الأرقام
def work_group(item):
    s = str(item or '').strip()
    for name, keys, _ in GROUPS:
        if any(s.startswith(k) if k.endswith('.') else s == k for k in keys): return name
    return 'Other'


def group_values(pairs):
    """[(رقم البند, قيمة)] ← {نوع الشغل: قيمة}"""
    g = {}
    for item, v in pairs:
        n = work_group(item); g[n] = g.get(n, 0.0) + (v or 0.0)
    return g


def _key(d, kind):
    if kind == 'M': return (d.year, d.month)
    if kind == 'Q': return (d.year, (d.month - 1) // 3 + 1)
    if kind == 'H': return (d.year, 1 if d.month <= 6 else 2)
    return (d.year, 0)


def _label(k, kind):
    y, n = k
    return {'M': f'{MON[n - 1]} {y}', 'Q': f'Q{n} {y}', 'H': f'H{n} {y}'}.get(kind, str(y))


def _end(k, kind):
    """آخر يوم في الفترة"""
    y, n = k
    m = {'M': n, 'Q': n * 3, 'H': n * 6}.get(kind, 12)
    return dt.date(y, m, calendar.monthrange(y, m)[1])


def buckets(ht, kind, ref=None):
    """ht = [(تاريخ, إجمالي)] ← [(الفترة, المنفّذ فيها, المتراكم في آخرها, لسه مخلصتش؟, اسمها)].
    آخر تحديث متسجل في الفترة هو اللي بيتاخد. ref = النهارده (أو تاريخ التحديث): الفترة اللي آخرها بعده لسه مخلصتش."""
    end = {}
    for d, v in sorted(ht, key=lambda x: x[0]): end[_key(d, kind)] = v
    out, prev = [], 0.0
    for k in sorted(end):
        lab = _label(k, kind)
        out.append((lab, end[k] - prev, end[k], bool(ref and _end(k, kind) > ref), lab)); prev = end[k]
    return out


def group_buckets(hg, kind, ref=None):
    """hg = [(تاريخ, {نوع الشغل: قيمة})] ← [(الفترة, {نوع الشغل: المنفّذ في الفترة}, لسه مخلصتش؟)]"""
    end = {}
    for d, g in sorted(hg, key=lambda x: x[0]): end[_key(d, kind)] = g
    out, prev = [], {}
    for k in sorted(end):
        cur = end[k]
        out.append((_label(k, kind), {n: cur.get(n, 0.0) - prev.get(n, 0.0) for n, _, _ in GROUPS}, bool(ref and _end(k, kind) > ref)))
        prev = cur
    return out


def month_of(ht, month=None):
    """الشهر اللي هيتعرض في This month: month لو فيه تحديثات متسجلة، وإلا آخر شهر فيه تحديثات"""
    ms = sorted({(d.year, d.month) for d, _ in ht})
    if not ms: return None
    return month if month in ms else ms[-1]


def _uname(d, names):
    n = (names or {}).get(d)
    return (f'{n} · ' if n else f'{WD[d.weekday()]} ') + d.strftime('%d-%m')


def update_bars(ht, month, names=None):
    """تحديثات الشهر: كل تحديث متسجل عمود بالمنفّذ من التحديث اللي قبله (حتى لو اللي قبله في الشهر اللي فات).
    names = {تاريخ: 'P1'} من تقويم التحديث. ← ([(اسم التحديث, المنفّذ, المتراكم, False, تاريخه)],
    before = (المنفّذ في التحديث اللي قبل أول تحديث في الشهر, تاريخه) – عشان أول تحديث يتقارن بيه)"""
    pts = sorted(ht, key=lambda x: x[0])
    rows, before = [], None
    for i, (d, v) in enumerate(pts):
        inc = v - (pts[i - 1][1] if i else 0.0)
        if (d.year, d.month) == month: rows.append((_uname(d, names), inc, v, False, d.strftime('%d-%m')))
        elif (d.year, d.month) < month: before = (inc, d.strftime('%d-%m'))
    return rows, before


def update_groups(hg, month, names=None):
    """نفس تحديثات الشهر متقسمة على نوع الشغل: [(اسم التحديث, {نوع الشغل: المنفّذ من التحديث اللي قبله}, False)]"""
    pts = sorted(hg, key=lambda x: x[0])
    out = []
    for i, (d, g) in enumerate(pts):
        if (d.year, d.month) != month: continue
        p = pts[i - 1][1] if i else {}
        out.append((_uname(d, names), {n: g.get(n, 0.0) - p.get(n, 0.0) for n, _, _ in GROUPS}, False))
    return out


def with_change(data, before=None):
    """[(اسم, منفّذ, متراكم, لسه مخلصتش, اسم قصير)] ← [(اسم, منفّذ, متراكم, %, اسم اللي قبله, لسه مخلصتش)].
    اللي لسه مخلصتش ملهاش نسبة (مش عدل تتقارن بفترة كاملة). before = (منفّذ, اسم) للي قبل أول واحد."""
    out, prev = [], before
    for lab, v, c, op, short in data:
        p = prev[0] if prev else None
        pct = None if op or p is None or p <= 0 else (v - p) / p * 100
        out.append((lab, v, c, pct, prev[1] if prev else '', op))
        prev = (v, short)
    return out


def change_text(pct, plab='', short=False):
    """النسبة بالكلام بدل السالب والموجب: «24% more than Oct 2026» / «15% less than 04-11» (short: من غير than)"""
    if abs(pct) < 0.5: return 'same' if short or not plab else f'same as {plab}'
    w = f'{abs(pct):,.0f}% {"more" if pct > 0 else "less"}'
    return w + (f' than {plab}' if plab and not short else '')


def _k(v):
    a = abs(v)
    if a >= 1e6: return f'{v / 1e6:,.2f}M'
    if a >= 1e3: return f'{v / 1e3:,.0f}K'
    return f'{v:,.0f}'


def _n(v): return f'{v:,.0f}'


def _nice(v):
    """أعلى المحور رقم مدوّر (التقسيمات أرباع)"""
    if v <= 0: return 1.0
    e = 10 ** math.floor(math.log10(v)); f = v / e
    for s in (1, 1.2, 1.6, 2, 2.4, 3, 4, 5, 6, 8, 10):
        if f <= s + 1e-9: return s * e
    return 10 * e


def _ink(col):
    """لون الكتابة على اللون ده (أبيض على الغامق، كحلي على الفاتح)"""
    r, g, b = (int(col[i:i + 2], 16) for i in (1, 3, 5))
    return '#1F2937' if 0.299 * r + 0.587 * g + 0.114 * b > 150 else 'white'


# ---------------------------------------------------------------- الرسم في البرنامج (Canvas)
def _frame(cv, title, sub):
    cv.delete('all')
    w, h = max(cv.winfo_width(), 320), max(cv.winfo_height(), 200)
    cv.create_text(16, 12, text=title, anchor='nw', fill=NAVY, font=(FONT, 12, 'bold'))
    if sub: cv.create_text(16, 36, text=sub, anchor='nw', fill=TXT, font=(FONT, 9))
    return w, h


def _empty(cv, w, h, msg=EMPTY):
    cv.create_text(w / 2, h / 2 + 10, text=msg, fill=TXT, font=(FONT, 10))


def _grid(cv, L, T, pw, ph, top):
    for i in range(5):
        y = T + ph - ph * i / 4
        cv.create_line(L, y, L + pw, y, fill=GRID)
        cv.create_text(L - 8, y, text=_k(top * i / 4), anchor='e', fill=TXT, font=(FONT, 8))
    cv.create_line(L, T + ph, L + pw, T + ph, fill='#AEB8C4')


def _fit(cv, rows, pw, w, R):
    """لو الأعمدة أكتر من اللي يساعه العرض: آخرها بس، ومكتوب فوق على اليمين."""
    fit = max(1, int(pw // 70))
    if len(rows) <= fit: return rows
    cv.create_text(w - R, 38, text=f'last {fit} of {len(rows)} shown – pick a longer period to see them all', anchor='ne',
                   fill=TXT, font=(FONT, 8))
    return rows[-fit:]


def _legend(cv, w, R, items, y=62):
    """مفتاح سطر واحد فوق على اليمين: [(الاسم، اللون)]"""
    x = w - R
    for name, col in reversed(items):
        t = cv.create_text(x, y, text=name, anchor='e', fill=TXT, font=(FONT, 9))
        x0 = cv.bbox(t)[0]
        cv.create_rectangle(x0 - 15, y - 5, x0 - 5, y + 5, fill=col, outline='')
        x = x0 - 26


def draw_done(cv, rows, title, sub='', empty=EMPTY, rep=None):
    """رسمة واحدة كبيرة: عمود لكل فترة (أو تحديث) عليه قيمته، وفوقه بكام % أكتر أو أقل من اللي قبله؛
    اللي لسه مخلصتش لونه فاتح ومكتوب عليه in progress. rows = with_change(...)
    rep = {اسم الفترة: المرفوع للشركة فيها} ← عمود برتقالي جنب الأزرق (Company بس – 15-11)"""
    w, h = _frame(cv, title, sub)
    if not rows: return _empty(cv, w, h, empty)
    L, R, T, B = 66, 20, 80, 30
    pw, ph = w - L - R, h - T - B
    rows = _fit(cv, rows, pw, w, R)
    rv = [(rep or {}).get(r[0]) for r in rows]
    top = _nice(max([r[1] for r in rows] + [v for v in rv if v is not None] + [1.0]) * 1.12)
    _grid(cv, L, T, pw, ph, top)
    slot = pw / len(rows); bw = min(slot * 0.56, 120); full = slot >= 92
    if rep is not None:
        _legend(cv, w, R, [('Calculated', BLUE), ('Reported to Company', ORANGE)])
        bw = min(slot * 0.42, 90); rw = bw * 0.5
    for i, (lab, v, _c, pct, plab, op) in enumerate(rows):
        x = L + slot * (i + 0.5)
        if rep is not None:
            x -= rw / 2 + 2
            r_ = rv[i]
            if r_ is not None:
                xr = x + bw / 2 + 4 + rw / 2; yr = T + ph - ph * max(r_, 0) / top
                cv.create_rectangle(xr - rw / 2, yr, xr + rw / 2, T + ph, fill=ORANGE, outline='')
                cv.create_text(xr + rw / 2 + 2, yr - 4, text=_n(r_) if full else _k(r_), anchor='sw', fill=ORANGE,
                               font=(FONT, 9 if full else 8, 'bold'))
        y = T + ph - ph * max(v, 0) / top
        cv.create_rectangle(x - bw / 2, y, x + bw / 2, T + ph, fill=PALE if op else BLUE, outline='')
        cv.create_text(x, y - 4, text=_n(v) if full else _k(v), anchor='s', fill=RED if v < 0 else NAVY,
                       font=(FONT, 10 if full else 9, 'bold'))
        if op:
            cv.create_text(x, y - 22, text='in progress', anchor='s', fill=TXT, font=(FONT, 8, 'italic'))
        elif pct is not None:
            cv.create_text(x, y - 22, text=change_text(pct, plab, short=slot < 170), anchor='s',
                           fill=GREEN if pct >= 0 else RED, font=(FONT, 8, 'bold'))
        cv.create_text(x + ((rw / 2 + 2) if rep is not None else 0), T + ph + 6, text=lab, anchor='n', fill=TXT, font=(FONT, 9))


def draw_stack(cv, data, title, sub='', empty=EMPTY):
    """نوع الشغل: عمود لكل فترة (أو تحديث) متقسم على أنواع الشغل، وفوقه الإجمالي. data = [(اسم, {نوع: قيمة}, لسه مخلصتش)]"""
    w, h = _frame(cv, title, sub)
    used = [(n, c) for n, _, c in GROUPS if any(g.get(n, 0) > 0.5 for _, g, _ in data)]
    if not data or not used: return _empty(cv, w, h, empty)
    L, R, T, B = 66, 20, 80, 30
    pw, ph = w - L - R, h - T - B
    rows = _fit(cv, data, pw, w, R)
    tot = [sum(g.get(n, 0) for n, _ in used) for _, g, _ in rows]
    top = _nice(max([sum(max(g.get(n, 0), 0) for n, _ in used) for _, g, _ in rows] + [1.0]) * 1.12)
    _grid(cv, L, T, pw, ph, top)
    slot = pw / len(rows); bw = min(slot * 0.56, 120); full = slot >= 92
    for i, (lab, g, op) in enumerate(rows):
        x = L + slot * (i + 0.5); base = T + ph
        for name, col in used:
            v = max(g.get(name, 0), 0)
            if v <= 0: continue
            hh = ph * v / top
            cv.create_rectangle(x - bw / 2, base - hh, x + bw / 2, base, fill=col, outline='white')
            if hh >= 16 and bw >= 40: cv.create_text(x, base - hh / 2, text=_k(v), fill=_ink(col), font=(FONT, 8, 'bold'))
            base -= hh
        cv.create_text(x, base - 4, text=_n(tot[i]) if full else _k(tot[i]), anchor='s', fill=NAVY, font=(FONT, 10 if full else 9, 'bold'))
        if op: cv.create_text(x, base - 22, text='in progress', anchor='s', fill=TXT, font=(FONT, 8, 'italic'))
        cv.create_text(x, T + ph + 6, text=lab, anchor='n', fill=TXT, font=(FONT, 9))
    x = w - R                                               # المفتاح: سطر واحد فوق على اليمين
    for name, col in reversed(used):
        t = cv.create_text(x, 62, text=name, anchor='e', fill=TXT, font=(FONT, 9))
        x0 = cv.bbox(t)[0]
        cv.create_rectangle(x0 - 15, 57, x0 - 5, 67, fill=col, outline='')
        x = x0 - 26


def draw_curve(cv, ht, contract, title, sub='', rep=None):
    """منحنى التقدّم: الإجمالي المتراكم مع كل تحديث متسجل، ونسبته من قيمة العقد.
    rep = [(تاريخ التقرير، المرفوع للشركة المتراكم)] ← خط برتقالي متقطع (Company بس – 15-11)"""
    w, h = _frame(cv, title, sub)
    pts = sorted(ht, key=lambda x: x[0])
    if not pts: return _empty(cv, w, h)
    rp = sorted(rep or [], key=lambda x: x[0])
    if contract:
        cv.create_text(w - 20, 8, text=f'{pts[-1][1] / contract * 100:.2f}%', anchor='ne', fill=ORANGE, font=(FONT, 17, 'bold'))
        cv.create_text(w - 20, 40, text=f'of the contract ({contract / 1e6:,.2f}M SAR)', anchor='ne', fill=ORANGE, font=(FONT, 8))
    L, R, T, B = 66, 30, 80, 30
    pw, ph = w - L - R, h - T - B
    top = _nice(max([v for _, v in pts] + [v for _, v in rp]) * 1.15)
    _grid(cv, L, T, pw, ph, top)
    d0, d1 = min([pts[0][0]] + [d for d, _ in rp]), max([pts[-1][0]] + [d for d, _ in rp])
    span = (d1 - d0).days
    X = lambda d: L + 28 + (pw - 56) * (d - d0).days / span if span else L + pw / 2
    Y = lambda v: T + ph - ph * max(v, 0) / top
    xy = [(X(d), Y(v)) for d, v in pts]
    if len(xy) > 1:
        cv.create_polygon(*[c for p in [(xy[0][0], T + ph)] + xy + [(xy[-1][0], T + ph)] for c in p], fill='#DCE9F6', outline='')
        cv.create_line(*[c for p in xy for c in p], fill=NAVY, width=3)
    r = 4 if len(xy) <= 40 else 2
    for x, y in xy: cv.create_oval(x - r, y - r, x + r, y + r, fill=NAVY, outline='white')
    fmt = '%d-%m' if d0.year == d1.year else '%d-%m-%y'
    keep, lastx = [], None                                  # الأرقام عند التحديثات اللي القيمة اتغيرت فيها، متباعدة، والأخير دايمًا
    for i in range(len(pts) - 1, -1, -1):
        changed = i == 0 or abs(pts[i][1] - pts[i - 1][1]) > 0.5
        if i != len(pts) - 1 and not changed: continue
        if lastx is not None and lastx - xy[i][0] < 66: continue
        keep.append(i); lastx = xy[i][0]
    for i in keep:
        (d, v), (x, y) = pts[i], xy[i]
        cv.create_text(x, y - 8, text=_n(v), anchor='s', fill=NAVY, font=(FONT, 9, 'bold'))
        cv.create_text(x, T + ph + 6, text=d.strftime(fmt), anchor='n', fill=TXT, font=(FONT, 8))
    if rep is not None:
        _legend(cv, w, R, [('Calculated', NAVY), ('Reported to Company', ORANGE)])
        rxy = [(X(d), Y(v)) for d, v in rp]
        if len(rxy) > 1: cv.create_line(*[c for p in rxy for c in p], fill=ORANGE, width=3, dash=(7, 4))
        shown = {round(xy[i][0]) for i in keep}
        for (d, v), (x, y) in zip(rp, rxy):
            cv.create_oval(x - 5, y - 5, x + 5, y + 5, fill=ORANGE, outline='white')
            cv.create_text(x + 9, y + 2, text=_n(v), anchor='nw', fill=ORANGE, font=(FONT, 9, 'bold'))     # يمين النقطة وتحتها
            if round(x) not in shown: cv.create_text(x, T + ph + 6, text=d.strftime(fmt), anchor='n', fill=ORANGE, font=(FONT, 8))


# ---------------------------------------------------------------- صفحات الرسومات في ملف التحديث (إكسل → PDF)
def rep_by_update(ht, rt, month, names=None):
    """المرفوع للشركة بين كل تحديث في الشهر واللي قبله (لو فيه تقرير إنتاج في الفترة دي): {اسم التحديث: قيمة}"""
    pts = sorted(ht, key=lambda x: x[0])
    rc = lambda x: max([v for d_, v in rt if d_ <= x] or [0.0])
    out = {}
    for i, (d, _v) in enumerate(pts):
        if (d.year, d.month) != month: continue
        p_ = pts[i - 1][0] if i else dt.date.min
        if any(p_ < d_ <= d for d_, _ in rt): out[_uname(d, names)] = rc(d) - rc(p_)
    return out


def add_sheet(wb, ht, hg, contract, title, ref=None, names=None, rep=None):
    """شيت «Charts» – 3 صفحات A4 بالعرض:
    1) منحنى التقدّم + تحديثات الشهر (كل تحديث قصاد اللي قبله)
    2) المنفّذ شهري (وتحت كل شهر بكام % أكتر أو أقل من اللي قبله) + المنفّذ شهري متقسم على نوع الشغل
    3) ربع سنوي / نص سنوي / سنوي.
    ref = تاريخ التحديث (الشهر بتاعه + الفترات اللي لسه مخلصتش)، names = {تاريخ: 'P1'} من تقويم التحديث.
    جداول الأرقام في الأعمدة من R (برا منطقة الطباعة)."""
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.legend import Legend
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.label import DataLabelList, DataLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.styles import Font
    from openpyxl.worksheet.pagebreak import Break
    from openpyxl.worksheet.page import PageMargins
    if 'Charts' in wb.sheetnames: del wb['Charts']
    ws = wb.create_sheet('Charts')
    pts = sorted(ht, key=lambda x: x[0])
    ref = ref or (pts[-1][0] if pts else None)
    ws['A1'] = title; ws['A1'].font = Font(bold=True, size=14, color=NAVY[1:])
    info = f'{len(pts)} updates' + (f'   ·   cumulative to {pts[-1][0]:%d-%m-%Y}: {pts[-1][1]:,.0f} SAR' if pts else '')
    if pts and contract: info += f'   ·   {pts[-1][1] / contract * 100:.2f}% of the contract ({contract:,.0f} SAR)'
    ws['A2'] = info; ws['A2'].font = Font(size=10, color='595959')
    nxt = [18]                                              # أول عمود فاضي للجداول (R)
    OFF = dict(showLegendKey=False, showCatName=False, showSerName=False, showPercent=False, showBubbleSize=False)

    def table(head, rows):
        c0 = nxt[0]
        for j, hd in enumerate(head): ws.cell(1, c0 + j, hd).font = Font(bold=True)
        for i, r in enumerate(rows, 2):
            for j, v in enumerate(r):
                x = ws.cell(i, c0 + j, round(v, 2) if isinstance(v, float) else v)
                if isinstance(v, float): x.number_format = '#,##0'
        nxt[0] += len(head) + 1
        return c0, len(rows)

    def labels(**kw):                                       # الرقم بس (من غير اسم السلسلة ولا الشهر)
        return DataLabelList(showVal=True, **OFF, **kw)

    def place(ch, cell, wcm, hcm):
        ch.width, ch.height = wcm, hcm; ch.y_axis.scaling.min = 0
        # إكسل: لو delete مش مكتوبة المحاور بتختفي، ولو overlay مش مكتوبة العنوان بيركب على الرسمة
        ch.x_axis.delete = False; ch.y_axis.delete = False; ch.x_axis.axPos = 'b'
        ch.y_axis.number_format = '#,##0'
        ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill='D9D9D9')))
        if ch.title is not None: ch.title.overlay = False
        ch.roundedCorners = False
        ws.add_chart(ch, cell)

    def bars(rows, head, ttl, cell, wcm, hcm, change=True, rmap=None):
        """أعمدة المنفّذ؛ تحت كل عمود بكام % أكتر أو أقل من اللي قبله، أو in progress.
        rmap = {اسم الفترة: المرفوع للشركة} ← عمود برتقالي جنب الأزرق (Company)"""
        cats = [lab + ('\nin progress' if op else f'\n{change_text(pct, short=True)}' if change and pct is not None else '')
                for lab, v, c, pct, plab, op in rows]
        if rmap is not None:
            c0, n = table([head, 'Calculated', 'Reported to Company'],
                          [(cat, float(r[1]), float(rmap[r[0]]) if r[0] in rmap else None) for cat, r in zip(cats, rows)])
        else:
            c0, n = table([head, 'Done (SAR)'], [(cat, float(r[1])) for cat, r in zip(cats, rows)])
        ch = BarChart(); ch.type = 'col'; ch.legend = None; ch.gapWidth = 60; ch.title = ttl
        ch.add_data(Reference(ws, min_col=c0 + 1, max_col=c0 + (2 if rmap is not None else 1), min_row=1, max_row=n + 1),
                    titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=c0, min_row=2, max_row=n + 1))
        ch.series[0].graphicalProperties.solidFill = BLUE[1:]
        if rmap is not None:
            ch.series[1].graphicalProperties.solidFill = ORANGE[1:]; ch.legend = Legend(); ch.legend.position = 't'
        if n <= 24: ch.dataLabels = labels(numFmt='#,##0,"K"') if n > 12 else labels()
        place(ch, cell, wcm, hcm)

    # ---- صفحة 1: منحنى التقدّم
    rp = sorted(rep or [], key=lambda x: x[0])
    if pts and rep is not None:                             # Company: خط تاني بالمرفوع للشركة (التواريخ مع بعض)
        cv_ = dict(pts); rv_ = dict(rp); ds_ = sorted(set(cv_) | set(rv_))
        c0, n = table(['Date', 'Calculated', 'Reported to Company'],
                      [(d.strftime('%d-%m-%Y'), float(cv_[d]) if d in cv_ else None, float(rv_[d]) if d in rv_ else None) for d in ds_])
        ch = LineChart(); ch.title = 'Progress – cumulative value (SAR)'
        ch.add_data(Reference(ws, min_col=c0 + 1, max_col=c0 + 2, min_row=1, max_row=n + 1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=c0, min_row=2, max_row=n + 1))
        ch.display_blanks = 'span'; ch.legend = Legend(); ch.legend.position = 't'
        for s, col, pos in zip(ch.series, (NAVY, ORANGE), ('t', 'b')):          # المحسوب فوق النقطة والمرفوع تحتها
            s.marker.symbol = 'circle'; s.marker.size = 6; s.smooth = False
            s.graphicalProperties.line.solidFill = col[1:]; s.graphicalProperties.line.width = 28575
            s.marker.graphicalProperties.solidFill = col[1:]; s.marker.graphicalProperties.line.solidFill = col[1:]
            s.dLbls = labels(dLblPos=pos)
        ch.series[1].graphicalProperties.line.dashStyle = 'dash'
        place(ch, 'A4', 25.4, 7.4)
    elif pts:
        c0, n = table(['Update', 'Cumulative (SAR)'], [(d.strftime('%d-%m-%Y'), float(v)) for d, v in pts])
        ch = LineChart(); ch.title = 'Progress – cumulative value at each update (SAR)'; ch.legend = None
        ch.add_data(Reference(ws, min_col=c0 + 1, min_row=1, max_row=n + 1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=c0, min_row=2, max_row=n + 1))
        s = ch.series[0]; s.marker.symbol = 'circle'; s.marker.size = 6; s.smooth = False
        s.graphicalProperties.line.solidFill = NAVY[1:]; s.graphicalProperties.line.width = 28575
        s.marker.graphicalProperties.solidFill = NAVY[1:]; s.marker.graphicalProperties.line.solidFill = NAVY[1:]
        gap, keep, last = max(1, math.ceil(n / 12)), [], None   # الرقم عند التحديثات اللي القيمة اتغيرت فيها (12 بالكتير) والأخير دايمًا
        for i in range(n - 1, -1, -1):
            if i != n - 1 and not (i == 0 or abs(pts[i][1] - pts[i - 1][1]) > 0.5): continue
            if last is not None and last - i < gap: continue
            keep.append(i); last = i
        s.dLbls = DataLabelList(dLbl=[DataLabel(idx=i, showVal=True, dLblPos='t', **OFF) for i in sorted(keep)], showVal=False, **OFF)
        place(ch, 'A4', 25.4, 7.4)
    # ---- صفحة 1: تحديثات الشهر – كل تحديث قصاد اللي قبله
    mon = month_of(ht, (ref.year, ref.month)) if ref else None
    if mon:
        ub, before = update_bars(ht, mon, names)
        bars(with_change(ub, before), 'Update', f'{MON[mon[1] - 1]} {mon[0]} updates – value done since the update before (SAR)',
             'A20', 25.4, 7.4, rmap=rep_by_update(ht, rp, mon, names) if rep is not None else None)
    ws.row_breaks.append(Break(id=34))
    # ---- صفحة 2: شهري + نوع الشغل
    ws['A35'] = 'Value done in each period'; ws['A35'].font = Font(bold=True, size=12, color=NAVY[1:])
    mrows = with_change(buckets(ht, 'M', ref))
    if mrows:
        bars(mrows, 'Monthly', 'Monthly (SAR)  –  under each month: change from the month before', 'A37', 25.4, 7.4,
             rmap={b[0]: b[1] for b in buckets(rp, 'M', ref)} if rep is not None else None)
    gb = group_buckets(hg, 'M', ref) if hg else []
    used = [(nm, c) for nm, _, c in GROUPS if any(g.get(nm, 0) > 0.5 for _, g, _ in gb)]
    if gb and used:                                          # إجمالي كل شهر مكتوب تحت اسمه
        rows = [(f'{lab}\n{sum(g.get(nm, 0) for nm, _ in used):,.0f}' + ('\nin progress' if op else ''),)
                + tuple(float(max(g.get(nm, 0), 0)) for nm, _ in used) for lab, g, op in gb]
        c0, n = table(['Month'] + [nm for nm, _ in used], rows)
        ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100; ch.gapWidth = 80
        ch.title = 'Value done each month – by type of work (SAR)'
        ch.add_data(Reference(ws, min_col=c0 + 1, max_col=c0 + len(used), min_row=1, max_row=n + 1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=c0, min_row=2, max_row=n + 1))
        for s, (_nm, c) in zip(ch.series, used):
            s.graphicalProperties.solidFill = c[1:]; s.graphicalProperties.line.solidFill = 'FFFFFF'
        ch.legend.position = 'r'
        place(ch, 'A53', 25.4, 7.4)
    ws.row_breaks.append(Break(id=68))
    # ---- صفحة 3: ربع سنوي / نص سنوي / سنوي
    ws['A69'] = 'Quarterly  ·  Half-year  ·  Yearly'; ws['A69'].font = Font(bold=True, size=12, color=NAVY[1:])
    for (kind, name), cell in zip([k for k in KINDS if k[0] in 'QHY'], ('A71', 'F71', 'K71')):
        rows = with_change(buckets(ht, kind, ref))
        if rows: bars(rows, name, f'{name} (SAR)', cell, 8.3, 7.0, change=False,
                      rmap={b[0]: b[1] for b in buckets(rp, kind, ref)} if rep is not None else None)
    # الطباعة: A4 بالعرض، مقاس ثابت (مع «Fit to» إكسل بيتجاهل فاصل الصفحة)، وارتفاع الصفوف ثابت عشان
    # الرسومات متتزحش (الأسماء اللي تحتها الفرق % فيها سطرين في الجدول)
    for i in range(1, 16): ws.column_dimensions[chr(64 + i)].width = 8.6
    for r in range(1, max(ws.max_row, 86) + 1): ws.row_dimensions[r].height = 15
    ws.row_dimensions[1].height = 22; ws.row_dimensions[35].height = 18; ws.row_dimensions[69].height = 18
    ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.scale = 90
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.3, footer=0.3)
    ws.print_options.horizontalCentered = True
    ws.print_area = 'A1:O86'
    return ws
