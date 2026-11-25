# -*- coding: utf-8 -*-
"""مواعيد التحديث: تحديث الشركة وتحديث المالك.

الشركة: تقويم بيتبعت كل شهر (PDF أو صورة) فيه الفترات ويوم «يتم الرفع على النظام» – الـ DC بيقراه وانت بتراجع وتوافق.
المالك: كل ثلاثاء، ولما التخطيط يبلغك بتغيير بتعدّل التاريخ بإيدك (Edit Dates في شاشة التحديث).
السجل: 00_Registers\\12-LOG UPDATE DATES.xlsx – سطر لكل تاريخ تحديث.
"""
import os, re, math, calendar, tempfile, datetime as dt, unicodedata
import qs_core

REG_REL = r'00_Registers\12-LOG UPDATE DATES.xlsx'
CAL_FOLDER_REL = r'08_Cost_Payments\QS\UPDATE_CALENDAR'          # ملفات تقويم الشركة بعد الموافقة
SHEET = 'DATES'
HDR = ['Update For', 'Month', 'No.', 'From', 'Update Date', 'Day', 'Working Days', 'Source', 'Notes']
COMPANY, OWNER = 'Company', 'Owner'
KINDS = (COMPANY, OWNER)
OWNER_DAY = 1                     # الثلاثاء (الاثنين = 0)
OFF_DAY = 4                       # الجمعة إجازة
DAY = ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')
MONTH_EN = ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
            'November', 'December')


# ---------------------------------------------------------------- مساعدات
def _n(s):
    """للمقارنة: الحروف المتشكّلة ← عادية، والألف/الياء/التاء المربوطة موحّدة."""
    s = unicodedata.normalize('NFKC', s or '').replace('ـ', '')
    return s.translate(str.maketrans('أإآىة', 'ااايه'))


ORD = [(_n(w), k) for w, k in (('الأولى', 1), ('الثانية', 2), ('الثالثة', 3), ('الرابعة', 4), ('الخامسة', 5), ('السادسة', 6))]
AR_MONTH = {_n(w): k for w, k in (('يناير', 1), ('فبراير', 2), ('مارس', 3), ('أبريل', 4), ('ابريل', 4), ('مايو', 5),
                                  ('يونيو', 6), ('يوليو', 7), ('أغسطس', 8), ('اغسطس', 8), ('سبتمبر', 9), ('أكتوبر', 10),
                                  ('اكتوبر', 10), ('نوفمبر', 11), ('ديسمبر', 12))}


def day_name(d): return DAY[d.weekday()] if d else ''


def fd(d): return d.strftime('%d-%m-%Y') if d else ''


def to_date(v):
    if v is None or v == '': return None
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    s = str(v).strip()
    m = re.fullmatch(r'(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4})', s)
    if m:
        try: return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError: return None
    m = re.fullmatch(r'(\d{4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})', s)
    if m:
        try: return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError: return None
    return None


def work_days(d0, d1):
    if not (d0 and d1) or d1 < d0: return 0
    return sum(1 for i in range((d1 - d0).days + 1) if (d0 + dt.timedelta(i)).weekday() != OFF_DAY)


def ym(d): return f'{d.year}-{d.month:02d}'


def add_month(y, m, k):
    t = y * 12 + (m - 1) + k
    return t // 12, t % 12 + 1


# ---------------------------------------------------------------- المواعيد المعتادة
def company_default(y, m, source='Usual pattern'):
    """من غير التقويم: كل فترة أسبوع لحد السبت (الجمعة إجازة)، والأخيرة لآخر الشهر – الحتة الصغيرة (أقل من 3 أيام) بتتضم للي جنبها."""
    n = calendar.monthrange(y, m)[1]
    parts, cur = [], []
    for k in range(1, n + 1):
        d = dt.date(y, m, k)
        if d.weekday() != OFF_DAY: cur.append(d)
        if (d.weekday() == 5 or k == n) and cur: parts.append(cur); cur = []
    if len(parts) > 1 and len(parts[-1]) < 3:
        last = parts.pop(); parts[-1] = parts[-1] + last
    if len(parts) > 1 and len(parts[0]) < 3:
        first = parts.pop(0); parts[0] = first + parts[0]
    return [dict(kind=COMPANY, month=f'{y}-{m:02d}', no=i, d0=p[0], d1=p[-1], days=len(p), source=source, notes='')
            for i, p in enumerate(parts, 1)]


def owner_default(y, m):
    """تحديث المالك: كل ثلاثاء في الشهر (الفترة من الأربع اللي قبله)."""
    out = []
    for k in range(1, calendar.monthrange(y, m)[1] + 1):
        d = dt.date(y, m, k)
        if d.weekday() == OWNER_DAY:
            d0 = d - dt.timedelta(days=6)
            out.append(dict(kind=OWNER, month=f'{y}-{m:02d}', no=len(out) + 1, d0=d0, d1=d, days=work_days(d0, d),
                            source='Every Tuesday', notes=''))
    return out


# ---------------------------------------------------------------- قراية تقويم الشركة
def month_of(text, fname=''):
    """(السنة، الشهر) من عنوان التقويم، وبعدين من اسم الملف. سطر «إجمالي الأيام لشهر …» مش بيتاخد (ساعات بيبقى غلط)."""
    lines = [l for l in (text or '').splitlines() if l.strip()]
    head = '\n'.join(lines[:3])
    y = m = None
    for i, name in enumerate(MONTH_EN, 1):
        if re.search(name, head, re.I): m = i; break
    if not m:
        h = _n(head)
        for w, k in AR_MONTH.items():
            if w in h: m = k; break
    yy = re.search(r'(20\d\d)', head)
    if yy: y = int(yy.group(1))
    base = os.path.basename(fname or '')
    if not m:
        for i, name in enumerate(MONTH_EN, 1):
            if re.search(rf'(?<![A-Za-z]){name[:3]}(?:{name[3:]})?(?![A-Za-z])', base, re.I): m = i; break
        mm = re.search(r'(20\d\d)\s*[-_. ]\s*(0?[1-9]|1[0-2])(?!\d)', base) or None
        if not m and mm: m = int(mm.group(2))
    if not y:
        yy = re.search(r'(20\d\d)', base)
        if yy: y = int(yy.group(1))
    return y, m


def _ords(s):
    """عدد مرات كل «فترة» في الكلام: {رقم الفترة: العدد}."""
    s = _n(s)
    return {k: s.count(w) for w, k in ORD if s.count(w)}


def weeks(y, m):
    """صفوف التقويم (الأسبوع من الأحد للسبت) – أيام الشهر بس."""
    out, cur = [], []
    n = calendar.monthrange(y, m)[1]
    for k in range(1, n + 1):
        d = dt.date(y, m, k); cur.append(d)
        if d.weekday() == 5 or k == n: out.append(cur); cur = []
    return out


def parse_calendar(text, fname='', year=None, month=None):
    """تقويم الشركة ← فترات التحديث. يرجّع dict(year, month, company=[صفوف], warn=[…], how='file'/'pattern').
    التقويم: صف أيام (أرقام) وتحته اسم الفترة لكل يوم شغل، و«يتم الرفع على النظام» تحت يوم الرفع."""
    y, m = month_of(text, fname)
    y, m = year or y, month or m
    warn = []
    if not (y and m):
        t = dt.date.today(); y, m = y or t.year, m or t.month
        warn.append('The month was not found in the file – check it')
    rows, summary, done = [], {}, False
    lab_lines = []                                       # احتياطي للصورة: سطور «الفترة …» من غير أرقام الأيام
    for ln in (text or '').splitlines():
        toks = ln.split()
        if not toks: continue
        nl = _n(ln)
        if re.search(r'\d+\s*(ايام|يوم)', nl) or 'اجما' in nl:          # ملخص تحت التقويم: «الفترة الأولى 6 أيام» / الإجمالي
            for w, k in ORD:
                for mm in re.finditer(w + r'\s*(\d{1,2})\s*(ايام|يوم)', nl): summary[k] = int(mm.group(1))
            done = True; continue
        if done: continue
        nums = [re.sub(r'\D', '', x) for x in toks]                     # صف الأيام (الـ OCR ساعات بيزوّد علامة)
        nums = [int(x) for x in nums if 0 < len(x) <= 2 and 1 <= int(x) <= 31]
        if len(nums) >= 5 and len(nums) >= 0.8 * len(toks):
            rows.append(dict(days=nums, block=[])); continue
        if rows: rows[-1]['block'].append(ln)
        if _ords(nl): lab_lines.append([ln])
        elif lab_lines: lab_lines[-1].append(ln)                         # «يتم الرفع على النظام» تحت سطر الفترة
    n_days = calendar.monthrange(y, m)[1]
    grid = []                                            # [(أيام الصف, كلام تحته)]
    if len(rows) >= 3:
        firsts = 0
        for r in rows:
            ds = []
            for d in r['days']:
                if d == 1: firsts += 1
                if firsts == 1 and d <= n_days: ds.append(dt.date(y, m, d))
            grid.append((ds, '\n'.join(r['block'])))
    elif lab_lines and len(lab_lines) == len(weeks(y, m)):              # صورة: كل سطر فترات = أسبوع من التقويم
        grid = [(w, '\n'.join(b)) for w, b in zip(weeks(y, m), lab_lines)]
        warn.append('Read from the picture without the day numbers – check the dates against it')
    elif summary and sorted(summary) == list(range(1, len(summary) + 1)):  # صورة: من ملخص «الفترة … X أيام» بالترتيب
        work = [d for w in weeks(y, m) for d in w if d.weekday() != OFF_DAY]
        need = sum(summary.values())
        if need <= len(work):
            company, i = [], 0
            for k in sorted(summary):
                ds = work[i:i + summary[k]]; i += summary[k]
                company.append(dict(kind=COMPANY, month=f'{y}-{m:02d}', no=k, d0=ds[0], d1=ds[-1], days=len(ds),
                                    source=os.path.basename(fname or '') or 'Calendar', notes=''))
            if need < len(work):
                company[-1]['d1'] = work[-1]                     # آخر فترة لآخر الشهر
                warn.append(f'The periods add up to {need} days but the month has {len(work)} working days – check')
            warn.append('Read from the days of each period (the grid itself was not clear) – check the dates against the picture')
            return dict(year=y, month=m, company=company, warn=warn, how='summary')
    if not grid:
        warn.append('The calendar grid could not be read – the usual pattern is filled in, check it against the file')
        return dict(year=y, month=m, company=company_default(y, m), warn=warn, how='pattern')
    lab = {}                                             # يوم ← (رقم الفترة, رقم الصف)
    upload_rows = set()
    for i, (dates, blk) in enumerate(grid):
        o = _ords(blk)
        if _n('الرفع') in _n(blk) or _n('النظام') in _n(blk): upload_rows.add(i)
        n_lab = sum(o.values())
        if not n_lab or not dates: continue
        work = [d for d in dates if d.weekday() != OFF_DAY]
        if n_lab == len(work): got = work
        elif n_lab == len(dates): got = dates
        else:
            got = work[:n_lab] if n_lab < len(work) else dates[:n_lab]
            warn.append(f'Week of {fd(dates[0])}: {n_lab} working days marked but {len(work)} expected – check')
        seq = [k for k in sorted(o) for _ in range(o[k])]
        for d, k in zip(got, seq):
            lab[d] = (k, i)
    if not lab:
        warn.append('No periods found in the calendar – the usual pattern is filled in, check it against the file')
        return dict(year=y, month=m, company=company_default(y, m), warn=warn, how='pattern')
    src = os.path.basename(fname or '') or 'Calendar'
    company = []
    if not upload_rows: warn.append('No "upload" marks were read – the last day of each period was taken as its update date')
    for k in sorted({v[0] for v in lab.values()}):
        ds = sorted(d for d, v in lab.items() if v[0] == k)
        last_row = lab[ds[-1]][1]
        if upload_rows and last_row not in upload_rows:
            warn.append(f'Period {k}: no "upload" mark found – its last day was taken as the update date')
        if k in summary and summary[k] != len(ds):
            warn.append(f'Period {k}: the file says {summary[k]} days, {len(ds)} were read – check')
        company.append(dict(kind=COMPANY, month=f'{y}-{m:02d}', no=k, d0=ds[0], d1=ds[-1], days=len(ds), source=src, notes=''))
    for a, b in zip(company, company[1:]):
        if b['d0'] <= a['d1']: warn.append(f"Periods {a['no']} and {b['no']} overlap – check")
    return dict(year=y, month=m, company=company, warn=warn, how='file')


# ---------------------------------------------------------------- السجل
def reg_path():
    return qs_core.project_path(REG_REL)


def load(path=None):
    """كل صفوف السجل: dict(kind, month, no, d0, d1, days, source, notes)."""
    p = path or reg_path()
    if not p or not os.path.isfile(p): return []
    import openpyxl
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    try:
        ws = wb[SHEET] if SHEET in wb.sheetnames else wb.worksheets[0]
        out = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            r = list(r) + [None] * 9
            k = str(r[0] or '').strip().capitalize()
            d1 = to_date(r[4])
            if k not in KINDS or not d1: continue
            try: no = int(float(r[2]))
            except (TypeError, ValueError): no = None
            try: days = int(float(r[6]))
            except (TypeError, ValueError): days = None
            out.append(dict(kind=k, month=str(r[1] or ym(d1)).strip(), no=no, d0=to_date(r[3]), d1=d1, days=days,
                            source=str(r[7] or ''), notes=str(r[8] or '')))
        return out
    finally:
        wb.close()


def month_rows(rows, kind, y, m):
    key = f'{y}-{m:02d}'
    return sorted((r for r in rows if r['kind'] == kind and r['month'] == key), key=lambda r: r['d1'])


def save_month(y, m, company=None, owner=None, path=None):
    """يبدّل تواريخ الشهر ده في السجل (الشركة و/أو المالك) – باقي الشهور زي ما هي. يرجّع مسار السجل."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    p = path or reg_path()
    key = f'{y}-{m:02d}'
    rows = load(p)
    for kind, new in ((COMPANY, company), (OWNER, owner)):
        if new is None: continue
        rows = [r for r in rows if not (r['kind'] == kind and r['month'] == key)]
        for i, r in enumerate(sorted(new, key=lambda r: r['d1']), 1):
            rows.append(dict(r, kind=kind, month=key, no=i))
    order = {COMPANY: 0, OWNER: 1}
    rows.sort(key=lambda r: (r['month'], order[r['kind']], r['d1']))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = SHEET
    ws.sheet_view.rightToLeft = False
    thin = Side(style='thin', color='A6A6A6'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    H = PatternFill('solid', fgColor='1F4E78'); FILL = {COMPANY: 'FFFFFF', OWNER: 'EEF4FB'}
    CEN = Alignment(horizontal='center', vertical='center')
    for c, h in enumerate(HDR, 1):
        x = ws.cell(1, c, h); x.font = Font(bold=True, color='FFFFFF'); x.fill = H; x.alignment = CEN; x.border = B
    for i, r in enumerate(rows, 2):
        vals = [r['kind'], r['month'], r['no'], r['d0'], r['d1'], day_name(r['d1']),
                r['days'] if r.get('days') is not None else work_days(r['d0'], r['d1']), r.get('source') or '', r.get('notes') or '']
        for c, v in enumerate(vals, 1):
            x = ws.cell(i, c, v); x.border = B; x.alignment = CEN; x.fill = PatternFill('solid', fgColor=FILL[r['kind']])
            if c in (4, 5): x.number_format = 'dd-mm-yyyy'
    for c, w in zip('ABCDEFGHI', (12, 10, 6, 12, 13, 7, 13, 34, 30)): ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = f'A1:I{max(len(rows) + 1, 1)}'
    os.makedirs(os.path.dirname(os.path.abspath(p)), exist_ok=True)
    fdh, tmp = tempfile.mkstemp(prefix='upd_', suffix='.xlsx', dir=os.path.dirname(os.path.abspath(p))); os.close(fdh)
    try:
        wb.save(tmp)
        os.replace(tmp, p)                                # لو السجل مفتوح في Excel ← PermissionError
    finally:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass
    return p


# ---------------------------------------------------------------- المواعيد للشاشات
def dates(kind, rows=None, today=None, back=1, ahead=1):
    """مواعيد النوع ده من الشهر اللي فات للشهر الجاي: الشركة من السجل بس، والمالك من السجل ولو الشهر مش متسجل = كل ثلاثاء."""
    rows = load() if rows is None else rows
    today = today or dt.date.today()
    months = [f'{y}-{m:02d}' for y, m in (add_month(today.year, today.month, k) for k in range(-back, ahead + 1))]
    out = [r for r in rows if r['kind'] == kind and r['month'] in months]
    if kind == OWNER:
        have = {r['month'] for r in out}
        for key in months:
            if key not in have: out += owner_default(*(int(x) for x in key.split('-')))
    return sorted(out, key=lambda r: r['d1'])


def next_due(kind, rows=None, today=None):
    today = today or dt.date.today()
    nx = [r for r in dates(kind, rows, today) if r['d1'] >= today]
    return nx[0] if nx else None


def label(r):
    """للقايمة في شاشة التحديث."""
    if r['kind'] == COMPANY:
        return f"P{r['no']}  {r['d0']:%d-%m} → {r['d1']:%d-%m}   ({day_name(r['d1'])} {fd(r['d1'])})"
    return f"{day_name(r['d1'])} {fd(r['d1'])}" + ('' if r.get('source') == 'Every Tuesday' else '   (changed)')


def due_text(today=None, rows=None):
    """سطر التذكير: التحديث الجاي للشركة وللمالك."""
    today = today or dt.date.today()
    rows = load() if rows is None else rows
    def when(d):
        n = (d - today).days
        return 'today' if n == 0 else ('tomorrow' if n == 1 else f'in {n} days')
    parts = []
    c = next_due(COMPANY, rows, today)
    if c: parts.append(f"Company: {day_name(c['d1'])} {c['d1']:%d-%m} (P{c['no']}) – {when(c['d1'])}")
    else: parts.append(f'Company: no calendar for {MONTH_EN[today.month - 1]} yet – put it in 00_INBOX')
    o = next_due(OWNER, rows, today)
    if o: parts.append(f"Owner: {day_name(o['d1'])} {o['d1']:%d-%m} – {when(o['d1'])}")
    return '   ·   '.join(parts)


# ---------------------------------------------------------------- شاشة المراجعة
def review(parent, y, m, loader, image=None, original=None, title='Update Dates', auto_open=True):
    """تراجع تواريخ الشهر وتعدّلها (والصورة جنبها). loader(y, m) ← (company, owner, warn).
    يرجّع (y, m, company, owner) لو دوست Save Dates، أو None."""
    import tkinter as tk
    from tkinter import ttk, messagebox
    NAVY = '#1F4E78'
    top = tk.Toplevel(parent); top.title(title)
    try: top.transient(parent.winfo_toplevel())
    except Exception: pass
    top.configure(padx=12, pady=10)
    res = {}
    # ---- الصورة
    left = ttk.Frame(top); left.grid(row=0, column=0, sticky='n', padx=(0, 14))
    shown = False
    if image and os.path.isfile(image):
        try:
            img = tk.PhotoImage(file=image)
            k = max(1, math.ceil(max(img.width() / 760, img.height() / 560)))
            if k > 1: img = img.subsample(k)
            lb = tk.Label(left, image=img, bd=1, relief='solid'); lb.image = img; lb.pack()
            shown = True
        except Exception:
            shown = False
    if not shown:
        ttk.Label(left, text='No picture here – press "Open File"\nto check the dates on the original.',
                  foreground='#7F7F7F', justify='center').pack(padx=20, pady=60)
    # ---- الجداول
    right = ttk.Frame(top); right.grid(row=0, column=1, sticky='nw')
    bar = ttk.Frame(right); bar.pack(fill='x')
    ttk.Label(bar, text='Month:', font=('Segoe UI', 10, 'bold')).pack(side='left')
    t = dt.date.today()
    months = sorted({f'{yy}-{mm:02d}' for yy, mm in [add_month(t.year, t.month, k) for k in range(-3, 4)] + [(y, m)]})
    v_month = tk.StringVar(value=f'{y}-{m:02d}')
    cb = ttk.Combobox(bar, textvariable=v_month, values=months, width=9, state='readonly'); cb.pack(side='left', padx=(6, 10))
    v_warn = tk.StringVar()
    wl = tk.Label(right, textvariable=v_warn, fg='#9C5700', bg='#FFF2CC', justify='left', anchor='w', wraplength=520, padx=6, pady=4)
    secs = {}
    for kind, head in ((COMPANY, 'Company update  –  periods from the company calendar'),
                       (OWNER, 'Owner update  –  every Tuesday (change a date when planning tells you)')):
        fr = ttk.LabelFrame(right, text=head, padding=6); fr.pack(fill='x', pady=(8, 0))
        tab = ttk.Frame(fr); tab.pack(fill='x')
        btn = ttk.Frame(fr); btn.pack(fill='x', pady=(4, 0))
        secs[kind] = dict(tab=tab, rows=[])
        ttk.Button(btn, text='+ Row', command=lambda k=kind: (add_row(k, None), redraw(k))).pack(side='left')
        ttk.Button(btn, text='− Last Row', command=lambda k=kind: (secs[k]['rows'].pop() if secs[k]['rows'] else None, redraw(k))).pack(side='left', padx=6)
    def add_row(kind, r):
        s = secs[kind]
        v0 = tk.StringVar(value=fd(r['d0']) if r else ''); v1 = tk.StringVar(value=fd(r['d1']) if r else '')
        vd = tk.StringVar(value='' if not r or r.get('days') is None else str(r['days']))
        s['rows'].append(dict(v0=v0, v1=v1, vd=vd, orig=r))
    def redraw(kind):
        s = secs[kind]
        for w in s['tab'].winfo_children(): w.destroy()
        cols = ['No.', 'From', 'Update Date', '', 'Working Days'] if kind == COMPANY else ['No.', 'Update Date', '']
        for c, h in enumerate(cols):
            ttk.Label(s['tab'], text=h, font=('Segoe UI', 9, 'bold'), foreground=NAVY).grid(row=0, column=c, padx=4, sticky='w')
        for i, x in enumerate(s['rows'], 1):
            ttk.Label(s['tab'], text=str(i)).grid(row=i, column=0, padx=4)
            ents = [x['v0'], x['v1']] if kind == COMPANY else [x['v1']]
            for c, v in enumerate(ents, 1):
                ttk.Entry(s['tab'], textvariable=v, width=12, justify='center').grid(row=i, column=c, padx=4, pady=1)
            dl = ttk.Label(s['tab'], width=4, foreground='#556'); dl.grid(row=i, column=len(ents) + 1, sticky='w')
            def upd(*_a, x=x, dl=dl):
                try: dl.configure(text=day_name(to_date(x['v1'].get())))
                except tk.TclError: pass
            if x.get('tid'):
                try: x['v1'].trace_remove('write', x['tid'])
                except tk.TclError: pass
            x['tid'] = x['v1'].trace_add('write', upd); upd()
            if kind == COMPANY:
                ttk.Entry(s['tab'], textvariable=x['vd'], width=6, justify='center').grid(row=i, column=4, padx=4)
    def load_month(*_a):
        yy, mm = (int(p) for p in v_month.get().split('-'))
        company, owner, warn = loader(yy, mm)
        for kind, lst in ((COMPANY, company), (OWNER, owner)):
            secs[kind]['rows'] = []
            for r in lst: add_row(kind, r)
            redraw(kind)
        v_warn.set('\n'.join(warn or []))
        if warn: wl.pack(fill='x', pady=(8, 0), after=bar)
        else: wl.pack_forget()
    cb.bind('<<ComboboxSelected>>', load_month)
    # ---- الزراير
    bb = ttk.Frame(right); bb.pack(fill='x', pady=(12, 0))
    def open_file():
        if not (original and os.path.exists(original)):
            messagebox.showinfo('Update Dates', 'The original file is not available.', parent=top); return
        try:
            if os.name == 'nt': os.startfile(original)
            else:
                import subprocess; subprocess.Popen(['xdg-open', original])
        except Exception as e:
            messagebox.showwarning('Update Dates', f'Could not open the file:\n{e}', parent=top)
    def collect():
        yy, mm = (int(p) for p in v_month.get().split('-'))
        out, bad = {COMPANY: [], OWNER: []}, []
        for kind in KINDS:
            prev = None
            items = []
            for i, x in enumerate(secs[kind]['rows'], 1):
                d1 = to_date(x['v1'].get())
                if not d1: bad.append(f'{kind} row {i}: update date (dd-mm-yyyy)'); continue
                d0 = to_date(x['v0'].get()) if kind == COMPANY else None
                if kind == COMPANY and not d0: bad.append(f'{kind} row {i}: from date (dd-mm-yyyy)'); continue
                if d0 and d0 > d1: bad.append(f'{kind} row {i}: from is after the update date'); continue
                o = x['orig'] or {}
                same = o.get('d1') == d1 and (kind == OWNER or o.get('d0') == d0)
                items.append(dict(d0=d0, d1=d1, vd=x['vd'].get().strip(), source=o.get('source') if same and o.get('source') else 'Manual',
                                  notes=o.get('notes', '')))
            items.sort(key=lambda r: r['d1'])
            for i, r in enumerate(items, 1):
                if kind == OWNER: r['d0'] = (prev + dt.timedelta(days=1)) if prev else r['d1'] - dt.timedelta(days=6)
                try: days = int(r['vd']) if r['vd'] else work_days(r['d0'], r['d1'])
                except ValueError: days = work_days(r['d0'], r['d1'])
                out[kind].append(dict(kind=kind, month=f'{yy}-{mm:02d}', no=i, d0=r['d0'], d1=r['d1'], days=days,
                                      source=r['source'], notes=r['notes']))
                prev = r['d1']
        return yy, mm, out, bad
    def save():
        yy, mm, out, bad = collect()
        if bad:
            messagebox.showwarning('Update Dates', 'Fix these first:\n\n' + '\n'.join(bad), parent=top); return
        odd = [fd(r['d1']) for k in KINDS for r in out[k] if (r['d1'].year, r['d1'].month) != (yy, mm)]
        if odd and not messagebox.askyesno('Update Dates', f'These dates are outside {v_month.get()}:\n' + ', '.join(odd) +
                                           '\n\nSave anyway?', parent=top): return
        res.update(y=yy, m=mm, company=out[COMPANY], owner=out[OWNER]); top.destroy()
    ttk.Button(bb, text='Open File', command=open_file).pack(side='left')
    ttk.Button(bb, text='Cancel', command=top.destroy).pack(side='right')
    ttk.Button(bb, text='Save Dates', command=save).pack(side='right', padx=6)
    load_month()
    if not shown and original and auto_open: top.after(300, open_file)     # من غير صورة هنا ← يفتح الأصل عشان تراجع
    top.grab_set(); top.focus_set()
    top.wait_window()
    return (res['y'], res['m'], res['company'], res['owner']) if res else None
