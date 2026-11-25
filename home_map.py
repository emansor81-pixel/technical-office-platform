# -*- coding: utf-8 -*-
"""الشاشة الرئيسية: البرنامج في النص، والأقسام كروت حواليه متوصلة بيه بخطوط رفيعة.

- على كل كارت نقط ملونة = السجلات اللي القسم بيشاركها، وأسامي السجلات مكتوبة تحت الخريطة.
- الماوس على قسم ← تفاصيله في المربع اللي على اليمين.  على اسم سجل ← الأقسام اللي بتستخدمه بس اللي تفضل منوّرة.
- كليك على قسم ← يفتحه.  كليك على اسم سجل ← يفتح ملفه (أو قايمة لو ليه ملفين).
- لو حصل أي خطأ هنا، main.py بيرجّع الكروت القديمة لوحده (والخطأ بيتسجل في platform_error.log).
"""
import os, sys, math, subprocess, traceback, zlib, struct, base64
import tkinter as tk
from tkinter import font as tkfont, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))

BG, CARD, LINE, INK, MUTED = '#F3F5F8', '#FFFFFF', '#D9E0E8', '#13293A', '#5D6878'
NAVY, NAVY2 = '#1F4E78', '#2E6496'
COL = {'dc': '#1F4E78', 'qs': '#0E6E66', 'wir': '#2C7536', 'ipc': '#4A4EA2', 'sub': '#8F461E', 'mat': '#66560F',
       'planned': '#94590A', 'wlog': '#B4234A', 'arch': '#56687A', 'upd': '#7A3D9A'}
ICON = {'doc': '✉', 'qs': '∑', 'wir': '☑', 'ipc': '✍', 'sub': '⚑', 'mar': '⚒'}

# (key زي main.MODULES، الاسم الكبير، السطر الصغير، اللون) – بالترتيب مع عقارب الساعة من فوق
# مركز الورق (DOC) بيفتح خريطته هو (DocMap تحت) – والسجلات (Letters In / Out) بقت جواه
SECTIONS = [('doc', 'DOC', 'Document Center', 'dc'), ('mar', 'Store', 'Stock · Deliveries · Issues', 'mat'),
            ('sub', 'SUB', 'Subcontractors', 'sub'), ('ipc', 'IPC', 'Payment Certificates', 'ipc'),
            ('wir', 'WIR', 'Inspection Requests', 'wir'), ('qs', 'QS', 'Quantities', 'qs')]

# السجلات المشتركة: مين بيستخدمها، والكليك بيفتح إيه (file / folder جوه فولدر المشروع، section، dcsheet)
REGISTERS = [
    dict(id='wlog', name='WIR Log', uses=['doc', 'sub', 'wir', 'qs'], col='wlog',
         opens=[('01-LOG WIR.xlsm', 'file', r'00_Registers\01-LOG WIR.xlsm')]),
    dict(id='qscalc', name='QS Quantities', uses=['qs', 'wir', 'ipc', 'sub'], col='qs',
         opens=[('the QS section', 'section', 'qs')]),
    dict(id='arch', name='Archive Folders', uses=['doc', 'mar', 'sub', 'ipc'], col='arch',
         opens=[('the project folder', 'folder', '')]),
    dict(id='ipclog', name='Client IPCs', uses=['doc', 'ipc', 'sub'], col='ipc',
         opens=[('11-LOG IPC.xlsx', 'file', r'00_Registers\11-LOG IPC.xlsx'),
                ('the CLIENT_IPC folder', 'folder', r'08_Cost_Payments\CLIENT_IPC')]),
    dict(id='upd', name='Update Dates', uses=['qs', 'doc'], col='upd',
         opens=[('12-LOG UPDATE DATES.xlsx', 'file', r'00_Registers\12-LOG UPDATE DATES.xlsx')]),
    dict(id='marstock', name='Materials & Stock', uses=['doc', 'mar'], col='mat',
         opens=[('02-LOG MAR+MIR.xlsm', 'file', r'00_Registers\02-LOG MAR+MIR.xlsm'),
                ('04-LOG STOCK.xlsx', 'file', r'00_Registers\04-LOG STOCK.xlsx')]),
    dict(id='dcx', name='DC.xlsx', uses=['doc'], col='planned',
         opens=[('DC.xlsx on OneDrive', 'dcsheet', '')]),
]

INFO = {
    'core': dict(text='The program is in the middle and the sections sit around it. The coloured dots on a card are '
                      'the registers that section shares – their names are under the map.',
                 hints=['Click the circle: show or hide the sections',
                        'Point at a section: what it takes and what it gives',
                        'Point at a register name: the sections that share it stay lit',
                        'Click a section: open it', 'Click a register name: open its file']),
    'doc': dict(text='Every paper of the project in one place: papers in (Distribute), papers out (New Document), '
                     'the registers, and what is still missing. Click to open its map.',
               takes=['00_INBOX – site, the Consultant & CLIENT, the company, Unifier, subcontractors',
                      'Claude chat – hard papers and scans with many papers',
                      'WIR Log – the name of every site WIR (S NAM)', 'MAR+MIR Log – the MIR name'],
               gives=['Archive folders', 'Update Dates – the company calendar',
                      'IPC Log – Certified when the VN invoice or the signed IPC arrives',
                      'Materials & Stock – delivery note items and the MIR', 'DC.xlsx – letters in / out']),
    'mar': dict(text='The store: the balance of every material, every delivery note with its items, and the issues to '
                     'site – with alerts for negative balance and deliveries that have no MIR yet. MAR and MIR are in '
                     'Document Center › Documents.',
                takes=['04-LOG STOCK – stock, deliveries, issues', 'Delivery notes – from Distribute'],
                gives=['Issue to Site – the issues sheet', 'Make MIR for a delivery that has none']),
    'sub': dict(text='Every subcontractor with his items and work stations, his certificate (V, Q and the follow-up '
                     'statement) with additions, deductions and retention, and his deduction memos.',
                takes=[r'00_Registers\10.1-SUBCONTRACTORS.xlsx – contracts, items, certificates',
                       r'13_PARTIES – one folder per subcontractor / supplier / plant / lab', 'WIR Log – the lines',
                       'QS – quantities within his stations', 'The client IPC with the same WIRs'],
                gives=['His folder (English name) – CERTIFICATES and DEDUCTION_MEMOS',
                       '10-LOG SUB – sheet DEDUCTION_MEMOS (DM-001 …)'],
                soon=['Addition memos', 'Final account and retention', 'Water contract items']),
    'ipc': dict(text='The whole IPC in one click: cover, certificate and full take-off, PDF from Excel, the approved '
                     'WIRs behind the take-off, and a zip for every WIR type.',
                takes=['QS – quantities up to the period end', 'IPC Log and the template',
                       'Approved WIRs from the archive'],
                gives=[r'08_Cost_Payments\CLIENT_IPC – IPC-NN', 'IPC Log',
                       'OUR_SHEET and consultant sheets – every hour', 'SUB – the client IPC with his work']),
    'wir': dict(text='The quantity of every request, ready to type into Unifier. Filters: uploaded & waiting, not '
                     'uploaded, approved, no Unifier No., missing levels.',
                takes=['WIR Log', 'QS quantities'],
                gives=['Unifier – you type the quantity', 'Uploaded mark', 'Export Excel']),
    'qs': dict(text='Quantities of every item (excavation, bedding, pipes, backfill, base, asphalt) from the WIRs up '
                    'to any date. The weekly and monthly update, and the charts.',
               takes=['WIR Log – delivery date (column M)', 'Update Dates'],
               gives=['Weekly update for the company and the owner (Excel + PDF)',
                      'QS_HISTORY – saved updates and charts', 'Quantities for WIR, IPC and SUB'],
               soon=['Approved 3-month plan – planned vs actual', 'Site productivity in the weekly update']),
    'wlog': dict(path=r'00_Registers\01-LOG WIR.xlsm',
                 text='The heart of the program – you fill it. Sheet SW: every request with its district, street and '
                      'line, from–to, request date, delivery date (M), Unifier No. (N), code and full name (S NAM).'),
    'qscalc': dict(text='Not a file: the quantities QS works out from the WIRs. WIR, IPC and SUB take them directly.'),
    'arch': dict(path='02_Correspondence · 04_Submittals · 05_Quality · 08_Cost_Payments · 09_Procurement_Materials · 13_PARTIES',
                 text='Where DC files every paper type: 1_WIP prepared, 2_SHARED submitted, 3_PUBLISHED approved, '
                      '4_ARCHIVED.'),
    'ipclog': dict(path=r'00_Registers\11-LOG IPC.xlsx  ·  08_Cost_Payments\CLIENT_IPC',
                   text='Every IPC with its number, period and status, the fixed template, and the client IPC files.'),
    'upd': dict(path=r'00_Registers\12-LOG UPDATE DATES.xlsx',
                text='One row for every update date: the company from its monthly calendar, the owner every Tuesday.'),
    'marstock': dict(path=r'00_Registers\02-LOG MAR+MIR.xlsm  ·  04-LOG STOCK.xlsx',
                     text='MAR LOG and MIR LOG with the MIR form, and the stock, deliveries and issues sheets. The '
                          'balances are Excel formulas – the program only writes rows.'),
    'dcx': dict(path=r'OneDrive\01-N.W.C PROJECT\CO WATER\DC.xlsx',
                text='Letters out to the the Consultant, CLIENT, suppliers and others, and letters in. DC writes to it when you '
                     'file a letter.'),
}


_OPENED = False      # اتفتحت مرة في التشغيلة دي؟ ← لما ترجع للرئيسية الكروت تفضل ظاهرة


def _family(widget, names):
    fams = set(tkfont.families(widget))
    for f in names:
        if f in fams: return f
    return tkfont.nametofont('TkDefaultFont').actual()['family']


def _blend(c, t, to=CARD):
    """اللون c ناحية to بنسبة t (0 = زي ما هو، 1 = to)."""
    a = [int(c[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(to[i:i + 2], 16) for i in (1, 3, 5)]
    return '#%02X%02X%02X' % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


# ---------------------------------------------------------------- رسم ناعم
# الكانفاس في ويندوز بيرسم الخطوط المايلة والدواير مسنّنة. فالخطوط والنقط والدايرة اللي في النص
# بتترسم هنا صورة PNG شفافة بحروف ناعمة (من غير أي مكتبة زيادة)، والكانفاس بيحطها زي ما هي.
def _rgb(c):
    return tuple(int(c[i:i + 2], 16) for i in (1, 3, 5))


def _photo(master, w, h, buf):
    raw, stride = bytearray(), w * 4
    for y in range(h):
        raw.append(0); raw += buf[y * stride:(y + 1) * stride]

    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xFFFFFFFF)
    png = (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)) +
           chunk(b'IDAT', zlib.compress(bytes(raw), 6)) + chunk(b'IEND', b''))
    return tk.PhotoImage(master=master, data=base64.b64encode(png).decode('ascii'))


def _aa_disc(master, diam, color, ring=None, ring_w=0.0):
    """دايرة مليانة ناعمة (ولو ring: حواليها حلقة باللون ده) ← (الصورة، المقاس)."""
    R = diam / 2.0
    Ro = R + (ring_w if ring else 0.0)
    size = int(math.ceil(2 * Ro + 2))
    c = size / 2.0
    ci, co = _rgb(color), (_rgb(ring) if ring else None)
    solid = bytes(ci + (255,))
    buf = bytearray(size * size * 4)
    for row in range(size):
        cy = row + 0.5 - c
        if abs(cy) > Ro + 1: continue
        off = row * size * 4
        inner = math.sqrt((R - 1) ** 2 - cy * cy) if abs(cy) < R - 1 else 0.0
        i0, i1 = (int(math.ceil(c - inner + 0.5)), int(math.floor(c + inner - 0.5))) if inner > 1 else (0, -1)
        if i1 >= i0:                              # الجزء اللي جوه خالص: لون واحد
            buf[off + i0 * 4:off + (i1 + 1) * 4] = solid * (i1 - i0 + 1)
        span = math.sqrt(max(0.0, (Ro + 1) ** 2 - cy * cy))
        for col in range(max(0, int(c - span) - 1), min(size, int(c + span) + 2)):
            if i0 <= col <= i1: continue
            cx = col + 0.5 - c
            d = math.sqrt(cx * cx + cy * cy)
            ai = min(1.0, max(0.0, R + 0.5 - d))
            if co:
                ao = min(1.0, max(0.0, Ro + 0.5 - d))
                a = ai + ao * (1 - ai)
                if a <= 0: continue
                rgb = tuple(int((ci[k] * ai + co[k] * ao * (1 - ai)) / a + 0.5) for k in range(3))
            else:
                a, rgb = ai, ci
                if a <= 0: continue
            buf[off + col * 4:off + col * 4 + 4] = bytes(rgb + (int(a * 255 + 0.5),))
    return _photo(master, size, size, buf), size


class HomeMap(tk.Frame):
    CORE, LANE = 60.0, 30.0                   # الدايرة اللي في النص، وبُعد الخطوط اللي بتطلع جنب بعض من فوق ومن تحت
    CW, CH, MARGIN = 272.0, 110.0, 24.0       # مقاس الكارت (بوحدات الرسم)
    # مكان نص كل كارت (النص = 0,0): نفس الشكل البيضاوي، متظبط عشان كل خط يبقى أفقي أو رأسي
    POS = {'doc': (-330, -160), 'mar': (330, -160), 'sub': (410, 30), 'ipc': (260, 240),
           'wir': (-260, 240), 'qs': (-410, 30)}
    SECTIONS, REGISTERS, INFO, ICONS = SECTIONS, REGISTERS, INFO, ICON
    CORE_TXT, TITLE, INBOX_CARD = ('Project', 'Platform'), 'Project Platform', 'doc'

    DRILL = {'doc'}                            # كروت ليها خريطة جوّاها: الدوسة بتعمل حركة «الدايرة تنتقل»

    def __init__(self, parent, open_section, ready, root, name='', number='', inbox_count=None, next_update=None,
                 on_fail=None, log=None, intro=False):
        super().__init__(parent, bg=BG)
        self._job = self._pending = self._fill_job = None
        self._intro = intro                    # جاي من خريطة تانية ← الكروت تطلع من الدايرة بحركة
        try:
            self._build_home(open_section, ready, root, name, number, inbox_count, next_update, on_fail, log)
        except Exception:
            try: self.destroy()
            except Exception: pass
            raise

    def _build_home(self, open_section, ready, root, name, number, inbox_count, next_update, on_fail, log):
        self.open_section, self.ready, self.root = open_section, set(ready or ()), root or ''
        self.on_fail, self.log, self.next_update = on_fail, log, next_update
        self.hot, self.shown, self._s, self._failed = None, None, 1.0, False   # hot = اللي متعلّم، shown = اللي تفاصيله ظاهرة
        self.expanded, self._anim, self._anim_job = self._start_open(), False, None          # الأقسام بتطلع لما تدوس على الدايرة
        self.items, self.cardbox, self.dots, self.legend_text, self.fonts = [], {}, {}, {}, []
        self._imgs, self._discs = [], {}                    # الصور الناعمة لازم تفضل محفوظة وإلا تختفي
        self._menu, self._last_click, self.next_short = None, -10 ** 9, ''
        self.fam = _family(self, ('Segoe UI', 'Segoe UI Variable Text', 'Calibri', 'Arial', 'DejaVu Sans'))
        self.mono = _family(self, ('Consolas', 'Cascadia Mono', 'Courier New', 'DejaVu Sans Mono'))
        self.sym = _family(self, ('Segoe UI Symbol', 'Segoe UI', 'DejaVu Sans'))
        try: self.inbox_n = inbox_count() if inbox_count else None
        except Exception: self.inbox_n = None
        self._geometry()
        self.side_w = 420                                   # العمود اللي على اليمين (بيضيق لو الشباك صغير)
        side = tk.Frame(self, bg=BG, width=self.side_w)
        side.pack(side='right', fill='y', padx=(0, 22), pady=16); side.grid_propagate(False)
        self.side = side
        self.cv = tk.Canvas(self, bg=CARD, highlightthickness=1, highlightbackground=LINE, bd=0, closeenough=2.0)
        self.cv.pack(side='left', fill='both', expand=True, padx=(22, 14), pady=16)
        self._side(side, name, number)
        self.cv.bind('<Configure>', self._on_resize)
        self.cv.bind('<Motion>', self._on_motion)
        self.cv.bind('<Button-1>', self._on_click)
        self.bind('<Configure>', self._on_frame_resize, add='+')
        self._details(None)

    def _start_open(self):
        return _OPENED and not self._intro

    # ------------------------------------------------------------------ الحسابات (بوحدات الرسم، النص = 0,0)
    def _geometry(self):
        self.sec = {}
        for i, (k, t, s, c) in enumerate(self.SECTIONS):
            x, y = self.POS[k]
            box = (x - self.CW / 2, y - self.CH / 2, x + self.CW / 2, y + self.CH / 2)
            self.sec[k] = dict(key=k, t=t, s=s, c=c, i=i, x=x, y=y, box=box, planned=(c == 'planned'))
        self.regs = {r['id']: dict(r) for r in self.REGISTERS}
        # لكل كارت خطه لوحده، أفقي ورأسي بس (من غير ما خطين يتدمجوا). الحتة اللي جوه الدايرة بتستخبى وراها
        g, L = 6, self.LANE
        for k, d in self.sec.items():
            x, y, (x1, y1, x2, y2) = d['x'], d['y'], d['box']
            if x1 < 0 < x2:                                   # فوق النص على طول: خط رأسي
                d['segs'] = [(0, 0, 0, (y2 + g) if y < 0 else (y1 - g))]
            elif abs(y) < self.CORE - 15:                     # في مستوى النص: خط أفقي
                d['segs'] = [(0, y, (x1 - g) if x > 0 else (x2 + g), y)]
            else:                                             # مايل: رأسي جنب النص وبعدين أفقي لجنب الكارت
                lx = L if x > 0 else -L
                d['segs'] = [(lx, 0, lx, y), (lx, y, (x1 - g) if x > 0 else (x2 + g), y)]
        xs = [v for d in self.sec.values() for v in (d['box'][0], d['box'][2])]
        ys = [v for d in self.sec.values() for v in (d['box'][1], d['box'][3])]
        self.legend_y = max(ys) + 36                      # سطر الشرح، وتحته أسامي السجلات
        m = self.MARGIN
        self.bbox = (min(xs) - m, min(ys) - m, max(xs) + m, self.legend_y + 32 + 14 + m)

    # ------------------------------------------------------------------ العمود اللي على اليمين
    def _side(self, side, name, number):
        side.columnconfigure(0, weight=1); side.rowconfigure(3, weight=1)
        tk.Label(side, text='Technical Office Platform', bg=BG, fg=MUTED, font=(self.fam, 10, 'bold'),
                 anchor='w').grid(row=0, column=0, sticky='ew')
        t1 = tk.Label(side, text=(name or 'Project').strip() or 'Project', bg=BG, fg=NAVY, font=(self.fam, 16, 'bold'),
                      anchor='w', justify='left', wraplength=self.side_w - 10)
        t1.grid(row=1, column=0, sticky='ew')
        ok = bool(self.root) and os.path.isdir(self.root)
        sub = (f'PO {number}   ·   ' if number else '') + \
              ('Project folder ✓' if ok else 'Project folder not found – edit PROJECT_ROOT in SETTINGS.json')
        t2 = tk.Label(side, text=sub, bg=BG, fg='#375623' if ok else '#C00000', font=(self.fam, 10), anchor='w',
                      justify='left', wraplength=self.side_w - 10)
        t2.grid(row=2, column=0, sticky='ew', pady=(2, 12))
        self._wraps = (t1, t2)
        card = tk.Frame(side, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        card.grid(row=3, column=0, sticky='nsew')
        card.pack_propagate(False)
        self.box = tk.Frame(card, bg=CARD, padx=16, pady=14); self.box.pack(fill='both', expand=True)
        # الورق اللي مستني في INBOX – دوس عليه يفتح DC
        n = self.inbox_n
        bar = tk.Frame(side, bg=NAVY, padx=18, pady=11); bar.grid(row=4, column=0, sticky='ew', pady=(10, 0))
        txt = ('All good – nothing waiting' if n == 0 else f'{n} paper(s) waiting in INBOX') if n is not None else 'Ready'
        l1 = tk.Label(bar, text=txt, bg=NAVY, fg='white', font=(self.fam, 12, 'bold'), anchor='w')
        l1.pack(side='left')
        l1b = tk.Label(bar, text='✔' if not n else '●', bg=NAVY, fg='#7CD99B' if not n else '#FFC000',
                       font=(self.fam, 15, 'bold'))
        l1b.pack(side='right')
        if n:
            for w in (bar, l1, l1b):
                w.configure(cursor='hand2'); w.bind('<Button-1>', lambda ev: self.open_section('dc'))   # التوزيع على طول
        # موعد التحديث الجاي – دوس عليه يفتح QS
        self.v_next = tk.StringVar(value='')
        b2 = tk.Frame(side, bg=NAVY2, padx=18, pady=9, cursor='hand2')
        l2 = tk.Label(b2, textvariable=self.v_next, bg=NAVY2, fg='white', font=(self.fam, 10, 'bold'),
                      cursor='hand2', anchor='w', justify='left', wraplength=self.side_w - 45)
        l2.pack(fill='x'); self.l2, self.b2 = l2, b2
        for w in (b2, l2): w.bind('<Button-1>', lambda ev: self.open_section('qs'))
        self._fill_job = self.after(300, self._fill_next)

    def _fill_next(self):
        """سطر التحديث الجاي (للشريط اللي تحت) + أقرب تاريخ (للشارة اللي على كارت QS) – من قراية واحدة للسجل."""
        self._fill_job = None
        txt, short = '', ''
        try:
            p = os.path.join(HERE, 'modules', 'qs')
            if p not in sys.path: sys.path.insert(0, p)
            import upd_dates as U
            rows = U.load()
            txt = U.due_text(rows=rows)
            ds = [r['d1'] for r in (U.next_due(U.COMPANY, rows), U.next_due(U.OWNER, rows)) if r]
            if ds:
                d = min(ds)
                short = f'Next: {U.day_name(d)} {d:%d-%m}'
        except Exception:
            try: txt = self.next_update() if self.next_update else ''
            except Exception: txt = ''
        try:
            if txt:
                self.v_next.set('Next update  ·  ' + txt)
                self.b2.grid(row=5, column=0, sticky='ew', pady=(8, 0))
            if short and short != self.next_short:
                self.next_short = short
                self._on_resize()
        except tk.TclError:
            pass

    def _on_frame_resize(self, ev):
        if ev.widget is not self: return
        w = 420 if ev.width >= 1280 else 340
        if w == self.side_w: return
        self.side_w = w
        self.side.configure(width=w)
        for lab in self._wraps: lab.configure(wraplength=w - 10)
        self.l2.configure(wraplength=w - 45)
        self._details(self.shown)

    def _details(self, obj):
        box = self.box
        for w in box.winfo_children(): w.destroy()
        W = self.side_w - 40

        def lab(text, font=None, fg=INK, pady=(0, 0), wrap=W):
            w = tk.Label(box, text=text, bg=CARD, fg=fg, font=font or (self.fam, 10), anchor='w', justify='left',
                         wraplength=wrap)
            w.pack(fill='x', pady=pady)
            return w

        def rows(head, items, fg=INK):
            if not items: return
            lab(head.upper(), (self.fam, 8, 'bold'), MUTED, (9, 2))
            for t in items: lab('•  ' + t, (self.fam, 10), fg, (1, 0), W - 10)

        if obj in self.sec:
            d, I = self.sec[obj], self.INFO.get(obj, {})
            planned = d['planned'] or obj not in self.ready
            lab(f"{d['t']}  –  {d['s']}", (self.fam, 15, 'bold'), COL[d['c']])
            lab('SOON  ·  not built yet' if planned else 'SECTION  ·  click to open', (self.fam, 8, 'bold'),
                COL['planned'] if planned else MUTED, (2, 8))
            lab(I.get('text', ''))
            rows('Takes from', I.get('takes'))
            rows('Gives to', I.get('gives'))
            rows('Soon', I.get('soon'), COL['planned'])
        elif obj in self.regs:
            r, I = self.regs[obj], self.INFO.get(obj, {})
            lab(r['name'], (self.fam, 15, 'bold'), COL[r['col']])
            many = len(r['opens']) > 1
            lab('SHARED REGISTER  ·  ' + ('click to choose what to open' if many else 'click to open'),
                (self.fam, 8, 'bold'), MUTED, (2, 6))
            if I.get('path'): lab(I['path'], (self.mono, 9), MUTED, (0, 8))
            lab(I.get('text', ''))
            rows('Used by', ['  ·  '.join(self.sec[m]['t'] for m in r['uses'])])
            rows('Opens', [o[0] for o in r['opens']])
        else:
            I = self.INFO['core']
            lab(self.TITLE, (self.fam, 15, 'bold'), INK)
            lab('HOW TO READ THE MAP', (self.fam, 8, 'bold'), MUTED, (2, 8))
            lab(I['text'])
            rows('Use it', I['hints'])
            self._key(box)

    def _key(self, box):
        tk.Label(box, text='KEY', bg=CARD, fg=MUTED, font=(self.fam, 8, 'bold'), anchor='w').pack(fill='x', pady=(9, 4))
        f = tkfont.Font(self, family=self.fam, size=10)
        fb = tkfont.Font(self, family=self.fam, size=8, weight='bold')
        lh = max(26, f.metrics('linespace') + 10)
        c = tk.Canvas(box, bg=CARD, highlightthickness=0, height=lh * 3, width=self.side_w - 40)
        c.pack(fill='x', anchor='w')
        y = lh / 2
        self._rrect(4, y - 10, 44, y + 10, 5, fill=CARD, outline=LINE, width=1, canvas=c)
        c.create_text(56, y, text='A section – click to open it', font=f, fill=INK, anchor='w')
        y += lh
        for i, col in enumerate((COL['wlog'], COL['qs'])):
            img, _ = _aa_disc(self, 11, col)
            c.create_image(13.5 + i * 16, y, image=img)
            c._imgs = getattr(c, '_imgs', []) + [img]     # بتتمسح مع الكانفاس الصغير
        c.create_text(56, y, text='The registers it shares', font=f, fill=INK, anchor='w')
        y += lh
        tint = _blend(COL['planned'], .88)
        self._rrect(4, y - 10, 46, y + 10, 10, fill=tint, outline=tint, canvas=c)
        c.create_text(25, y, text='Soon', font=fb, fill=COL['planned'])
        c.create_text(56, y, text='Not built yet', font=f, fill=INK, anchor='w')
        c._fonts = (f, fb)

    # ------------------------------------------------------------------ الرسم
    def _on_resize(self, ev=None):
        if self._job: self.after_cancel(self._job)
        self._job = self.after(40, self._redraw)

    def _redraw(self):
        self._job = None
        if self._failed: return
        if self._anim_job:
            try: self.after_cancel(self._anim_job)
            except Exception: pass
            if self._anim and getattr(self, '_anim_target', None) is not None:
                self.expanded = self._anim_target           # الشباك اتغيّر وسط الحركة ← نكمّلها على آخرها
            self._anim_job, self._anim = None, False
        try:
            self._draw()
            self._apply(self.hot)
            if self._intro and not self.expanded:       # أول رسمة بعد الانتقال ← الكروت تطلع من الدايرة
                self._anim_job = self.after(60, self._intro_open)
        except Exception:
            self._failed = True
            if self.log:
                try: self.log(traceback.format_exc())
                except Exception: pass
            if self.on_fail: self.on_fail(self)

    def _font(self, px, s, bold=False, mono=False, sym=False):
        fam = self.mono if mono else (self.sym if sym else self.fam)
        f = tkfont.Font(self, family=fam, size=-max(8, int(round(px * s))), weight='bold' if bold else 'normal')
        self.fonts.append(f)
        return f

    def _rrect(self, x1, y1, x2, y2, r, canvas=None, **kw):
        r = min(r, (x2 - x1) / 2, (y2 - y1) / 2)
        p = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2, x1, y2,
             x1, y2 - r, x1, y1 + r, x1, y1]
        return (canvas or self.cv).create_polygon(p, smooth=True, splinesteps=16, **kw)

    def _add(self, item, group, normal, faded=None):
        self.items.append((item, group, normal, faded))

    def _disc(self, diam, color, ring=None):
        """صورة نقطة ناعمة (محفوظة عشان تتعاد من غير ما تترسم تاني)."""
        key = (int(round(diam)), color, ring)
        if key not in self._discs:
            self._discs[key] = _aa_disc(self, key[0], color, ring, 2.0 if ring else 0.0)
        return self._discs[key][0]

    def _dot(self, cx, cy, diam, color, tags=(), group=None, ring=False):
        """نقطة ناعمة في (cx, cy) بإحداثيات الكانفاس. لو group: بتبهت مع الكارت."""
        img = self._disc(diam, color)
        it = self.cv.create_image(cx, cy, image=img, tags=tags)
        if group:
            self._add(it, group, {'image': img}, {'image': self._disc(diam, _blend(color, .8))})
        if ring:
            return it, self._disc(diam, color, INK)
        return it

    def _badge(self, k):
        """الشارة اللي على يمين الكارت: (الكلام، لون الكلام، الخلفية، نقطة)."""
        if k == self.INBOX_CARD and self.inbox_n: return (f'{self.inbox_n} waiting', '#6B4A00', '#FFF4D6', '#E8A400')
        if k == 'qs' and self.next_short: return (self.next_short, COL['qs'], _blend(COL['qs'], .88), None)
        if k not in self.ready: return ('Soon', COL['planned'], _blend(COL['planned'], .88), None)
        return None

    def _draw(self):
        cv = self.cv
        cv.delete('all'); self.items, self.cardbox, self.dots, self.legend_text, self.fonts = [], {}, {}, {}, []
        self._imgs, self._discs = [], {}
        W, H = cv.winfo_width(), cv.winfo_height()
        if W < 80 or H < 80: return
        x0, y0, x1, y1 = self.bbox
        s = max(0.5, min(W / (x1 - x0), H / (y1 - y0), 1.35))
        self._s = s
        ox, oy = W / 2 - (x0 + x1) / 2 * s, H / 2 - (y0 + y1) / 2 * s
        X = lambda x: ox + x * s
        Y = lambda y: oy + y * s
        F = dict(name=self._font(22, s, True), sub=self._font(14.5, s), badge=self._font(12.5, s, True),
                 small=self._font(12.5, s), icon=self._font(21, s, sym=True))

        # كل كارت وخطه لوحده: أفقي ورأسي بس (حاد من غير تسنين)
        lw = max(2, int(round(3 * s)))
        for k, d in self.sec.items():
            c = '#C9A774' if d['planned'] else COL[d['c']]
            for a, b, e, f in d['segs']:
                it = cv.create_line(round(X(a)), round(Y(b)), round(X(e)), round(Y(f)), fill=c, width=lw,
                                    capstyle='projecting', tags=('sec', 'line'))
                self._add(it, 'card:' + k, {'fill': c})

        for k, d in self.sec.items():
            self._card(k, d, X, Y, s, F)
        self._legend(X, Y, s)

        # الدايرة اللي في النص (صغيرة): دوس عليها تطلّع الأقسام أو ترجّعها
        f_core = self._font(15, s, True)
        img, _ = _aa_disc(self, 2 * self.CORE * s, INK)
        self._imgs.append(img)
        tg = ('obj:core', 'core')
        cv.create_image(X(0), Y(0), image=img, tags=tg)
        cv.create_text(X(0), Y(-9), text=self.CORE_TXT[0], font=f_core, fill='#FFFFFF', tags=tg)
        cv.create_text(X(0), Y(10), text=self.CORE_TXT[1], font=f_core, fill='#FFFFFF', tags=tg)
        cv.create_text(X(0), Y(self.CORE + 26), text='Click to show the sections', font=self._font(14, s),
                       fill=MUTED, tags=('hint',))
        self._show(self.expanded)

    def _intro_open(self):
        self._anim_job = None
        self._intro = False
        if not self.expanded and not self._anim: self._toggle()

    def fly_to(self, key, done):
        """«الدايرة تنتقل»: الكروت التانية ترجع ورا الدايرة، والكارت ده يروح مكانها – وبعدين done() تفتح خريطته."""
        if self._anim or key not in self.sec: return done()
        cv, s = self.cv, self._s
        self._focus(None)
        cv.itemconfigure('line', state='hidden'); cv.itemconfigure('legend', state='hidden')
        cv.tag_raise('core'); cv.tag_raise('mv:' + key)
        vec = {k: (-d['x'] * s, -d['y'] * s) for k, d in self.sec.items()}
        self._anim, N, self._anim_target = True, 14, None

        def step(i=1):
            self._fly_job = None
            a, b = 1 - (1 - (i - 1) / N) ** 3, 1 - (1 - i / N) ** 3
            for k, (vx, vy) in vec.items(): cv.move('mv:' + k, (b - a) * vx, (b - a) * vy)
            if i < N:
                self._fly_job = self.after(16, step, i + 1); return
            self._anim = False
            if self.winfo_ismapped(): done()             # لو روحت لقسم تاني وسط الحركة ← متشدّكش
        step()

    def _show(self, on):
        """الأقسام وخطوطها وأسامي السجلات: ظاهرة ولا مستخبية (والجملة اللي تحت الدايرة بالعكس)."""
        self.cv.itemconfigure('sec', state='normal' if on else 'hidden')
        self.cv.itemconfigure('hint', state='hidden' if on else 'normal')

    def _toggle(self):
        """دوسة على الدايرة: الأقسام تطلع منها (أو ترجعلها) بحركة."""
        if self._anim: return
        cv, s, opening = self.cv, self._s, not self.expanded
        vec = {k: (-d['x'] * s, -d['y'] * s) for k, d in self.sec.items()}   # من مكان الكارت لحد النص
        if opening:
            self._focus(None)
            for k, (vx, vy) in vec.items(): cv.move('mv:' + k, vx, vy)
            cv.itemconfigure('card', state='normal'); cv.itemconfigure('hint', state='hidden')
        cv.itemconfigure('line', state='hidden'); cv.itemconfigure('legend', state='hidden')
        cv.tag_raise('core')                         # الكروت بتطلع من ورا الدايرة
        self._anim, N, st = True, 12, {'p': 0.0 if opening else 1.0}
        self._anim_target = opening

        def step(i=1):
            self._anim_job = None
            t = 1 - (1 - i / N) ** 3
            p = t if opening else 1 - t
            dp = st['p'] - p
            for k, (vx, vy) in vec.items(): cv.move('mv:' + k, dp * vx, dp * vy)
            st['p'] = p
            if i < N:
                self._anim_job = self.after(16, step, i + 1); return
            self._anim, self.expanded = False, opening
            if opening:
                global _OPENED
                _OPENED = True
            if not opening:                          # رجعت للنص: تستخبى وترجع مكانها الأصلي
                for k, (vx, vy) in vec.items(): cv.move('mv:' + k, -vx, -vy)
            self._show(opening)
        step()

    def _card(self, k, d, X, Y, s, F):
        cv, c = self.cv, COL[d['c']]
        x1, y1, x2, y2 = d['box']
        tg, g = ('obj:' + k, 'mv:' + k, 'card', 'sec'), 'card:' + k
        border = '#E8DCCB' if d['planned'] else LINE
        bw = max(1, 1.2 * s)
        bg = self._rrect(X(x1), Y(y1), X(x2), Y(y2), 14 * s, fill=CARD, outline=border, width=bw, tags=tg)
        self._add(bg, g, {'fill': CARD, 'outline': border, 'width': bw})
        self.cardbox[k] = bg
        tint = _blend(c, .86)
        it = self._rrect(X(x1 + 16), Y(y1 + 16), X(x1 + 56), Y(y1 + 56), 10 * s, fill=tint, outline=tint, tags=tg)
        self._add(it, g, {'fill': tint, 'outline': tint})
        it = cv.create_text(X(x1 + 36), Y(y1 + 36), text=self.ICONS.get(k, '▣'), font=F['icon'], fill=c, tags=tg)
        self._add(it, g, {'fill': c})
        it = cv.create_text(X(x1 + 68), Y(y1 + 29), text=d['t'], font=F['name'], fill=c, anchor='w', tags=tg)
        self._add(it, g, {'fill': c})
        it = cv.create_text(X(x1 + 68), Y(y1 + 51), text=d['s'], font=F['sub'], fill=MUTED, anchor='w', tags=tg)
        self._add(it, g, {'fill': MUTED})
        badge = self._badge(k)
        if badge:
            text, fg, bgc, dot = badge
            w = F['badge'].measure(text) / s + 18 + (14 if dot else 0)
            bx2 = x2 - 14; bx1 = bx2 - w; by = y1 + 28
            it = self._rrect(X(bx1), Y(by - 12), X(bx2), Y(by + 12), 12 * s, fill=bgc, outline=bgc, tags=tg)
            self._add(it, g, {'fill': bgc, 'outline': bgc})
            tx = bx1 + 9
            if dot:
                self._dot(X(tx + 4), Y(by), 8 * s, dot, tg, g); tx += 14
            it = cv.create_text(X(tx), Y(by), text=text, font=F['badge'], fill=fg, anchor='w', tags=tg)
            self._add(it, g, {'fill': fg})
        # النقط = السجلات اللي القسم بيشاركها (الماوس على النقطة = اسم السجل)
        it = cv.create_text(X(x1 + 16), Y(y1 + 88), text='Shares', font=F['small'], fill=MUTED, anchor='w', tags=tg)
        self._add(it, g, {'fill': MUTED})
        dx = x1 + 16 + F['small'].measure('Shares') / s + 10
        for rid, r in self.regs.items():
            if k not in r['uses']: continue
            o, ring = self._dot(X(dx + 5.5), Y(y1 + 88), 11 * s, COL[r['col']], ('obj:' + rid, 'mv:' + k, 'card', 'sec'),
                                g, ring=True)
            self.dots[(k, rid)] = (o, ring)
            dx += 17

    def _legend(self, X, Y, s):
        cv = self.cv
        f_leg, f_name = self._font(14, s), self._font(15, s, True)
        y = self.legend_y
        cv.create_text(X(0), Y(y), font=f_leg, fill=MUTED, tags=('legend', 'sec'),
                       text='The coloured dots on a card are the registers that section shares.  Click a name to open its file.')
        items = [(rid, r, f_name.measure(r['name']) / s) for rid, r in self.regs.items()]
        gap, dot = 26, 11
        total = sum(dot + 8 + w for _, _, w in items) + gap * (len(items) - 1)
        x, yy = -total / 2, y + 32
        for rid, r, w in items:
            c, tg = COL[r['col']], ('obj:' + rid, 'legend', 'sec')
            cv.create_rectangle(X(x - 4), Y(yy - 13), X(x + dot + 8 + w + 4), Y(yy + 13), fill=CARD, outline='',
                                tags=tg)
            self._dot(X(x + dot / 2), Y(yy), dot * s, c, tg)
            t = cv.create_text(X(x + dot + 8), Y(yy), text=r['name'], font=f_name, fill=INK, anchor='w', tags=tg)
            self._add(t, 'legend:' + rid, {'fill': INK})
            self.legend_text[rid] = t
            x += dot + 8 + w + gap

    # ------------------------------------------------------------------ التعليم (الماوس)
    def _apply(self, obj):
        cv, s = self.cv, self._s
        lit = set(self.regs[obj]['uses']) if obj in self.regs else None
        for it, g, normal, faded in self.items:
            k = g.split(':', 1)[1]
            if lit is not None and g.startswith('card:') and k not in lit:
                cv.itemconfigure(it, **(faded or {a: (_blend(v, .8) if isinstance(v, str) and v.startswith('#')
                                                      and v != CARD else v) for a, v in normal.items()}))
            else:
                cv.itemconfigure(it, **normal)
        if obj in self.cardbox:                         # الكارت اللي عليه الماوس: حرف بلونه
            cv.itemconfigure(self.cardbox[obj], outline=COL[self.sec[obj]['c']], width=max(2, 2.5 * s))
        if obj in self.regs:                            # السجل: نقطته تتعلّم في كل كارت، واسمه تحت بلونه
            for (k, rid), (o, ring) in self.dots.items():
                if rid == obj: cv.itemconfigure(o, image=ring)
            t = self.legend_text.get(obj)
            if t: cv.itemconfigure(t, fill=COL[self.regs[obj]['col']])

    def _obj_at(self, x, y):
        for it in reversed(self.cv.find_overlapping(x - 2, y - 2, x + 2, y + 2)):
            if self.cv.itemcget(it, 'state') == 'hidden': continue
            for t in self.cv.gettags(it):
                if t.startswith('obj:'): return t[4:]
        return None

    def _focus(self, obj):
        self.shown = obj
        self.hot = obj if (obj in self.sec or obj in self.regs) else None
        self._apply(self.hot)
        self._details(obj)

    def _on_motion(self, ev):
        if self._anim: return
        obj = self._obj_at(ev.x, ev.y)
        if obj:
            if self._pending: self.after_cancel(self._pending); self._pending = None
            if obj != self.shown: self._focus(obj)
            cur = 'hand2' if (obj in self.regs or obj in self.ready or obj == 'core') else ''
        else:
            cur = ''
            if self.shown is not None and not self._pending: self._pending = self.after(300, self._maybe_clear)
        if self.cv.cget('cursor') != cur: self.cv.configure(cursor=cur)

    def _maybe_clear(self):
        """الماوس بقاله شوية على مكان فاضي جوه الخريطة ← نرجع للشكل العادي. لو خرج برّه الخريطة سيبها زي ما هي."""
        self._pending = None
        try:
            px, py = self.cv.winfo_pointerxy()
            x, y = px - self.cv.winfo_rootx(), py - self.cv.winfo_rooty()
        except tk.TclError:
            return
        if 0 <= x < self.cv.winfo_width() and 0 <= y < self.cv.winfo_height() and not self._obj_at(x, y):
            self._focus(None)

    def _on_click(self, ev):
        gap = ev.time - self._last_click
        if 0 <= gap < 600: return                      # الدبل كليك ميفتحش الحاجة مرتين
        self._last_click = ev.time
        if self._anim: return
        obj = self._obj_at(ev.x, ev.y)
        if obj == 'core':
            self._toggle()
        elif obj in self.sec:
            if obj in self.ready and obj in self.DRILL and self.expanded:
                self.fly_to(obj, lambda o=obj: self.open_section(o))
            elif obj in self.ready: self.open_section(obj)
        elif obj in self.regs:
            self._open_register(obj, ev)
        elif obj is None:
            self._focus(None)

    # ------------------------------------------------------------------ فتح ملفات السجلات
    def _open_register(self, rid, ev):
        opens = self.regs[rid]['opens']
        if len(opens) == 1: return self._open_target(*opens[0])
        if self._menu is None: self._menu = tk.Menu(self, tearoff=0)
        m = self._menu
        m.delete(0, 'end')
        for label, kind, val in opens:
            m.add_command(label='Open ' + label, command=lambda l=label, k=kind, v=val: self._open_target(l, k, v))
        try: m.tk_popup(ev.x_root, ev.y_root)
        finally: m.grab_release()

    def _dc_sheet(self):
        try:
            p = os.path.join(HERE, 'modules', 'dc')
            if p not in sys.path: sys.path.insert(0, p)
            import dc_sheet
            return dc_sheet.sheet_path()
        except Exception:
            return os.path.join(os.path.expanduser('~'), 'OneDrive', '01-N.W.C PROJECT', 'CO WATER', 'DC.xlsx')

    def _open_target(self, label, kind, val):
        if kind == 'section': return self.open_section(val)
        if kind == 'dcsheet': path = self._dc_sheet()
        else: path = os.path.join(self.root, val.replace('\\', os.sep)) if val else self.root
        if not path or not os.path.exists(path):
            messagebox.showwarning('Not found', f'Could not find {label}:\n\n{path or "(no project folder)"}', parent=self)
            return
        try:
            path = os.path.normpath(path)
            if hasattr(os, 'startfile'): os.startfile(path)
            else: subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror('Open', f'Could not open {label}:\n\n{path}\n\n{e}', parent=self)

    def destroy(self):
        for j in (self._job, self._pending, self._fill_job, self._anim_job, getattr(self, '_fly_job', None),
                  getattr(self, '_home_job', None)):
            if j:
                try: self.after_cancel(j)
                except Exception: pass
        super().destroy()


# ================================================================ Document Center (DOC) – خريطة جوه الخريطة
# نفس شكل الرئيسية وكودها: دايرة في النص وكروت حواليها. التوزيع اللي كان DC بقى كارت هنا (Distribute).
def inbox_dirs(base):
    """فولدر المراجعة وفولدر Claude جوه INBOX: الأسامي بالأرقام (01_REVIEW / 02_CLAUDE)، أو القديمة لو لسه ماتغيّرتش."""
    def pick(new, old):
        return old if not os.path.isdir(os.path.join(base, new)) and os.path.isdir(os.path.join(base, old)) else new
    return pick('01_REVIEW', 'مراجعة'), pick('02_CLAUDE', 'Claude')


DOC_SECTIONS = [('dist', 'Distribute', 'Papers in  ·  00_INBOX', 'wir'), ('new', 'Documents', 'See  ·  Preview  ·  New', 'dc'),
                ('regs', 'Registers', 'DC sheet  ·  Logs', 'qs'), ('claude', 'Claude Folder', 'Hard papers read in chat', 'arch'),
                ('comp', 'Company Docs', 'CR  ·  Zakat  ·  VAT', 'sub'), ('cycle', 'Dashboard', 'Submittals  ·  status  ·  overdue', 'upd')]      # 21-11: كان Submittals Cycle (Soon)
DOC_ICON = {'dist': '⇩', 'new': '✎', 'regs': '☷', 'claude': '☁', 'comp': '⚑', 'cycle': '▦'}
DOC_READY = {'dist', 'regs', 'claude'}
DOC_REGISTERS = [
    dict(id='dcx', name='DC.xlsx', uses=['dist', 'new', 'regs'], col='dc', opens=[('DC.xlsx on OneDrive', 'dcsheet', '')]),
    dict(id='logs', name='Approval Logs', uses=['dist', 'new', 'cycle', 'regs'], col='mat',
         opens=[('02-LOG MAR+MIR.xlsm', 'file', r'00_Registers\02-LOG MAR+MIR.xlsm'),
                ('01-LOG WIR.xlsm', 'file', r'00_Registers\01-LOG WIR.xlsm'),
                ('05-LOG SD.xlsx', 'file', r'00_Registers\05-LOG SD.xlsx'),
                ('08-LOG TESTS.xlsx', 'file', r'00_Registers\08-LOG TESTS.xlsx')]),
    dict(id='sublog', name='10-LOG SUB', uses=['dist', 'comp', 'regs'], col='sub',
         opens=[('10-LOG SUB.xlsx', 'file', r'00_Registers\10-LOG SUB.xlsx')]),
    dict(id='ipclog', name='11-LOG IPC', uses=['dist', 'regs'], col='ipc',
         opens=[('11-LOG IPC.xlsx', 'file', r'00_Registers\11-LOG IPC.xlsx')]),
    dict(id='arch', name='Archive Folders', uses=['dist', 'cycle', 'claude'], col='arch',
         opens=[('the project folder', 'folder', '')]),
    dict(id='inbox', name='00_INBOX', uses=['dist', 'claude'], col='wlog', opens=[('00_INBOX', 'folder', '00_INBOX')]),
]
DOC_INFO = {
    'core': dict(text='Every paper of the project: papers in, papers out, the registers, and what is still missing. '
                      'The coloured dots on a card are the registers it writes to.',
                 hints=['Point at a card: what it does', 'Click a card: open it',
                        'Click a register name: open its file', '← Home: back to the project map']),
    'dist': dict(text='Drop any paper, scan or zip in 00_INBOX. The program reads it, knows its type and number, names '
                      'it, suggests the folder and files it when you press Approve – then writes it in the registers.',
                 takes=['00_INBOX and its review folder', 'Claude folder – hard papers read in chat',
                        'DC.xlsx – the number and title of every outgoing letter'],
                 gives=['Archive folders – 1_WIP / 2_SHARED / 3_PUBLISHED', 'DC.xlsx – letters in / out',
                        'Logs – MAR, MIR, WIR, SD', '10-LOG SUB – contracts and their approvals',
                        '11-LOG IPC – Certified, payrolls to the IPC folder']),
    'new': dict(text='Every paper by type on its own tab: the list with its status, the paper itself in the preview, '
                     'and «＋ New» to make a letter or prepare a submittal with all its requirements.',
                takes=['DC.xlsx – the letters', '02-LOG MAR+MIR – MAR and MIR', 'The submittal folders – SD, MS, ITP, RFI'],
                gives=['Letters out – Word + PDF, numbered and registered', 'Submittal package – one PDF in its 1_WIP folder']),
    'regs': dict(text='All the registers in one place. Click to choose which one to open.',
                 takes=['00_Registers', 'DC.xlsx on OneDrive']),
    'claude': dict(text='Papers the program could not read well. Press «Send to Claude» in Distribute, then write '
                        '«شوف فولدر Claude» in the chat – the result comes back next to the paper.',
                   takes=[r'00_INBOX\02_CLAUDE']),
    'comp': dict(text='The company papers that go with every IPC – commercial registration, Zakat, VAT, articles of '
                      'association – with their expiry dates and a warning before they end.',
                 soon=['Register with expiry dates', 'The valid copy goes to the IPC folder for upload']),
    'cycle': dict(text='Every paper of the project at a glance: what waits in INBOX, where every submittal is now '
                       '(in preparation → sent → approved), and the MAR that got no reply after 14 days.',
                  takes=['The project folders – 1_WIP / 2_SHARED / 3_PUBLISHED / 4_ARCHIVED', '02-LOG MAR+MIR – sent and reply dates',
                         '00_INBOX and the Claude folder']),
    'dcx': dict(path=r'OneDrive\01-N.W.C PROJECT\CO WATER\DC.xlsx',
                text='Letters out to the the Consultant, CLIENT, suppliers and others, and letters in. The outgoing number and title '
                     'give the file its name.'),
    'logs': dict(path=r'00_Registers\02-LOG MAR+MIR · 01-LOG WIR · 05-LOG SD · 08-LOG TESTS',
                 text='The approval logs – the status of every MAR, MIR, WIR, shop drawing and test.'),
    'sublog': dict(path=r'00_Registers\10-LOG SUB.xlsx', text='Subcontractors, their contracts and the 3 approvals, '
                                                             'and the deduction memos.'),
    'ipclog': dict(path=r'00_Registers\11-LOG IPC.xlsx', text='Every client IPC with its status.'),
    'arch': dict(path='02_Correspondence · 04_Submittals · 05_Quality · 08_Cost_Payments · 09_Procurement_Materials · 13_PARTIES',
                 text='Where every paper is filed: 1_WIP prepared, 2_SHARED submitted, 3_PUBLISHED approved, 4_ARCHIVED.'),
    'inbox': dict(path='00_INBOX', text='Drop the papers here – Distribute reads them.'),
}
DOC_EXT = ('.pdf', '.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp', '.zip', '.docx', '.doc', '.xlsx', '.xls')


class DocMap(HomeMap):
    """خريطة Document Center: نفس الرئيسية، بكروت المركز. ← Home يرجّع للخريطة الكبيرة."""
    SECTIONS, REGISTERS, INFO, ICONS = DOC_SECTIONS, DOC_REGISTERS, DOC_INFO, DOC_ICON
    CORE_TXT, TITLE, INBOX_CARD = ('Document', 'Center'), 'Document Center', 'dist'
    CW = 312.0                                         # أسامي الكروت هنا أطول (New Document، Submittals Cycle)
    POS = {'dist': (-330, -160), 'new': (330, -160), 'regs': (410, 30), 'claude': (260, 240),
           'comp': (-260, 240), 'cycle': (-410, 30)}

    DRILL = set()

    def __init__(self, parent, open_section, root, name='', number='', on_fail=None, log=None, intro=False):
        self._claude_n = self._claude_ready = 0
        super().__init__(parent, open_section, DOC_READY, root, name, number, inbox_count=self._count,
                         on_fail=on_fail, log=log, intro=intro)

    def _start_open(self):
        return not self._intro

    def go_home(self):
        """← Home: الكروت ترجع للدايرة، وبعدين الخريطة الكبيرة تطلع من دايرتها."""
        if self._anim: return
        if getattr(self, '_home_job', None): return       # دوسة تانية وهي راجعة
        if self.expanded:
            self._toggle()
            self._home_job = self.after(16 * 13 + 40, self._home_now)
        else:
            self.open_section('home')

    def _home_now(self):
        self._home_job = None
        if self.winfo_ismapped(): self.open_section('home')

    def _count(self):
        """الورق اللي مستني: INBOX + «مراجعة» + فولدر Claude (والـ Claude لوحده للكارت بتاعه)."""
        base = os.path.join(self.root, '00_INBOX')
        n = 0
        rv, cl = inbox_dirs(base)
        for sub in ('', rv, cl):
            d = os.path.join(base, sub)
            if not os.path.isdir(d): continue
            for f in os.listdir(d):
                p = os.path.join(d, f)
                if os.path.isfile(p) and f.lower().endswith(DOC_EXT) and not f.startswith(('~', '.')):
                    n += 1
                    if sub == cl:
                        self._claude_n += 1
                        if os.path.isfile(p + '.claude.json'): self._claude_ready += 1
        return n

    def _badge(self, k):
        if k == 'dist' and self.inbox_n: return (f'{self.inbox_n} waiting', '#6B4A00', '#FFF4D6', '#E8A400')
        if k == 'claude' and self._claude_n:
            return (f'{self._claude_ready}/{self._claude_n} read', COL['arch'], _blend(COL['arch'], .86), None)
        if k not in self.ready: return ('Soon', COL['planned'], _blend(COL['planned'], .88), None)
        return None

    def _side(self, side, name, number):
        side.columnconfigure(0, weight=1); side.rowconfigure(4, weight=1)
        back = tk.Label(side, text='←  Home', bg=CARD, fg=NAVY, font=(self.fam, 10, 'bold'), padx=12, pady=5,
                        cursor='hand2', highlightthickness=1, highlightbackground=LINE)
        back.grid(row=0, column=0, sticky='w', pady=(0, 8))
        back.bind('<Button-1>', lambda ev: self.go_home())
        tk.Label(side, text='Technical Office Platform  ›  Document Center', bg=BG, fg=MUTED,
                 font=(self.fam, 10, 'bold'), anchor='w').grid(row=1, column=0, sticky='ew')
        t1 = tk.Label(side, text='Document Center', bg=BG, fg=NAVY, font=(self.fam, 16, 'bold'), anchor='w',
                      justify='left', wraplength=self.side_w - 10)
        t1.grid(row=2, column=0, sticky='ew')
        t2 = tk.Label(side, text='Papers in  ·  papers out  ·  registers', bg=BG, fg=MUTED, font=(self.fam, 10),
                      anchor='w', justify='left', wraplength=self.side_w - 10)
        t2.grid(row=3, column=0, sticky='ew', pady=(2, 12))
        self._wraps = (t1, t2)
        card = tk.Frame(side, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        card.grid(row=4, column=0, sticky='nsew'); card.pack_propagate(False)
        self.box = tk.Frame(card, bg=CARD, padx=16, pady=14); self.box.pack(fill='both', expand=True)
        n = self.inbox_n
        bar = tk.Frame(side, bg=NAVY, padx=18, pady=11); bar.grid(row=5, column=0, sticky='ew', pady=(10, 0))
        txt = ('Nothing waiting in INBOX' if n == 0 else f'{n} paper(s) waiting – Distribute') if n is not None else 'Ready'
        l1 = tk.Label(bar, text=txt, bg=NAVY, fg='white', font=(self.fam, 12, 'bold'), anchor='w'); l1.pack(side='left')
        l1b = tk.Label(bar, text='✔' if not n else '●', bg=NAVY, fg='#7CD99B' if not n else '#FFC000',
                       font=(self.fam, 15, 'bold')); l1b.pack(side='right')
        for w in (bar, l1, l1b):
            w.configure(cursor='hand2'); w.bind('<Button-1>', lambda ev: self.open_section('dc'))
        # «Waiting for you»: العقود الناقصة والمستخلص اللي عليه الدور (بيتملي بعد ما الشاشة تظهر)
        self.v_next = tk.StringVar(value='')
        b2 = tk.Frame(side, bg=NAVY2, padx=18, pady=9)
        l2 = tk.Label(b2, textvariable=self.v_next, bg=NAVY2, fg='white', font=(self.fam, 10, 'bold'), anchor='w',
                      justify='left', wraplength=self.side_w - 45)
        l2.pack(fill='x'); self.l2, self.b2 = l2, b2
        self._fill_job = self.after(300, self._fill_next)

    def _fill_next(self):
        self._fill_job = None
        lines = []
        mods = os.path.join(HERE, 'modules')
        for p in ('qs', 'sub', 'ipc'):
            q = os.path.join(mods, p)
            if q not in sys.path: sys.path.append(q)
        try:
            import sub_core
            for c in sub_core.contracts():
                if c['missing']: lines.append(f"{c['sub']} contract {c['no']} – missing {', '.join(c['missing'])}")
        except Exception:
            pass
        try:
            import ipc_core
            out = ipc_core.qs_core.project_path(ipc_core.OUT_REL)
            no = ipc_core.next_ipc(out_dir=out)
            _, res = ipc_core.upload_check(no, out)
            miss = [k for k, f in res if not f]
            lines.append(f'IPC-{no:02d} for upload – ' + (f"missing {', '.join(miss)}" if miss else 'all files there'))
        except Exception:
            pass
        if self._claude_n:
            lines.append(f'Claude folder – {self._claude_ready} of {self._claude_n} read, ready to distribute')
        try:
            if lines:
                self.v_next.set('Waiting for you\n' + '\n'.join('•  ' + t for t in lines))
                self.b2.grid(row=6, column=0, sticky='ew', pady=(8, 0))
        except tk.TclError:
            pass

    def _on_click(self, ev):
        gap = ev.time - self._last_click
        if 0 <= gap < 600: return
        if self._anim: return
        obj = self._obj_at(ev.x, ev.y)
        if obj == 'dist':
            self._last_click = ev.time; return self.open_section('dc')
        if obj == 'regs':
            self._last_click = ev.time
            opens = [o for r in self.REGISTERS if r['id'] != 'inbox' for o in r['opens']]
            opens.append(('the 00_Registers folder', 'folder', '00_Registers'))
            self.regs['regs'] = dict(opens=opens)        # مؤقت للقايمة بس
            try: return self._open_register('regs', ev)
            finally: self.regs.pop('regs', None)
        if obj == 'claude':
            self._last_click = ev.time
            return self._open_target('the Claude folder', 'folder', os.path.join('00_INBOX', inbox_dirs(os.path.join(self.root, '00_INBOX'))[1]))
        return super()._on_click(ev)
