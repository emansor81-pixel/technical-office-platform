# -*- coding: utf-8 -*-
"""Project Platform | one window for all project tools.
Project name, number and data folder come from SETTINGS.json (next to the program folder – see SETTINGS.example.json).
Each module lives in modules\\<name> and is loaded when its button is pressed.
"""
import os, sys, json, traceback, importlib
import tkinter as tk
from tkinter import ttk, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, 'platform_error.log')
NAVY, NAVY2, SIDE, SIDE_HOVER, SIDE_ACT, BG = '#1F4E78', '#2E6496', '#16324F', '#21466B', '#2E6496', '#F3F5F8'


def find_settings():
    d = HERE
    for _ in range(4):
        f = os.path.join(d, 'SETTINGS.json')
        if os.path.exists(f): return f
        d = os.path.dirname(d)
    return None


def load_settings():
    f = find_settings()
    if not f: return {}
    try:
        with open(f, encoding='utf-8-sig') as fh: return json.load(fh)
    except Exception:
        return {}


S = load_settings()
PROJECT_NAME = S.get('PROJECT_NAME', 'Project')
PROJECT_NO = str(S.get('PROJECT_NO', ''))
PROJECT_ROOT = S.get('PROJECT_ROOT', '')

# (key, label, module folder, module file, class) ; module None = not built yet
MODULES = [
    ('home', 'Home', None, None, None),
    ('doc', 'DOC  |  Document Center', 'DOCMAP', None, None),          # خريطة المركز (home_map.DocMap)
    ('dc', 'Distribute  |  Document Center', 'dc', 'dc_app', 'App'),
    ('qs', 'QS  |  Quantities', 'qs', 'qs_app', 'App'),
    ('wir', 'WIR  |  Inspection Requests', 'wir', 'wir_app', 'App'),
    ('ipc', 'IPC  |  Payment Certificates', 'ipc', 'ipc_app', 'App'),
    ('sub', 'SUB  |  Subcontractors', 'sub', 'sub_app', 'App'),
    ('mar', 'MAR / MIR  |  Materials', 'mat', 'mat_app', 'App'),
]


def write_log(err):
    try:
        with open(LOG, 'w', encoding='utf-8') as f: f.write(err)
    except Exception:
        pass


UI_FILE = os.path.join(HERE, 'ui.json')
def ui_get(k, d=None):
    try:
        import json
        with open(UI_FILE, encoding='utf-8') as f: return json.load(f).get(k, d)
    except Exception: return d
def ui_set(k, v):
    try:
        import json
        try:
            with open(UI_FILE, encoding='utf-8') as f: s = json.load(f)
        except Exception: s = {}
        s[k] = v
        with open(UI_FILE, 'w', encoding='utf-8') as f: json.dump(s, f)
    except Exception: pass


class Platform(tk.Tk):
    def __init__(self):
        super().__init__()
        self.base_title = f'{PROJECT_NAME}  |  PO {PROJECT_NO}' if PROJECT_NO else PROJECT_NAME
        self.title(self.base_title)
        self.geometry('1400x800'); self.minsize(1100, 640)
        try: self.state('zoomed')
        except tk.TclError: pass
        ico = os.path.join(HERE, 'assets', 'app.ico')
        if os.path.exists(ico):
            try: self.iconbitmap(ico)
            except tk.TclError: pass
        self.configure(bg=BG)
        self.report_callback_exception = self.on_tk_error
        self.frames, self.buttons, self.current, self._intro = {}, {}, None, {}
        self._build()
        self.protocol('WM_DELETE_WINDOW', self.on_close)
        self.show('home')

    def on_tk_error(self, *args):
        err = ''.join(traceback.format_exception(*args))
        write_log(err); print(err)
        messagebox.showerror('Error', err[-1500:])

    # ---------------- layout
    def _build(self):
        head = tk.Frame(self, bg=NAVY, height=48); head.pack(fill='x'); head.pack_propagate(False)
        tg = tk.Label(head, text='☰', bg=NAVY, fg='white', font=('Segoe UI', 16, 'bold'), padx=12, cursor='hand2')
        tg.pack(side='left'); self.tg = tg
        tg.bind('<Button-1>', lambda ev: self.toggle_side())
        tg.bind('<Enter>', lambda ev: (tg.configure(bg=SIDE_HOVER), self.open_drawer(pin=False)))
        tg.bind('<Leave>', lambda ev: tg.configure(bg=NAVY))
        tk.Label(head, text=PROJECT_NAME, bg=NAVY, fg='white', font=('Segoe UI', 14, 'bold')).pack(side='left', padx=16)
        if PROJECT_NO:
            tk.Label(head, text=f'PO {PROJECT_NO}', bg=NAVY, fg='#BDD7EE', font=('Segoe UI', 11)).pack(side='left')
        self.v_where = tk.StringVar(value='')
        tk.Label(head, textvariable=self.v_where, bg=NAVY, fg='#BDD7EE', font=('Segoe UI', 10)).pack(side='right', padx=16)

        body = tk.Frame(self, bg=BG); body.pack(fill='both', expand=True)
        self.body = body
        self.content = tk.Frame(body, bg=BG); self.content.pack(fill='both', expand=True)
        # القايمة: شباك صغير قدام الشاشة + طبقة غامقة فوق الصفحة بتقفلها (مفيش أي حاجة تشتغل وراها)
        self.dim = tk.Toplevel(self); self.dim.withdraw(); self.dim.overrideredirect(True)
        self.dim.configure(bg='#0B1724', cursor='arrow'); self.dim.transient(self)
        try: self.dim.attributes('-alpha', 0.35)
        except tk.TclError: pass
        self.dim.bind('<Button-1>', lambda ev: self.close_drawer())
        self.drawer = tk.Toplevel(self.dim); self.drawer.withdraw(); self.drawer.overrideredirect(True)
        self.drawer.configure(bg=SIDE); self.drawer.transient(self)
        side = tk.Frame(self.drawer, bg=SIDE); side.pack(side='left', fill='both', expand=True)
        tk.Frame(self.drawer, bg='#0A1622', width=3).pack(side='right', fill='y')           # حافة/ظل على اليمين
        self.side = side
        top = tk.Frame(side, bg=SIDE); top.pack(fill='x', padx=18, pady=(18, 10))
        tk.Label(top, text=PROJECT_NAME, bg=SIDE, fg='white', font=('Segoe UI', 12, 'bold'), anchor='w', justify='left', wraplength=190).pack(fill='x')
        if PROJECT_NO:
            tk.Label(top, text=f'PO {PROJECT_NO}', bg=SIDE, fg='#9FB8D0', font=('Segoe UI', 9), anchor='w').pack(fill='x')
        tk.Frame(side, bg='#2C4A68', height=1).pack(fill='x', padx=14, pady=(4, 8))
        self.marks = {}
        for key, label, folder, *_ in MODULES:
            ready = key == 'home' or folder is not None
            row = tk.Frame(side, bg=SIDE); row.pack(fill='x', pady=1)
            mk = tk.Frame(row, bg=SIDE, width=4); mk.pack(side='left', fill='y')
            b = tk.Label(row, text=label if ready else label + '   ·  soon', anchor='w',
                         bg=SIDE, fg='white' if ready else '#7F93A8', font=('Segoe UI', 10, 'bold' if ready else 'normal'),
                         padx=16, pady=12, cursor='hand2' if ready else 'arrow')
            b.pack(side='left', fill='x', expand=True)
            if ready:
                b.bind('<Button-1>', lambda ev, k=key: self.show(k))
                b.bind('<Enter>', lambda ev, k=key: self._hover(k, True))
                b.bind('<Leave>', lambda ev, k=key: self._hover(k, False))
            self.buttons[key] = b; self.marks[key] = mk
        tk.Label(side, text=f'Data:\n{PROJECT_ROOT}', bg=SIDE, fg='#7F93A8', font=('Segoe UI', 8),
                 justify='left', wraplength=200, anchor='w').pack(side='bottom', fill='x', padx=14, pady=12)
        for w in (self.drawer, self.dim): w.bind('<Escape>', lambda ev: self.close_drawer())
        self.bind('<Escape>', lambda ev: self.close_drawer(), add='+')
        self.bind('<Configure>', lambda ev: self._place_drawer() if ev.widget is self else None, add='+')
        self.bind('<Unmap>', lambda ev: self.close_drawer() if ev.widget is self else None, add='+')
        self.pinned, self._open, self._poll, self._anim = False, False, None, None

    # ---------------- القايمة: تطلع بالماوس على ☰ – كليك على ☰ يثبتها – وهي مفتوحة الصفحة مقفولة
    DRAWER_W = 230

    def _geom(self):
        self.update_idletasks()
        return self.body.winfo_rootx(), self.body.winfo_rooty(), self.body.winfo_width(), self.body.winfo_height()

    def _place_drawer(self, w=None):
        if not self._open: return
        x, y, bw, bh = self._geom()
        self.dim.geometry(f'{bw}x{bh}+{x}+{y}')
        self.drawer.geometry(f'{w or self.DRAWER_W}x{bh}+{x}+{y}')

    def open_drawer(self, pin=False):
        if pin: self.pinned = True
        if self._open:
            return
        self._open = True
        x, y, bw, bh = self._geom()
        self.dim.geometry(f'{bw}x{bh}+{x}+{y}'); self.dim.deiconify(); self.dim.lift()
        self.drawer.geometry(f'1x{bh}+{x}+{y}'); self.drawer.deiconify(); self.drawer.lift()
        self.drawer.focus_force()
        steps = [60, 120, 170, 205, 225, self.DRAWER_W]
        def anim(i=0):
            if not self._open: return
            self._place_drawer(steps[i])
            if i + 1 < len(steps): self._anim = self.after(14, anim, i + 1)
        anim()
        self._watch()

    def close_drawer(self):
        if not self._open: return
        self._open = False; self.pinned = False
        for j in (self._poll, self._anim):
            if j: self.after_cancel(j)
        self._poll = self._anim = None
        self.drawer.withdraw(); self.dim.withdraw()

    def toggle_side(self):
        if self._open and self.pinned: self.close_drawer()
        else:
            self.pinned = True
            self.open_drawer(pin=True)

    def _inside(self, w, x, y):
        return w.winfo_ismapped() and w.winfo_rootx() <= x < w.winfo_rootx() + w.winfo_width() and \
            w.winfo_rooty() <= y < w.winfo_rooty() + w.winfo_height()

    def _watch(self):
        """لو مش متثبتة: تتقفل لما الماوس يبعد عنها وعن ☰"""
        self._poll = None
        if not self._open: return
        if not self.pinned:
            x, y = self.winfo_pointerxy()
            if not (self._inside(self.drawer, x, y) or self._inside(self.tg, x, y)):
                self._away = getattr(self, '_away', 0) + 1
                if self._away >= 3: self._away = 0; self.close_drawer(); return
            else:
                self._away = 0
        self._poll = self.after(150, self._watch)

    def _hover(self, key, on):
        if key == self.current: return
        self.buttons[key].configure(bg=SIDE_HOVER if on else SIDE)
        self.buttons[key].master.configure(bg=SIDE_HOVER if on else SIDE)

    def _home(self):
        """الشاشة الرئيسية: خريطة البرنامج (home_map.py). لو فيها أي خطأ ← الكروت القديمة، والخطأ في platform_error.log."""
        try:
            if HERE not in sys.path: sys.path.insert(0, HERE)
            import home_map
            ready = {m[0] for m in MODULES if m[2]}
            return home_map.HomeMap(self.content, open_section=self.go, ready=ready, root=PROJECT_ROOT,
                                    name=PROJECT_NAME, number=PROJECT_NO, inbox_count=self._inbox_count,
                                    next_update=self._next_updates, on_fail=self._home_failed, log=write_log,
                                    intro=self._intro.pop('home', False))
        except Exception:
            write_log(traceback.format_exc())
            return self._home_cards()

    def _home_failed(self, frame):
        """الخريطة وقعت وهي مفتوحة ← الكروت مكانها."""
        def swap():
            if self.frames.get('home') is not frame: return
            try: frame.destroy()
            except tk.TclError: pass
            self.frames['home'] = self._home_cards()
            if self.current == 'home': self.frames['home'].pack(fill='both', expand=True)
        self.after_idle(swap)

    def _home_cards(self):
        """الشاشة الرئيسية القديمة: كارت لكل قسم بأيقونة – دوس عليه يفتحه."""
        f = tk.Frame(self.content, bg=BG)
        wrap = tk.Frame(f, bg=BG); wrap.pack(fill='both', expand=True, padx=48, pady=(34, 20))
        tk.Label(wrap, text='Technical Office Platform', bg=BG, fg='#5D6878', font=('Segoe UI', 11, 'bold')).pack(anchor='w')
        tk.Label(wrap, text=PROJECT_NAME.strip() or 'Project', bg=BG, fg=NAVY, font=('Segoe UI', 24, 'bold')).pack(anchor='w')
        ok = bool(PROJECT_ROOT) and os.path.isdir(PROJECT_ROOT)
        sub = (f'PO {PROJECT_NO}   ·   ' if PROJECT_NO else '') + ('Project folder ✓' if ok else 'Project folder not found – edit PROJECT_ROOT in SETTINGS.json')
        tk.Label(wrap, text=sub, bg=BG, fg='#375623' if ok else '#C00000', font=('Segoe UI', 10)).pack(anchor='w', pady=(2, 20))
        grid = tk.Frame(wrap, bg=BG); grid.pack(fill='x')
        cards = [m for m in MODULES if m[0] != 'home']
        cols = 3
        for c in range(cols): grid.columnconfigure(c, weight=1, uniform='c')
        for i, (key, label, folder, *_) in enumerate(cards):
            self._card(grid, key, label, folder is not None).grid(row=i // cols, column=i % cols, sticky='nsew', padx=10, pady=10)
        # الشريط اللي تحت: الحالة
        bar = tk.Frame(wrap, bg=NAVY, padx=22, pady=14); bar.pack(fill='x', pady=(22, 0), padx=10)
        n = self._inbox_count()
        txt = ('All good – nothing waiting' if n == 0 else f'{n} paper(s) waiting in INBOX') if n is not None else 'Ready'
        tk.Label(bar, text=txt, bg=NAVY, fg='white', font=('Segoe UI', 13, 'bold')).pack(side='left')
        tk.Label(bar, text='✔' if not n else '●', bg=NAVY, fg='#7CD99B' if not n else '#FFC000', font=('Segoe UI', 16, 'bold')).pack(side='right')
        if n: bar.bind('<Button-1>', lambda ev: self.show('dc'))
        # مواعيد التحديث الجاية: الشركة (تقويمها الشهري) والمالك (كل ثلاثاء) – دوس عليه يفتح الحصر
        v_next = tk.StringVar(value='')
        b2 = tk.Frame(wrap, bg=NAVY2, padx=22, pady=10, cursor='hand2')
        l2 = tk.Label(b2, textvariable=v_next, bg=NAVY2, fg='white', font=('Segoe UI', 11, 'bold'), cursor='hand2')
        l2.pack(side='left')
        for w in (b2, l2): w.bind('<Button-1>', lambda ev: self.show('qs'))
        def fill():
            t = self._next_updates()
            try:
                if t: v_next.set('Next update  ·  ' + t); b2.pack(fill='x', pady=(8, 0), padx=10)
            except tk.TclError: pass
        self.after(300, fill)
        return f

    def _next_updates(self):
        """سطر مواعيد التحديث (من 00_Registers\\12-LOG UPDATE DATES.xlsx) – فاضي لو فيه مشكلة."""
        try:
            p = os.path.join(HERE, 'modules', 'qs')
            if p not in sys.path: sys.path.insert(0, p)
            import upd_dates
            return upd_dates.due_text()
        except Exception:
            return ''

    # أيقونات من رموز ويندوز العادية (بتظهر في أي نسخة Tk)
    CARD_INFO = {'dc': ('✉', 'Read & file incoming papers'), 'qs': ('∑', 'Quantities from approved WIRs'),
                 'wir': ('☑', 'Quantity & status of every request'), 'ipc': ('✍', 'The full IPC in one click'), 'sub': ('⚑', 'Contracts, certificates & final account'),
                 'mar': ('⚒', 'MAR, MIR, deliveries & stock'), 'doc': ('✉', 'Papers in, papers out, registers')}

    def _card(self, parent, key, label, ready):
        icon, desc = self.CARD_INFO.get(key, ('▣', ''))
        title = label.split('|')[-1].strip() if '|' in label else label
        code = label.split('|')[0].strip() if '|' in label else ''
        bg, hov = 'white', '#EEF4FB'
        c = tk.Frame(parent, bg=bg, highlightbackground='#D9E0E8', highlightthickness=1, padx=20, pady=18,
                     cursor='hand2' if ready else 'arrow')
        top = tk.Frame(c, bg=bg); top.pack(fill='x')
        ic = tk.Label(top, text=icon, bg='#E6EEF8', fg=NAVY if ready else '#9AA5B4', font=('Segoe UI Symbol', 20), width=3, pady=4); ic.pack(side='left')
        st = tk.Label(top, text='✓' if ready else 'soon', bg='#E2EFDA' if ready else '#EDEFF2',
                      fg='#2E7D4F' if ready else '#7F8A99', font=('Segoe UI', 10, 'bold'), padx=8, pady=2)
        st.pack(side='right', anchor='n')
        t1 = tk.Label(c, text=title, bg=bg, fg=NAVY if ready else '#7F8A99', font=('Segoe UI', 14, 'bold'), anchor='w')
        t1.pack(fill='x', pady=(14, 0))
        t2 = tk.Label(c, text=(code + '  ·  ' if code else '') + desc, bg=bg, fg='#5D6878', font=('Segoe UI', 10), anchor='w')
        t2.pack(fill='x', pady=(2, 0))
        if ready:
            parts = (c, top, t1, t2, ic)
            def hl(on):
                for w in (c, top, t1, t2): w.configure(bg=hov if on else bg)
            for w in parts:
                w.bind('<Button-1>', lambda ev, k=key: self.show(k))
                w.bind('<Enter>', lambda ev: hl(True)); w.bind('<Leave>', lambda ev: hl(False))
        return c

    def _inbox_count(self):
        try:
            d = os.path.join(PROJECT_ROOT, '00_INBOX')
            return sum(1 for n in os.listdir(d) if os.path.isfile(os.path.join(d, n)) and not n.lower().startswith('readme'))
        except Exception:
            return None

    def _docmap(self):
        """خريطة Document Center (نفس شكل الرئيسية). لو وقعت ← شاشة التوزيع على طول."""
        try:
            if HERE not in sys.path: sys.path.insert(0, HERE)
            import home_map
            return home_map.DocMap(self.content, open_section=self.go, root=PROJECT_ROOT, name=PROJECT_NAME,
                                   number=PROJECT_NO, on_fail=self._doc_failed, log=write_log,
                                   intro=self._intro.pop('doc', False))
        except Exception:
            write_log(traceback.format_exc())
            return self._load('dc')

    def _doc_failed(self, frame):
        def swap():
            if self.frames.get('doc') is not frame: return
            try: frame.destroy()
            except tk.TclError: pass
            self.frames.pop('doc', None)
            if self.current == 'doc': self.current = None; self.show('dc')
        self.after_idle(swap)

    def go(self, key):
        """فتح من جوه الخرايط: الرئيسية وخريطة المركز بيطلعوا بحركة (الكروت من الدايرة)."""
        if key in ('home', 'doc') and key != self.current: self._intro[key] = True
        self.show(key)

    def _load(self, key):
        if key == 'home': return self._home()
        if key == 'doc': return self._docmap()
        _, label, folder, modname, cls = next(m for m in MODULES if m[0] == key)
        path = os.path.join(HERE, 'modules', folder)
        if path not in sys.path: sys.path.insert(0, path)
        mod = importlib.import_module(modname)
        return getattr(mod, cls)(self.content)

    def show(self, key):
        if key == self.current: return
        if key in ('home', 'doc') and key in self.frames:   # الرئيسية وخريطة المركز تتبني من جديد عشان العدد يبقى محدّث
            self.frames.pop(key).destroy()
        try:
            if key not in self.frames: self.frames[key] = self._load(key)
        except Exception:
            err = traceback.format_exc(); write_log(err)
            messagebox.showerror('Error', f'Could not open this section:\n\n{err[-1200:]}'); return
        if self.current:
            self.frames[self.current].pack_forget(); self.buttons[self.current].configure(bg=SIDE)
            self.buttons[self.current].master.configure(bg=SIDE); self.marks[self.current].configure(bg=SIDE)
        self.frames[key].pack(fill='both', expand=True)
        self.buttons[key].configure(bg=SIDE_ACT); self.buttons[key].master.configure(bg=SIDE_ACT)
        self.marks[key].configure(bg='#4DA3FF')
        self.current = key
        self.v_where.set(next(m[1] for m in MODULES if m[0] == key))
        self.close_drawer()

    def on_close(self):
        for k, f in self.frames.items():
            fn = getattr(f, 'on_close', None)
            if fn and fn() is False:
                self.show(k); return
        self.destroy()


# ---------------- مراقب التهنيج (مؤقت): أي توقف أكتر من ربع ثانية بيتسجل في PERF_LOG.txt مع السطر اللي كان شغال
PERF_WATCH = True
PERF_LOG = os.path.join(HERE, 'PERF_LOG.txt')


class PerfWatch:
    def __init__(self, app, limit=0.25):
        import threading, time
        self.app, self.limit, self.time = app, limit, time
        self.main_id = threading.get_ident()
        self.last = time.perf_counter()
        self.stats = {'ticks': 0, 'stalls': 0}
        self._log(f'===== بدأ {__import__("datetime").datetime.now():%Y-%m-%d %H:%M:%S} =====')
        app.after(50, self._tick)
        threading.Thread(target=self._watch, daemon=True).start()

    def _log(self, txt):
        try:
            with open(PERF_LOG, 'a', encoding='utf-8') as f: f.write(txt + '\n')
        except Exception: pass

    def _tick(self):
        self.last = self.time.perf_counter(); self.stats['ticks'] += 1
        n = self.stats['ticks']
        if n % 4 == 0: self.where = self._where()        # بيتحسب في البرنامج نفسه (مش من المراقب)
        # if n == 60: self._bench()      # قياس سرعة الرسم – متوقف (علّق البرنامج 97 ثانية على الجهاز)
        try: self.app.after(50, self._tick)
        except Exception: pass

    def _bench(self):
        """قياس سرعة الرسم على الجهاز ده: جدول 56 عمود × 40 صف + 300 خانة"""
        try:
            from tkinter import ttk
            t = self.time.perf_counter
            top = tk.Toplevel(self.app); top.geometry('1200x700+0+0')
            t0 = t(); cols = [f'c{i}' for i in range(56)]
            tv = ttk.Treeview(top, columns=cols, show='headings'); tv.pack(fill='both', expand=True)
            for c in cols: tv.heading(c, text='عمق الحفر'); tv.column(c, width=70)
            for r in range(40): tv.insert('', 'end', values=[f'{r}.250' if i % 3 else 'صخرية' for i in range(56)])
            top.update(); t1 = t()
            for k in range(10): tv.xview_scroll(3, 'units'); top.update()
            t2 = t()
            labs = [tk.Label(top, text='0.850') for _ in range(300)]
            for i, l in enumerate(labs): l.place(x=(i % 20) * 50, y=(i // 20) * 22)
            top.update(); t3 = t()
            top.destroy()
            self._log(f'      [قياس الجهاز] جدول: رسم أول مرة {t1 - t0:.2f} ث | 10 سكرول {t2 - t1:.2f} ث | 300 خانة {t3 - t2:.2f} ث')
            a = self.app
            try:
                import ctypes; sf = ctypes.windll.shcore.GetScaleFactorForDevice(0)
            except Exception: sf = '?'
            try:
                import ctypes; mon = ctypes.windll.user32.GetSystemMetrics(80)
            except Exception: mon = '?'
            self._log(f'      [الشاشة] {a.winfo_screenwidth()}x{a.winfo_screenheight()} | DPI {a.winfo_fpixels("1i"):.0f} | '
                      f'تكبير ويندوز {sf}% | عدد الشاشات {mon} | tk scaling {a.tk.call("tk", "scaling"):.2f} | '
                      f'البرنامج {a.winfo_width()}x{a.winfo_height()}')
        except Exception as e:
            self._log(f'      [قياس الجهاز] فشل: {e}')

    def _where(self):
        """الإيه اللي شغال دلوقتي في البرنامج (أي أيقونة وأي تبويب)"""
        try:
            k = self.app.current or ''
            f = self.app.frames.get(k)
            nb = getattr(f, 'nb', None)
            tab = nb.tab(nb.select(), 'text').strip() if nb is not None and nb.select() else ''
            return f'{k} / {tab}' if tab else k
        except Exception:
            return '?'

    def _watch(self):
        import sys, traceback, collections
        t = self.time
        while True:
            t.sleep(0.05)
            gap = t.perf_counter() - self.last
            if gap < self.limit: continue
            start = self.last; samples = collections.Counter(); where = getattr(self, 'where', '?')
            while t.perf_counter() - self.last > 0.06 and self.last == start:
                fr = sys._current_frames().get(self.main_id)
                if fr is not None:
                    st = traceback.extract_stack(fr)
                    ours = [f'{os.path.relpath(x.filename, HERE) if x.filename.startswith(HERE) else os.path.basename(x.filename)}:{x.lineno} {x.name}'
                            for x in st if 'tkinter' not in x.filename and 'threading' not in x.filename][-4:]
                    top = f'{os.path.basename(st[-1].filename)}:{st[-1].lineno} {st[-1].name}'
                    samples[' <- '.join(reversed(ours)) + f'   [آخر سطر: {top}]'] += 1
                t.sleep(0.05)
            dur = t.perf_counter() - start
            if dur < self.limit: continue
            self.stats['stalls'] += 1
            lines = [f'{__import__("datetime").datetime.now():%H:%M:%S}  هنّج {dur:.2f} ثانية  |  في: {where}']
            for s, n in samples.most_common(3): lines.append(f'      {n * 50} ms  {s}')
            self._log('\n'.join(lines))


if __name__ == '__main__':
    try:
        app = Platform()
        if PERF_WATCH:
            try: PerfWatch(app)
            except Exception: pass
        app.mainloop()
    except Exception:
        err = traceback.format_exc(); write_log(err)
        try:
            r = tk.Tk(); r.withdraw(); messagebox.showerror('Error', err[-1500:])
        except Exception:
            print(err); input('Press Enter...')
