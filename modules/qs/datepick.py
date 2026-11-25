# -*- coding: utf-8 -*-
"""اختيار التاريخ من نتيجة (من غير كتابة). tkinter بس – مفيش مكتبات زيادة.
DateField(parent, var, on_change=None): خانة التاريخ + زرار 📅 ، الضغط على أي واحد فيهم يفتح النتيجة.
القيمة في var بتفضل بالشكل dd-mm-yyyy زي ما البرنامج متعود."""
import calendar, datetime as dt
import tkinter as tk
from tkinter import ttk

NAVY = '#1F4E78'
FMT = '%d-%m-%Y'
DAYS = ['أحد', 'اثنين', 'ثلاثاء', 'أربعاء', 'خميس', 'جمعة', 'سبت']   # الأسبوع بيبدأ الأحد
MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر']


def parse(s):
    s = str(s or '').strip().replace('/', '-')
    for fm in (FMT, '%Y-%m-%d', '%d-%m-%y'):
        try: return dt.datetime.strptime(s, fm).date()
        except ValueError: pass
    return None


class Picker(tk.Toplevel):
    def __init__(self, anchor, var, on_change=None):
        super().__init__(anchor)
        self.var, self.on_change = var, on_change
        self.overrideredirect(True); self.configure(bg=NAVY, padx=1, pady=1)
        self.sel = parse(var.get()) or dt.date.today()
        self.y, self.m = self.sel.year, self.sel.month
        body = tk.Frame(self, bg='white', padx=6, pady=6); body.pack()
        head = tk.Frame(body, bg='white'); head.pack(fill='x')
        tk.Button(head, text='◀', width=3, relief='flat', bg='#EEF3F8', command=lambda: self.move(-1)).pack(side='left')
        self.lbl = tk.Label(head, bg='white', fg=NAVY, font=('Segoe UI', 11, 'bold'), width=16); self.lbl.pack(side='left', expand=True)
        tk.Button(head, text='▶', width=3, relief='flat', bg='#EEF3F8', command=lambda: self.move(1)).pack(side='left')
        self.grid_f = tk.Frame(body, bg='white'); self.grid_f.pack(pady=(6, 4))
        foot = tk.Frame(body, bg='white'); foot.pack(fill='x')
        tk.Button(foot, text='النهارده', relief='flat', bg='#EEF3F8', command=lambda: self.pick(dt.date.today())).pack(side='left')
        tk.Button(foot, text='مسح', relief='flat', bg='#EEF3F8', command=lambda: self.pick(None)).pack(side='left', padx=4)
        tk.Button(foot, text='إلغاء', relief='flat', bg='#EEF3F8', command=self.destroy).pack(side='right')
        self.draw()
        x = anchor.winfo_rootx(); y = anchor.winfo_rooty() + anchor.winfo_height() + 2
        self.geometry(f'+{x}+{y}')
        self.bind('<Escape>', lambda e: self.destroy())
        self.bind('<FocusOut>', self._focus_out)
        self.after(50, self.focus_force)
        try: self.grab_set()
        except tk.TclError: pass

    def _focus_out(self, ev):
        if ev.widget is self:
            self.after(150, lambda: (self.winfo_exists() and self.focus_get() is None) and self.destroy())

    def move(self, k):
        self.m += k
        if self.m == 0: self.m, self.y = 12, self.y - 1
        if self.m == 13: self.m, self.y = 1, self.y + 1
        self.draw()

    def draw(self):
        for w in self.grid_f.winfo_children(): w.destroy()
        self.lbl.config(text=f'{MONTHS[self.m - 1]} {self.y}')
        # الأحد على اليمين زي النتيجة العربي
        cols = list(range(7))[::-1]
        for i, d in enumerate(DAYS):
            tk.Label(self.grid_f, text=d, bg='white', fg='#7F8C9A', font=('Segoe UI', 8), width=5).grid(row=0, column=cols[i])
        cal = calendar.Calendar(firstweekday=6)   # 6 = الأحد
        today = dt.date.today()
        for r, week in enumerate(cal.monthdatescalendar(self.y, self.m), 1):
            for i, day in enumerate(week):
                other = day.month != self.m
                bg = NAVY if day == self.sel else ('#DDEBF7' if day == today else 'white')
                fg = 'white' if day == self.sel else ('#B0B8C2' if other else '#222')
                b = tk.Button(self.grid_f, text=str(day.day), width=4, relief='flat', bg=bg, fg=fg,
                              activebackground='#BDD7EE', command=lambda d=day: self.pick(d))
                b.grid(row=r, column=cols[i], padx=1, pady=1)

    def pick(self, d):
        self.var.set(d.strftime(FMT) if d else '')
        try: self.grab_release()
        except tk.TclError: pass
        self.destroy()
        if self.on_change: self.on_change()


class DateField(ttk.Frame):
    def __init__(self, parent, var, on_change=None, width=11, allow_typing=False):
        super().__init__(parent)
        self.var, self.on_change = var, on_change
        self.e = ttk.Entry(self, textvariable=var, width=width, state='normal' if allow_typing else 'readonly')
        self.e.pack(side='left')
        self.b = ttk.Button(self, text='📅', width=3, command=self.open)
        self.b.pack(side='left', padx=(2, 0))
        if not allow_typing: self.e.bind('<Button-1>', lambda ev: self.open())

    def open(self):
        Picker(self.e, self.var, self.on_change)
