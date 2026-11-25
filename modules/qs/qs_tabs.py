# -*- coding: utf-8 -*-
"""Extra QS tabs: Unifier, Chambers, Connections, Boring.
Manual entries are kept in p.extra and saved to QS_INPUT.xlsx automatically."""
import tkinter as tk
from tkinter import ttk, messagebox
import qs_sheets as S

EDIT_BG, RO_BG = '#FFF2CC', '#F4F4F4'


class Grid(ttk.Frame):
    """Treeview with Excel-like editing on chosen columns.
    cols: list of (id, title, width, editable, choices or None)"""

    def __init__(self, master, cols, on_edit=None, height_rows=None):
        super().__init__(master)
        self.cols = cols; self.on_edit = on_edit; self.rows = []; self.editor = None
        self.edit_ids = [c[0] for c in cols if c[3]]
        self.tv = ttk.Treeview(self, columns=[c[0] for c in cols], show='headings', selectmode='browse')
        for cid, txt, w, ed, ch in cols:
            self.tv.heading(cid, text=('✎ ' if ed else '') + txt)
            self.tv.column(cid, width=w, minwidth=30, anchor='center', stretch=False)
        ys = ttk.Scrollbar(self, orient='vertical', command=self.tv.yview)
        xs = ttk.Scrollbar(self, orient='horizontal', command=self._xview)
        self.tv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tv.grid(row=0, column=0, sticky='nsew'); ys.grid(row=0, column=1, sticky='ns'); xs.grid(row=1, column=0, sticky='ew')
        self.rowconfigure(0, weight=1); self.columnconfigure(0, weight=1)
        self.tv.tag_configure('ok', background='white'); self.tv.tag_configure('grey', foreground='#8A8A8A', background=RO_BG)
        self.tv.tag_configure('warn', background='#FDE9D9'); self.tv.tag_configure('miss', background='#FFF8E1')
        self.tv.bind('<Double-1>', self._dbl); self.tv.bind('<Return>', self._key_edit); self.tv.bind('<F2>', self._key_edit)
        self.tv.bind('<Delete>', self._clear)
        self.cur_col = self.edit_ids[0] if self.edit_ids else None

    def _xview(self, *a):
        self._close(True); self.tv.xview(*a)

    def set_rows(self, rows, values_fn, tag_fn=None):
        self._close(False)
        self.rows = rows
        sel = self.tv.selection()
        self.tv.delete(*self.tv.get_children())
        for i, r in enumerate(rows):
            self.tv.insert('', 'end', iid=str(i), values=values_fn(r), tags=(tag_fn(r) if tag_fn else 'ok',))
        if sel and self.tv.exists(sel[0]): self.tv.selection_set(sel[0])

    def _col_at(self, x):
        c = self.tv.identify_column(x)
        try: return self.cols[int(c[1:]) - 1][0]
        except (ValueError, IndexError): return None

    def _dbl(self, ev):
        iid = self.tv.identify_row(ev.y); cid = self._col_at(ev.x)
        if iid and cid in self.edit_ids: self.cur_col = cid; self.open(iid, cid)

    def _key_edit(self, ev=None):
        sel = self.tv.selection()
        if sel and self.cur_col: self.open(sel[0], self.cur_col)

    def _clear(self, ev=None):
        sel = self.tv.selection()
        if sel and self.cur_col and self.on_edit: self.on_edit(int(sel[0]), self.cur_col, '')

    def open(self, iid, cid):
        self._close(True)
        bbox = self.tv.bbox(iid, cid)
        if not bbox: self.tv.see(iid); self.update_idletasks(); bbox = self.tv.bbox(iid, cid)
        if not bbox: return
        x, y, w, h = bbox
        spec = next(c for c in self.cols if c[0] == cid)
        cur = self.tv.set(iid, cid)
        if spec[4]:
            e = ttk.Combobox(self.tv, values=spec[4], font=('Segoe UI', 10)); e.set(cur)
        else:
            e = tk.Entry(self.tv, font=('Segoe UI', 10), justify='center', bg=EDIT_BG, relief='solid', bd=1)
            e.insert(0, cur); e.select_range(0, 'end')
        e.place(x=x, y=y, width=max(w, 60), height=h); e.focus_set()
        e.bind('<Return>', lambda ev: self._close(True, 1)); e.bind('<Tab>', lambda ev: (self._close(True, 0, 1), 'break')[1])
        e.bind('<Escape>', lambda ev: self._close(False)); e.bind('<FocusOut>', lambda ev: self._close(True))
        self.editor = (e, iid, cid)

    def _close(self, commit, down=0, right=0):
        if not self.editor: return
        e, iid, cid = self.editor; self.editor = None
        txt = e.get().strip(); e.destroy()
        if commit and self.on_edit and txt != self.tv.set(iid, cid): self.on_edit(int(iid), cid, txt)
        if down or right:
            i = int(iid) + down
            if right:
                k = self.edit_ids.index(cid) + right
                if k >= len(self.edit_ids): k = 0; i += 1
                cid = self.edit_ids[k]; self.cur_col = cid
            if 0 <= i < len(self.rows):
                self.tv.selection_set(str(i)); self.tv.see(str(i)); self.after(10, lambda: self.open(str(i), cid))
        else:
            self.tv.focus_set()


def _hint(parent, text):
    ttk.Label(parent, text=text, foreground='#556', font=('Segoe UI', 9)).pack(anchor='w', pady=4)


class ExtraTabs:
    def __init__(self, app):
        self.app = app
        nb = app.nb
        self._build_unifier(nb); self._build_chambers(nb); self._build_connections(nb); self._build_boring(nb)

    @property
    def p(self): return getattr(self.app, 'p', None)

    def _changed(self):
        if self.p is not None: self.p.dirty = True
        t = self.app.title()
        if not t.endswith('*'): self.app.title(t + ' *')

    # ---------------- Unifier
    def _build_unifier(self, nb):
        f = ttk.Frame(nb); nb.add(f, text='Unifier')
        cols = [('n', '#', 36), ('line', 'Line', 50), ('x0', 'From', 70), ('x1', 'To', 70), ('need', 'Layers Req.', 75)]
        self.u_cols = []
        for gi, (title, codes) in enumerate(S.GROUPS, 1):
            short = title.split('\n')[-1]
            if codes == 'BF':
                cols += [(f'u{gi}', f'{short} – UN', 230), (f'c{gi}', 'Layers', 60), (f'i{gi}', 'INV', 60)]
            else:
                cols += [(f'u{gi}', f'{short} – UN', 120), (f'i{gi}', 'INV', 55)]
        spec = [(c, t, w, c.startswith('i'), None) for c, t, w in cols]
        self.ug = Grid(f, spec, on_edit=self._u_edit); self.ug.pack(fill='both', expand=True)
        _hint(f, 'UN = Unifier no. of the approved WIR (from column "رقم اليونيفاير" in the WIR log) – edit it in the log.   '
                 '✎ INV = invoice no. (type it here).   Red layers = a missing layer below an approved one.   '
                 '(WIR n) = approved WIR without Unifier no.   Saved automatically.')
        self.u_data = []

    def refresh_unifier(self):
        p = self.p
        if p is None: return
        try:
            rows = p.summary_rows(self.app.items)
            self.u_data = S.unifier_rows(p, rows, self.app.items)
        except Exception as e:
            self.u_data = []; print('unifier', e)
        def vals(r):
            v = [r['n'], r['line'], f"{r['x0']:g}", f"{r['x1']:g}", r['need'] or '']
            for gi, (txt, miss, cnt, gap) in enumerate(r['groups'], 1):
                inv = S.inv_of(p, r['line'], r['dist'], r['x0'], r['x1'], gi)
                if S.GROUPS[gi - 1][1] == 'BF': v += [txt, (f'{cnt} ⚠' if gap else cnt) or '', inv]
                else: v += [txt, inv]
            return v
        def tag(r):
            if any(g[3] for g in r['groups']): return 'warn'
            if any(g[1] for g in r['groups']): return 'miss'
            return 'ok'
        self.ug.set_rows(self.u_data, vals, tag)

    def _u_edit(self, i, cid, txt):
        r = self.u_data[i]; gi = int(cid[1:])
        k = (r['line'], r['dist'], round(r['x0'], 3), round(r['x1'], 3), gi)
        inv = self.p.extra.setdefault('INV', {})
        if txt: inv[k] = txt
        else: inv.pop(k, None)
        self._changed(); self.refresh_unifier()

    # ---------------- generic list tab (chambers / connections / boring)
    def _list_tab(self, nb, title, key, cols, hint, calc=None):
        f = ttk.Frame(nb); nb.add(f, text=title)
        bar = ttk.Frame(f); bar.pack(fill='x', pady=(4, 4))
        g = Grid(f, cols, on_edit=lambda i, c, t, key=key: self._l_edit(key, i, c, t))
        ttk.Button(bar, text='+ Add row', style='Sec.TButton', command=lambda: self._l_add(key)).pack(side='left')
        ttk.Button(bar, text='Delete row', style='Sec.TButton', command=lambda: self._l_del(key)).pack(side='left', padx=6)
        lbl = tk.StringVar(); ttk.Label(bar, textvariable=lbl, font=('Segoe UI', 10, 'bold'), foreground='#1F4E78').pack(side='right', padx=8)
        g.pack(fill='both', expand=True)
        _hint(f, hint)
        setattr(self, f'g_{key}', g); setattr(self, f'l_{key}', lbl); setattr(self, f'calc_{key}', calc)

    def refresh_list(self, key):
        p = self.p
        if p is None: return
        data = p.extra.setdefault(key, [])
        g = getattr(self, f'g_{key}'); calc = getattr(self, f'calc_{key}')
        def vals(r):
            i = data.index(r) + 1
            out = []
            for c in g.cols:
                cid = c[0]
                if cid == 'n': out.append(i)
                elif calc and cid in calc: out.append(calc[cid](r))
                else: out.append(r.get(cid, ''))
            return out
        g.set_rows(data, vals)
        tot = getattr(self, f'total_{key}', None)
        if tot: getattr(self, f'l_{key}').set(tot(data))

    def _l_edit(self, key, i, cid, txt):
        self.p.extra[key][i][cid] = txt; self._changed(); self.refresh_list(key)

    def _l_add(self, key):
        if self.p is None: return
        data = self.p.extra.setdefault(key, [])
        new = dict(data[-1]) if data else {}
        for k in list(new):
            if k.startswith(('un', 'inv')) or k in ('station', 'x0', 'x1', 'un', 'inv'): new[k] = ''
        data.append(new); self._changed(); self.refresh_list(key)
        g = getattr(self, f'g_{key}'); last = str(len(data) - 1); g.tv.selection_set(last); g.tv.see(last)

    def _l_del(self, key):
        g = getattr(self, f'g_{key}'); sel = g.tv.selection()
        if not sel or self.p is None: return
        if not messagebox.askyesno('QS', 'Delete the selected row?'): return
        del self.p.extra[key][int(sel[0])]; self._changed(); self.refresh_list(key)

    # ---------------- chambers
    def _build_chambers(self, nb):
        cols = [('n', '#', 36, False, None), ('zone', 'Zone', 90, True, None), ('line', 'Line', 60, True, None),
                ('type', 'Type', 70, True, list(S.CH_RATES)), ('dia', 'Dia', 55, True, ['300', '400', '600']),
                ('station', 'Station', 70, True, None), ('valves', 'Valves', 90, True, None),
                ('pct', 'Progress %', 75, False, None), ('val', 'Value (SAR)', 95, False, None)]
        for i, (h, pc) in enumerate(S.STAGES, 1):
            short = h.split('(')[0].strip()[:18]
            cols += [(f'un{i}', f'{i}. {short} – UN ({pc:.1%})', 150, True, None), (f'inv{i}', f'{i}. INV', 55, True, None)]
        calc = {'pct': lambda r: f'{S.ch_pct(r):.1%}',
                'val': lambda r: f"{S.ch_pct(r) * S.CH_RATES.get(r.get('type', ''), 0):,.0f}"}
        self._list_tab(nb, 'Chambers', 'CHAMBERS', cols,
                       'One row per chamber.  Type: ' + '  '.join(f'{t} = {it}' for t, it, _ in S.CH_TYPES) +
                       '.   Type the Unifier no. in each stage when its WIR is approved – progress % uses the agreed stage weights.', calc)
        self.total_CHAMBERS = lambda d: f"Chambers: {len(d)}   |   Value: {sum(S.ch_pct(r) * S.CH_RATES.get(r.get('type', ''), 0) for r in d):,.0f} SAR"

    # ---------------- connections
    def _build_connections(self, nb):
        cols = [('n', '#', 36, False, None), ('zone', 'Zone', 100, True, None), ('line', 'Line', 70, True, None),
                ('station', 'Station', 80, True, None), ('prop', 'Proposed Dia', 95, True, [str(x) for x in S.CON_PROP]),
                ('exist', 'Existing Dia', 95, True, [str(x) for x in S.CON_EXIST]), ('un', 'UN NO', 110, True, None), ('inv', 'INV', 70, True, None)]
        self._list_tab(nb, 'Connections', 'CONNECTIONS', cols,
                       'One row per connection (proposed line on existing line).  Export groups them by zone & line into the consultant matrix.')
        self.total_CONNECTIONS = lambda d: f'Connections: {len(d)}'

    # ---------------- boring
    def _build_boring(self, nb):
        def lm(r):
            try: return f"{abs(float(r.get('x1')) - float(r.get('x0'))):.2f}"
            except (TypeError, ValueError): return ''
        cols = [('n', '#', 36, False, None), ('zone', 'Zone', 110, True, None), ('line', 'Line', 70, True, None),
                ('dia', 'Dia', 60, True, ['300', '400', '500', '600']), ('x0', 'From', 80, True, None), ('x1', 'To', 80, True, None),
                ('lm', 'LM', 70, False, None), ('un', 'UN NO', 110, True, None), ('inv', 'INV', 70, True, None)]
        self._list_tab(nb, 'Boring', 'BORING', cols, 'Horizontal boring (item 33.1.2 – 1000008 – 11,200 SAR/LM).  LM = |To − From|.', {'lm': lm})
        def tot(d):
            s = 0.0
            for r in d:
                try: s += abs(float(r.get('x1')) - float(r.get('x0')))
                except (TypeError, ValueError): pass
            return f'Boring: {s:,.2f} LM   |   {s * 11200:,.0f} SAR'
        self.total_BORING = tot

    def refresh_all(self):
        self.refresh_unifier()
        for k in ('CHAMBERS', 'CONNECTIONS', 'BORING'): self.refresh_list(k)
