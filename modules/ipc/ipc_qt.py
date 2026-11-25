# -*- coding: utf-8 -*-
"""IPC | Payment Certificates بـ Qt – نفس شاشة ipc_app.py في الشغل (ipc_core / ipc_ho زي ما هما)، والشكل الجديد الغامق.
تابات: IPC – Summary – QS Summary – QS Total – Notes.  الجداول العربي من اليمين للشمال زي شيت VO."""
import os, sys, time, threading, traceback, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
MODS = os.path.dirname(HERE); PLAT = os.path.dirname(MODS)
for p in (PLAT, os.path.join(MODS, 'dc'), os.path.join(MODS, 'qs'), os.path.join(MODS, 'wir')):
    if p not in sys.path: sys.path.insert(0, p)

from PySide6.QtCore import Qt, QTimer, QDate
from PySide6.QtGui import QColor, QFont, QGuiApplication
from PySide6.QtWidgets import (QStyledItemDelegate, QStyle, QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem,
                               QAbstractItemView, QMessageBox, QTabWidget, QDialog, QLineEdit, QDateEdit, QCheckBox, QSpinBox,
                               QListWidget, QListWidgetItem, QFileDialog, QInputDialog, QFrame)
import home_qt as HQ
import dc_qt as DQ
import qs_core
import ipc_core as C
import ipc_defs as OLDI                     # نفس الأعمدة

INK, MUTED, NAVY = DQ.INK, DQ.MUTED, DQ.NAVY
TT = HQ.T
CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#ipcq') + f"""
QTabWidget::pane {{ border: none; border-top: 2px solid {NAVY}; }}
QTabBar::tab {{ background: {TT['chip']}; color: {MUTED}; padding: 7px 16px; margin-right: 2px; border-top-left-radius: 8px;
               border-top-right-radius: 8px; font: 600 10pt "Segoe UI"; }}
QTabBar::tab:selected {{ background: {NAVY}; color: white; }}
QDateEdit, QSpinBox {{ background: {TT['card2']}; border: 1px solid {DQ.SOFT}; border-radius: 8px; padding: 5px 9px; color: {INK}; }}
QListWidget {{ background: {TT['card2']}; border: none; color: {INK}; font: 10pt "Segoe UI"; }}
QListWidget::item {{ padding: 4px 6px; }}
QListWidget::item:selected {{ background: #2B5F8E; color: white; }}
"""
D = HQ.DARK
TAG = {'div': ('#1E3550' if D else '#DDEBF7', None, True), 'head': (None, '#9AA5B4' if D else '#595959', False),
       'tot': ('#3A3018' if D else '#FFF2CC', None, True), 'grand': (NAVY, '#FFFFFF', True),
       'err': (None, '#FF8A8A' if D else '#C00000', False), 'warn': (None, '#F0C860' if D else '#9A5B00', False)}
OKC, AMB, RED = DQ.GREEN, ('#F0C860' if D else '#B7791F'), ('#FF8A8A' if D else '#C00000')
money, q2 = OLDI.money, OLDI.q2


class BgDelegate(QStyledItemDelegate):
    """لون خلفية الصف (الستايل الغامق بيتجاهل لون الخلية) – بيترسم قبل الكلام"""
    def paint(self, p, opt, idx):
        bg = idx.data(Qt.BackgroundRole)
        if bg is not None and not (opt.state & QStyle.StateFlag.State_Selected): p.fillRect(opt.rect, bg)
        super().paint(p, opt, idx)


def header(title, sub, color, back, extra=()):
    t = QHBoxLayout(); t.setSpacing(14)
    b = DQ.btn('←  Back', back); b.setObjectName('back'); t.addWidget(b)
    tv = QVBoxLayout(); tv.setSpacing(0)
    c = QLabel(f'Technical Office Platform  ›  {title}'); c.setFont(HQ.font(12.5, True)); c.setStyleSheet(f'color: {MUTED};')
    tv.addWidget(c)
    tv.addWidget(QLabel(f'<span style="color:{color}; font-size:20pt; font-weight:bold;">{title}</span>'
                        f'<span style="color:{MUTED}; font-size:11pt;">   {sub}</span>'))
    t.addLayout(tv, 1)
    for w in extra: t.addWidget(w)
    return t


class Grid:
    """جدول عربي من اليمين للشمال – put(dict, tag) زي الـ Treeview القديم"""

    def __init__(self, cols, rtl=True):
        self.cols = cols; self.ids = [c for c, _, _ in cols]
        t = self.t = QTableWidget(0, len(cols))
        t.setHorizontalHeaderLabels([x for _, x, _ in cols]); t.verticalHeader().setVisible(False)
        if rtl: t.setLayoutDirection(Qt.RightToLeft)
        t.setSelectionBehavior(QAbstractItemView.SelectRows); t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.setShowGrid(False); t.setWordWrap(False); t.verticalHeader().setDefaultSectionSize(28)
        t.horizontalHeader().setHighlightSections(False)
        for i, (_, _, w) in enumerate(cols): t.setColumnWidth(i, int(w * 1.1))
        t.horizontalHeader().setStretchLastSection(True)
        t.setItemDelegate(BgDelegate(t))

    def clear(self): self.t.setRowCount(0)

    def put(self, d, tags=()):
        t = self.t; r = t.rowCount(); t.insertRow(r)
        tag = tags[0] if tags else None
        bg, fg, bold = TAG.get(tag, (None, None, False))
        for j, c in enumerate(self.ids):
            it = QTableWidgetItem(str(d.get(c, '') if d.get(c) is not None else ''))
            it.setTextAlignment(Qt.AlignVCenter | (Qt.AlignRight if c in ('desc', 'words', 'det') else Qt.AlignCenter))
            if bg: it.setBackground(QColor(bg))
            if fg: it.setForeground(QColor(fg))
            if bold: f = it.font(); f.setBold(True); it.setFont(f)
            t.setItem(r, j, it)


def qd(d): return QDate(d.year, d.month, d.day)
def pyd(q): return dt.date(q.year(), q.month(), q.day())


class DateDlg(QDialog):
    """شباك صغير: تاريخ (+ رقم فاتورة اختياري)"""

    def __init__(self, parent, title, head, date_label, optional=False, invoice=False, note=''):
        super().__init__(parent); self.setWindowTitle(title); self.setStyleSheet(CSS + f'QDialog {{ background: {HQ.BG}; }}')
        g = QGridLayout(self); g.setContentsMargins(18, 16, 18, 14); g.setVerticalSpacing(8)
        h = QLabel(head); h.setFont(HQ.font(13, True)); g.addWidget(h, 0, 0, 1, 2)
        self.c_d = QCheckBox(date_label) if optional else None
        self.e_d = QDateEdit(qd(dt.date.today())); self.e_d.setCalendarPopup(True); self.e_d.setDisplayFormat('dd-MM-yyyy')
        if optional:
            g.addWidget(self.c_d, 1, 0); self.e_d.setEnabled(False); self.c_d.toggled.connect(self.e_d.setEnabled)
        else: g.addWidget(QLabel(date_label), 1, 0)
        g.addWidget(self.e_d, 1, 1)
        self.e_i = QLineEdit() if invoice else None
        if invoice: g.addWidget(QLabel('Invoice No. – optional'), 2, 0); g.addWidget(self.e_i, 2, 1)
        if note:
            n = QLabel(note); n.setStyleSheet(f'color: {RED};'); n.setWordWrap(True); g.addWidget(n, 3, 0, 1, 2)
        bar = QHBoxLayout(); bar.addStretch(1); bar.addWidget(DQ.btn('Cancel', self.reject)); bar.addWidget(DQ.btn('Save', self.accept, True))
        g.addLayout(bar, 4, 0, 1, 2)

    def date(self):
        if self.c_d is not None and not self.c_d.isChecked(): return None
        d = pyd(self.e_d.date()); return dt.datetime(d.year, d.month, d.day)


class IpcQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        self.setObjectName('ipcq'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.p = None; self.calc = None; self.calc_for = None; self._up_no = None; self._lb_paths = []
        self.out_dir = qs_core.project_path(C.OUT_REL)
        self.template = qs_core.project_path(C.TEMPLATE_REL)
        self._build()
        QTimer.singleShot(150, self.reset_defaults)

    # ------------------------------------------------------------ الواجهة
    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 14, 22, 16); v.setSpacing(10)
        self.l_info = QLabel(''); self.l_info.setStyleSheet(f'color: {OKC};'); self.l_info.setWordWrap(True); self.l_info.setMaximumWidth(520)
        v.addLayout(header('IPC', 'payment certificates – the IPC, its summary and the take-off', HQ.COL.get('ipc', '#8E92F0'),
                           lambda: self.open_section('doc'),
                           [self.l_info, DQ.btn('Open IPC folder', lambda: self.open_path(self.out_dir)),
                            DQ.btn('Open template', lambda: self.open_path(self.template))]))
        # شريط الإدخال
        top = DQ.card(); th = QHBoxLayout(top); th.setContentsMargins(14, 10, 14, 10); th.setSpacing(8)
        th.addWidget(QLabel('IPC No.')); self.e_no = QSpinBox(); self.e_no.setRange(1, 200); th.addWidget(self.e_no)
        th.addSpacing(10); th.addWidget(QLabel('Period from'))
        self.e_from = QDateEdit(); self.e_from.setCalendarPopup(True); self.e_from.setDisplayFormat('dd-MM-yyyy'); self.e_from.setMinimumWidth(130); th.addWidget(self.e_from)
        th.addWidget(QLabel('to'))
        self.e_to = QDateEdit(qd(dt.date.today())); self.e_to.setCalendarPopup(True); self.e_to.setDisplayFormat('dd-MM-yyyy'); self.e_to.setMinimumWidth(130); th.addWidget(self.e_to)
        th.addSpacing(10); th.addWidget(QLabel('Plan %'))
        self.e_plan = QLineEdit(); self.e_plan.setMaximumWidth(70); th.addWidget(self.e_plan)
        self.l_plan = QLabel(''); self.l_plan.setStyleSheet('font-weight: bold;'); th.addWidget(self.l_plan)
        th.addStretch(1)
        th.addWidget(DQ.btn('Calculate', self.do_calc))
        self.b_make = DQ.btn('Make IPC (Excel)', self.do_make, True); th.addWidget(self.b_make)
        th.addWidget(DQ.btn('Update consultant', self.do_ho)); th.addWidget(DQ.btn('…', lambda: self.ho_path(change=True)))
        self.e_plan.textChanged.connect(self.plan_changed); self.e_no.valueChanged.connect(self.load_plan)
        v.addWidget(top)
        # المستخلص الجاي للرفع
        up = QFrame(); up.setObjectName('kpi'); up.setStyleSheet(f'QFrame#kpi {{ background: {"#3A3018" if D else "#FFF8E6"}; border-radius: 10px; }}')
        uh = QHBoxLayout(up); uh.setContentsMargins(12, 6, 12, 6)
        self.l_up = QLabel(''); self.l_up.setFont(HQ.font(12, True)); self.l_up.setWordWrap(True); uh.addWidget(self.l_up, 1)
        uh.addWidget(DQ.btn('Open upload folder', self.open_upload)); uh.addWidget(DQ.btn('Mark uploaded…', self.mark_uploaded))
        v.addWidget(up)
        # التابات
        c = DQ.card(); cv = QVBoxLayout(c); cv.setContentsMargins(14, 10, 14, 10)
        self.tabs = QTabWidget(); self.tabs.setDocumentMode(True); cv.addWidget(self.tabs)
        # 1) IPC
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
        b = QHBoxLayout(); b.addStretch(1); self.c_only = QCheckBox('Items with quantities only'); self.c_only.setChecked(True)
        self.c_only.toggled.connect(self.fill_ipc); b.addWidget(self.c_only); wv.addLayout(b)
        self.g_ipc = Grid(OLDI.IPC_COLS); wv.addWidget(self.g_ipc.t, 1); self.tabs.addTab(w, '  IPC  ')
        # 2) Summary
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
        self.l_head = QLabel(''); self.l_head.setFont(HQ.font(14, True)); self.l_head.setAlignment(Qt.AlignCenter)
        self.l_head.setStyleSheet(f'color: {INK}; background: {TT["chip"]}; border-radius: 8px; padding: 6px;'); wv.addWidget(self.l_head)
        self.g_amt = Grid(OLDI.AMT_COLS); self.g_amt.t.setMaximumHeight(230); wv.addWidget(self.g_amt.t)
        self.l_warn = QLabel(''); self.l_warn.setStyleSheet(f'color: {RED}; font-weight: bold;'); wv.addWidget(self.l_warn)
        self.g_div = Grid(OLDI.DIV_COLS); wv.addWidget(self.g_div.t, 1); self.tabs.addTab(w, '  Summary  ')
        # 3) QS Summary
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
        b = QHBoxLayout(); b.addStretch(1); self.c_sum_only = QCheckBox('Items with this-period qty only (as printed)'); self.c_sum_only.setChecked(True)
        self.c_sum_only.toggled.connect(self.fill_qsum); b.addWidget(self.c_sum_only); wv.addLayout(b)
        self.g_sum = Grid(OLDI.QSUM_COLS); wv.addWidget(self.g_sum.t, 1); self.tabs.addTab(w, '  QS Summary  ')
        # 4) QS Total
        w = QWidget(); self.qs_v = QVBoxLayout(w); self.qs_v.setContentsMargins(0, 8, 0, 0)
        self.l_qs = QLabel(''); self.l_qs.setStyleSheet(f'color: {INK}; font-weight: bold;'); self.l_qs.setAlignment(Qt.AlignRight)
        self.qs_v.addWidget(self.l_qs); self.g_qs = None; self.tabs.addTab(w, '  QS Total  ')
        # 5) Notes
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
        self.l_notes = QLabel(''); self.l_notes.setStyleSheet(f'color: {INK}; font-weight: bold;'); wv.addWidget(self.l_notes)
        self.g_notes = Grid(OLDI.NOTE_COLS, rtl=False); wv.addWidget(self.g_notes.t, 1)
        self.tab_notes = self.tabs.addTab(w, '  Notes  ')
        v.addWidget(c, 1)
        # المستخلصات اللي طلعت
        bc = DQ.card(); bv = QVBoxLayout(bc); bv.setContentsMargins(14, 8, 14, 8); bv.setSpacing(4)
        l = QLabel('IPCs made  (double-click to open)'); l.setStyleSheet(f'color: {MUTED}; font-weight: bold;'); bv.addWidget(l)
        self.lb = QListWidget(); self.lb.setMaximumHeight(110); self.lb.itemDoubleClicked.connect(self.open_selected); bv.addWidget(self.lb)
        lr = QHBoxLayout(); lr.addWidget(DQ.btn('Mark certified…', self.mark_certified))
        x = QLabel('select the IPC in the list first – the status goes to 11-LOG IPC'); x.setStyleSheet(f'color: {MUTED};'); lr.addWidget(x); lr.addStretch(1)
        bv.addLayout(lr)
        v.addWidget(bc)

    # ------------------------------------------------------------ مساعدات
    def info(self, t, col=None):
        self.l_info.setText(t); self.l_info.setStyleSheet(f'color: {col or OKC};')

    def open_path(self, p):
        try:
            if not os.path.exists(p): raise FileNotFoundError(p)
            os.startfile(p)
        except Exception as e:
            QMessageBox.critical(self, 'IPC', f'Cannot open:\n{e}')

    def project(self):
        p = qs_core.Project(qs_core.DEFAULT_INPUT); p.load_requests(); return p

    def ask(self, q): return QMessageBox.question(self, 'IPC', q) == QMessageBox.Yes

    def check_invoices(self):
        skip = self.__dict__.setdefault('_inv_skip', set())
        try:
            for no, b, ok, msg in C.scan_invoices(self.out_dir, ask=self.ask, skip=skip):
                if not ok: skip.add((no, b))
                self.info(msg)
        except PermissionError:
            self.info('11-LOG IPC is open – close it so the IPC approval is recorded', AMB)
        except Exception:
            traceback.print_exc()

    def list_files(self):
        self.lb.clear()
        files = C.ipc_files(self.out_dir) if os.path.isdir(self.out_dir) else {}
        if files: self.check_invoices()
        self.refresh_upload()
        try: st = C.ipc_status()
        except Exception: st = {}
        self._lb_paths = []
        for k in sorted(files):
            s = st.get(k, {}); tag = s.get('status') or ''
            if isinstance(s.get('cert'), (dt.date, dt.datetime)): tag += f"  {s['cert']:%d-%m-%Y}"
            if s.get('invoice'): tag += f"  |  Invoice {s['invoice']}"
            it = QListWidgetItem(os.path.relpath(files[k], self.out_dir) + (f'      —  {tag}' if tag else ''))
            if s.get('status') in C.CERTIFIED: it.setForeground(QColor(OKC))
            self.lb.addItem(it); self._lb_paths.append(files[k])
        return files

    def open_selected(self, *_):
        r = self.lb.currentRow()
        if 0 <= r < len(self._lb_paths): self.open_path(self._lb_paths[r])

    # ------------------------------------------------------------ الرفع على سيستم الجهة المالكة
    def refresh_upload(self):
        try:
            no = C.next_ipc(out_dir=self.out_dir)
            folder, res = C.upload_check(no, self.out_dir)
        except Exception as e:
            self._up_no = None; self.l_up.setText(f'Upload list: could not read 11-LOG IPC ({e})'); return
        self._up_no, self._up_dir = no, folder
        have = [k for k, f in res if f]; miss = [k for k, f in res if not f]
        txt = f'IPC-{no:02d} for upload (CLIENT system):  {len(have)}/{len(res)} ready'
        txt += ('   ·   missing: ' + ', '.join(miss)) if miss else '   ·   all there ✔ – upload it, then press «Mark uploaded»'
        self.l_up.setText(txt); self.l_up.setStyleSheet(f'color: {AMB if miss else OKC};')

    def open_upload(self):
        if self._up_no is None: return
        os.makedirs(self._up_dir, exist_ok=True); self.open_path(self._up_dir)

    def mark_uploaded(self):
        no = self._up_no
        if no is None: return
        _, res = C.upload_check(no, self.out_dir)
        miss = [k for k, f in res if not f]
        d = DateDlg(self, f'IPC No. {no}', f'IPC No. {no} was uploaded to the CLIENT system', 'Upload date',
                    note=('Not in the folder: ' + ', '.join(miss)) if miss else '')
        if not d.exec(): return
        try: ok = C.mark_uploaded(no, d.date())
        except PermissionError: return QMessageBox.critical(self, 'IPC', '11-LOG IPC is open in Excel – close it and try again.')
        except Exception as e: return QMessageBox.critical(self, 'IPC', f'Could not update 11-LOG IPC:\n{e}')
        self.refresh_upload()
        self.info(f'IPC No. {no} is Uploaded (11-LOG IPC)' if ok else f'IPC No. {no} is already past upload')

    def mark_certified(self):
        r = self.lb.currentRow()
        if r < 0 or r >= len(self._lb_paths): return QMessageBox.information(self, 'IPC', 'Select the IPC in the list first.')
        path = self._lb_paths[r]
        files = C.ipc_files(self.out_dir) if os.path.isdir(self.out_dir) else {}
        no = next((k for k, v in files.items() if v == path), None)
        if no is None: return
        d = DateDlg(self, f'IPC No. {no}', f'Mark IPC No. {no} as Certified', 'Certified date', optional=True, invoice=True)
        if not d.exec(): return
        try: C.mark_certified(no, d.date(), d.e_i.text().strip() or None)
        except PermissionError: return QMessageBox.critical(self, 'IPC', '11-LOG IPC is open in Excel – close it and try again.')
        except Exception as e: return QMessageBox.critical(self, 'IPC', f'Could not update 11-LOG IPC:\n{e}')
        self.list_files(); self.info(f'IPC No. {no} is Certified (11-LOG IPC)')

    def reset_defaults(self):
        files = self.list_files()
        self.e_no.setValue((max(files) + 1) if files else 1)
        start = None
        if files:
            try:
                _, prev_to = C.read_todate(files[max(files)])
                if prev_to: start = (prev_to.date() if isinstance(prev_to, dt.datetime) else prev_to) + dt.timedelta(days=1)
            except Exception:
                pass
        self.e_from.setDate(qd(start or dt.date.today()))
        self.e_to.setDate(qd(dt.date.today()))
        self.load_plan()
        if not os.path.exists(self.template): self.info('Template not found: 00_Registers\\11.1-TEMPLATE IPC.xlsx', RED)

    # ------------------------------------------------------------ نسبة التخطيط
    def _plans(self):
        import ipc_ho
        return ipc_ho.load_settings().get('PLAN', {})

    def load_plan(self, *_):
        v = self._plans().get(str(self.e_no.value()))
        self._loading = True; self.e_plan.setText('' if v is None else f'{v:g}'); self._loading = False
        self.show_plan()

    def plan_changed(self, *_):
        if getattr(self, '_loading', False): return
        import ipc_ho
        no = str(self.e_no.value()); txt = self.e_plan.text().replace('%', '').strip()
        st = ipc_ho.load_settings(); plans = st.setdefault('PLAN', {})
        if not txt: plans.pop(no, None)
        else:
            try: plans[no] = float(txt)
            except ValueError: self.show_plan(); return
        ipc_ho.save_settings(st); self.show_plan()

    def plan_check(self):
        try: pct = float(self.e_plan.text().replace('%', '').strip())
        except Exception: return 'none', None, None
        if pct <= 0: return 'none', None, None
        if not self.calc: return 'nocalc', None, None
        pv = self.calc['contract'] * pct / 100
        diff = (self.calc['todate'] - pv) / pv
        d = round(abs(diff) * 100, 2)
        return ('over' if d > 10 else 'ten' if d == 10 else 'ok'), diff, pv

    def show_plan(self):
        st, diff, pv = self.plan_check()
        if st == 'none': txt, fg = 'Enter Plan %', MUTED
        elif st == 'nocalc': txt, fg = 'Press Calculate', MUTED
        else:
            txt = f'= {money(pv)}  |  Diff {diff * 100:+.2f}%' + ('  ⛔' if st == 'over' else '')
            fg = {'ok': OKC, 'ten': AMB, 'over': RED}[st]
        self.l_plan.setText(txt); self.l_plan.setStyleSheet(f'color: {fg}; font-weight: bold;')
        if self.calc: self.fill_notes()

    def inputs(self):
        no = self.e_no.value(); d1, d2 = pyd(self.e_from.date()), pyd(self.e_to.date())
        if d2 < d1: raise ValueError('"To" date is before "From" date')
        return no, d1, d2

    # ------------------------------------------------------------ الحساب والعرض
    def do_calc(self):
        try:
            no, d1, d2 = self.inputs()
            QGuiApplication.setOverrideCursor(Qt.WaitCursor)
            self.p = self.project()
            self.calc = C.compute(self.p, no, d2, self.out_dir, self.template, d1)
            self.calc_for = (no, d1, d2)
        except Exception as e:
            traceback.print_exc(); QGuiApplication.restoreOverrideCursor(); QMessageBox.critical(self, 'IPC', str(e)); return False
        QGuiApplication.restoreOverrideCursor()
        self.fill_all()
        c = self.calc
        self.info(('Previous from: ' + os.path.basename(c['prev_file'])) if c['prev_file'] else 'No previous IPC – previous = 0')
        self.show_plan()
        return True

    def fill_all(self):
        self.fill_ipc(); self.fill_summary(); self.fill_qsum(); self.fill_qs(); self.fill_notes()

    def note_list(self):
        c = self.calc; out = []
        if not c: return out
        order = set()
        for x in c['rows']:
            for w in str(x.get('wirs') or '').split(','):
                w = w.strip()
                if w.isdigit(): order.add(int(w))
        info = {}
        for q in getattr(self.p, 'reqs_all', []) or []:
            try: info[int(float(str(q['no']).strip()))] = q
            except (ValueError, TypeError): pass
        umap = C.unifier_map(self.p) if self.p else {}
        try: uni = C.wir_pdfs()
        except Exception: uni = {}
        try: paper = C.wir_pdfs(qs_core.project_path(C.WIR_PAPER_REL))
        except Exception: paper = {}
        for n in sorted(order):
            q = info.get(n, {}); what = f"{q.get('type', '')}  {q.get('a_raw', '')}–{q.get('b_raw', '')}  {q.get('street', '')}"
            u = umap.get(n)
            if not u: out.append(('err', 'Unifier', str(n), f'No Unifier no. in WIR log (SW, column N) – {what}'))
            elif u not in uni: out.append(('err', 'Unifier PDF', f'{n}  (UN-{u:05d})', f'No file {u:05d}.pdf in 05_Quality\\WIR\\5_UNIFIER_APPROVED – {what}'))
            if n not in paper: out.append(('warn', 'Site copy', str(n), f'No approved site copy in 05_Quality\\WIR\\3_PUBLISHED – {what}'))
        for x in c.get('less_than_prev') or []:
            out.append(('err', 'Quantity', '', f'Item {x}: to-date quantity is less than the previous IPC'))
        st, diff, pv = self.plan_check()
        if st in ('over', 'ten'): out.append(('err' if st == 'over' else 'warn', 'Plan %', '', f'IPC vs planning update: {diff * 100:+.2f}%'))
        elif st == 'none': out.append(('warn', 'Plan %', '', 'Enter the cumulative Plan % before making the IPC'))
        for t in c.get('notes') or []: out.append(('warn', 'QS', '', str(t)))
        return out

    def fill_notes(self):
        g = self.g_notes; g.clear(); notes = self.note_list()
        for lev, topic, req, det in notes:
            g.put(dict(lev='Fix' if lev == 'err' else 'Check', topic=topic, req=req, det=det), (lev,))
        n_err = sum(1 for x in notes if x[0] == 'err')
        self.tabs.setTabText(4, f'  Notes ({len(notes)})  ' if notes else '  Notes  ')
        self.l_notes.setText(f'{n_err} to fix, {len(notes) - n_err} to check – resolve them before making the IPC' if notes
                             else 'No notes – ready to print')

    def fill_ipc(self, *_):
        g = self.g_ipc; g.clear()
        if not self.calc: return
        only = self.c_only.isChecked(); grand = 0
        for d in self.calc['divs']:
            rows = [x for x in self.calc['rows'] if x['div'] == d['no']]
            if only: rows = [x for x in rows if x['prev'] or x['todate']]
            if only and not rows: continue
            g.put(dict(item=str(d['no']), desc=f"{d['no']} {d['name']}"), ('div',))
            for x in rows:
                g.put(dict(item=x['item'], desc=x['desc'], unit=x['unit'], qty=q2(x['qty']), prev=q2(x['prev']), period=q2(x['period']),
                           todate=q2(x['todate']), rate=money(x['rate']), value=money(x['vtodate'])))
            g.put(dict(desc=f"إجمالي {d['name']}", value=money(d['todate'])), ('tot',))
            grand += d['todate']
        g.put(dict(desc='إجمالى قيمة الأعمال', value=money(grand)), ('grand',))

    def fill_summary(self):
        c = self.calc; self.g_amt.clear(); self.g_div.clear()
        if not c: return
        no, d1, d2 = self.calc_for
        self.l_head.setText(f'Interim Payment Certificate No. ({no})   |   Period {d1:%d-%m-%Y}  to  {d2:%d-%m-%Y}')
        for i, (lab, amt) in enumerate(c['amounts']):
            self.g_amt.put(dict(desc=lab, amt=money(amt), words=C.tafqit(amt)), ('tot',) if i == 5 else ())
        for d in c['divs']:
            self.g_div.put(dict(no=f"رقم ({d['no']})", desc=f"إجمالي {d['name']}", contract=money(d['contract']),
                                prev=money(d['prev']), period=money(d['period']), todate=money(d['todate'])))
        self.g_div.put(dict(desc='إجمالى قيمة الأعمال', contract=money(c['contract']), prev=money(c['prev']), period=money(c['period']),
                            todate=money(c['todate'])), ('tot',))
        pc = lambda v: f'{v / c["contract"] * 100:.2f}%' if c['contract'] else ''
        self.g_div.put(dict(desc='النسبة المئـــــــوية %', contract='100.00%', prev=pc(c['prev']), period=pc(c['period']),
                            todate=pc(c['todate'])), ('tot',))
        self.l_warn.setText('Warning: interim IPCs exceeded 90% of the contract value – the rest goes to the final IPC'
                            if c['pct'] > 0.9 else f"Interim IPCs to date: {c['pct'] * 100:.2f}% of contract value (limit 90%)")
        self.l_warn.setStyleSheet(f"color: {RED if c['pct'] > 0.9 else MUTED}; font-weight: bold;")

    def fill_qsum(self, *_):
        g = self.g_sum; g.clear(); c = self.calc
        if not c: return
        try: umap = C.unifier_map(self.p) if self.p else {}
        except Exception: umap = {}
        only = self.c_sum_only.isChecked(); tot = dict(vtodate=0.0, vperiod=0.0, vprev=0.0)
        for x in c['rows']:
            for k in tot: tot[k] += x[k]
            if only and abs(x['period']) <= 0.004: continue
            bad = x['period'] < 0 or (x['qty'] and x['todate'] > x['qty'])
            g.put(dict(item=x['item'], code=x['code'], desc=x['desc'], unit=C._unit_ar(x['unit']), qty=q2(x['qty']), prev=q2(x['prev']),
                       period=q2(x['period']), todate=q2(x['todate']), rate=money(x['rate']), vtodate=q2(x['vtodate']),
                       vperiod=q2(x['vperiod']), vprev=q2(x['vprev']), wirs=C.to_unifier(x['wirs'], umap) if x.get('wirs') else '',
                       chk='NOT OK' if bad else 'OK'), ('err',) if bad else ())
        vat = {k: round(v * C.VAT, 2) for k, v in tot.items()}
        g.put(dict(desc='القيمة الإجمالية', **{k: money(v) for k, v in tot.items()}), ('tot',))
        g.put(dict(desc='ضريبة القيمة المضافة 15 %', **{k: money(v) for k, v in vat.items()}), ('tot',))
        g.put(dict(desc='إجمالي قيمة الأعمال المنفذة بعد إضافة الضريبة', **{k: money(tot[k] + vat[k]) for k in tot}), ('grand',))

    def fill_qs(self):
        if not self.calc or not self.p: return
        try: hdr = qs_core.template_header(os.path.join(MODS, 'qs', 'QS-TOTAL_TEMPLATE.xlsx'))
        except Exception: hdr = None
        rows = self.p.summary_rows(self.calc['items'])
        if hdr:
            letters = hdr['letters']; names = {}
            for (r1, c1, r2, c2, txt, col) in hdr['cells']:
                for ci in range(c1, c2 + 1):
                    if ci - 1 < len(letters) and txt: names.setdefault(letters[ci - 1], []).append(str(txt).strip())
            cols = [(L, names.get(L, [L])[-1], max(60, int(w * 7))) for L, w in zip(letters, hdr['widths'])]
        else:
            letters = sorted({k for R in rows for k in R if isinstance(k, str) and k.isupper() and len(k) <= 2})
            cols = [(L, L, 80) for L in letters]
        if self.g_qs is not None: self.g_qs.t.setParent(None); self.g_qs.t.deleteLater()
        self.g_qs = Grid(cols, rtl=False); self.g_qs.t.horizontalHeader().setStretchLastSection(False)
        self.qs_v.addWidget(self.g_qs.t, 1)
        SUMC = ('J', 'K', 'V', 'W', 'X', 'Y', 'AA', 'AC', 'AD', 'AE', 'AF', 'AH', 'AJ', 'AN', 'AP', 'AR')
        tot = {c: 0.0 for c in SUMC}
        for R in rows:
            d = {}
            for c, _, _ in cols:
                v = R.get(c)
                if c in ('E', 'F'): v = qs_core.sta(v)
                elif c in ('AG', 'AI', 'AK', 'AM') and v is not None: v = f'{v * 100:.0f}%'
                elif isinstance(v, float): v = f'{v:,.3f}'
                d[c] = '' if v is None else v
                if c in tot and isinstance(R.get(c), float): tot[c] += R[c]
            self.g_qs.put(d)
        self.g_qs.put({c: ('الإجمالي' if c == 'B' else (f'{tot[c]:,.3f}' if c in tot else '')) for c, _, _ in cols}, ('tot',))
        no, d1, d2 = self.calc_for
        self.l_qs.setText(f'QS Total – approved requests up to {d2:%d-%m-%Y}  ({len(rows)} rows)')

    # ------------------------------------------------------------ عمل المستخلص
    def do_make(self):
        try: no, d1, d2 = self.inputs()
        except Exception as e: return QMessageBox.critical(self, 'IPC', str(e))
        try: certified = C.ipc_status().get(no, {}).get('status') in C.CERTIFIED
        except Exception: certified = False
        if certified: return QMessageBox.critical(self, 'IPC', f"IPC No. {no} is Certified ⛔\nAn approved IPC can't be made again.")
        if not self.calc or self.calc_for != (no, d1, d2):
            if not self.do_calc(): return
        fix = [x for x in self.note_list() if x[0] == 'err' and x[1] != 'Plan %']
        if fix and not self.ask(f'{len(fix)} note(s) still need fixing (see the Notes tab):\n\n'
                                + '\n'.join(f'• {x[1]} {x[2]}'.strip() for x in fix[:10]) + ('\n…' if len(fix) > 10 else '')
                                + '\n\nMake the IPC anyway?'):
            self.tabs.setCurrentIndex(4); return
        pst, pdiff, pv = self.plan_check()
        if pst == 'none': return QMessageBox.warning(self, 'IPC', "Enter the cumulative Plan % from the planning engineer's update first.")
        ptxt = f'IPC to date: {money(self.calc["todate"])}\nPlan ({self.e_plan.text().strip()}%): {money(pv)}\nDifference: {pdiff * 100:+.2f}%'
        if pst == 'over':
            return QMessageBox.critical(self, 'IPC', 'IPC blocked ⛔\nThe difference between the IPC and the planning update is more than 10%.\n\n' + ptxt)
        if pst == 'ten' and not self.ask('The difference between the IPC and the planning update is exactly 10%.\n\n' + ptxt + '\n\nIssue the IPC?'):
            return
        files = C.ipc_files(self.out_dir) if os.path.isdir(self.out_dir) else {}
        if no in files and not self.ask(f'IPC No. {no} already exists:\n{os.path.basename(files[no])}\n\nMake it again? (the old one is kept as OLD_)'):
            return
        c = self.calc
        if not self.ask(f'Make IPC No. {no}\nfrom {d1:%d-%m-%Y} to {d2:%d-%m-%Y}\n\nThis period: {money(c["period"])}\n'
                        f'Due incl. VAT: {money(c["total"])}\n\nContinue?'):
            return
        res = {}
        p, tmpl, out_dir, plan_txt = self.p, self.template, self.out_dir, self.e_plan.text()

        def work():
            try:
                wir_rows = None
                try:
                    import wir_core
                    wir_rows = wir_core.wir_quantities(p)
                except Exception:
                    traceback.print_exc()
                if no in files:
                    try: os.replace(files[no], os.path.join(os.path.dirname(files[no]), 'OLD_' + os.path.basename(files[no])))
                    except Exception: pass
                out, s = C.make_ipc(p, no, d1, d2, out_dir, tmpl, wir_rows=wir_rows, calc=c)
                try: C.log_plan(qs_core.project_path(C.LOG_REL), no, float(plan_txt.replace('%', '')), pv, pdiff)
                except PermissionError: s['log_error'] = '11-LOG IPC is open – Plan % was not recorded'
                except Exception: traceback.print_exc()
                res['ok'] = (out, s)
            except PermissionError:
                res['err'] = 'The file is open in Excel – close it and try again.'
            except Exception as e:
                traceback.print_exc(); res['err'] = f'Something went wrong:\n{e}'
        th = threading.Thread(target=work, daemon=True); th.start()
        self.b_make.setEnabled(False); t0 = time.time()

        def wait():
            if th.is_alive():
                self.info(f'Making IPC and PDFs… {int(time.time() - t0)} s (may take a few minutes)', AMB)
                QTimer.singleShot(500, wait); return
            self.b_make.setEnabled(True)
            self.raise_(); self.activateWindow()
            if 'err' in res:
                self.info('IPC not created', RED); return QMessageBox.critical(self, 'IPC', res['err'])
            self._made(no, *res['ok'])
        QTimer.singleShot(500, wait)

    def _made(self, no, out, s):
        self.list_files()
        msg = f"IPC No. {no} created\n\nDue: {money(s['total'])} SAR\n\n{os.path.basename(out)}"
        for k in ('qs_file', 'pdf', 'sum_pdf', 'qs_pdf'):
            if s.get(k): msg += f"\n{os.path.basename(s[k])}"
        wp = s.get('wir_pack')
        if wp:
            if wp.get('error'): msg += f"\n\n⚠ WIR PDF not created: {wp['error']}"
            else:
                if wp.get('file'): msg += f"\n{os.path.basename(wp['file'])}  ({len(wp['added'])} WIRs: Unifier + site copy)"
                if wp.get('missing'): msg += f"\n⚠ No Unifier copy (5_UNIFIER_APPROVED): {', '.join(map(str, wp['missing']))}"
                if wp.get('missing_paper'): msg += f"\n⚠ No site copy (3_PUBLISHED): {', '.join(map(str, wp['missing_paper']))}"
                if wp.get('bad'): msg += f"\n⚠ Could not open: {', '.join(wp['bad'])}"
        wz = s.get('wir_zips') or {}
        if wz.get('error'): msg += f"\n\n⚠ WIR zip files not created: {wz['error']}"
        elif wz.get('zips'):
            msg += f"\n\nMerged WIRs (zip) in 6_MERG_UN_SIT_WIR, copy in \"{os.path.basename(wz['copy_dir'])}\":"
            for f, nos in wz['zips'].items(): msg += f"\n   {f}: {len(nos)} WIRs ({', '.join(map(str, nos))})"
            if wz.get('bad'): msg += f"\n⚠ Could not open: {', '.join(wz['bad'])}"
        if s.get('pdf_error'): msg += f"\n\n⚠ PDF not created: {s['pdf_error']}"
        if s.get('log_error'): msg += f"\n\n⚠ {s['log_error']}"
        if s.get('less_than_prev'): msg += f"\n\n⚠ Items with less quantity than the previous IPC: {', '.join(map(str, s['less_than_prev']))}"
        self.info(f'{os.path.basename(out)} created')
        if self.ask(msg + '\n\nOpen the file?'): self.open_path(s.get('pdf') or out)

    # ------------------------------------------------------------ شيت الاستشاري
    def ho_path(self, change=False):
        import ipc_ho
        st = ipc_ho.load_settings(); p = st.get('HO_SHEET')
        if change:
            v, ok = QInputDialog.getText(self, 'Consultant sheet', 'Consultant sheet link (in Excel: File → Info → Copy path)\n'
                                         'or leave empty to pick a file:', text=p or '')
            if not ok: return None
            v = v.strip()
            if not v:
                v, _ = QFileDialog.getOpenFileName(self, 'Select consultant sheet', '', 'Excel (*.xlsx *.xlsm)')
                if not v: return None
            st['HO_SHEET'] = v; ipc_ho.save_settings(st); p = v
            self.info('Consultant sheet: ' + p.replace('\\', '/').split('/')[-1])
        return p

    def do_ho(self):
        import ipc_ho
        try: no, d1, d2 = self.inputs()
        except Exception as e: return QMessageBox.critical(self, 'IPC', str(e))
        if not self.calc or self.calc_for != (no, d1, d2):
            if not self.do_calc(): return
        try: certified = C.ipc_status().get(no, {}).get('status') in C.CERTIFIED
        except Exception: certified = False
        if certified and not self.ask(f'IPC No. {no} is already Certified.\n\nUpdate the consultant sheet for it anyway?'): return
        p = self.ho_path()
        if not p: return
        st = ipc_ho.load_settings(); loc = ipc_ho.full_path(st['LOCAL_SHEET']); both = ipc_ho.HO_ENABLED
        where = f'1) Our sheet: {loc}\n2) Consultant sheet: {p}' if both else f'Our sheet: {loc}\n(online consultant sheet will not be touched)'
        if not self.ask(f'Update the approved QS up to {d2:%d-%m-%Y} in:\n{where}\n\n'
                        '• QS sheet: manual cells and WIR numbers only – their formulas do the math\n'
                        '• Rows are added automatically if needed\n• IPC summary sheet: previous quantity and WIR numbers\n\n'
                        'The files must be closed. A backup is taken first. Continue?'):
            return
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            n, notes = ipc_ho.update_ho(self.p, self.calc, no, d1, d2, backup_dir=os.path.join(self.out_dir, 'HO_BACKUP'), ho=p)
        except Exception as e:
            traceback.print_exc(); QGuiApplication.restoreOverrideCursor(); return QMessageBox.critical(self, 'IPC', f'Update failed:\n{e}')
        QGuiApplication.restoreOverrideCursor()
        self.info(f'{"Both sheets updated" if both else "OUR_SHEET updated"} ({n} sections)')
        msg = (f'{n} sections written to our sheet and the consultant sheet; IPC summary updated.\nOneDrive will sync the change.' if both
               else f'{n} sections written to OUR_SHEET; IPC summary updated.')
        if notes: msg += '\n\nNotes:\n' + '\n'.join('• ' + x for x in notes[:15])
        QMessageBox.information(self, 'IPC', msg)
