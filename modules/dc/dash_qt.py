# -*- coding: utf-8 -*-
"""Document Center Dashboard (طلب المستخدم 21-11-2026) – كارت «Dashboard» في خريطة Document Center.

كل ورق المشروع في نظرة واحدة: INBOX، مراحل الـ submittals، الـ MAR اللي مستني رد، واللي مستنيك.
قراية بس من الفولدرات وسجل MAR+MIR (dash_data.py) – مفيش حاجة بتتكتب.
"""
import os, sys, threading

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QHBoxLayout, QVBoxLayout, QGridLayout,
                               QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QMessageBox)

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import dc_core as D
import dc_qt as DQ
import dash_data as X
from dc_io import LOG

PLAT = os.path.dirname(os.path.dirname(HERE))
if PLAT not in sys.path: sys.path.insert(0, PLAT)
import home_qt as HQ

BG, CARD, INK, MUTED, SOFT = HQ.BG, HQ.CARD, HQ.INK, HQ.MUTED, HQ.SOFT_LINE
TT = HQ.T
CARD2 = TT.get('card2', CARD)
GREEN = '#7FD98A' if HQ.DARK else '#2C7536'
AMBER = '#F0C860' if HQ.DARK else '#8A6100'
REDC = '#FF8A8A' if HQ.DARK else '#C00000'
BLUE = '#5EA3E6' if HQ.DARK else '#1F5C99'
TEAL = '#3CC2B4' if HQ.DARK else '#1C7C72'
GREY = '#3A4450' if HQ.DARK else '#C9D1DA'
ACCENT = '#7FB3E8' if HQ.DARK else '#1F5C99'
STAGE_COL = (GREY, AMBER, GREEN, '#5D6A78' if HQ.DARK else '#9AA5B4')      # WIP / Shared / Published / Archived

CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#dash') + f"""
QFrame#kpi, QFrame#box {{ background: {CARD2}; border: 1px solid {SOFT}; border-radius: 10px; }}
QLabel#sec {{ color: {ACCENT}; font: bold 9.5pt "Segoe UI"; }}
QLabel#k {{ color: {MUTED}; font: 10pt "Segoe UI"; }}
QTableWidget {{ background: {CARD2}; color: {INK}; border: none; font: 10.5pt "Segoe UI"; }}
QTableWidget::item {{ border-bottom: 1px solid {SOFT}; padding-left: 6px; }}
QHeaderView::section {{ background: {CARD2}; color: {MUTED}; border: none; border-bottom: 1px solid {SOFT};
                        padding: 6px; font: bold 8.5pt "Segoe UI"; }}
"""
OLD_COLS = [('NO.', 60), ('SUBMITTAL', 170), ('MATERIAL', 300), ('SUPPLIER', 190), ('SENT', 95), ('DAYS', 60)]


class DashQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None, embedded=False):
        """embedded=True ← تابة Overview جوه Documents: من غير زرار الرجوع، وبيقرا أول ما التابة تتفتح."""
        super().__init__()
        self.embedded = embedded; self.loaded = False
        self.setObjectName('dash'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section = open_section or (lambda k: None)
        self.run_old, self.log = run_old, log
        self.root = D.project_root()
        self.busy = False
        self._build()
        if not embedded: QTimer.singleShot(60, self.refresh)

    # ------------------------------------------------------------ الشكل
    def _build(self):
        v = QVBoxLayout(self); v.setSpacing(12)
        v.setContentsMargins(*((0, 0, 0, 0) if self.embedded else (22, 14, 22, 16)))
        t = QHBoxLayout(); t.setSpacing(14)
        tv = QVBoxLayout(); tv.setSpacing(0)
        if not self.embedded:
            b = DQ.btn('←  Document Center', lambda: self.open_section('doc')); b.setObjectName('back'); t.addWidget(b)
            c = QLabel('Technical Office Platform  ›  Document Center  ›  Dashboard'); c.setFont(HQ.font(12.5, True))
            c.setStyleSheet(f'color: {MUTED};'); tv.addWidget(c)
        self.l_title = QLabel(); tv.addWidget(self.l_title); self._title('reading…')
        t.addLayout(tv, 1)
        t.addWidget(DQ.btn('Open 04_Submittals', lambda: self.open_path(os.path.join(self.root, '04_Submittals'))))
        t.addWidget(DQ.btn('MAR+MIR log', lambda: self.open_path(os.path.join(self.root, '00_Registers', '02-LOG MAR+MIR.xlsm'))))
        self.b_ref = DQ.btn('↻  Refresh', self.refresh, True); t.addWidget(self.b_ref)
        v.addLayout(t)

        # الأرقام
        kr = QHBoxLayout(); kr.setSpacing(11); self.kpi = {}
        for key, label, col in (('inbox', 'Papers waiting in 00_INBOX', AMBER), ('mar', 'MAR in the log', BLUE),
                                ('mir', 'MIR in the log', TEAL), ('late', f'MAR over {X.OVERDUE_DAYS} days, no reply', REDC),
                                ('wip', 'Submittals in preparation (1_WIP)', '#D2BC4A' if HQ.DARK else '#7A6A10')):
            f = QFrame(); f.setObjectName('kpi'); fl = QVBoxLayout(f); fl.setContentsMargins(14, 11, 14, 12); fl.setSpacing(2)
            n = QLabel('–'); n.setStyleSheet(f'color: {col}; font: bold 24pt "Segoe UI";')
            l = QLabel(label); l.setObjectName('k')
            s = QLabel(''); s.setStyleSheet(f'color: {MUTED}; font: 9pt "Segoe UI";'); s.setWordWrap(True)
            fl.addWidget(n); fl.addWidget(l); fl.addWidget(s)
            self.kpi[key] = (n, s); kr.addWidget(f, 1)
        v.addLayout(kr)

        row = QHBoxLayout(); row.setSpacing(12)
        # مراحل الـ submittals
        b1 = QFrame(); b1.setObjectName('box'); l1 = QVBoxLayout(b1); l1.setContentsMargins(15, 12, 15, 13); l1.setSpacing(9)
        l1.addWidget(self._sec('SUBMITTALS BY STAGE  ·  from the folders'))
        self.g_st = QGridLayout(); self.g_st.setHorizontalSpacing(10); self.g_st.setVerticalSpacing(10)
        self.g_st.setColumnStretch(1, 1)
        l1.addLayout(self.g_st)
        lg = QHBoxLayout(); lg.setSpacing(14)
        for txt, col in zip(X.STAGE_EN, STAGE_COL):
            q = QLabel(f'<span style="color:{col};">■</span>&nbsp;{txt}'); q.setStyleSheet(f'color: {MUTED}; font: 9pt "Segoe UI";')
            lg.addWidget(q)
        lg.addStretch(1); l1.addLayout(lg); l1.addStretch(1)
        row.addWidget(b1, 5)
        # أقدم MAR من غير رد
        b2 = QFrame(); b2.setObjectName('box'); l2 = QVBoxLayout(b2); l2.setContentsMargins(15, 12, 15, 13); l2.setSpacing(8)
        l2.addWidget(self._sec('MAR SENT – NO REPLY YET  ·  oldest first'))
        self.tv = QTableWidget(0, len(OLD_COLS)); self.tv.setHorizontalHeaderLabels([c[0] for c in OLD_COLS])
        self.tv.verticalHeader().setVisible(False); self.tv.setShowGrid(False)
        self.tv.setSelectionBehavior(QAbstractItemView.SelectRows); self.tv.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tv.verticalHeader().setDefaultSectionSize(32)
        hh = self.tv.horizontalHeader(); hh.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter); hh.setHighlightSections(False)
        for i, (_, w) in enumerate(OLD_COLS): self.tv.setColumnWidth(i, w)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        l2.addWidget(self.tv, 1)
        row.addWidget(b2, 8)
        # اللي مستنيك
        b3 = QFrame(); b3.setObjectName('box'); l3 = QVBoxLayout(b3); l3.setContentsMargins(15, 12, 15, 13); l3.setSpacing(8)
        l3.addWidget(self._sec('WAITING FOR YOU'))
        self.l_todo = QLabel(''); self.l_todo.setWordWrap(True); self.l_todo.setTextFormat(Qt.RichText)
        self.l_todo.setStyleSheet(f'color: {INK}; font: 10.5pt "Segoe UI";'); self.l_todo.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        l3.addWidget(self.l_todo, 1)
        row.addWidget(b3, 4)
        v.addLayout(row, 1)

    def _sec(self, t):
        l = QLabel(t); l.setObjectName('sec'); return l

    def first_show(self):
        if not self.loaded: self.refresh()

    def _title(self, sub):
        if self.embedded:
            return self.l_title.setText(f'<span style="color:{GREEN}; font-size:13pt; font-weight:bold;">Overview</span>'
                                        f'<span style="color:{MUTED}; font-size:10.5pt;">   every paper of the project at a glance  ·  {sub}</span>')
        self.l_title.setText(f'<span style="color:{GREEN}; font-size:20pt; font-weight:bold;">Dashboard</span>'
                             f'<span style="color:{MUTED}; font-size:11pt;">   every paper of the project at a glance  ·  {sub}</span>')

    # ------------------------------------------------------------ القراية (في thread – السجل بياخد ثواني)
    def refresh(self):
        if self.busy: return
        self.busy = True; self.loaded = True; self.b_ref.setEnabled(False); self._title('reading…')
        res = {}
        review, claude = HQ.inbox_dirs(os.path.join(self.root, '00_INBOX'))
        exts = getattr(HQ.HM, 'DOC_EXT', ('.pdf',))

        def work():
            try:
                res['d'] = X.collect(self.root, review, claude, exts, load_marmir=lambda: D._mat().load_marmir())
            except Exception as e:
                LOG.exception('Dashboard read failed'); res['e'] = e

        th = threading.Thread(target=work, daemon=True); th.start()

        def wait():
            if th.is_alive(): return QTimer.singleShot(200, wait)
            self.busy = False
            try: self.b_ref.setEnabled(True)
            except RuntimeError: return                  # الشاشة اتقفلت وهو بيقرا
            if 'e' in res: return self._title(f'could not read: {res["e"]}')
            try: self.show_data(res['d'])
            except Exception as e:
                LOG.exception('Dashboard show failed'); self._title(f'could not show: {e}')

        QTimer.singleShot(200, wait)

    # ------------------------------------------------------------ العرض
    def show_data(self, d):
        top, rv, c_ready, c_all = d['inbox']
        stages = d['stages']; mar, mir = d.get('mar'), d.get('mir')
        wip = sum(s[1][0] for s in stages)

        def kpi(key, n, sub):
            self.kpi[key][0].setText(str(n)); self.kpi[key][1].setText(sub)

        kpi('inbox', top + rv, f'{rv} need your choice (review)' if rv else 'press Distribute to read them' if top else 'nothing waiting')
        if mar:
            kpi('mar', mar['total'], f"{mar['replied']} replied  ·  {mar['waiting']} waiting  ·  {mar['not_sent']} not sent")
            kpi('late', mar['overdue'], 'PMM: reply is due in 14 days' if mar['overdue'] else 'none – all inside 14 days')
        else:
            kpi('mar', '?', 'log not read'); kpi('late', '?', 'log not read')
        if mir:
            ld = f"  ·  {mir['last_date']:%d-%m}" if mir.get('last_date') else ''
            kpi('mir', mir['total'], f"last: {mir['last']}{ld}\n{mir['waiting']} without approval date")
        else:
            kpi('mir', '?', 'log not read')
        kpi('wip', wip, '  ·  '.join(f'{n} {c[0]}' for n, c in stages if c[0]) or 'nothing in 1_WIP')

        # المراحل
        while self.g_st.count():
            w = self.g_st.takeAt(0).widget()
            if w: w.deleteLater()
        for r, (name, c) in enumerate(stages):
            nm = QLabel(name); nm.setStyleSheet(f'color: {INK}; font: 10.5pt "Segoe UI";'); nm.setMinimumWidth(52)
            self.g_st.addWidget(nm, r, 0)
            bar = QWidget(); bl = QHBoxLayout(bar); bl.setContentsMargins(0, 0, 0, 0); bl.setSpacing(1)
            if sum(c):
                for n, col in zip(c, STAGE_COL):
                    if not n: continue
                    f = QFrame(); f.setFixedHeight(11); f.setStyleSheet(f'background: {col}; border-radius: 3px;')
                    f.setToolTip(str(n)); bl.addWidget(f, n)
            else:
                f = QFrame(); f.setFixedHeight(11); f.setStyleSheet(f'background: {BG}; border-radius: 3px;'); bl.addWidget(f, 1)
            self.g_st.addWidget(bar, r, 1)
            cnt = QLabel('  '.join(f'<span style="color:{col};">{n}</span>' for n, col in zip(c, STAGE_COL)))
            cnt.setStyleSheet('font: 9.5pt "Segoe UI";'); cnt.setAlignment(Qt.AlignRight | Qt.AlignVCenter); cnt.setMinimumWidth(92)
            self.g_st.addWidget(cnt, r, 2)

        # أقدم MAR
        self.tv.setRowCount(0)
        for w in (mar['oldest'] if mar else [])[:40]:
            i = self.tv.rowCount(); self.tv.insertRow(i)
            late = w['days'] > X.OVERDUE_DAYS
            vals = (w['no'], w['ref'], w['desc'], w['supplier'], f"{w['sent']:%d-%m-%Y}", str(w['days']))
            for cix, val in enumerate(vals):
                it = QTableWidgetItem(str(val))
                if cix == 5: it.setForeground(QColor(REDC if late else AMBER))
                self.tv.setItem(i, cix, it)

        # اللي مستنيك
        lines = []

        def add(col, mark, text):
            lines.append(f'<p style="margin:0 0 9px 0;"><span style="color:{col}; font-weight:bold;">{mark}</span>&nbsp;&nbsp;{text}</p>')

        if rv: add(AMBER, str(rv), 'paper(s) in INBOX need your choice (review folder)')
        if top: add(AMBER, str(top), 'paper(s) in INBOX not distributed yet')
        if mar and mar['overdue']: add(REDC, str(mar['overdue']), f'MAR over {X.OVERDUE_DAYS} days with no reply – follow up')
        if mar and mar['not_sent']: add(MUTED, str(mar['not_sent']), 'MAR in the log with no sent date')
        if c_all: add(MUTED, f'{c_ready}/{c_all}', 'Claude folder – read and ready to distribute')
        if mir and mir['waiting']: add(MUTED, str(mir['waiting']), 'MIR without an approval date in the log')
        for e in d.get('errors', []): add(REDC, '!', str(e).replace('<', '&lt;'))
        self.l_todo.setText(''.join(lines) or f'<span style="color:{GREEN};">Nothing waiting ✔</span>')
        self._title(f"updated {d['when']:%d-%m-%Y  %H:%M}")

    # ------------------------------------------------------------ helpers
    def open_path(self, p):
        try:
            if os.name == 'nt': os.startfile(p)
            else:
                import subprocess; subprocess.Popen(['xdg-open', p])
        except Exception as e:
            QMessageBox.warning(self, 'Dashboard', str(e))

    def on_close(self): return True
