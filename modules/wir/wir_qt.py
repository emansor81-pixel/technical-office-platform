# -*- coding: utf-8 -*-
"""WIR | Inspection Requests بـ Qt – نفس شاشة wir_app.py في الشغل (الكميات من wir_core)، والشكل الجديد الغامق بس.
كليك على خانة = تتنسخ.  كليك على «اترفع» = علامة اترفع على اليونيفاير.  اختار صفوف + Ctrl+C = تتنسخ كاملة."""
import os, sys, traceback, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
PLAT = os.path.dirname(os.path.dirname(HERE))
for p in (PLAT, os.path.join(PLAT, 'modules', 'dc')):
    if p not in sys.path: sys.path.insert(0, p)

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QGuiApplication, QKeySequence, QShortcut
from PySide6.QtWidgets import (QStyledItemDelegate, QStyle, QWidget, QLabel, QHBoxLayout, QVBoxLayout, QTableWidget, QTableWidgetItem, QAbstractItemView,
                               QComboBox, QMessageBox, QFileDialog, QMenu, QLineEdit)
import home_qt as HQ
import dc_qt as DQ
import wir_core as W
import wir_defs as OLDW                      # نفس الأعمدة والعلامات (wir_status.json) بالظبط

INK, MUTED, NAVY = DQ.INK, DQ.MUTED, DQ.NAVY
TT = HQ.T
CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#wirq')
COLS = OLDW.COLS
TAG_BG = ({'miss': '#3A3018', 'upl': '#1F3A24', 'pend': '#222E3A', 'ok': None} if HQ.DARK
          else {'miss': '#FFF2CC', 'upl': '#E2EFDA', 'pend': '#EEF3F8', 'ok': None})
NOUNF = '#FF8A8A' if HQ.DARK else '#C00000'
STATUSES = ['All', 'Uploaded – waiting approval', 'Not uploaded yet', 'Pending', 'Approved', 'No Unifier No.', 'Missing Levels']


class BgDelegate(QStyledItemDelegate):
    """لون خلفية الصف (الستايل الغامق بيتجاهل لون الخلية) – بيترسم قبل الكلام"""
    def paint(self, p, opt, idx):
        bg = idx.data(Qt.BackgroundRole)
        if bg is not None and not (opt.state & QStyle.StateFlag.State_Selected): p.fillRect(opt.rect, bg)
        super().paint(p, opt, idx)


def header(title, sub, color, back, extra=()):
    """رأس الشاشة زي SUB: ← Back + اسم القسم بلونه + زراير يمين"""
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


class WirQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        self.setObjectName('wirq'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.rows, self.view, self.marks = [], [], OLDW.load_marks()
        self._build()
        QTimer.singleShot(100, self.refresh)

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 14, 22, 16); v.setSpacing(12)
        self.l_info = QLabel(''); self.l_info.setStyleSheet(f'color: {DQ.GREEN};')
        v.addLayout(header('WIR', 'inspection requests – the quantity of every request, ready for Unifier', HQ.COL.get('wir', '#6CC070'),
                           lambda: self.open_section('doc'),
                           [self.l_info, DQ.btn('⟳  Refresh from WIR Log', self.refresh, True), DQ.btn('Export Excel', self.export)]))
        c = DQ.card(); cv = QVBoxLayout(c); cv.setContentsMargins(14, 12, 14, 12); cv.setSpacing(8)
        f = QHBoxLayout(); f.setSpacing(8)
        self.cb = {}
        for k, cap, w in (('code', 'Code', 130), ('st', 'Status', 230), ('line', 'Line', 110)):
            l = QLabel(cap); l.setStyleSheet(f'color: {MUTED};'); f.addWidget(l)
            cb = QComboBox(); cb.setMinimumWidth(w); cb.addItem('All'); cb.currentIndexChanged.connect(self.fill)
            f.addWidget(cb); self.cb[k] = cb
        self.cb['st'].clear(); self.cb['st'].addItems(STATUSES)
        self.q = QLineEdit(); self.q.setPlaceholderText('Search (request no., street, Unifier no.…)'); self.q.setMinimumWidth(260)
        self.q.textChanged.connect(self.fill); f.addWidget(self.q)
        f.addStretch(1)
        self.l_count = QLabel(''); self.l_count.setStyleSheet(f'color: {MUTED};'); f.addWidget(self.l_count)
        cv.addLayout(f)
        self.tv = QTableWidget(0, len(COLS))
        self.tv.setHorizontalHeaderLabels([t for _, t, _ in COLS])
        self.tv.verticalHeader().setVisible(False)
        self.tv.setSelectionBehavior(QAbstractItemView.SelectRows); self.tv.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tv.setEditTriggers(QAbstractItemView.NoEditTriggers); self.tv.setShowGrid(False); self.tv.setWordWrap(False)
        self.tv.verticalHeader().setDefaultSectionSize(28)
        for i, (_, _, w) in enumerate(COLS): self.tv.setColumnWidth(i, int(w * 1.1))
        self.tv.horizontalHeader().setHighlightSections(False)
        self.tv.setItemDelegate(BgDelegate(self.tv))
        self.tv.cellClicked.connect(self.on_click)
        self.tv.setContextMenuPolicy(Qt.CustomContextMenu); self.tv.customContextMenuRequested.connect(self.menu)
        QShortcut(QKeySequence.Copy, self.tv, activated=lambda: self.copy(True))
        cv.addWidget(self.tv, 1)
        self.l_det = QLabel('Click a cell to copy it.   Click «اترفع» to mark a request uploaded.   Select rows + Ctrl+C to copy whole rows.')
        self.l_det.setFont(HQ.font(12.5, True)); self.l_det.setStyleSheet(f'color: {INK}; background: {TT["chip"]}; border-radius: 8px; padding: 8px 12px;')
        self.l_det.setTextInteractionFlags(Qt.TextSelectableByMouse)
        cv.addWidget(self.l_det)
        v.addWidget(c, 1)

    # ------------------------------------------------------------ data
    def refresh(self):
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.p = W.load()
            self.rows = W.wir_quantities(self.p)
        except Exception as e:
            if self.log: self.log(traceback.format_exc())
            QMessageBox.critical(self, 'WIR', f'Cannot read data:\n{e}'); return
        finally:
            QGuiApplication.restoreOverrideCursor()
        for k, vals in (('code', sorted({r['code'] for r in self.rows})), ('line', sorted({str(r['line']) for r in self.rows if r['line']}))):
            cb = self.cb[k]; cur = cb.currentText()
            cb.blockSignals(True); cb.clear(); cb.addItems(['All'] + vals)
            i = cb.findText(cur); cb.setCurrentIndex(max(i, 0)); cb.blockSignals(False)
        self.fill()
        self.l_info.setText(f'{len(self.rows)} requests  ·  refreshed {dt.datetime.now():%H:%M}')

    def visible(self):
        c, s, l = self.cb['code'].currentText(), self.cb['st'].currentText(), self.cb['line'].currentText()
        q = self.q.text().strip().lower()
        out = []
        for r in self.rows:
            if c != 'All' and r['code'] != c: continue
            if l != 'All' and str(r['line']) != l: continue
            if s == 'Pending' and r['status'] != 'Pending': continue
            if s == 'Uploaded – waiting approval' and not (r.get('upd') and r['status'] == 'Pending'): continue
            if s == 'Not uploaded yet' and (r.get('upd') or r['status'] != 'Pending'): continue
            if s == 'Approved' and r['status'] != 'Approved': continue
            if s == 'No Unifier No.' and r['unf']: continue
            if s == 'Missing Levels' and not r.get('miss'): continue
            if q and q not in ' '.join(str(r.get(k) or '') for k in ('no', 'unf', 'street', 'dist', 'line', 'title', 'code')).lower(): continue
            out.append(r)
        return out

    def fill(self):
        if not hasattr(self, 'tv'): return
        for r in self.rows: r['upd'] = self.marks.get(OLDW.ukey(r))
        self.view = self.visible()
        t = self.tv; t.setUpdatesEnabled(False); t.setRowCount(len(self.view))
        for i, r in enumerate(self.view):
            c = OLDW.cells(r)
            tag = 'miss' if r.get('miss') else ('upl' if r.get('upd') and r['status'] == 'Pending' else ('pend' if r['status'] == 'Pending' else 'ok'))
            for j, (k, _, _) in enumerate(COLS):
                it = QTableWidgetItem(c[k][0])
                it.setTextAlignment(Qt.AlignVCenter | (Qt.AlignRight if k in ('title', 'note', 'desc', 'street', 'dist') else Qt.AlignCenter))
                if TAG_BG[tag]: it.setBackground(QColor(TAG_BG[tag]))
                if r['status'] == 'Approved' and not r['unf']: it.setForeground(QColor(NOUNF))
                if k == 'up': it.setForeground(QColor(DQ.GREEN if r.get('upd') else MUTED))
                t.setItem(i, j, it)
        t.setUpdatesEnabled(True)
        self.l_count.setText(f'{len(self.view)} shown')

    # ------------------------------------------------------------ actions
    def on_click(self, row, col):
        if row < 0 or row >= len(self.view): return
        r, cid = self.view[row], COLS[col][0]
        if cid == 'up': return self.toggle(r)
        txt = OLDW.cells(r)[cid][1]
        if txt == '': return
        QGuiApplication.clipboard().setText(txt)
        self.l_det.setText(f"Copied  ✔   {r['title'] or r['no']}   |   {OLDW.COLS_T[cid]}:   {txt}")

    def toggle(self, r):
        k = OLDW.ukey(r)
        if self.marks.get(k): self.marks.pop(k)
        else: self.marks[k] = dt.date.today().strftime('%d-%m-%Y')
        try: OLDW.save_marks(self.marks)
        except Exception as e:
            QMessageBox.critical(self, 'WIR', f'Could not save the mark:\n{e}'); return
        sv, sh = self.tv.verticalScrollBar().value(), self.tv.horizontalScrollBar().value()
        self.fill()
        self.tv.verticalScrollBar().setValue(sv); self.tv.horizontalScrollBar().setValue(sh)
        self.l_det.setText(f"{'Uploaded ✔  ' + self.marks[k] if k in self.marks else 'Mark removed'}   |   {r['title'] or r['no']}")

    def _sel(self):
        return sorted({i.row() for i in self.tv.selectedIndexes()})

    def copy(self, row):
        rows = self._sel()
        if not rows: return
        lines = []
        for i in rows:
            c = OLDW.cells(self.view[i])
            lines.append('\t'.join(c[k][1] for k, _, _ in COLS) if row else c['qty'][1])
        QGuiApplication.clipboard().setText('\n'.join(lines))
        self.l_det.setText(f'Copied  ✔   {len(lines)} row(s)' + ('' if row else '  (quantity)'))

    def menu(self, pos):
        r = self.tv.rowAt(pos.y())
        if r >= 0 and r not in self._sel(): self.tv.selectRow(r)
        m = QMenu(self)
        m.addAction('Copy quantity', lambda: self.copy(False))
        m.addAction('Copy row(s)', lambda: self.copy(True))
        m.exec(self.tv.viewport().mapToGlobal(pos))

    def export(self):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        out, _ = QFileDialog.getSaveFileName(self, 'Export WIR quantities',
                                             os.path.join(W.core.QS_DIR, f'WIR_Quantities_{dt.date.today():%Y-%m-%d}.xlsx'), 'Excel (*.xlsx)')
        if not out: return
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'WIR Quantities'
        thin = Side(style='thin', color='808080'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
        for j, (_, t, w) in enumerate(COLS, 1):
            c = ws.cell(1, j, t); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=OLDW.NAVY.lstrip('#')); c.border = B
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = max(8, w / 7)
        for i, r in enumerate(self.view, 2):
            c = OLDW.cells(r)
            for j, (k, _, _) in enumerate(COLS, 1):
                v = c[k][2]
                cell = ws.cell(i, j, v); cell.border = B
                if isinstance(v, float): cell.number_format = '#,##0.000'
                if isinstance(v, dt.date): cell.number_format = 'dd-mm-yyyy'
                if r.get('miss'): cell.fill = PatternFill('solid', fgColor='FFF2CC')
        ws.freeze_panes = 'A2'
        try: wb.save(out)
        except PermissionError:
            QMessageBox.critical(self, 'WIR', 'The file is open in Excel. Close it and try again.'); return
        if QMessageBox.question(self, 'WIR', f'Exported:\n{out}\n\nOpen it now?') == QMessageBox.Yes: os.startfile(out)
