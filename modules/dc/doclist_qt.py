# -*- coding: utf-8 -*-
"""Documents – صفحة القايمة: كل ورق النوع + بحث وفلاتر + معاينة + زرار «＋ New» (طلب المستخدم 21-11-2026).

صفحة واحدة بتتملى حسب التابة المختارة: show(spec) – spec فيه الأعمدة ودالة القراية (بتشتغل في thread).
"""
import os, sys, threading

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QHBoxLayout, QVBoxLayout, QTableWidget, QTableWidgetItem,
                               QHeaderView, QAbstractItemView, QLineEdit, QPushButton, QMessageBox)

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import dc_qt as DQ
import doclist_data as X
from dc_io import LOG
try:
    from dc_preview import PreviewPane
except Exception:
    PreviewPane = None

PLAT = os.path.dirname(os.path.dirname(HERE))
if PLAT not in sys.path: sys.path.insert(0, PLAT)
import home_qt as HQ

INK, MUTED, SOFT = HQ.INK, HQ.MUTED, HQ.SOFT_LINE
KIND_COL = {'ok': '#7FD98A' if HQ.DARK else '#2C7536', 'wait': '#F0C860' if HQ.DARK else '#8A6100',
            'bad': '#FF8A8A' if HQ.DARK else '#C00000', 'dim': MUTED}
FILTERS = [('all', 'All'), ('open', 'Open'), ('delivered', 'Delivered'), ('waiting', 'Waiting'), ('unsigned', 'Waiting signature'),
           ('signed', 'Signed'), ('overdue', f'Over {X.OVERDUE_DAYS} days'), ('approved', 'Approved'), ('unifier', 'Not on Unifier')]

CSS = f"""
QPushButton#chip {{ background: {HQ.T.get('card2', HQ.CARD)}; color: {MUTED}; border: 1px solid {SOFT}; border-radius: 12px;
                    padding: 4px 11px; font: 9.5pt "Segoe UI"; }}
QPushButton#chip:checked {{ background: #1E3448; border-color: #3D7DB8; color: #9CC8F0; }}
QLineEdit#find {{ background: {HQ.BG}; color: {INK}; border: 1px solid {SOFT}; border-radius: 7px; padding: 6px 9px; font: 10.5pt "Segoe UI"; }}
"""


class ListPage(QWidget):
    def __init__(self, on_new=None, on_edit=None, on_deliver=None, parent=None):
        super().__init__(parent)
        self.setStyleSheet(CSS)
        self.on_new = on_new or (lambda: None)
        self.on_edit = on_edit or (lambda row: None)         # «Edit» على الصف المختار (لو مش مقفول)
        self.on_deliver = on_deliver or (lambda row: None)   # «Delivered / Undo delivered» (الخطابات بس)
        self.spec, self.rows, self.view = None, [], []
        self.cache, self.loading = {}, set()
        self.flag = 'all'
        self._build()

    # ------------------------------------------------------------ الشكل
    def _build(self):
        h = QHBoxLayout(self); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(12)
        c = DQ.card(); v = QVBoxLayout(c); v.setContentsMargins(12, 10, 12, 12); v.setSpacing(9)
        bar = QHBoxLayout(); bar.setSpacing(8)
        self.b_new = DQ.btn('＋  New', lambda: self.on_new(), True); bar.addWidget(self.b_new)
        self.b_edit = DQ.btn('✎  Edit', self._edit); bar.addWidget(self.b_edit)
        self.b_del = DQ.btn('✔  Delivered', self._deliver); bar.addWidget(self.b_del)
        self.b_unf = DQ.btn('☁  Unifier ✔', self._unifier); bar.addWidget(self.b_unf); self.b_unf.hide()
        self.e_find = QLineEdit(); self.e_find.setObjectName('find'); self.e_find.setPlaceholderText('Search…')
        self.e_find.setFixedWidth(190); self.e_find.textChanged.connect(lambda *_: self.fill()); bar.addWidget(self.e_find)
        self.chips = {}
        for key, text in FILTERS:
            b = QPushButton(text); b.setObjectName('chip'); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, k=key: self.set_flag(k)); bar.addWidget(b); self.chips[key] = (b, text)
        bar.addStretch(1)
        self.l_sum = QLabel(''); self.l_sum.setStyleSheet(f'color: {MUTED}; font: 10pt "Segoe UI";'); bar.addWidget(self.l_sum)
        bar.addWidget(DQ.btn('↻', self.reload))
        v.addLayout(bar)
        self.tv = QTableWidget(0, 0)
        self.tv.verticalHeader().setVisible(False); self.tv.setShowGrid(False)
        self.tv.setSelectionBehavior(QAbstractItemView.SelectRows); self.tv.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tv.setEditTriggers(QAbstractItemView.NoEditTriggers); self.tv.setWordWrap(False)
        self.tv.verticalHeader().setDefaultSectionSize(34)
        hh = self.tv.horizontalHeader(); hh.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter); hh.setHighlightSections(False)
        self.tv.itemSelectionChanged.connect(self.show_sel)
        self.tv.itemDoubleClicked.connect(lambda *_: self.open_file())
        v.addWidget(self.tv, 1)
        self.l_note = QLabel(''); self.l_note.setStyleSheet(f'color: {MUTED}; font: 9.5pt "Segoe UI";'); self.l_note.setWordWrap(True)
        v.addWidget(self.l_note)
        h.addWidget(c, 1)
        # المعاينة
        right = QVBoxLayout(); right.setSpacing(8)
        self.pv = PreviewPane(self, width=420) if PreviewPane else None
        if self.pv: right.addWidget(self.pv, 1)
        rb = QHBoxLayout(); rb.setSpacing(8)
        rb.addWidget(DQ.btn('Open file', self.open_file)); rb.addWidget(DQ.btn('Open folder', self.open_folder)); rb.addStretch(1)
        right.addLayout(rb)
        if not self.pv: right.addStretch(1)
        h.addLayout(right)

    # ------------------------------------------------------------ التابة
    def open(self, spec):
        """spec = dict(key, new, cols, fetch, filters=(…), note='', folder='', can_new=True)"""
        self.spec = spec
        self.b_new.setText(f"＋  {spec.get('new', 'New')}")
        self.b_new.setVisible(bool(spec.get('can_new', True)))
        self.b_del.setVisible(bool(spec.get('deliver')))
        self.b_unf.setVisible(bool(spec.get('unifier')))
        cols = spec['cols']
        self.tv.setRowCount(0); self.tv.setColumnCount(len(cols))
        self.tv.setHorizontalHeaderLabels([c[1] for c in cols])
        hh = self.tv.horizontalHeader()
        for i, (_, _, w) in enumerate(cols):
            hh.setSectionResizeMode(i, QHeaderView.Interactive); self.tv.setColumnWidth(i, w)
        hh.setSectionResizeMode(spec.get('stretch', 2 if len(cols) > 2 else 0), QHeaderView.Stretch)
        want = set(spec.get('filters', ('all',)))
        for k, (b, _) in self.chips.items(): b.setVisible(k in want and len(want) > 1)
        self.flag = 'all'; self.e_find.blockSignals(True); self.e_find.clear(); self.e_find.blockSignals(False)
        self.l_note.setText(spec.get('note', ''))
        if self.pv: self.pv.clear('Select a paper from the list')
        if spec['key'] in self.cache:
            self.rows = self.cache[spec['key']]; self.fill()
        else:
            self.rows = []; self.fill(); self.load()

    def invalidate(self, key=None):
        if key is None: self.cache.clear()
        else: self.cache.pop(key, None)

    def reload(self):
        if not self.spec: return
        self.invalidate(self.spec['key']); self.rows = []; self.fill(); self.load()

    def load(self):
        spec = self.spec; key = spec['key']
        if key in self.loading: return
        self.loading.add(key); self.l_sum.setText('reading…')
        res = {}

        def work():
            try: res['rows'] = spec['fetch']()
            except Exception as e:
                LOG.exception('Documents list failed: %s', key); res['e'] = e

        th = threading.Thread(target=work, daemon=True); th.start()

        def wait():
            if th.is_alive(): return QTimer.singleShot(200, wait)
            self.loading.discard(key)
            try:
                if 'e' in res:
                    if self.spec and self.spec['key'] == key: self.l_sum.setText(''); self.l_note.setText(f'Could not read: {res["e"]}')
                    return
                self.cache[key] = res.get('rows') or []
                if self.spec and self.spec['key'] == key:
                    self.rows = self.cache[key]; self.fill()
            except RuntimeError:
                pass                                     # الشاشة اتقفلت وهو بيقرا

        QTimer.singleShot(200, wait)

    # ------------------------------------------------------------ الجدول
    def set_flag(self, k):
        self.flag = k; self.fill()

    def fill(self):
        if not self.spec: return
        cols = self.spec['cols']
        for k, (b, text) in self.chips.items():
            n = sum(1 for r in self.rows if k in r.get('flags', {'all'}))
            b.setText(f'{text}  {n}'); b.setChecked(k == self.flag)
        self.view = [r for r in self.rows if X.match(r, cols, self.e_find.text(), self.flag)]
        self.tv.setRowCount(0)
        for r in self.view:
            i = self.tv.rowCount(); self.tv.insertRow(i)
            for c, (key, _, _) in enumerate(cols):
                it = QTableWidgetItem(str(r.get(key, '') if r.get(key) is not None else ''))
                if key in ('status', 'days') or (key == 'filename' and not r.get('path')):
                    it.setForeground(QColor(KIND_COL.get(r.get('kind'), MUTED)))
                self.tv.setItem(i, c, it)
        self._buttons()
        if self.spec['key'] not in self.loading:
            self.l_sum.setText(f'{len(self.view)} of {len(self.rows)}' if len(self.view) != len(self.rows) else f'{len(self.rows)} paper(s)')

    def sel(self):
        i = self.tv.currentRow()
        return self.view[i] if 0 <= i < len(self.view) and self.tv.selectionModel().hasSelection() else None

    def _buttons(self):
        """Edit شغال على الصف المفتوح بس؛ المقفول 🔒 يتشاف ومايتعدّلش."""
        r = self.sel()
        self.b_edit.setEnabled(bool(r) and not r.get('locked') and bool(r.get('editable', True)))
        self.b_edit.setToolTip('' if not r else 'Locked – delivered / approved' if r.get('locked')
                               else '' if r.get('editable', True) else 'No Word file made here for this line')
        self.b_del.setEnabled(bool(r))
        self.b_del.setText('↶  Undo delivered' if r and r.get('delivered') else '✔  Delivered')
        on = bool(r) and str(r.get('unf', '')).startswith('✔')
        self.b_unf.setEnabled(bool(r) and bool(r.get('mark')))
        self.b_unf.setText('↶  Not on Unifier' if on else '☁  Uploaded to Unifier')

    def _edit(self):
        r = self.sel()
        if r and not r.get('locked'): self.on_edit(r)

    def _deliver(self):
        r = self.sel()
        if r: self.on_deliver(r)

    def _unifier(self):
        """علامة «اترفع على Unifier» للـ MAR / MIR المختار (نفس العلامة اللي كانت في Materials)."""
        r = self.sel(); fn = (self.spec or {}).get('unifier')
        if not r or not r.get('mark') or not fn: return
        try: r['unf'] = fn(r)
        except Exception as e:
            LOG.exception('Unifier mark'); return QMessageBox.warning(self, 'Documents', f'Could not save the mark:\n{e}')
        if str(r['unf']).startswith('✔'): r['flags'].discard('unifier')
        elif r.get('kind') == 'ok' or 'signed' in r['flags']: r['flags'].add('unifier')
        self.fill()

    def show_sel(self):
        self._buttons()
        if not self.pv: return
        r = self.sel()
        p = (r or {}).get('path')
        if p and os.path.isfile(p) and p.lower().endswith(X.PREVIEW_EXT): self.pv.load(p)
        elif r: self.pv.clear('No file found for this line' if not (r.get('path') or r.get('open')) else 'No preview for this kind of file – press «Open file»')
        else: self.pv.clear('Select a paper from the list')

    # ------------------------------------------------------------ فتح
    def _target(self):
        r = self.sel()
        if not r: return None
        return r.get('open') or r.get('path') or None

    def open_file(self):
        p = self._target()
        if not p: return
        self._start(p)

    def open_folder(self):
        p = self._target()
        if p: p = p if os.path.isdir(p) else os.path.dirname(p)
        elif self.spec and self.spec.get('folder'): p = self.spec['folder']
        if p: self._start(p)

    def _start(self, p):
        try:
            if os.name == 'nt': os.startfile(p)
            else:
                import subprocess; subprocess.Popen(['xdg-open', p])
        except Exception as e:
            QMessageBox.warning(self, 'Documents', str(e))
