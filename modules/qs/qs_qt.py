# -*- coding: utf-8 -*-
"""QS | Quantities بـ Qt – نفس شاشة qs_app.py في الشغل (qs_core / qs_sheets / qs_charts / upd_dates زي ما هما)،
والشكل الجديد الغامق.

التابات (نفس ترتيب القديم): Summary – Weekly & Monthly Update – Update Charts – Excavation – Bedding – Pipe Laying –
Backfill – Road Base – Asphalt – Notes – Unifier – Chambers – Connections – Boring.
- أساس الحساب (Approved / Requested) و«Up to date» واحد فوق لكل التابات.
- الخانات الصفرا في Excavation بتتكتب زي الإكسيل: Enter = الصف اللي تحت، Tab = اللي جنبه، Delete = تمسح، R / O للتربة.
- الحفظ أوتوماتيك في QS_INPUT.xlsx بعد ما الكتابة تقف ثانيتين (Ctrl+S يحفظ على طول).
- الرسومات بنفس كود qs_charts (اللي بيرسم في الإكسيل كمان) – بيترسم هنا على لوحة Qt.
- «Edit Dates…» لسه شباك Tk (بيفتح فوق البرنامج ويقفل لما تحفظ)."""
import os, sys, time, json, traceback, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
MODS = os.path.dirname(HERE); PLAT = os.path.dirname(MODS)
for p in (PLAT, os.path.join(MODS, 'dc')):
    if p not in sys.path: sys.path.insert(0, p)

from PySide6.QtCore import Qt, QTimer, QDate, QRectF, QPointF, QEvent
from PySide6.QtGui import (QColor, QFont, QGuiApplication, QPainter, QPen, QBrush, QPolygonF, QFontMetricsF, QKeySequence,
                           QShortcut)
from PySide6.QtWidgets import (QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem,
                               QAbstractItemView, QMessageBox, QTabWidget, QComboBox, QDateEdit, QPlainTextEdit, QFrame,
                               QStyledItemDelegate, QAbstractItemDelegate, QStyle, QFileDialog, QInputDialog, QPushButton, QSplitter, QHeaderView)
import home_qt as HQ
import dc_qt as DQ
import qs_core as core
import qs_defs as OLDQ                    # نفس الأعمدة والصيغ (f2 / fd / STATUSES …)

INK, MUTED, NAVY = DQ.INK, DQ.MUTED, DQ.NAVY
TT = HQ.T
D = HQ.DARK
CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#qsq') + f"""
QTabWidget::pane {{ border: none; border-top: 2px solid {NAVY}; }}
QTabBar::tab {{ background: {TT['chip']}; color: {MUTED}; padding: 7px 13px; margin-right: 2px; border-top-left-radius: 8px;
               border-top-right-radius: 8px; font: 600 9.5pt "Segoe UI"; }}
QTabBar::tab:selected {{ background: {NAVY}; color: white; }}
QDateEdit {{ background: {TT['card2']}; border: 1px solid {DQ.SOFT}; border-radius: 8px; padding: 5px 9px; color: {INK}; }}
QPlainTextEdit {{ background: {TT['card2']}; color: {INK}; border: none; font: 10.5pt "Segoe UI"; padding: 8px; }}
QFrame#kpi {{ background: {TT['chip']}; border-radius: 10px; }}
QPushButton#tog {{ padding: 5px 14px; }}
QPushButton#tog:checked {{ background: {NAVY}; color: white; border: none; }}
"""
f2, fd = OLDQ.f2, OLDQ.fd
INPUT_COLS = OLDQ.INPUT_COLS
TAG = ({'grey': ('#262C33', '#7C8896'), 'miss': ('#3A3018', '#F0C860'), 'ok': (None, None), 'part': ('#1F3A24', None),
        'warn': ('#3E1E22', None), 'tot': ('#3A3018', None), 'net': (None, '#F4A261'), 'chg': ('#1F3A24', None)}
       if D else {'grey': ('#F4F4F4', '#8A8A8A'), 'miss': ('#FFF8E1', '#7F6000'), 'ok': (None, None), 'part': ('#EAF4E4', None),
                  'warn': ('#FDE9D9', None), 'tot': ('#FFF2CC', None), 'net': (None, '#C55A11'), 'chg': ('#EAF4E4', None)})
IN_BG, IN_BG_OFF = ('#4A3C12', '#2E2A1E') if D else ('#FFF2CC', '#FBF6E9')
OKC, AMB, RED = DQ.GREEN, ('#F0C860' if D else '#B7791F'), ('#FF8A8A' if D else '#C00000')
UI = OLDQ.load_ui()


def save_ui(): OLDQ.save_ui(UI)


class BgDelegate(QStyledItemDelegate):
    """لون خلفية الصف (الستايل الغامق بيتجاهل لون الخلية) + محرر للخانات اللي بتتكتب"""

    def __init__(self, parent, choices=None):
        super().__init__(parent); self.choices = choices or {}

    def paint(self, p, opt, idx):
        bg = idx.data(Qt.BackgroundRole)
        if bg is not None and not (opt.state & QStyle.StateFlag.State_Selected): p.fillRect(opt.rect, bg)
        super().paint(p, opt, idx)

    def createEditor(self, parent, opt, idx):
        ch = self.choices.get(idx.column())
        if ch:
            c = QComboBox(parent); c.setEditable(True); c.addItems([str(x) for x in ch]); return c
        return super().createEditor(parent, opt, idx)

    def setEditorData(self, ed, idx):
        raw = idx.data(Qt.UserRole)
        txt = '' if raw is None else str(raw)
        if isinstance(ed, QComboBox): ed.setCurrentText(txt)
        elif hasattr(ed, 'setText'): ed.setText(txt); ed.selectAll()
        else: super().setEditorData(ed, idx)

    def setModelData(self, ed, model, idx):
        txt = ed.currentText() if isinstance(ed, QComboBox) else ed.text()
        model.setData(idx, txt.strip(), Qt.UserRole + 5)          # القيمة الجديدة – الشاشة بتتعامل معاها (مش بتتكتب في الخانة)


class Table(QTableWidget):
    """جدول زي الإكسيل: الخانات اللي بتتكتب (edit_cols) – Enter ينزل، Tab يروح يمين، Delete يمسح.
    on_edit(row, col, text) بيتنده لما خانة تتغير."""

    def __init__(self, cols, edit_cols=(), on_edit=None, choices=None, name=None):
        super().__init__(0, len(cols))
        self.cols, self.ids = cols, [c[0] for c in cols]
        self.edit_cols = [self.ids.index(c) for c in edit_cols if c in self.ids]
        self.on_edit, self.name = on_edit, name
        self.setHorizontalHeaderLabels([('✎ ' if c[0] in edit_cols and not str(c[1]).startswith('✎') else '') + c[1] for c in cols])
        self.verticalHeader().setVisible(False); self.verticalHeader().setDefaultSectionSize(int(UI.get('row_height', 26)) + 2)
        self.setSelectionBehavior(QAbstractItemView.SelectItems if self.edit_cols else QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed
                             if self.edit_cols else QAbstractItemView.NoEditTriggers)
        self.setShowGrid(False); self.setWordWrap(False); self.horizontalHeader().setHighlightSections(False)
        saved = UI.get('cols', {}).get(name, {}) if name else {}
        for i, (cid, _, w) in enumerate(cols): self.setColumnWidth(i, int(saved.get(cid, w * 1.12)))
        ch = {self.ids.index(k): v for k, v in (choices or {}).items() if k in self.ids}
        self.setItemDelegate(BgDelegate(self, ch))
        self.model().dataChanged.connect(self._changed)
        self._busy = False
        if name: self.horizontalHeader().sectionResized.connect(self._remember)
        self._move = None

    def _remember(self, i, old, new):
        if self._busy: return
        UI.setdefault('cols', {}).setdefault(self.name, {})[self.ids[i]] = new; save_ui()

    def set_rows(self, rows):
        """rows: [(values, tag)] – القيم بالترتيب. القيمة ممكن تبقى (المكتوب, الرقم الخام) للخانات اللي بتتكتب."""
        self._busy = True
        cur = (self.currentRow(), self.currentColumn())
        sv, sh = self.verticalScrollBar().value(), self.horizontalScrollBar().value()
        self.setUpdatesEnabled(False); self.setRowCount(len(rows))
        for i, (vals, tag) in enumerate(rows):
            bg, fg = TAG.get(tag, (None, None))
            for j, v in enumerate(vals):
                raw = None
                if isinstance(v, tuple): v, raw = v
                it = QTableWidgetItem('' if v is None else str(v)); it.setData(Qt.UserRole, raw if raw is not None else v)
                it.setTextAlignment(Qt.AlignCenter)
                if j in self.edit_cols:
                    it.setFlags(it.flags() | Qt.ItemIsEditable)
                    it.setBackground(QColor(IN_BG_OFF if tag == 'grey' else IN_BG))
                    f = it.font(); f.setBold(True); it.setFont(f)
                else:
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                    if bg: it.setBackground(QColor(bg))
                if fg: it.setForeground(QColor(fg))
                if tag in ('tot', 'net'): f = it.font(); f.setBold(True); it.setFont(f)
                self.setItem(i, j, it)
        self.setUpdatesEnabled(True)
        self.verticalScrollBar().setValue(sv); self.horizontalScrollBar().setValue(sh)
        if 0 <= cur[0] < len(rows) and cur[1] >= 0: self.setCurrentCell(cur[0], cur[1])
        self._busy = False
        if self._move: r, c = self._move; self._move = None; QTimer.singleShot(0, lambda: self._go(r, c))

    def _go(self, r, c):
        if 0 <= r < self.rowCount(): self.setCurrentCell(r, c); self.scrollToItem(self.item(r, c))

    def _changed(self, a, b, roles=()):
        if self._busy or (roles and (Qt.UserRole + 5) not in roles): return
        txt = a.data(Qt.UserRole + 5)
        if txt is None: return
        self.model().setData(a, None, Qt.UserRole + 5)
        if self.on_edit: QTimer.singleShot(0, lambda r=a.row(), c=a.column(), t=txt: self.on_edit(r, self.ids[c], t))

    def closeEditor(self, editor, hint):
        super().closeEditor(editor, hint)
        r, c = self.currentRow(), self.currentColumn()
        if hint == QAbstractItemDelegate.EndEditHint.EditNextItem:                       # Tab ← الخانة اللي بتتكتب بعدها
            nxt = [x for x in self.edit_cols if x > c]
            self._move = (r, nxt[0]) if nxt else (r + 1, self.edit_cols[0])
        elif hint == QAbstractItemDelegate.EndEditHint.SubmitModelCache:                 # Enter ← الصف اللي تحت
            self._move = (r + 1, c)
        if self._move:
            m = self._move
            QTimer.singleShot(60, lambda: (self._go(*m), setattr(self, '_move', None)) if self._move == m else None)

    def keyPressEvent(self, ev):
        k = ev.key(); c = self.currentColumn(); r = self.currentRow()
        if self.state() != QAbstractItemView.State.EditingState and c in self.edit_cols and r >= 0:
            if k in (Qt.Key_Delete, Qt.Key_Backspace):
                if self.on_edit: self.on_edit(r, self.ids[c], '')
                return
            if k in (Qt.Key_Return, Qt.Key_Enter): self.editItem(self.item(r, c)); return
            if self.ids[c] == 'soil' and ev.text().lower() in ('r', 'o'):
                if self.on_edit: self.on_edit(r, 'soil', core.ROCK if ev.text().lower() == 'r' else core.OTHER)
                return
        super().keyPressEvent(ev)


# ---------------------------------------------------------------- لوحة الرسومات: نفس أوامر Canvas بتاعة Tk، والرسم بـ QPainter
class ChartCanvas(QWidget):
    """qs_charts بيرسم بأوامر Tk Canvas (create_text / create_rectangle …) – هنا بتتسجل وبتترسم بـ QPainter."""
    ANCH = {'n': (0.5, 0), 's': (0.5, 1), 'e': (1, 0.5), 'w': (0, 0.5), 'nw': (0, 0), 'ne': (1, 0), 'sw': (0, 1), 'se': (1, 1),
            'center': (0.5, 0.5), None: (0.5, 0.5)}

    def __init__(self, redraw):
        super().__init__(); self.ops = []; self.redraw = redraw; self.setMinimumHeight(220)
        self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet('background: white; border-radius: 10px;')

    # ---- واجهة Tk
    def delete(self, *_): self.ops = []
    def winfo_width(self): return self.width()
    def winfo_height(self): return self.height()

    def _font(self, f):
        if not f: return QFont('Segoe UI', 9)
        q = QFont(f[0], int(f[1]) if len(f) > 1 else 9)
        mods = f[2] if len(f) > 2 else ''
        q.setBold('bold' in mods); q.setItalic('italic' in mods)
        return q

    def _text_rect(self, x, y, text, font, anchor, width=None):
        fm = QFontMetricsF(font)
        w = fm.horizontalAdvance(text) if not width else min(fm.horizontalAdvance(text), width)
        h = fm.height()
        ax, ay = self.ANCH.get(anchor, (0.5, 0.5))
        return QRectF(x - ax * w, y - ay * h, w, h)

    def create_text(self, x, y, text='', anchor=None, fill='black', font=None, width=None, justify=None, **k):
        f = self._font(font); r = self._text_rect(x, y, str(text), f, anchor, width)
        self.ops.append(('t', r, str(text), f, fill)); return len(self.ops) - 1

    def bbox(self, i):
        op = self.ops[i]; r = op[1]
        return (r.left(), r.top(), r.right(), r.bottom())

    def create_rectangle(self, x0, y0, x1, y1, fill='', outline='', width=1, **k):
        self.ops.append(('r', QRectF(min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0)), fill, outline, width)); return len(self.ops) - 1

    def create_oval(self, x0, y0, x1, y1, fill='', outline='', width=1, **k):
        self.ops.append(('o', QRectF(min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0)), fill, outline, width)); return len(self.ops) - 1

    def create_line(self, *c, fill='black', width=1, dash=None, **k):
        pts = [QPointF(c[i], c[i + 1]) for i in range(0, len(c) - 1, 2)]
        self.ops.append(('l', pts, fill, width, dash)); return len(self.ops) - 1

    def create_polygon(self, *c, fill='', outline='', **k):
        pts = [QPointF(c[i], c[i + 1]) for i in range(0, len(c) - 1, 2)]
        self.ops.append(('p', pts, fill, outline)); return len(self.ops) - 1

    def configure(self, **k): pass

    # ---- الرسم
    def resizeEvent(self, ev):
        super().resizeEvent(ev); QTimer.singleShot(0, self.redraw)

    def paintEvent(self, ev):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing); p.setRenderHint(QPainter.TextAntialiasing)
        p.fillRect(self.rect(), QColor('white'))
        for op in self.ops:
            k = op[0]
            if k == 't':
                _, r, txt, f, fill = op
                p.setFont(f); p.setPen(QColor(fill or 'black'))
                p.drawText(r.adjusted(-2, 0, 2, 0), Qt.AlignCenter, txt)
            elif k in ('r', 'o'):
                _, r, fill, outline, w = op
                p.setBrush(QBrush(QColor(fill)) if fill else Qt.NoBrush)
                p.setPen(QPen(QColor(outline), w) if outline else Qt.NoPen)
                (p.drawRect if k == 'r' else p.drawEllipse)(r)
            elif k == 'l':
                _, pts, fill, w, dash = op
                pen = QPen(QColor(fill or 'black'), w)
                if dash: pen.setDashPattern([max(1, d / max(w, 1)) for d in dash])
                p.setPen(pen); p.setBrush(Qt.NoBrush)
                for a, b in zip(pts[:-1], pts[1:]): p.drawLine(a, b)
            elif k == 'p':
                _, pts, fill, outline = op
                p.setBrush(QBrush(QColor(fill)) if fill else Qt.NoBrush)
                p.setPen(QPen(QColor(outline)) if outline else Qt.NoPen)
                p.drawPolygon(QPolygonF(pts))
        p.end()


def kpi_box(lbl, val):
    f = QFrame(); f.setObjectName('kpi'); v = QVBoxLayout(f); v.setContentsMargins(12, 6, 12, 6); v.setSpacing(1)
    a = QLabel(lbl); a.setStyleSheet(f'color: {MUTED}; font-size: 9pt;'); b = QLabel(val); b.setFont(HQ.font(15, True))
    b.setStyleSheet(f'color: {INK};'); v.addWidget(a); v.addWidget(b)
    return f


def clear_layout(lay):
    while lay.count():
        it = lay.takeAt(0)
        if it.widget(): it.widget().deleteLater()


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


def qd(d): return QDate(d.year, d.month, d.day)
def pyd(q): return dt.date(q.year(), q.month(), q.day())


class QsQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        self.setObjectName('qsq'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.items = dict(exc=[], bed=[], pip=[], bf=[]); self.views = {}; self.notes = []; self.view = []
        self.wk_cmp, self._wk_list, self._wk_auto, self._gc_data, self._contract = None, [], False, None, None
        self._last_edit = 0
        self._build()
        QShortcut(QKeySequence.Save, self, activated=lambda: self.autosave(manual=True))
        self._tick = QTimer(self); self._tick.timeout.connect(self.autosave); self._tick.start(3000)
        QTimer.singleShot(100, self.first_load)

    # ------------------------------------------------------------ UI
    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 12, 22, 12); v.setSpacing(8)
        self.l_info = QLabel(''); self.l_info.setStyleSheet(f'color: {OKC};')
        v.addLayout(header('QS', 'quantities – excavation, bedding, pipes, backfill, road base, asphalt', HQ.COL.get('qs', '#3CC2B4'),
                           lambda: self.open_section('doc'),
                           [self.l_info]))
        # أساس الحساب – واحد لكل التابات
        bar = DQ.card(); bh = QHBoxLayout(bar); bh.setContentsMargins(14, 8, 14, 8); bh.setSpacing(8)
        l = QLabel('Calculation basis'); l.setStyleSheet('font-weight: bold;'); bh.addWidget(l)
        self.cb_basis = QComboBox(); self.cb_basis.addItems(['Approved', 'Requested']); bh.addWidget(self.cb_basis)
        bh.addSpacing(10); l = QLabel('Up to date'); l.setStyleSheet('font-weight: bold;'); bh.addWidget(l)
        self.e_cut = QDateEdit(qd(dt.date.today())); self.e_cut.setCalendarPopup(True); self.e_cut.setDisplayFormat('dd-MM-yyyy')
        self.e_cut.setMinimumWidth(130); bh.addWidget(self.e_cut)
        bh.addWidget(DQ.btn('⟳  Refresh from WIR Log', self.refresh))
        bh.addStretch(1)
        bh.addWidget(DQ.btn('Export Excel', self.export)); bh.addWidget(DQ.btn('Preview && Print', self.preview, True))
        self.cb_basis.currentIndexChanged.connect(lambda *_: self.basis_changed())
        self.e_cut.dateChanged.connect(lambda *_: self.basis_changed())
        v.addWidget(bar)
        # التابات
        c = DQ.card(); cv = QVBoxLayout(c); cv.setContentsMargins(12, 8, 12, 8)
        self.tabs = QTabWidget(); self.tabs.setDocumentMode(True); cv.addWidget(self.tabs)
        self.tab_key = {}
        self._build_summary(); self._build_weekly(); self._build_charts()
        # Excavation
        w, wv = self._page('Excavation', 'exc')
        self.tv = Table(OLDQ.EXC_COLS, INPUT_COLS, self.apply_value, {'soil': [core.ROCK, core.OTHER]}, 'exc')
        self.tv.currentCellChanged.connect(lambda r, c, *_: self.on_select(r))
        wv.addWidget(self.tv, 1)
        h = QLabel('✎ Click a yellow cell and type (like Excel): Enter = next row, Tab = next yellow cell, Delete = clear, double-click / F2 = edit   |   '
                   'Soil: press R = Rock, O = Other   |   Asphalt in cm   |   Grey = not counted   |   Yellow = missing levels   |   ~ = interpolated')
        h.setStyleSheet(f'color: {MUTED}; font-size: 9pt;'); h.setWordWrap(True); wv.addWidget(h)
        self.l_det = QLabel('Select a row to see the calculation.'); self.l_det.setWordWrap(True)
        self.l_det.setStyleSheet(f'color: {INK}; background: {TT["chip"]}; border-radius: 8px; padding: 7px 10px;')
        self.l_det.setTextInteractionFlags(Qt.TextSelectableByMouse); wv.addWidget(self.l_det)
        # الباقي (عرض بس)
        hints = {'bed': 'Below = length × width × (Rock 0.20 / Other 0.15) from BED-D-P.   Around = length × width × (pipe OD + 0.30) − pipe volume from COV-P.   '
                        '% Done = done ÷ (below + around).   Done Value = done × 95 SAR (progress).   Invoice = below + around only when COV-P is approved '
                        '(item closed) – this is what goes to the Summary.',
                 'pip': 'Pipe laying (m.l.) – PIP approved = 85%  |  + approved pressure test (PIP -T-1) = 90%  |  + approved flushing & chlorination = 100%.   '
                        'Only on approved excavation.',
                 'bf': 'Layers H = Avg depth − asphalt − 0.15 road base − bedding below − (pipe OD + 0.30)  →  Layers Req. = Layers H ÷ 0.30 (rounded up).   '
                       'Qty H (the Consultant) = same but − 0.45 (consultant formula)  →  Backfill Qty.   Paid = Backfill Qty × approved layers ÷ layers required.',
                 'rb': 'Road base = 2 layers of 15 cm (RB-1, RB-2).  Each layer = length × trench width × 0.15, counted when its WIR is approved (or requested, '
                       'by the calculation basis).  BOQ 31.2.4.5 – 84 SAR/m³.  Not deducted here – the 0.45 deduction is in the backfill quantity only.',
                 'asp': '32.4.1 (ASP-1) = length × (trench width + 0.30) × thickness – 1040 SAR/m³.   32.4.2 (milling, ASP-2-ML) = length × (3.60 − asphalt width) '
                        '× 0.04 – 1270 SAR/m³.   Wearing course = same quantity and same WIR as milling, not paid.   Prime coats MC / RC: no quantity.'}
        self.t = {}
        for key, title, cols, fk in (('bedp', 'Bedding', OLDQ.BEDP_COLS, 'bed'), ('pip', 'Pipe Laying', OLDQ.PIP_COLS, 'pip'),
                                     ('bf', 'Backfill', OLDQ.BF_COLS, 'bf'), ('rb', 'Road Base', OLDQ.RB_COLS, 'rb'),
                                     ('asp', 'Asphalt', OLDQ.ASP_COLS, 'asp')):
            w, wv = self._page(title, fk, key)
            t = Table(cols, name=key); self.t[key] = t; wv.addWidget(t, 1)
            h = QLabel(hints[fk]); h.setWordWrap(True); h.setStyleSheet(f'color: {MUTED}; font-size: 9pt;'); wv.addWidget(h)
        # Notes
        self.tx = QPlainTextEdit(); self.tx.setReadOnly(True)
        self.i_notes = self.tabs.addTab(self.tx, 'Notes'); self.tab_key[self.i_notes] = 'none'
        self._build_extra()
        self.tabs.currentChanged.connect(self.tab_changed)
        v.addWidget(c, 1)
        self.kbar = QHBoxLayout(); self.kbar.setSpacing(10)
        kw = QWidget(); kw.setLayout(self.kbar); v.addWidget(kw)

    def _page(self, title, fkey, item=None):
        """تاب فيه شريط فلتر (عرض بس) + جدول"""
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0); wv.setSpacing(6)
        fb = QHBoxLayout(); fb.setSpacing(6); cbs = []
        for lbl in ('Line', 'District', 'Status'):
            l = QLabel(lbl); l.setStyleSheet(f'color: {MUTED};'); fb.addWidget(l)
            cb = QComboBox(); cb.setMinimumWidth(130); cb.addItems(OLDQ.STATUSES if lbl == 'Status' else ['All'])
            cb.currentIndexChanged.connect(lambda *_: self.fill()); fb.addWidget(cb); cbs.append(cb)
        b = DQ.btn('Clear filter', lambda cbs=cbs: ([x.setCurrentIndex(0) for x in cbs], self.fill())); fb.addWidget(b)
        n = QLabel('(view only – Summary and export always include everything)'); n.setStyleSheet(f'color: {MUTED}; font-size: 9pt;')
        fb.addWidget(n); fb.addStretch(1)
        wv.addLayout(fb)
        self.flt = getattr(self, 'flt', {}); self.flt[fkey] = cbs
        i = self.tabs.addTab(w, title); self.tab_key[i] = item or fkey
        return w, wv

    # ---- Summary (QS-TOTAL): أول 6 أعمدة ثابتة
    def _build_summary(self):
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
        tp = os.path.join(HERE, 'QS-TOTAL_TEMPLATE.xlsx')
        try: self.hdr = core.template_header(tp)
        except Exception:
            self.hdr = dict(ncol=25, letters=[c[0] for c in OLDQ.SUM_COLS], widths=[c[2] / 7 for c in OLDQ.SUM_COLS],
                            cells=[(1, i + 1, 4, i + 1, c[1], '#D9E1F2') for i, c in enumerate(OLDQ.SUM_COLS)], rows=4)
        L = self.hdr['letters']; names = {}
        for (r1, c1, r2, c2, txt, col) in sorted(self.hdr['cells'], key=lambda x: x[0]):
            for ci in range(c1, c2 + 1):
                if ci - 1 < len(L) and txt: names.setdefault(L[ci - 1], []).append(' '.join(str(txt).split()))
        saved = UI.get('cols', {}).get('sum', {})
        cols = []
        for Lt, wd in zip(L, self.hdr['widths']):
            nm = names.get(Lt, [Lt])
            title = nm[-1] if len(nm) == 1 or nm[-1] != nm[0] else nm[-1]
            if len(nm) > 1 and len(nm[-1]) < 12 and nm[-2] != nm[-1]: title = f'{nm[-2]}\n{nm[-1]}'
            cols.append((Lt, title, int(saved.get(Lt, max(34, wd * 7 + 5)) / 1.12)))
        NF = min(OLDQ.SUM_FREEZE, len(cols))
        sp = QSplitter(Qt.Horizontal); sp.setHandleWidth(3)
        self.tsf = Table(cols[:NF], name='sumf'); self.ts = Table(cols[NF:], name='sum')
        for t in (self.tsf, self.ts):
            t.horizontalHeader().setMinimumHeight(52); t.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.TextWordWrap)
            t.horizontalHeader().setStretchLastSection(False)
        self.tsf.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.ts.verticalScrollBar().valueChanged.connect(self.tsf.verticalScrollBar().setValue)
        self.tsf.verticalScrollBar().valueChanged.connect(self.ts.verticalScrollBar().setValue)
        self.ts.currentCellChanged.connect(lambda r, *_: self.tsf.selectRow(r) if r >= 0 and self.tsf.currentRow() != r else None)
        self.tsf.currentCellChanged.connect(lambda r, *_: self.ts.selectRow(r) if r >= 0 and self.ts.currentRow() != r else None)
        sp.addWidget(self.tsf); sp.addWidget(self.ts)
        fw = sum(self.tsf.columnWidth(i) for i in range(NF)) + 4
        sp.setSizes([fw, 2000]); sp.setStretchFactor(1, 1)
        wv.addWidget(sp, 1)
        self.i_sum = self.tabs.addTab(w, '∑  Summary (QS-TOTAL)'); self.tab_key[self.i_sum] = 'sum'

    # ---- Weekly & Monthly Update
    WK_COLS = OLDQ.WK_COLS

    def _toggles(self, lay, items, cmd):
        out = {}
        for k, txt in items:
            b = QPushButton(txt); b.setObjectName('tog'); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _c=False, k=k: cmd(k)); lay.addWidget(b); out[k] = b
        return out

    def _build_weekly(self):
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0); wv.setSpacing(6)
        self.wtrack = UI.get('wk_track', 'Company'); self.wper = 'Weekly'
        bar = QHBoxLayout(); bar.setSpacing(6)
        lab = lambda t: (lambda l: (l.setStyleSheet('font-weight: bold;'), bar.addWidget(l)))(QLabel(t))
        lab('1  Update for'); self.wk_tbtn = self._toggles(bar, [('Company', 'Company'), ('Owner', 'Owner')], self.wk_set_track)
        bar.addSpacing(14); lab('2  Compare with')
        self.wk_pbtn = self._toggles(bar, [('Weekly', 'Previous update'), ('Monthly', 'Month end')], self.wk_set_period)
        bar.addSpacing(14); lab('3  Update date')
        self.cb_wsel = QComboBox(); self.cb_wsel.setMinimumWidth(330); self.cb_wsel.activated.connect(lambda *_: self.wk_pick())
        bar.addWidget(self.cb_wsel); bar.addStretch(1)
        bar.addWidget(DQ.btn('Edit dates…', self.wk_edit_dates))
        wv.addLayout(bar)
        r2 = QHBoxLayout()
        self.l_wnote = QLabel('Choose the update, then its date.'); self.l_wnote.setWordWrap(True)
        self.l_wnote.setStyleSheet(f'color: {INK}; font-weight: bold;'); r2.addWidget(self.l_wnote, 1)
        r2.addWidget(DQ.btn('Save this update', self.wk_save, True)); r2.addWidget(DQ.btn('Export update', self.wk_export))
        wv.addLayout(r2)
        self.tw = Table(self.WK_COLS, name='wk'); wv.addWidget(self.tw, 3)
        self.wkbar = QHBoxLayout(); self.wkbar.setSpacing(10); kw = QWidget(); kw.setLayout(self.wkbar); wv.addWidget(kw)
        hb = QHBoxLayout()
        for attr, title, cols in (('th_w', 'Weekly – saved updates', [('d', 'Date', 95), ('v', 'Total Value (SAR)', 140), ('dv', 'This Week (SAR)', 130), ('days', 'Days', 50)]),
                                  ('th_m', 'Monthly – month end', [('m', 'Month', 90), ('d', 'Last Update', 95), ('v', 'Total Value (SAR)', 140), ('dv', 'This Month (SAR)', 130)])):
            col = QVBoxLayout(); l = QLabel(title); l.setStyleSheet('font-weight: bold;'); col.addWidget(l)
            t = Table(cols); t.setMaximumHeight(190); col.addWidget(t); hb.addLayout(col, 1); setattr(self, attr, t)
        wv.addLayout(hb, 1)
        self.i_wk = self.tabs.addTab(w, 'Weekly & Monthly Update'); self.tab_key[self.i_wk] = 'none'

    # ---- Update Charts
    def _build_charts(self):
        import qs_charts
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0); wv.setSpacing(6)
        bar = QHBoxLayout(); bar.setSpacing(6)
        l = QLabel('Update for'); l.setStyleSheet('font-weight: bold;'); bar.addWidget(l)
        self.gc_tbtn = self._toggles(bar, [('Company', 'Company'), ('Owner', 'Owner')], self.gc_set_track)
        bar.addSpacing(14); l = QLabel('Period'); l.setStyleSheet('font-weight: bold;'); bar.addWidget(l)
        self.gc_pbtn = self._toggles(bar, qs_charts.KINDS, self.gc_set_period)
        bar.addStretch(1); self.l_gc = QLabel(''); self.l_gc.setStyleSheet(f'color: {MUTED};'); bar.addWidget(self.l_gc)
        wv.addLayout(bar)
        self.gc_kind = UI.get('chart_period') if UI.get('chart_period') in dict(qs_charts.KINDS) else 'M'
        g = QGridLayout(); g.setSpacing(8)
        self.gc_done, self.gc_curve, self.gc_stack = (ChartCanvas(self.charts_draw_later) for _ in range(3))
        g.addWidget(self.gc_done, 0, 0, 1, 2); g.addWidget(self.gc_curve, 1, 0); g.addWidget(self.gc_stack, 1, 1)
        g.setRowStretch(0, 1); g.setRowStretch(1, 1)
        wv.addLayout(g, 1)
        self._gc_job = QTimer(self); self._gc_job.setSingleShot(True); self._gc_job.timeout.connect(lambda: self.charts_draw())
        self.i_gc = self.tabs.addTab(w, 'Update Charts'); self.tab_key[self.i_gc] = 'none'

    # ---- Unifier / Chambers / Connections / Boring (من qs_tabs)
    def _build_extra(self):
        try:
            import qs_sheets as S
        except Exception:
            traceback.print_exc(); self.S = None; return
        self.S = S
        w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
        cols = [('n', '#', 36), ('line', 'Line', 50), ('x0', 'From', 70), ('x1', 'To', 70), ('need', 'Layers Req.', 75)]
        for gi, (title, codes) in enumerate(S.GROUPS, 1):
            short = title.split('\n')[-1]
            cols += ([(f'u{gi}', f'{short} – UN', 230), (f'c{gi}', 'Layers', 60), (f'i{gi}', 'INV', 60)] if codes == 'BF'
                     else [(f'u{gi}', f'{short} – UN', 120), (f'i{gi}', 'INV', 55)])
        self.ug = Table(cols, [c for c, _, _ in cols if c.startswith('i')], self._u_edit, name='unifier'); wv.addWidget(self.ug, 1)
        h = QLabel('UN = Unifier no. of the approved WIR (from the WIR log) – edit it in the log.   ✎ INV = invoice no. (type it here).   '
                   'Red layers = a missing layer below an approved one.   (WIR n) = approved WIR without Unifier no.   Saved automatically.')
        h.setWordWrap(True); h.setStyleSheet(f'color: {MUTED}; font-size: 9pt;'); wv.addWidget(h)
        i = self.tabs.addTab(w, 'Unifier'); self.tab_key[i] = 'none'; self.u_data = []
        # الجداول اللي بتتكتب كلها
        self.lists = {}
        stages = []
        for k, (hh, pc) in enumerate(S.STAGES, 1):
            short = hh.split('(')[0].strip()[:18]
            stages += [(f'un{k}', f'{k}. {short} – UN ({pc:.1%})', 150, True, None), (f'inv{k}', f'{k}. INV', 55, True, None)]
        ch_calc = {'pct': lambda r: f'{S.ch_pct(r):.1%}', 'val': lambda r: f"{S.ch_pct(r) * S.CH_RATES.get(r.get('type', ''), 0):,.0f}"}

        def lm(r):
            try: return f"{abs(float(r.get('x1')) - float(r.get('x0'))):.2f}"
            except (TypeError, ValueError): return ''

        def bor_tot(d):
            s = 0.0
            for r in d:
                try: s += abs(float(r.get('x1')) - float(r.get('x0')))
                except (TypeError, ValueError): pass
            return f'Boring: {s:,.2f} LM   |   {s * 11200:,.0f} SAR'
        specs = [('CHAMBERS', 'Chambers',
                  [('n', '#', 36, False, None), ('zone', 'Zone', 90, True, None), ('line', 'Line', 60, True, None),
                   ('type', 'Type', 70, True, list(S.CH_RATES)), ('dia', 'Dia', 55, True, ['300', '400', '600']),
                   ('station', 'Station', 70, True, None), ('valves', 'Valves', 90, True, None),
                   ('pct', 'Progress %', 75, False, None), ('val', 'Value (SAR)', 95, False, None)] + stages,
                  'One row per chamber.  Type: ' + '  '.join(f'{t} = {it}' for t, it, _ in S.CH_TYPES) +
                  '.   Type the Unifier no. in each stage when its WIR is approved – progress % uses the agreed stage weights.', ch_calc,
                  lambda d: f"Chambers: {len(d)}   |   Value: {sum(S.ch_pct(r) * S.CH_RATES.get(r.get('type', ''), 0) for r in d):,.0f} SAR"),
                 ('CONNECTIONS', 'Connections',
                  [('n', '#', 36, False, None), ('zone', 'Zone', 100, True, None), ('line', 'Line', 70, True, None),
                   ('station', 'Station', 80, True, None), ('prop', 'Proposed Dia', 95, True, [str(x) for x in S.CON_PROP]),
                   ('exist', 'Existing Dia', 95, True, [str(x) for x in S.CON_EXIST]), ('un', 'UN NO', 110, True, None), ('inv', 'INV', 70, True, None)],
                  'One row per connection (proposed line on existing line).  Export groups them by zone & line into the consultant matrix.', {},
                  lambda d: f'Connections: {len(d)}'),
                 ('BORING', 'Boring',
                  [('n', '#', 36, False, None), ('zone', 'Zone', 110, True, None), ('line', 'Line', 70, True, None),
                   ('dia', 'Dia', 60, True, ['300', '400', '500', '600']), ('x0', 'From', 80, True, None), ('x1', 'To', 80, True, None),
                   ('lm', 'LM', 70, False, None), ('un', 'UN NO', 110, True, None), ('inv', 'INV', 70, True, None)],
                  'Horizontal boring (item 33.1.2 – 1000008 – 11,200 SAR/LM).  LM = |To − From|.', {'lm': lm}, bor_tot)]
        for key, title, spec, hint, calc, tot in specs:
            w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0)
            bar = QHBoxLayout(); bar.addWidget(DQ.btn('+ Add row', lambda k=key: self._l_add(k)))
            bar.addWidget(DQ.btn('Delete row', lambda k=key: self._l_del(k))); bar.addStretch(1)
            lt = QLabel(''); lt.setStyleSheet('font-weight: bold;'); bar.addWidget(lt); wv.addLayout(bar)
            g = Table([(c, t, wd) for c, t, wd, _, _ in spec], [c for c, _, _, ed, _ in spec if ed],
                      lambda r, c, t, k=key: self._l_edit(k, r, c, t), {c: ch for c, _, _, _, ch in spec if ch}, name=key.lower())
            wv.addWidget(g, 1)
            h = QLabel(hint); h.setWordWrap(True); h.setStyleSheet(f'color: {MUTED}; font-size: 9pt;'); wv.addWidget(h)
            i = self.tabs.addTab(w, title); self.tab_key[i] = 'none'
            self.lists[key] = dict(g=g, spec=spec, calc=calc, tot=tot, lbl=lt)

    # ------------------------------------------------------------ data
    def info(self, t, col=None):
        self.l_info.setText(t); self.l_info.setStyleSheet(f'color: {col or OKC};')

    def first_load(self):
        try:
            self.p = core.Project(core.DEFAULT_INPUT)
        except Exception as e:
            return QMessageBox.critical(self, 'QS', f'Cannot open input file:\n{core.DEFAULT_INPUT}\n\n{e}')
        self.refresh()

    def refresh(self):
        if not hasattr(self, 'p'): return
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.p.load_requests()
        except Exception as e:
            QGuiApplication.restoreOverrideCursor()
            QMessageBox.critical(self, 'QS', f'Cannot read WIR log:\n{self.p.S.get("WIR_LOG_PATH")}\n\n{e}')
        QGuiApplication.restoreOverrideCursor()
        lines = sorted({core.kstr(q['line']) for q in self.p.reqs_all if core.kstr(q['line'])})
        dists = sorted({core.kstr(q['dist']) for q in self.p.reqs_all if core.kstr(q['dist'])})
        for cbs in self.flt.values():
            for cb, vals in ((cbs[0], lines), (cbs[1], dists)):
                cur = cb.currentText(); cb.blockSignals(True); cb.clear(); cb.addItems(['All'] + vals)
                cb.setCurrentIndex(max(cb.findText(cur), 0)); cb.blockSignals(False)
        self.set_basis()
        try: self.fill_history()
        except Exception: traceback.print_exc()
        self.wk_fill_dates()
        self.info(f'WIR log read {dt.datetime.now():%H:%M}')

    def cutoff(self): return pyd(self.e_cut.date())

    def basis_changed(self):
        if not hasattr(self, 'p'): return
        self.set_basis()
        if self.wk_cmp is not None or self.tabs.currentIndex() == self.i_wk: self.wk_compare()
        if self.tabs.currentIndex() == self.i_gc: self.charts_draw(True)

    def set_basis(self):
        self.p.mode = 'requested' if self.cb_basis.currentText() == 'Requested' else 'approved'
        self.tabs.setTabText(self.i_sum, '∑  Summary (QS-REQUESTED)' if self.p.mode == 'requested' else '∑  Summary (QS-TOTAL)')
        self.recompute()

    def st(self, status):
        if getattr(self, 'p', None) is not None and self.p.mode == 'requested' and status == 'Approved': return 'Requested'
        return status

    def recompute(self):
        self.items, self.notes = self.p.build(self.cutoff())
        self.rows = self.items['exc']
        self.fill()
        try: self.refresh_extra()
        except Exception: traceback.print_exc()

    def visible(self, rows, name):
        cbs = self.flt.get(name)
        L, Dd, S = (c.currentText() for c in cbs) if cbs else ('All', 'All', 'All')
        if S == 'Requested': S = 'Approved'
        return [r for r in rows if (L == 'All' or core.kstr(r['q']['line']) == L)
                and (Dd == 'All' or core.kstr(r['q']['dist']) == Dd) and (S == 'All' or r['status'] == S)]

    def fill(self, *_):
        if not hasattr(self, 'p') or not self.items: return
        tag_of = OLDQ.tag_of
        # الحفر
        self.view = self.visible(self.items['exc'], 'exc'); rows = []
        for i, r in enumerate(self.view):
            q = r['q']
            d0 = f2(r['d0']) + (' ~' if r['interp'][0] and r['d0'] is not None else '')
            d1 = f2(r['d1']) + (' ~' if r['interp'][1] and r['d1'] is not None else '')
            if r['d0'] is None: d0 = '✎'
            if r['d1'] is None: d1 = '✎'
            asph = '✎' if r['asph'] is None else (f"{r['asph']:g} ~" if r.get('asph_def') else f"{r['asph']:g}")
            vals = [i + 1, q['dist'], q['street'], q['line'], core.sta(r['x0']), core.sta(r['x1']), f2(r['L']),
                    (d0, '' if r['d0'] is None else f"{r['d0']:.3f}"), (d1, '' if r['d1'] is None else f"{r['d1']:.3f}"),
                    (r['soil'], r['soil']), (asph, '' if r['asph'] is None else f"{r['asph']:g}"),
                    f2(r['avg']), f2(r['w']), *[f2(x) for x in r['v']], f2(r['total']), fd(q['appdate']), self.st(r['status']), q['no']]
            rows.append((vals, tag_of(r['status'])))
        self.tv.set_rows(rows); self.views['exc'] = self.view
        # الفرشة
        self.views['bed'] = self.visible(self.items['bed'], 'bed')
        self.views['bedp'] = self.visible(self.items.get('bedp', []), 'bed'); rows = []
        for i, r in enumerate(self.views['bedp']):
            q = r['q']
            rows.append(([i + 1, q['dist'], q['street'], q['line'], q['dia_raw'], core.sta(r['x0']), core.sta(r['x1']), f2(r['L']), r['soil'],
                          f2(r['w']), f2(r['t']), f2(r['below']), '✔' if r['below_ok'] else '–', f2(r['h']), f2(r['net']),
                          '✔' if r['cov_ok'] else '–', f2(r['total']), f2(r['done']), '–' if r['pct'] is None else f"{r['pct'] * 100:.0f}%",
                          f"{r['value']:,.2f}", f2(r['inv']), r['status'], r['wirs'] or '–'],
                         'ok' if r['status'] == 'Closed' else ('part' if r['status'] == 'Below Done' else 'grey')))
        self.t['bedp'].set_rows(rows)
        # المواسير
        self.views['pip'] = self.visible(self.items['pip'], 'pip'); rows = []
        for i, r in enumerate(self.views['pip']):
            q = r['q']
            rows.append(([i + 1, q['dist'], q['street'], q['line'], q['dia_raw'], core.sta(r['x0']), core.sta(r['x1']), f2(r['L']),
                          r['test'] or '–', r['flush'] or '–', f"{r['pct']:.0f}%", f2(r['paid']), fd(q['appdate']), self.st(r['status']), q['no']],
                         tag_of(r['status'])))
        self.t['pip'].set_rows(rows)
        # الردم
        self.views['bf'] = self.visible(self.items.get('bf', []), 'bf'); rows = []
        for i, r in enumerate(self.views['bf']):
            q = r['q']
            rows.append(([i + 1, q['dist'], q['street'], q['line'], q['dia_raw'], core.sta(r['x0']), core.sta(r['x1']), f2(r['L']),
                          f2(r['avg']), '–' if r['asph'] is None else f"{r['asph']:g}", r['soil'], f2(r['t']), f2(r['ch'], 3),
                          f2(r.get('Hn'), 3), f2(r['H'], 3), '–' if r['n'] is None else r['n'], ', '.join(f'BF-{k}' for k in r['layers']) or '–',
                          '–' if r['pct'] is None else f"{r['pct'] * 100:.0f}%", f2(r['w']), f2(r['qty']), f2(r['paid']),
                          self.st(r['status']), r['wirs'] or '–'], tag_of(r['status'])))
        self.t['bf'].set_rows(rows)
        # القاعدة
        self.views['rb'] = self.visible(self.items.get('rb', []), 'rb'); rows = []
        for i, r in enumerate(self.views['rb']):
            q = r['q']
            rows.append(([i + 1, q['dist'], q['street'], q['line'], q['dia_raw'], core.sta(r['x0']), core.sta(r['x1']), f2(r['L']),
                          f2(r['w']), f2(r['th'], 2), '✔' if 1 in r['layers'] else '–', f2(r['q1']), '✔' if 2 in r['layers'] else '–',
                          f2(r['q2']), f2(r['qty']), self.st(r['status']), r['wirs'] or '–'], tag_of(r['status'])))
        self.t['rb'].set_rows(rows)
        # الأسفلت
        self.views['asp'] = self.visible(self.items.get('asp', []), 'asp'); rows = []
        for i, r in enumerate(self.views['asp']):
            q = r['q']
            th = '–' if r['asph'] is None else (f"{r['asph']:g} ~" if r.get('asph_def') else f"{r['asph']:g}")
            rows.append(([i + 1, q['dist'], q['street'], q['line'], q['dia_raw'], core.sta(r['x0']), core.sta(r['x1']), f2(r['L']),
                          f2(r['w']), f2(r['AS'], 2), th, r['w1'] or '–', f2(r['q1']), f2(r['AV'], 2), r['w2'] or '–', f2(r['q2']),
                          r['w3'] or '–', f2(r['q3']), r['wmc'] or '–', r['wrc'] or '–', self.st(r['status'])], tag_of(r['status'])))
        self.t['asp'].set_rows(rows)
        self.fill_summary(); self.fill_totals(); self.fill_notes()

    def fill_summary(self):
        rows = self.p.summary_rows(self.items)
        L = self.hdr['letters']
        SUMC = ('J', 'K', 'V', 'W', 'X', 'Y', 'AA', 'AC', 'AD', 'AE', 'AF', 'AH', 'AJ', 'AN', 'AP', 'AR', 'AU', 'AX', 'AY')
        tot = {c: 0.0 for c in SUMC}; out = []
        for R in rows:
            vals = []
            for c in L:
                v = R.get(c)
                if c in ('E', 'F'): v = core.sta(v)
                elif c in ('AG', 'AI', 'AK', 'AM') and v is not None: v = f'{v * 100:.0f}%'
                elif c in ('AL', 'AD') and v is not None: v = f2(v, 3)
                elif c in ('AT', 'AW') and v is not None: v = f'{v:g}'
                elif isinstance(v, float): v = f2(v)
                vals.append('' if v is None else v)
                if c in tot and isinstance(R.get(c), float): tot[c] += R[c]
            out.append((vals, 'ok'))
        out.append(([('الإجمالي' if c == 'B' else (f2(tot[c]) if c in tot else '')) for c in L], 'tot'))
        NF = self.tsf.columnCount()
        self.tsf.set_rows([(v[:NF], t) for v, t in out]); self.ts.set_rows([(v[NF:], t) for v, t in out])

    def fill_totals(self, *_):
        if not hasattr(self, 'p'): return
        item = self.tab_key.get(self.tabs.currentIndex(), 'sum')
        clear_layout(self.kbar)
        if item == 'none': return
        try: tots = self.p.totals(item, self.views.get(item, []))
        except Exception: traceback.print_exc(); tots = []
        for lbl, val in tots: self.kbar.addWidget(kpi_box(lbl, val))
        self.kbar.addStretch(1)

    def fill_notes(self):
        self.tx.setPlainText('\n'.join(self.notes) if self.notes else 'No notes.')
        self.tabs.setTabText(self.i_notes, f'Notes ({len(self.notes)})')

    def on_select(self, i):
        if i < 0 or i >= len(self.view): return
        r = self.view[i]
        if r['avg'] is None or not r['w']:
            self.l_det.setText(f"WIR {r['q']['no']} – {core.sta(r['x0'])} → {core.sta(r['x1'])}  |  {r['status']}: enter Depth Start and Depth End to calculate.")
            return
        names = ['0–1.5', '1.5–2.5', '2.5–3.5', '>3.5']
        parts = [f"{n}: {r['L']:.3f}×{r['w']:.3f}×{h:.3f} = {v:.3f}" for n, h, v in zip(names, r['h'], r['v']) if h > 0]
        self.l_det.setText(f"WIR {r['q']['no']} – {core.sta(r['x0'])} → {core.sta(r['x1'])}  |  Avg depth ({r['d0']:.3f}+{r['d1']:.3f})/2 = "
                           f"{r['avg']:.3f}  |  " + '  |  '.join(parts) + f"  |  Total {r['total']:.3f} m³   [{self.st(r['status'])}]")

    def tab_changed(self, i):
        self.fill_totals()
        if i == self.i_wk and not self._wk_auto and hasattr(self, 'p'):
            self._wk_auto = True; self.wk_fill_dates(); self.wk_auto_pick()
        if i == self.i_gc: self.charts_draw(True)

    # ------------------------------------------------------------ الكتابة في الخانات الصفرا
    def apply_value(self, row, cid, txt):
        self._last_edit = time.time()
        try: r = self.view[row]
        except IndexError: return False
        key = r['q']['key']
        if cid == 'soil':
            if txt.lower().startswith('r'): txt = core.ROCK
            elif txt.lower().startswith('o'): txt = core.OTHER
            if txt in (core.ROCK, core.OTHER) and txt != r['soil']:
                self.p.set_soil(key, r['x0'], r['x1'], txt); self._changed()
            return True
        if txt == '':
            val = None
        else:
            try: val = float(txt.replace(',', '.').replace('~', '').strip())
            except ValueError:
                QMessageBox.warning(self, 'QS', 'Please type a number'); return False
            if cid == 'asph':
                if 0 < val < 1: val = round(val * 100, 2)
                if not 0 <= val <= 50: QMessageBox.warning(self, 'QS', 'Asphalt thickness: number in cm (e.g. 5 or 7)'); return False
            elif not 0 <= val <= 15:
                QMessageBox.warning(self, 'QS', 'Depth must be a number between 0 and 15 m'); return False
        if cid == 'asph':
            if val != r['asph'] or r.get('asph_def'): self.p.set_asph(key, r['x0'], r['x1'], val); self._changed()
            return True
        x = r['x0'] if cid == 'd0' else r['x1']
        old = r['d0'] if cid == 'd0' else r['d1']
        if val != old or r['interp'][0 if cid == 'd0' else 1]:
            self.p.set_depth(key, x, val); self._changed()
        return True

    def _changed(self):
        self._last_edit = time.time()
        self.recompute()
        self.info('● changes not saved yet – saving by itself…', AMB)

    def autosave(self, manual=False):
        p = getattr(self, 'p', None)
        if p is None or not getattr(p, 'dirty', False): return True
        if self.tv.state() == QAbstractItemView.State.EditingState: return True
        if not manual and time.time() - self._last_edit < 2: return True
        try:
            p.save_points(); self.info(f'Saved {dt.datetime.now():%H:%M:%S}'); return True
        except PermissionError:
            self.info('QS_INPUT.xlsx is open in Excel – NOT saved. Close it, saving will retry by itself.', RED)
        except Exception as e:
            self.info(f'Save failed: {e}', RED)
        return False

    def on_close(self):
        p = getattr(self, 'p', None)
        while p is not None and p.dirty and not self.autosave(manual=True):
            a = QMessageBox.question(self, 'QS', 'Levels could not be saved (QS_INPUT.xlsx open in Excel?).\n'
                                     'Yes = try again   No = close without saving   Cancel = stay',
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if a == QMessageBox.Cancel: return False
            if a == QMessageBox.No: break
        return True

    # ------------------------------------------------------------ Unifier / القوايم
    def refresh_extra(self):
        S, p = self.S, getattr(self, 'p', None)
        if not S or p is None: return
        try:
            self.u_data = S.unifier_rows(p, p.summary_rows(self.items), self.items)
        except Exception as e:
            self.u_data = []; print('unifier', e)
        rows = []
        for r in self.u_data:
            v = [r['n'], r['line'], f"{r['x0']:g}", f"{r['x1']:g}", r['need'] or '']
            for gi, (txt, miss, cnt, gap) in enumerate(r['groups'], 1):
                inv = S.inv_of(p, r['line'], r['dist'], r['x0'], r['x1'], gi)
                v += ([txt, (f'{cnt} ⚠' if gap else cnt) or '', inv] if S.GROUPS[gi - 1][1] == 'BF' else [txt, inv])
            tag = 'warn' if any(g[3] for g in r['groups']) else ('miss' if any(g[1] for g in r['groups']) else 'ok')
            rows.append((v, tag))
        self.ug.set_rows(rows)
        for k in self.lists: self.refresh_list(k)

    def _extra_changed(self):
        if getattr(self, 'p', None) is not None: self.p.dirty = True
        self._last_edit = time.time()
        self.info('● changes not saved yet – saving by itself…', AMB)

    def _u_edit(self, i, cid, txt):
        r = self.u_data[i]; gi = int(cid[1:])
        k = (r['line'], r['dist'], round(r['x0'], 3), round(r['x1'], 3), gi)
        inv = self.p.extra.setdefault('INV', {})
        if txt: inv[k] = txt
        else: inv.pop(k, None)
        self._extra_changed(); self.refresh_extra()

    def refresh_list(self, key):
        L = self.lists[key]; data = self.p.extra.setdefault(key, [])
        rows = []
        for i, r in enumerate(data, 1):
            vals = []
            for c, *_ in L['spec']:
                if c == 'n': vals.append(i)
                elif c in L['calc']: vals.append(L['calc'][c](r))
                else: vals.append(r.get(c, ''))
            rows.append((vals, 'ok'))
        L['g'].set_rows(rows); L['lbl'].setText(L['tot'](data))

    def _l_edit(self, key, i, cid, txt):
        self.p.extra[key][i][cid] = txt; self._extra_changed(); self.refresh_list(key)

    def _l_add(self, key):
        if getattr(self, 'p', None) is None: return
        data = self.p.extra.setdefault(key, [])
        new = dict(data[-1]) if data else {}
        for k in list(new):
            if k.startswith(('un', 'inv')) or k in ('station', 'x0', 'x1', 'un', 'inv'): new[k] = ''
        data.append(new); self._extra_changed(); self.refresh_list(key)
        g = self.lists[key]['g']; g.setCurrentCell(len(data) - 1, 1)

    def _l_del(self, key):
        g = self.lists[key]['g']; r = g.currentRow()
        if r < 0 or getattr(self, 'p', None) is None: return
        if QMessageBox.question(self, 'QS', 'Delete the selected row?') != QMessageBox.Yes: return
        del self.p.extra[key][r]; self._extra_changed(); self.refresh_list(key)

    # ------------------------------------------------------------ التحديث (الشركة / المالك)
    def _upd(self):
        try:
            import upd_dates
            return upd_dates
        except Exception:
            traceback.print_exc(); return None

    def _paint_toggles(self, btns, cur):
        for k, b in btns.items(): b.setChecked(k == cur)

    def wk_set_track(self, k):
        self.wtrack = k; UI['wk_track'] = k; save_ui()
        self.wk_fill_dates(); self.wk_auto_pick()

    def wk_set_period(self, k):
        self.wper = k; self._paint_toggles(self.wk_pbtn, k); self.wk_compare()

    def wk_fill_dates(self):
        self._paint_toggles(self.wk_tbtn, self.wtrack); self._paint_toggles(self.wk_pbtn, self.wper)
        U = self._upd()
        if not U: return
        tr = self.wtrack
        try: rows = U.load()
        except Exception: rows = []
        self._wk_list = U.dates(tr, rows)
        try: hist = self.p.load_history(tr) if hasattr(self, 'p') else {}
        except Exception: hist = {}
        today = dt.date.today()
        nxt = next((r['d1'] for r in self._wk_list if r['d1'] >= today and r['d1'] not in hist), None)
        vals = []
        for r in self._wk_list:
            d = r['d1']
            st = 'saved ✔' if d in hist else 'due today' if d == today else 'not saved' if d < today else 'next' if d == nxt else ''
            name = f"Period {r['no']}" if tr == U.COMPANY else f"Week {r['no']}"
            vals.append(f"{name} · {U.day_name(d)} {d:%d-%m-%Y}" + (f' · {st}' if st else ''))
        self.cb_wsel.blockSignals(True); self.cb_wsel.clear(); self.cb_wsel.addItems(vals)
        d = self.cutoff()
        i = next((k for k, r in enumerate(self._wk_list) if r['d1'] == d), None)
        if i is not None: self.cb_wsel.setCurrentIndex(i)
        else:
            msg = (f'No company calendar for {U.MONTH_EN[today.month - 1]} – put it in 00_INBOX'
                   if tr == U.COMPANY and not U.month_rows(rows, tr, today.year, today.month) else '— choose the date —')
            self.cb_wsel.insertItem(0, msg); self.cb_wsel.setCurrentIndex(0)
        self.cb_wsel.blockSignals(False)

    def wk_auto_pick(self):
        lst = self._wk_list
        if not lst: return self.wk_compare()
        d, today = self.cutoff(), dt.date.today()
        r = next((x for x in lst if x['d1'] == d), None) or next((x for x in lst if x['d1'] >= today), None) or lst[-1]
        self.wk_pick_date(r['d1'])

    def wk_pick(self):
        t = self.cb_wsel.currentText()
        i = next((k for k, r in enumerate(self._wk_list) if f"{r['d1']:%d-%m-%Y}" in t), None)
        if i is not None: self.wk_pick_date(self._wk_list[i]['d1'])

    def wk_pick_date(self, d):
        if self.cutoff() == d: self.basis_changed(); self.wk_compare()
        else: self.e_cut.setDate(qd(d))                     # ← basis_changed ← wk_compare

    def wk_mode(self): return 'requested' if self.cb_basis.currentText() == 'Requested' else 'approved'

    def wk_edit_dates(self):
        U = self._upd()
        if not U: return
        d = self.cutoff(); y, m = d.year, d.month
        try: rows = U.load()
        except PermissionError: return QMessageBox.warning(self, 'QS', '12-LOG UPDATE DATES.xlsx is open in Excel. Close it and try again.')

        def loader(yy, mm):
            company = U.month_rows(rows, U.COMPANY, yy, mm)
            warn = [] if company else ['No company calendar saved for this month – the usual pattern is filled in. '
                                       'Put the company calendar in 00_INBOX (DC reads it), or type the dates here.']
            return company or U.company_default(yy, mm), U.month_rows(rows, U.OWNER, yy, mm) or U.owner_default(yy, mm), warn
        cal = None
        try:
            fol = core.project_path(U.CAL_FOLDER_REL)
            cal = next((os.path.join(fol, n) for n in sorted(os.listdir(fol), reverse=True) if n.startswith(f'{y}-{m:02d}')), None)
        except Exception: pass
        import tkinter as tk                                 # شباك مراجعة التواريخ لسه Tk – بيفتح فوق ويقفل لما تحفظ
        root = tk.Tk(); root.withdraw()
        try: got = U.review(root, y, m, loader, original=cal, title='Update Dates', auto_open=False)
        finally:
            try: root.destroy()
            except Exception: pass
        if not got: return
        yy, mm, company, owner = got
        try: U.save_month(yy, mm, company, owner)
        except PermissionError: return QMessageBox.critical(self, 'QS', '12-LOG UPDATE DATES.xlsx is open in Excel. Close it and try again.')
        self.wk_fill_dates()
        QMessageBox.information(self, 'QS', f'Update dates saved for {yy}-{mm:02d}:\ncompany {len(company)} periods, owner {len(owner)} dates.')

    def wk_compare(self):
        if not hasattr(self, 'p'): return
        d = self.cutoff(); tr = self.wtrack or ''
        prow, since = None, None
        U = self._upd()
        if U and self.wper != 'Monthly':
            try: regs = U.load()
            except Exception: regs = []
            lst = U.month_rows(regs, tr, d.year, d.month) or (U.owner_default(d.year, d.month) if tr == U.OWNER else [])
            prow = next((r for r in lst if r['d1'] == d), None)
            since = prow['d0'] if prow and prow.get('d0') else None
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            cm = self.p.compare(d, self.wk_mode(), 'month' if self.wper == 'Monthly' else 'week', track=tr or None, since=since)
        except PermissionError:
            QGuiApplication.restoreOverrideCursor(); return QMessageBox.critical(self, 'QS', 'QS_HISTORY.xlsx is open in Excel. Close it and try again.')
        except Exception as e:
            QGuiApplication.restoreOverrideCursor(); traceback.print_exc(); return self.info(f'Compare failed: {e}', RED)
        QGuiApplication.restoreOverrideCursor()
        self.wk_cmp = cm
        try: self.wk_fill_dates()
        except Exception: traceback.print_exc()
        try: rp = core.reported_at(d, cm['rows'])
        except Exception: traceback.print_exc(); rp = dict(items={}, value=0.0, n=0, loose=[])
        cm['reported'] = rp
        comp = tr == 'Company'
        try: ccodes = core.company_codes() if comp else {}
        except Exception: traceback.print_exc(); ccodes = {}
        nocode = []
        for r in cm['rows']:
            r['cc'] = ccodes.get(core._item_key(r['item']), '')
            if comp and not r['cc'] and (abs(r['cq'] or 0) > 1e-9 or abs(r['pq'] or 0) > 1e-9): nocode.append(str(r['item']))
        ids = [c[0] for c in self.WK_COLS]
        hide = ('code', 'cv', 'pv', 'dv') if comp else ('rq', 'tq')
        for j, c in enumerate(ids): self.tw.setColumnHidden(j, c in hide)
        self.tw.horizontalHeaderItem(0).setText('Company Code' if comp else 'Item No.')

        def rq(r):
            q = rp['items'].get(r['item'])
            return '' if q is None else f2(q) + ('  ▲' if q - (r['cq'] or 0) > 0.005 else '')
        tq = lambda r: max((r['cq'] or 0) - (rp['items'].get(r['item']) or 0), 0.0)
        tq_val = sum(tq(r) * (r['rate'] or 0) for r in cm['rows']); cm['to_report'] = tq_val
        rows = []
        for r in cm['rows']:
            if abs(r['cq'] or 0) < 1e-9 and abs(r['pq'] or 0) < 1e-9 and r['item'] not in rp['items']: continue
            rows.append(([r['cc'] if comp else r['item'], r['code'], r['desc'], r['unit'], f"{r['rate']:,.2f}" if r['rate'] else '–',
                          f2(r['cq']), rq(r), f2(tq(r)), f"{r['cv']:,.2f}", f2(r['pq']), f"{r['pv']:,.2f}", f2(r['dq']), f"{r['dv']:,.2f}"],
                         'chg' if abs(r['dv']) > 0.005 else 'ok'))
        dv = cm['cur_total'] - cm['prev_total']
        if comp:
            rows.append((['', '', 'TOTAL VALUE (SAR)', '', '', f"{cm['cur_total']:,.2f}", f"{rp['value']:,.2f}" if rp['n'] else '',
                          f'{tq_val:,.2f}', '', f"{cm['prev_total']:,.2f}", '', f'{dv:,.2f}', ''], 'tot'))
            if rp['n']:
                rows.append((['', '', 'NET – after items reported above the calculated', '', '', '', '',
                              f"{cm['cur_total'] - rp['value']:,.2f}", '', '', '', '', ''], 'net'))
        else:
            rows.append((['', '', 'TOTAL', '', '', '', '', '', f"{cm['cur_total']:,.2f}", '', f"{cm['prev_total']:,.2f}", '', f'{dv:,.2f}'], 'tot'))
        self.tw.set_rows(rows)
        for j in range(self.tw.columnCount()):
            it = self.tw.item(0, j) if self.tw.rowCount() else None
        di = ids.index('desc')
        for i in range(self.tw.rowCount()):
            it = self.tw.item(i, di)
            if it: it.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.fill_history()
        pd_ = cm['prev_date']
        day = lambda x: ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')[x.weekday()] + x.strftime('  %d-%m-%Y')
        exact = bool(prow and pd_ and since and pd_ == since - dt.timedelta(days=1))
        days, dlbl = (prow['days'], 'Working Days') if exact and prow.get('days') else (cm['days'], 'Days')
        what = (f"{tr}  ·  {'Period' if tr == 'Company' else 'Week'} {prow['no']}:   {day(prow['d0'])}   →   {day(d)}" if prow
                else f"{tr} update  {day(d)}")
        if self.wper == 'Monthly':
            base = f"compared with end of last month:  {day(pd_)}" if pd_ else 'no update saved before this month'
        elif pd_:
            base = f"compared with the saved update of  {day(pd_)}" + ('' if exact or not since else f"   (nothing saved on {day(since - dt.timedelta(days=1))})")
        else:
            base = 'no update saved before it'
        rnote = (f"     |   Reported to Company: {rp['n']} report(s) {rp['first']:%d-%m} → {rp['last']:%d-%m}" if rp['n'] else
                 f"     |   Reported to Company: no production report in {core.PROD_REL} up to this date")
        if rp.get('loose'): rnote += f"   ·   not linked: {'; '.join(rp['loose'][:3])}"
        if nocode: rnote += f"   ·   no company code in {core.PLAN_REL}: {', '.join(nocode[:5])}"
        self.l_wnote.setText(f"{what}     {base}" + ('     ✔ saved' if cm['saved'] else '     ●  not saved yet – press Save this update') + rnote)
        cm['days'], cm['label'] = days, what
        clear_layout(self.wkbar)
        kp = [('Previous Value', f"{cm['prev_total']:,.0f} SAR"), ('Current Value', f"{cm['cur_total']:,.0f} SAR"),
              ('Value This Period', f"{dv:,.0f} SAR")]
        if days: kp += [(dlbl, f'{days}'), ('Average per Day', f'{dv / days:,.0f} SAR')]
        if rp['n']: kp += [('Reported to Company', f"{rp['value']:,.0f} SAR")]
        if comp: kp += [('To Report to CTR', f'{tq_val:,.0f} SAR')]
        for l, v in kp: self.wkbar.addWidget(kpi_box(l, v))
        self.wkbar.addStretch(1)

    def fill_history(self):
        try: ht = self.p.history_totals(self.wk_mode(), self.wtrack)
        except Exception: ht = []
        rw = []
        for i, (d, v) in enumerate(ht):
            dv = v - ht[i - 1][1] if i else v
            rw.insert(0, ([f'{d:%d-%m-%Y}', f'{v:,.0f}', f'{dv:,.0f}', (d - ht[i - 1][0]).days if i else ''], 'ok'))
        self.th_w.set_rows(rw)
        months = {}
        for d, v in ht: months[(d.year, d.month)] = (d, v)
        rm, last = [], 0.0
        for k in sorted(months):
            d, v = months[k]; rm.insert(0, ([f'{k[1]:02d}-{k[0]}', f'{d:%d-%m-%Y}', f'{v:,.0f}', f'{v - last:,.0f}'], 'ok')); last = v
        self.th_m.set_rows(rm)

    def wk_save(self):
        if not hasattr(self, 'p'): return
        d, tr = self.cutoff(), self.wtrack
        hist = self.p.load_history(tr)
        if d in hist and QMessageBox.question(self, 'QS', f'A {tr} update for {d:%d-%m-%Y} is already saved. Replace it?') != QMessageBox.Yes: return
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            path = self.p.save_update(d, tr)
        except PermissionError:
            QGuiApplication.restoreOverrideCursor(); return QMessageBox.critical(self, 'QS', 'QS_HISTORY.xlsx is open in Excel. Close it and try again.')
        except Exception as e:
            QGuiApplication.restoreOverrideCursor(); return QMessageBox.critical(self, 'QS', f'Save failed:\n{e}')
        QGuiApplication.restoreOverrideCursor()
        self.recompute(); self.wk_compare()
        QMessageBox.information(self, 'QS', f'{tr} update {d:%d-%m-%Y} saved (Approved + Requested) to:\n{path}')

    def _ipc(self):
        p = os.path.join(MODS, 'ipc')
        if p not in sys.path: sys.path.append(p)
        import ipc_core
        return ipc_core

    def contract_value(self):
        if self._contract is None:
            v = 0.0
            try:
                ipc = self._ipc()
                items, _ = ipc.template_layout(core.project_path(ipc.TEMPLATE_REL))
                divs = {n for n, _ in ipc.DIVISIONS}
                v = sum((it['qty'] or 0) * (it['rate'] or 0) for it in items if it['div'] in divs)
            except Exception:
                traceback.print_exc()
            self._contract = v
        return self._contract or None

    def wk_export(self):
        if not self.wk_cmp: self.wk_compare()
        if not self.wk_cmp: return
        cm = self.wk_cmp; d = cm['date']; tr = self.wtrack or 'Company'; mode = self.wk_mode()
        try: n = self.p.update_number(d, tr)
        except Exception: n = 1
        n, ok = QInputDialog.getInt(self, 'Export update', f'{tr} update  ·  {mode.capitalize()}  ·  {d:%d-%m-%Y}\n\n'
                                    f'Update number (running, from the first {tr.lower()} update):', n, 1, 999)
        if not ok: return
        name = f"{n:02d}-UQ-{tr}-{mode.capitalize()}-{d:%d-%m-%Y}"
        base = next((p for p in (os.path.join(core.QS_DIR, 'QS', 'UP-Q'), os.path.join(core.QS_DIR, 'UP-Q')) if os.path.isdir(p)), core.QS_DIR)
        folder = os.path.join(base, name)
        xlsx, pdf = os.path.join(folder, name + '.xlsx'), os.path.join(folder, name + '.pdf')
        if os.path.exists(xlsx) and QMessageBox.question(self, 'QS', f'{name} is already there. Replace it?') != QMessageBox.Yes: return
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            os.makedirs(folder, exist_ok=True)
            import qs_charts
            ht, hg = self.p.history_series(mode, tr)
            if not cm['saved']:
                ht = [x for x in ht if x[0] != d] + [(d, cm['cur_total'])]
                hg = [x for x in hg if x[0] != d] + [(d, qs_charts.group_values((r['item'], r['cv']) for r in cm['rows']))]
            self.p.export_compare(cm, mode, xlsx, ht=ht, hg=hg, contract=self.contract_value(), names=self.upd_names(tr, d.year, d.month),
                                  title=f'{tr} updates  ·  {mode.capitalize()}  ·  up to {d:%d-%m-%Y}')
        except PermissionError:
            QGuiApplication.restoreOverrideCursor(); return QMessageBox.critical(self, 'QS', 'The file is open in Excel. Close it and try again.')
        warn = ''
        try:
            self._ipc().export_pdf(xlsx, ['Weekly Update', 'Monthly Update', 'Charts'], pdf=pdf)
        except Exception as e:
            warn = f'\n\nThe PDF was not made:\n{e}'
        QGuiApplication.restoreOverrideCursor()
        self.raise_(); self.activateWindow()
        files = f'{name}.xlsx' + ('' if warn else f'\n{name}.pdf')
        if QMessageBox.question(self, 'QS', f'Update exported to:\n{folder}\n\n{files}{warn}\n\nOpen the folder?') == QMessageBox.Yes:
            os.startfile(folder)

    # ------------------------------------------------------------ الرسومات
    def gc_set_track(self, k):
        self.wk_set_track(k); self.charts_draw(True)

    def gc_set_period(self, k):
        self.gc_kind = k; UI['chart_period'] = k; save_ui(); self.charts_draw()

    def charts_draw_later(self):
        if hasattr(self, '_gc_job'): self._gc_job.start(120)

    def upd_names(self, tr, y, m):
        U = self._upd()
        if not U: return {}
        try: rows = U.dates(tr, U.load(), today=dt.date(y, m, 15), back=0, ahead=0)
        except Exception: return {}
        return {r['d1']: ('P' if tr == U.COMPANY else 'W') + str(r['no']) for r in rows if r.get('d1')}

    def charts_draw(self, reload=False):
        if not hasattr(self, 'gc_done') or not hasattr(self, 'p'): return
        import qs_charts as C
        tr, mode = self.wtrack, self.wk_mode()
        self._paint_toggles(self.gc_tbtn, tr); self._paint_toggles(self.gc_pbtn, self.gc_kind)
        hp = str(self.p.S.get('HISTORY_PATH'))
        try: mt = os.path.getmtime(hp)
        except OSError: mt = None
        today = dt.date.today(); key = (tr, mode, hp, mt, today)
        if reload or not self._gc_data or self._gc_data[0] != key:
            try: ht, hg = self.p.history_series(mode, tr)
            except Exception: traceback.print_exc(); ht, hg = [], []
            mon = C.month_of(ht, (today.year, today.month))
            rt = None
            if tr == 'Company':
                try: rt = core.reported_series()
                except Exception: traceback.print_exc(); rt = []
            self._gc_data = (key, ht, hg, mon, self.upd_names(tr, *mon) if mon else {}, rt)
        _, ht, hg, mon, names, rt = self._gc_data
        k = self.gc_kind; sub = f'{tr} updates  ·  {mode.capitalize()}'
        self.l_gc.setText(f'{len(ht)} saved {tr.lower()} updates' + (f'  ·  last {ht[-1][0]:%d-%m-%Y}' if ht else ''))
        try:
            if k == 'U':
                mname = f'{C.MON[mon[1] - 1]} {mon[0]}' if mon else ''
                ub, before = C.update_bars(ht, mon, names) if mon else ([], None)
                rmap = None
                if rt is not None:
                    rmap, pts = {}, sorted(ht, key=lambda x: x[0])
                    rc = lambda x: max([v for d_, v in rt if d_ <= x] or [0.0])
                    for i, (d, _v) in enumerate(pts):
                        p_ = pts[i - 1][0] if i else dt.date.min
                        if any(p_ < d_ <= d for d_, _ in rt):
                            lab = next((u[0] for u in ub if u[4] == d.strftime('%d-%m')), None)
                            if lab: rmap[lab] = rc(d) - rc(p_)
                C.draw_done(self.gc_done, C.with_change(ub, before), f'{mname} updates – value done since the update before (SAR)',
                            sub + '  ·  above each bar: change from the update before', rep=rmap)
                C.draw_stack(self.gc_stack, C.update_groups(hg, mon, names) if mon else [], f'{mname} updates – by type of work (SAR)', sub)
            else:
                name = dict(C.KINDS)[k].lower()
                rmap = {b_[0]: b_[1] for b_ in C.buckets(rt, k, today)} if rt is not None else None
                C.draw_done(self.gc_done, C.with_change(C.buckets(ht, k, today)), f'Value done – {name} (SAR)',
                            sub + '  ·  above each bar: change from the period before  ·  the period not finished yet: in progress', rep=rmap)
                C.draw_stack(self.gc_stack, C.group_buckets(hg, k, today), f'By type of work – {name} (SAR)', sub)
            C.draw_curve(self.gc_curve, ht, self.contract_value(), 'Progress – cumulative value (SAR)',
                         sub + '  ·  a point for each saved update' + ('  ·  orange: each production report' if rt is not None else ''), rep=rt)
        except Exception:
            traceback.print_exc()
        for cv in (self.gc_done, self.gc_curve, self.gc_stack): cv.update()

    # ------------------------------------------------------------ التصدير
    def _export(self, path=None):
        try:
            return self.p.export(self.items, self.notes, path)
        except PermissionError:
            QMessageBox.critical(self, 'QS', 'The output file is open in Excel.\nClose it and try again.')
        except Exception as e:
            QMessageBox.critical(self, 'QS', f'Export failed:\n{e}\n\n{traceback.format_exc()[-600:]}')
        return None

    def export(self):
        if not hasattr(self, 'p'): return
        out = str(self.p.S.get('OUTPUT_PATH'))
        path, _ = QFileDialog.getSaveFileName(self, 'Export QS', out, 'Excel (*.xlsx)')
        if not path: return
        if self._export(path) and QMessageBox.question(self, 'QS', f'Exported:\n{path}\n\nOpen it now?') == QMessageBox.Yes:
            os.startfile(path)

    def preview(self):
        if not hasattr(self, 'p'): return
        path = self._export()
        if path: os.startfile(path)
