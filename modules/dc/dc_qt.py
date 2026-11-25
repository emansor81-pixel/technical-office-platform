# -*- coding: utf-8 -*-
"""Distribute (DC) بـ Qt – نفس شاشة dc_app.py بالظبط في الشغل، والشكل الجديد بس.

ملفات الشغل زي ما هي: dc_core / dc_rules / dc_read / dc_sheet / dc_claude.
- إذن التوريد (DN) وتقويم تحديث الشركة (UPD_CAL) شبابيك مراجعتهم لسه Tkinter ← بيفتحوا في شاشة التوزيع القديمة.
- الشغل التقيل (القراية والشيت) في thread، والشاشة بتتحدّث من الطابور كل 150ms (زي القديم).
"""
import os, re, sys, queue, shutil, hashlib, threading, subprocess, datetime as dt

from PySide6.QtCore import Qt, QTimer, QRectF, QSize
from PySide6.QtGui import QColor, QPainter, QFontMetricsF, QGuiApplication
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout, QTableWidget,
                               QTableWidgetItem, QHeaderView, QAbstractItemView, QComboBox, QLineEdit, QMessageBox,
                               QStyledItemDelegate, QDialog, QCheckBox, QPlainTextEdit, QFileDialog, QApplication, QInputDialog)

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import dc_core as D
from dc_io import LOG
import dc_rules as C
import dc_read as R
import dc_sheet as SH
import dc_claude as CL
try:
    from dc_preview import PreviewPane
except Exception:                                   # من غير الـ preview الشاشة تفضل شغالة
    PreviewPane = None

PLAT = os.path.dirname(os.path.dirname(HERE))
if PLAT not in sys.path: sys.path.insert(0, PLAT)
import home_qt as HQ

NAVY, BG, INK, MUTED, SOFT = HQ.NAVY, HQ.BG, HQ.INK, HQ.MUTED, HQ.SOFT_LINE
GREEN = '#7FD98A' if HQ.DARK else '#2C7536'
TT = HQ.T
COLS = [('file', 'FILE', 250), ('st', 'STATUS', 120), ('type', 'UNDERSTOOD AS', 190), ('no', 'NO.', 70), ('date', 'DATE', 100),
        ('folder', 'GOES TO', 300), ('name', 'NEW NAME', 330)]
CODES = list(C.TYPES)
# قايمة Type مجموعات بعناوين (طلب المستخدم 21-11-2026) – أي نوع جديد مش مكتوب هنا بينزل تحت «OTHERS» لوحده
TYPE_GROUPS = [('LETTERS', ('IN_CLIENT', 'IN_YA', 'IN_SUB', 'IN_OTH', 'OUT_YA', 'OUT_CLIENT', 'OUT_SUB', 'OUT_OTH')),
               ('SUBMITTALS', ('MAR', 'SD', 'MS', 'ITP', 'RFI')),
               ('QUALITY', ('SITE_WIR', 'UNF_WIR', 'MIR', 'NCR', 'TEST', 'CAL')),
               ('PARTIES  ·  subcontractors / suppliers', ('SUB', 'QUOTE', 'SUB_CONTRACT', 'SUB_IPC')),
               ('PAYMENTS', ('IPC', 'IPC_INV', 'IPC_SIGNED', 'PAYROLL', 'GIS', 'PRICE')),
               ('SITE', ('DN', 'PROD', 'PERMIT', 'UPD_CAL'))]


def type_rows():
    """[(code أو None للعنوان, الكلام)] بترتيب المجموعات."""
    rows, used = [], set()
    for title, codes in TYPE_GROUPS:
        cs = [c for c in codes if c in C.TYPES and c not in used]
        if not cs: continue
        rows.append((None, f'──  {title}  ──'))
        for c in cs: rows.append((c, f'     {C.EN.get(c, c)}  ({c})')); used.add(c)
    rest = [c for c in CODES if c not in used]
    if rest:
        rows.append((None, '──  OTHERS  ──'))
        rows += [(c, f'     {C.EN.get(c, c)}  ({c})') for c in rest]
    return rows
OLD_ONLY = {'DN': 'Delivery note', 'UPD_CAL': 'Company update calendar'}   # شبابيك مراجعتهم لسه Tkinter
PILL = {'dup': ('#EDE7F6', '#5E35B1'), 'sure': ('#E2EFDA', GREEN), 'ask': ('#FFF2CC', '#8A6100'), 'err': ('#FBE3E3', '#C00000'), 'wait': ('#EEF1F4', '#7F8A99')}
if HQ.DARK:
    PILL = {'dup': ('#2E2445', '#C3A6FF'), 'sure': ('#1F3A24', GREEN), 'ask': ('#3A3018', '#F0C860'), 'err': ('#3E1E22', '#FF8A8A'), 'wait': ('#222E3A', '#8C9AAA')}
C_ASK, C_ERR, C_DUP, C_DIM, C_OK = ((PILL['ask'][1], PILL['err'][1], PILL['dup'][1], '#7C8896', '#8FD18F') if HQ.DARK
                                   else ('#8A6100', '#C00000', '#5E35B1', '#7F7F7F', '#375623'))
ROLE_TAG = int(getattr(Qt.UserRole, 'value', Qt.UserRole)) + 1

CSS = f"""
QWidget#dist {{ background: {BG}; }}
QLabel, QCheckBox {{ color: {INK}; }}
QDialog {{ background: white; }}
QFrame#card {{ background: white; border: 1px solid {SOFT}; border-radius: 14px; }}
QPushButton {{ background: #F6F8FB; border: 1px solid {SOFT}; border-radius: 9px; padding: 7px 14px;
              font: 600 10pt "Segoe UI"; color: {INK}; }}
QPushButton:hover {{ background: #EAF1F8; border-color: #C9D6E3; }}
QPushButton:pressed {{ background: #DCE7F2; }}
QPushButton#pri {{ background: {NAVY}; color: white; border: none; padding: 7px 18px; }}
QPushButton#pri:hover {{ background: #2E6496; }}
QPushButton#back {{ background: white; color: {NAVY}; font-weight: bold; }}
QTableWidget {{ background: white; color: {INK}; border: none; gridline-color: transparent; font: 10pt "Segoe UI";
               selection-background-color: #EAF1F8; selection-color: {INK}; outline: 0; }}
QTableWidget::item {{ border-bottom: 1px solid #EEF2F6; padding-left: 8px; }}
QHeaderView::section {{ background: #F8FAFC; color: {MUTED}; font: bold 8pt "Segoe UI"; border: none;
                       border-bottom: 1px solid {SOFT}; padding: 8px 10px; }}
QComboBox, QLineEdit {{ background: white; border: 1px solid #D4DCE5; border-radius: 8px; padding: 5px 9px;
                       font: 10pt "Segoe UI"; color: {INK}; min-height: 20px; }}
QComboBox:focus, QLineEdit:focus {{ border-color: {NAVY}; }}
QComboBox:disabled, QLineEdit:disabled {{ background: #F3F5F8; color: #9AA5B4; }}
QComboBox QAbstractItemView {{ background: white; border: 1px solid {SOFT}; selection-background-color: #EAF1F8;
                              selection-color: {NAVY}; outline: 0; padding: 4px; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: #CBD5E1; border-radius: 4px; min-height: 30px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
"""


# ------------------------------------------------------------------ المكررات (طلب المستخدم 12-11)
# نفس الملف بالظبط موجود في فولدره (أو فولدرات المراحل اللي جنبه) ← يروح 00_INBOX\_duplicates من غير أسئلة.
# نفس الرقم بس ورقة مختلفة ← يسأل: نسخة أحدث (القديمة لـ 4_ARCHIVED) ولا مكررة.
DUP = getattr(D, 'DUP_DIR', '_duplicates')        # 05_DUPLICATES (أو _duplicates لو لسه بالاسم القديم)
STAGE = re.compile(r'^\d_[A-Z_]+$', re.I)               # 1_WIP, 2_SHARED, 3_PUBLISHED, 4_ARCHIVED, 5_UNIFIER_APPROVED…


def C_num(v):
    """رقم من قيمة في JSON (أو 0)"""
    try: return float(str(v).replace(',', ''))
    except (TypeError, ValueError): return 0.0


def _sha(p):
    h = hashlib.sha1()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''): h.update(b)
    return h.hexdigest()


def _family(folder_rel):
    """فولدر الورقة + فولدرات المراحل اللي جنبه (لو هو مرحلة)."""
    f = os.path.normpath(os.path.join(D.project_root(), folder_rel or ''))
    if STAGE.match(os.path.basename(f)):
        base = os.path.dirname(f)
        try: return [os.path.join(base, n) for n in os.listdir(base) if STAGE.match(n) and os.path.isdir(os.path.join(base, n))]
        except OSError: return [f]
    return [f]


def _files(folder_rel):
    for d in _family(folder_rel):
        try: names = os.listdir(d)
        except OSError: continue
        for n in names:
            p = os.path.join(d, n)
            if os.path.isfile(p): yield p


def same_file(src, folder_rel):
    """نفس المحتوى بالظبط (حتى لو الاسم مختلف) ← مسار النسخة الموجودة، أو None."""
    try: size = os.path.getsize(src)
    except OSError: return None
    if size == 0: return None
    h = None
    for p in _files(folder_rel):
        try:
            if os.path.getsize(p) != size: continue
            h = h or _sha(src)
            if _sha(p) == h: return p
        except OSError:
            continue
    return None


def same_number(out, folder_rel):
    """ورق تاني بنفس رقم الصادر في فولدره / فولدرات المراحل."""
    if not out: return []
    pat = re.compile(r'(?<![A-Za-z0-9])' + re.escape(out) + r'(?!\d)', re.I)   # 0003 مايلقطش 00036
    return [p for p in _files(folder_rel) if pat.search(os.path.basename(p))]


def safe_move(p, q):
    """نقل من غير ما يفضل نسختين لو الملف مفتوح في برنامج تاني (زي D.approve)."""
    try:
        shutil.move(p, q)
    except PermissionError:
        if os.path.isfile(p) and os.path.isfile(q):
            try: os.remove(q)
            except OSError: pass
        raise PermissionError(f'The file is open in another program (PDF viewer?) – close it and try again.\n\n{p}')


def rel(p):
    try: return os.path.relpath(p, D.project_root())
    except ValueError: return p


if HQ.DARK:                                          # نفس الستايل بألوان الغامق
    for a, b in (('background: white', f'background: {TT["card2"]}'), ('#F6F8FB', '#1C2834'), ('#EAF1F8', '#1E3448'),
                 ('#C9D6E3', '#33506B'), ('#DCE7F2', '#284A66'), ('#2E6496', '#3A72A8'), ('#F8FAFC', '#1A2530'),
                 ('#EEF2F6', '#22303C'), ('#D4DCE5', '#33424F'), ('#F3F5F8', '#141C24'), ('#9AA5B4', '#5D6A78'),
                 ('#CBD5E1', '#3A4A5A')):
        CSS = CSS.replace(a, b)
    CSS += f"QPushButton#back {{{{ color: {TT['title']}; }}}} QPushButton#pri {{{{ color: white; }}}}".replace('{{', '{').replace('}}', '}')


# السطر المختار لازم يبان (طلب المستخدم 12-11): قاعدة ::item كانت بتغطي على selection-background-color
CSS += ('QTableWidget::item:selected { background: #2B5F8E; color: white; }' if HQ.DARK
        else f'QTableWidget::item:selected {{ background: #CFE0F1; color: {INK}; }}')


class PillDelegate(QStyledItemDelegate):
    """عمود الحالة: شارة ملونة (Sure / Check / Not understood …)."""
    def paint(self, p, opt, idx):
        super().paint(p, opt, idx)          # الخلفية (التحديد)
        text = idx.data(Qt.DisplayRole) or ''
        tag = idx.data(ROLE_TAG) or 'wait'
        if not text: return
        bg, fg = PILL.get(tag, PILL['wait'])
        p.save(); p.setRenderHint(QPainter.Antialiasing)
        f = HQ.font(12, True); p.setFont(f)
        w = QFontMetricsF(f).horizontalAdvance(text) + 20
        r = QRectF(opt.rect.left() + 8, opt.rect.center().y() - 11, min(w, opt.rect.width() - 12), 22)
        p.setPen(Qt.NoPen); p.setBrush(QColor(bg)); p.drawRoundedRect(r, 11, 11)
        p.setPen(QColor(fg)); p.drawText(r, Qt.AlignCenter, text)
        p.restore()

    def initStyleOption(self, opt, idx):
        super().initStyleOption(opt, idx)
        opt.text = ''                         # الكلام بيترسم جوه الشارة بس


def card():
    f = QFrame(); f.setObjectName('card')            # من غير ظل (الظل بيبهّت الكلام اللي جوّاه)
    return f


def btn(text, slot, primary=False):
    b = QPushButton(text)
    if primary: b.setObjectName('pri')
    b.setCursor(Qt.PointingHandCursor); b.clicked.connect(slot)
    return b


class DistQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        self.setObjectName('dist'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.items, self._next_id = {}, 1        # iid -> item dict
        self.busy, self._subs = False, []
        self._pcache, self._alias, self._add_for = {}, {}, None      # الجهات: نتيجة البحث بالاسم · أسامي اتربطت في الجلسة دي · خانة الإضافة معمولة لأنهي اسم
        self.q = queue.Queue()
        self._filling = False                     # بنملا الخانات من الكود ← متشغّلش type_changed
        self._build()
        self.timer = QTimer(self); self.timer.timeout.connect(self._poll); self.timer.start(150)
        QTimer.singleShot(600, self.refresh_next)
        QTimer.singleShot(1500, lambda: threading.Thread(target=D.clean_work, daemon=True).start())
        QTimer.singleShot(150, self.scan)

    # ------------------------------------------------------------ UI
    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 14, 22, 16); v.setSpacing(12)
        # العنوان
        t = QHBoxLayout(); t.setSpacing(14)
        b = btn('←  Document Center', lambda: self.open_section('doc')); b.setObjectName('back'); t.addWidget(b)
        tv = QVBoxLayout(); tv.setSpacing(0)
        c = QLabel('Technical Office Platform  ›  Document Center  ›  Distribute'); c.setFont(HQ.font(12.5, True))
        c.setStyleSheet(f'color: {MUTED};'); tv.addWidget(c)
        tt = QLabel(f'<span style="color:{GREEN}; font-size:20pt; font-weight:bold;">Distribute</span>'
                    f'<span style="color:{MUTED}; font-size:11pt;">  ·  papers in 00_INBOX</span>'); tv.addWidget(tt)
        t.addLayout(tv); t.addStretch(1)
        self.l_next = QLabel('Next YA letter: …'); self.l_next.setFont(HQ.font(13, True))
        self.l_next.setStyleSheet(f'color: {TT["title"]}; background: {TT["card2"]}; border: 1px solid {SOFT}; border-radius: 14px; padding: 5px 12px;')
        t.addWidget(self.l_next)
        self.l_claude = QLabel(''); self.l_claude.setFont(HQ.font(12.5, True))
        self.l_claude.setStyleSheet('color: #56687A; background: #EEF1F4; border-radius: 14px; padding: 5px 12px;')
        self.l_claude.hide()                             # حالة مفتاح الـ API – مش مستخدمة
        v.addLayout(t)
        # شريط الأزرار
        bar = card(); h = QHBoxLayout(bar); h.setContentsMargins(12, 10, 12, 10); h.setSpacing(8)
        def sep():
            s = QFrame(); s.setFixedSize(1, 24); s.setStyleSheet(f'background: {SOFT};'); return s
        h.addWidget(btn('⇩  Distribute', self.distribute, True))
        h.addWidget(btn('↻  Refresh list', self.scan))
        h.addWidget(btn('Open INBOX', lambda: self.open_path(D.inbox())))
        h.addWidget(sep())
        h.addWidget(btn('↶  Undo last move', self.undo)); h.addWidget(btn('Moves log', self.show_log))
        h.addWidget(sep())
        h.addWidget(btn('DC Register…', self.register))
        self.b_pend = btn('Update DC sheet', self.flush); h.addWidget(self.b_pend)
        # زرار Claude… (مفتاح الـ API) اتشال – المستخدم مش هيشترك في الـ API. «Send to Claude» (الشات) فاضل.
        h.addStretch(1)
        self.l_info = QLabel(''); self.l_info.setFont(HQ.font(12.5, True)); self.l_info.setStyleSheet(f'color: {C_OK};')
        self.l_info.setWordWrap(True); self.l_info.setMaximumWidth(620)
        self.l_info.setAlignment(Qt.AlignRight | Qt.AlignVCenter); h.addWidget(self.l_info)
        v.addWidget(bar)
        # الصف الرئيسي: (جدول + مربع التأكيد) على الشمال، والـ Preview على اليمين
        mainrow = QHBoxLayout(); mainrow.setSpacing(12)
        lcol = QVBoxLayout(); lcol.setSpacing(12)
        # الجدول
        tb = card(); tl = QVBoxLayout(tb); tl.setContentsMargins(1, 1, 1, 1)
        self.tv = QTableWidget(0, len(COLS))
        self.tv.setHorizontalHeaderLabels([c[1] for c in COLS])
        self.tv.verticalHeader().setVisible(False); self.tv.setShowGrid(False)
        self.tv.setSelectionBehavior(QAbstractItemView.SelectRows); self.tv.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tv.setEditTriggers(QAbstractItemView.NoEditTriggers); self.tv.setWordWrap(False)
        self.tv.setTextElideMode(Qt.ElideMiddle)
        self.tv.verticalHeader().setDefaultSectionSize(38)
        hh = self.tv.horizontalHeader(); hh.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter); hh.setHighlightSections(False)
        for i, (cid, _, w) in enumerate(COLS):
            self.tv.setColumnWidth(i, w)
        hh.setSectionResizeMode(len(COLS) - 1, QHeaderView.Stretch)
        self.tv.setItemDelegateForColumn(1, PillDelegate(self.tv))
        self.tv.itemSelectionChanged.connect(self.show_item)
        self.tv.itemDoubleClicked.connect(lambda *_: self.open_sel())
        tl.addWidget(self.tv)
        lcol.addWidget(tb, 1)
        # مربع التأكيد
        box = card(); bh = QHBoxLayout(box); bh.setContentsMargins(16, 12, 16, 14); bh.setSpacing(18)
        left = QVBoxLayout(); left.setSpacing(4)
        k = QLabel('CONFIRM'); k.setFont(HQ.font(11, True)); k.setStyleSheet(f'color: {MUTED};'); left.addWidget(k)
        self.l_guess = QLabel(''); self.l_guess.setTextFormat(Qt.RichText); self.l_guess.setWordWrap(True); self.l_guess.setFont(HQ.font(13)); left.addWidget(self.l_guess)
        g = QGridLayout(); g.setHorizontalSpacing(10); g.setVerticalSpacing(8)
        def lab(t):
            l = QLabel(t); l.setStyleSheet(f'color: {MUTED};'); l.setAlignment(Qt.AlignRight | Qt.AlignVCenter); return l
        self.cb_type = QComboBox()
        for code_, text_ in type_rows():
            self.cb_type.addItem(text_, code_)
            if code_ is None:                              # عنوان مجموعة – مابيتختارش
                mi = self.cb_type.model().item(self.cb_type.count() - 1)
                if mi is not None:
                    mi.setFlags(Qt.NoItemFlags); mi.setForeground(QColor(C_ASK)); mi.setFont(HQ.font(10.5, True))
        self.cb_type.setCurrentIndex(-1)
        self.cb_type.setMaxVisibleItems(24)
        self.cb_type.activated.connect(lambda *_: self.type_changed())
        self.e_no = QLineEdit(); self.e_no.setFixedWidth(190)
        self.e_no.textEdited.connect(lambda *_: self.type_changed(keep_folder=True))
        self.cb_stat = QComboBox(); self.cb_stat.addItems(['Submitted', 'Approved', 'Approved with comments', 'Rejected'])
        self.cb_stat.activated.connect(lambda *_: self.type_changed())
        self.cb_sub = QComboBox(); self.cb_sub.setEditable(True); self.cb_sub.setInsertPolicy(QComboBox.NoInsert)
        self.cb_sub.setCompleter(None)                   # من غير إكمال تلقائي (زي القديم) – الفولدر يفضل مطابق للمكتوب
        self.cb_sub.activated.connect(lambda *_: self.sub_changed())
        self.cb_sub.lineEdit().textEdited.connect(lambda *_: self.sub_changed())
        self.l_subhow = QLabel(''); self.l_subhow.setStyleSheet(f'color: {C_DIM};')
        self.e_folder = QLineEdit(); self.e_folder.setFont(HQ.font(12.5, fam=HQ.MONO))
        self.e_name = QLineEdit()
        # خانة «شركة جديدة» – الفولدر مابيتعملش غير لما المستخدم يدوس Add party (طلبه 21-11-2026)
        self.add_box = QFrame(); self.add_box.setObjectName('addbox')
        self.add_box.setStyleSheet(
            'QFrame#addbox { border: 2px solid #F0C860; border-radius: 8px; background: %s; }'
            'QFrame#addbox QLabel { border: none; background: transparent; }'
            'QPushButton#cat { padding: 6px 12px; border: 1px solid %s; border-radius: 7px; }'
            'QPushButton#cat:checked { background: #1E3448; border-color: #3D7DB8; color: #9CC8F0; font-weight: bold; }'
            % (('#26261A' if HQ.DARK else '#FFF8E1'), HQ.SOFT_LINE))
        av = QVBoxLayout(self.add_box); av.setContentsMargins(10, 8, 10, 9); av.setSpacing(6)
        self.l_add = QLabel(''); self.l_add.setFont(HQ.font(11, True))
        self.l_add.setStyleSheet(f"color: {'#F0C860' if HQ.DARK else '#8A6100'};"); av.addWidget(self.l_add)
        ah = QHBoxLayout(); ah.setSpacing(7)
        self.e_new = QLineEdit(); self.e_new.setPlaceholderText('Company name (English)'); ah.addWidget(self.e_new, 1)
        self.cat_btns = {}
        for cat in D.PT.CATS:
            cb_ = QPushButton(cat); cb_.setObjectName('cat'); cb_.setCheckable(True); cb_.setCursor(Qt.PointingHandCursor)
            cb_.setToolTip(D.PT.CAT_NAMES.get(cat, cat))
            cb_.clicked.connect(lambda _=False, c=cat: self.set_cat(c)); ah.addWidget(cb_); self.cat_btns[cat] = cb_
        ah.addWidget(btn('＋  Add party', self.add_party, True))
        av.addLayout(ah)
        self.add_box.setVisible(False)
        g.addWidget(lab('Type'), 0, 0); g.addWidget(self.cb_type, 0, 1)
        g.addWidget(lab('No.'), 0, 2); g.addWidget(self.e_no, 0, 3)
        g.addWidget(lab('Status'), 0, 4); g.addWidget(self.cb_stat, 0, 5)
        g.addWidget(lab('Company'), 1, 0); g.addWidget(self.cb_sub, 1, 1, 1, 3); g.addWidget(self.l_subhow, 1, 4, 1, 2)
        g.addWidget(self.add_box, 2, 1, 1, 5)
        g.addWidget(lab('Folder'), 3, 0); g.addWidget(self.e_folder, 3, 1, 1, 5)
        g.addWidget(lab('Name'), 4, 0); g.addWidget(self.e_name, 4, 1, 1, 5)
        g.setColumnStretch(1, 1); g.setColumnMinimumWidth(5, 220)
        left.addSpacing(6); left.addLayout(g)
        bh.addLayout(left, 1)
        right = QVBoxLayout(); right.setSpacing(7); right.addSpacing(18)
        for text, slot, pri in (('✓  Approve && move', self.approve, True), ('Open file', self.open_sel, False),
                                ('Show text read', self.show_text, False), ('Split scan', self.split_sel, False),
                                ('Send to Claude', self.send_claude, False)):
            b = btn(text, slot, pri); b.setMinimumWidth(190); right.addWidget(b)
            if slot == self.approve: self.b_approve = b
        right.addStretch(1)
        bh.addLayout(right)
        lcol.addWidget(box)
        mainrow.addLayout(lcol, 1)
        # الـ Preview (يمين)
        self.pv = PreviewPane(self) if PreviewPane else None
        if self.pv: mainrow.addWidget(self.pv)
        v.addLayout(mainrow, 1)
        self.claude_state()

    # ------------------------------------------------------------ helpers (رسايل وجدول)
    def info(self, t): self.l_info.setText(t)
    def msg(self, t, kind='info'):
        {'info': QMessageBox.information, 'warn': QMessageBox.warning, 'err': QMessageBox.critical}[kind](self, 'DC', t)
    def ask(self, t, parent=None):
        return QMessageBox.question(parent or self, 'DC', t) == QMessageBox.Yes

    def _row(self, iid):
        for r in range(self.tv.rowCount()):
            it = self.tv.item(r, 0)
            if it and it.data(Qt.UserRole) == iid: return r
        return -1

    def _set_row(self, iid, values, tag):
        r = self._row(iid)
        if r < 0: return
        for c, val in enumerate(values):
            cell = QTableWidgetItem(str(val if val is not None else ''))
            if c == 0: cell.setData(Qt.UserRole, iid); cell.setToolTip(str(val))
            if c == 1: cell.setData(ROLE_TAG, tag)
            if c in (5, 6) and val: cell.setToolTip(str(val))
            if c == 5: cell.setFont(HQ.font(12, fam=HQ.MONO)); cell.setForeground(QColor(MUTED))
            if tag == 'err' and c == 2: cell.setForeground(QColor(C_ERR))
            if tag == 'wait' and c != 1: cell.setForeground(QColor(C_DIM))
            self.tv.setItem(r, c, cell)

    def _add_row(self, iid, values, tag):
        r = self.tv.rowCount(); self.tv.insertRow(r)
        self.tv.setItem(r, 0, QTableWidgetItem('')); self.tv.item(r, 0).setData(Qt.UserRole, iid)
        self._set_row(iid, values, tag)

    def sel(self):
        r = self.tv.currentRow()
        if r < 0 or not self.tv.selectionModel().hasSelection(): return None, None
        iid = self.tv.item(r, 0).data(Qt.UserRole)
        return iid, self.items.get(iid)

    def select_row(self, r):
        if not (0 <= r < self.tv.rowCount()): return
        if self.tv.selectionModel().isRowSelected(r, self.tv.rootIndex()): self.show_item()
        else: self.tv.selectRow(r)

    def set_cursor_busy(self, on):
        if on: QApplication.setOverrideCursor(Qt.WaitCursor)
        else:
            while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()

    # ------------------------------------------------------------ list
    def scan(self):
        if self.busy: return
        self.tv.setRowCount(0); self.items.clear()
        for it in D.list_inbox():
            st = 'Claude ✓' if it.get('claude') else ('Review' if it['review'] else 'Not read')
            iid = self._next_id; self._next_id += 1
            self.items[iid] = it
            self._add_row(iid, (it['name'], st, '', '', '', '', ''), 'wait')
        ok, why = R.ocr_status()
        try: w = D.claude_waiting()
        except Exception: w = 0
        self.info(f'{len(self.items)} file(s) in INBOX' + (f'   |   {w} with Claude (hidden until read)' if w else '')
                  + ('' if ok else f'   |   OCR off: {why}'))
        self.clear_box()

    def distribute(self):
        if self.busy: return
        self.scan()
        todo = [i for i, it in self.items.items() if 'sugg' not in it and not (it.get('claude') and not it.get('claude_ready'))]
        if not todo:
            try: waiting = D.claude_waiting()
            except Exception: waiting = 0
            self.msg(f'{waiting} file(s) are waiting for Claude in the chat.' if waiting else f'INBOX is empty.\n\nDrop papers in:\n{D.inbox()}')
            return
        self.busy = True; self.set_cursor_busy(True)
        threading.Thread(target=self._work, args=(todo,), daemon=True).start()

    def _poll(self):
        if getattr(self, '_polling', False): return       # رسالة مفتوحة من جوه الطابور ← متدخلش تاني
        self._polling = True
        try:
            while True:
                fn = self.q.get_nowait()
                try: fn()
                except Exception: LOG.exception('UI update failed')
        except queue.Empty:
            pass
        finally:
            self._polling = False

    def _work(self, todo):
        for k, iid in enumerate(todo, 1):
            it = self.items.get(iid)
            if not it: continue
            self.q.put(lambda iid=iid, k=k: (self._status(iid, 'Reading…'), self.info(f'Reading {k} of {len(todo)}…')))
            try:
                r = D.analyse(it['path'])
                it.update(text=r.get('text', ''), method=r.get('method'), warn=r.get('warn'), ocr_pdf=r.get('ocr_pdf'), sugg=r.get('sugg', []),
                          sub=r.get('sub'), work=r.get('work'), key=r.get('key'), cached=r.get('cached'),
                          parts=r.get('parts'), claude_err=r.get('claude_err'), claude_doc=r.get('claude'))
                top = it['sugg'][0] if it['sugg'] else None
                it['dup'] = None
                if top and top.get('folder') and not it.get('parts'):
                    try: it['dup'] = same_file(it['path'], top['folder'])
                    except Exception: LOG.exception('Duplicate check failed: %s', it.get('path'))
                if it.get('dup') and it.get('part'):        # جزء من سكان اتقسم ومتخزن قبل كده ← 05_DUPLICATES لوحده
                    try:
                        dd = os.path.join(D.inbox(), DUP); os.makedirs(dd, exist_ok=True)
                        dst = D.free_name(dd, os.path.basename(it['path'])); safe_move(it['path'], dst)
                        D._log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='move', src=it['path'], dst=dst, type='DUPLICATE',
                                    guessed=top['code'], same_as=it['dup'], fields=top.get('fields') or {}))
                        try: D._chat_done(it['path'])
                        except Exception: pass
                        it['gone'] = dst
                    except Exception:
                        LOG.exception('Auto duplicate move failed: %s', it.get('path'))
                if (not top or not top['sure']) and not it.get('claude') and not it.get('dup'):
                    it['path'] = D.to_review(it['path']); it['review'] = True
            except Exception as e:
                it.update(sugg=[], warn=str(e)); LOG.exception('Reading failed: %s', it.get('path'))
            self.q.put(lambda iid=iid: self.paint(iid))
        self.q.put(self._done)

    def _status(self, iid, text):
        r = self._row(iid)
        if r >= 0:
            c = QTableWidgetItem(text); c.setData(ROLE_TAG, 'wait'); self.tv.setItem(r, 1, c)

    def _done(self):
        self.busy = False; self.set_cursor_busy(False)
        n = sum(1 for it in self.items.values() if it.get('sugg') and it['sugg'][0]['sure'] and not it.get('dup'))
        gone = [i for i, it in self.items.items() if it.get('gone')]
        for i in gone:                                   # أجزاء اتنقلت للمكررات لوحدها ← تتشال من القايمة
            r = self._row(i)
            if r >= 0: self.tv.removeRow(r)
            D.drop_work(self.items.pop(i))
        dups = sum(1 for it in self.items.values() if it.get('dup'))
        m = sum(1 for it in self.items.values() if it.get('cached'))
        errs = [it['claude_err'] for it in self.items.values() if it.get('claude_err')]
        self.info(f'{len(self.items)} read  |  {n} sure' + (f'  |  {dups} duplicate(s)' if dups else '')
                  + f'  |  {len(self.items) - n - dups} need your choice (in «{D.REVIEW}»)'
                  + (f'  |  {len(gone)} part(s) already filed → {DUP}' if gone else '')
                  + (f'  |  {m} from memory' if m else '') + (f'  |  Claude: {errs[0]}' if errs else ''))
        self.select_row(0)
        split = False
        for iid in [i for i, it in self.items.items() if it.get('parts')]:
            split = self.split_item(iid, ask=False) or split          # يتقسم لوحده (طلب المستخدم 12-11)
        if split: QTimer.singleShot(300, self.distribute)

    def paint(self, iid):
        it = self.items.get(iid)
        if not it or it.get('gone'): return
        s = it['sugg'][0] if it.get('sugg') else None
        if it.get('parts'):
            return self._set_row(iid, (it['name'], f"{len(it['parts'])} papers", 'Scan with several papers (Claude)', '', '', 'split first', ''), 'ask')
        if not s:
            return self._set_row(iid, (it['name'], 'Not understood', it.get('warn') or '—', '', '', '', ''), 'err')
        f = s['fields']
        if it.get('dup'):
            return self._set_row(iid, (it['name'], 'Duplicate', C.EN[s['code']] + '  – already filed', f.get('no', ''), f.get('date', ''),
                                       '00_INBOX\\' + DUP, 'same as  ' + rel(it['dup'])), 'dup')
        if s['code'] in C.SUB_CODES and not (it.get('sub') or {}).get('folder'):     # شركة مش في القايمة ← مفيش فولدر لحد ما تتضاف
            comp = (it.get('sub') or {}).get('label') or f.get('company') or ''
            pty = self._party(comp)
            if pty:
                return self._set_row(iid, (it['name'], 'Check', C.EN[s['code']], f.get('no', ''), f.get('date', ''),
                                           D.sub_folder(s['code'], pty['folder'], pty['label']) or '', s['name']), 'ask')
            return self._set_row(iid, (it['name'], 'New company', C.EN[s['code']], f.get('no', ''), f.get('date', ''),
                                       '— choose / add the company —', s['name']), 'ask')
        self._set_row(iid, (it['name'], 'Sure' if s['sure'] else 'Check', C.EN[s['code']], f.get('no', ''), f.get('date', ''),
                            s['folder'], s['name']), 'sure' if s['sure'] else 'ask')

    # ------------------------------------------------------------ confirm box
    def clear_box(self):
        if getattr(self, 'pv', None): self.pv.clear()
        self._filling = True
        self.l_guess.setText(''); self.cb_type.setCurrentIndex(-1); self.e_no.clear()
        self.e_folder.clear(); self.e_name.clear(); self.cb_sub.setEditText(''); self.l_subhow.clear()
        self._filling = False
        self._sync_add()

    def set_type(self, code):
        i = self.cb_type.findData(code) if code in CODES else -1
        self.cb_type.setCurrentIndex(i)

    def preview_sel(self):
        """يعرض الورقة المختارة في الـ Preview (لو الـ pane موجود)."""
        if not getattr(self, 'pv', None): return
        iid, it = self.sel()
        p = (it or {}).get('path')
        if p: self.pv.load(p)
        else: self.pv.clear()

    def show_item(self):
        self.preview_sel()
        iid, it = self.sel()
        if not it: return
        self._filling = True
        try:
            sg = it.get('sugg') or []
            if not sg:
                self.l_guess.setText(str(it["name"]).replace('&', '&amp;').replace('<', '&lt;') + ':  ' + ('not read yet – press Distribute.' if 'sugg' not in it else 'could not understand it – choose the type yourself.'))
                self.cb_type.setCurrentIndex(-1); self.e_no.clear(); self.e_folder.clear(); self.e_name.clear(); return
            s = sg[0]; f = s['fields']
            alt = ',  '.join(f"{C.EN[x['code']]} ({x['score']})" for x in sg[1:3])
            if it.get('parts'): alt = f"this scan has {len(it['parts'])} papers – press «Split scan»" + (f'  |  {alt}' if alt else '')
            esc = lambda t: str(t).replace('&', '&amp;').replace('<', '&lt;')
            col = GREEN if s['sure'] else C_ASK
            self.l_guess.setText(
                f"Understood: <b style='color:{col}'>{esc(C.EN[s['code']])}</b>{'  (Claude)' if s.get('claude') else ''}  –  score {s['score']}"
                f"{'  (sure)' if s['sure'] else ''}"
                + (f"<span style='color:{MUTED}'>   |  other options: {esc(alt)}</span>" if alt else '')
                + (f"<span style='color:{C_ERR}'>   |  ⚠ {esc(s['note'])}</span>" if s.get('note') else '')
                + (f"<span style='color:{C_ERR}'>   |  ⚠ {esc(it['warn'])}</span>" if it.get('warn') else '')
                + (f"<br><span style='color:{C_DUP}'><b>Already filed:</b> {esc(rel(it['dup']))}  –  Approve moves this copy to 00_INBOX\\{DUP}</span>"
                   if it.get('dup') else ''))
            self.set_type(s['code']); self.e_no.setText(str(f.get('out') or f.get('no', '') or ''))
            self.cb_stat.setCurrentText({'approved': 'Approved', 'approved_comments': 'Approved with comments',
                                         'rejected': 'Rejected'}.get(f.get('status'), 'Submitted'))
            self.e_folder.setText(str(s.get('folder') or '')); self.e_name.setText(str(s.get('name') or ''))
            self.sub_list(); self.set_sub_state(s['code'])
            if s['code'] in C.SUB_CODES:
                sb = it.get('sub'); comp = f.get('company') or ''
                if sb and not sb.get('folder'): comp, sb = (sb.get('label') or comp), None      # اسم قراه Claude ومالوش فولدر ← مش جهة عندنا
                self.cb_sub.setEditText(sb['label'] if sb else '')
                self.l_subhow.setText(f"found by {sb['how']}" if sb else 'choose the company from the list')
                if not sb: self.e_folder.clear()
                if not sb and comp:
                    self.cb_sub.setEditText(comp); self._filling = False; self.sub_changed()
        finally:
            self._filling = False
            self._sync_add()

    def sub_list(self):
        self._subs = D.subs(); self._pcache = {}
        cur = self.cb_sub.currentText()
        self.cb_sub.blockSignals(True)
        self.cb_sub.clear(); self.cb_sub.addItems([s['label'] for s in self._subs])
        self.cb_sub.setEditText(cur)
        self.cb_sub.blockSignals(False)

    def set_sub_state(self, code):
        on = code in C.SUB_CODES
        self.cb_sub.setEnabled(on)
        if not on: self.cb_sub.setEditText(''); self.l_subhow.clear()

    def _party(self, v):
        """الجهة الموجودة فعلاً في 13_PARTIES بالاسم ده (اللي في القايمة، أو اسمها الإنجليزي/العربي/اسم اتربط بيها)، أو None.
        مفيش فولدر بيتعمل هنا."""
        v = str(v or '').strip()
        if not v: return None
        if v in self._pcache: return self._pcache[v]
        if not self._subs:
            try: self._subs = D.subs()
            except Exception: LOG.exception('parties list'); return None
        s = next((x for x in self._subs if x['label'] == v), None)
        if not s:
            rel_ = self._alias.get(D.PT.key(v))
            if not rel_:
                try:
                    en = D.en_company(v)
                    rel_ = D.party_rel(v) or (D.party_rel(en) if en else None)
                except Exception:
                    LOG.exception('party lookup'); rel_ = None
            if rel_:
                k = os.path.normcase(os.path.normpath(rel_))
                s = next((x for x in self._subs if os.path.normcase(os.path.normpath(x['folder'])) == k), None)
        self._pcache[v] = s
        return s

    def sub_changed(self):
        if self._filling: return
        code = self.cur_code(); v = self.cb_sub.currentText().strip()
        if code not in C.SUB_CODES: return self._sync_add()
        s = self._party(v)
        if s: folder = D.sub_folder(code, s['folder'], s['label']); self.l_subhow.setText(f"existing folder  ·  {s['dir']}")
        elif v: folder = ''; self.l_subhow.setText('not in your list')
        else: folder = ''; self.l_subhow.setText('choose the company from the list')
        self.e_folder.setText(folder or '')
        self._sync_add()

    # ------------------------------------------------------------ شركة جديدة (بتتضاف بإيد المستخدم بس)
    def set_cat(self, cat):
        for c, b in self.cat_btns.items(): b.setChecked(c == cat)

    def cur_cat(self):
        return next((c for c, b in self.cat_btns.items() if b.isChecked()), None)

    def _sync_add(self):
        """خانة الإضافة تظهر لما النوع بيروح فولدر جهة والشركة المكتوبة مش عندنا – وساعتها Approve مقفول."""
        if not hasattr(self, 'b_approve') or not hasattr(self, 'tv'): return
        try:
            iid, it = self.sel(); code = self.cur_code(); v = self.cb_sub.currentText().strip()
            need = bool(it) and code in C.SUB_CODES and not self._party(v)
            self.add_box.setVisible(need)
            self.b_approve.setEnabled(not need)
            self.b_approve.setToolTip('Choose the company from the list, or add it with «Add party» first.' if need else '')
            if not need: self._add_for = None; return
            if self._add_for != (iid, v):
                self._add_for = (iid, v)
                self.e_new.setText(D.en_company(v) or v)
                self.set_cat('SUB' if code in ('SUB_CONTRACT', 'SUB_IPC') else None)
            self.l_add.setText(f'NEW  ·  «{v}» is not in your list – add it?' if v
                               else 'Company not in the list?  Type its name, choose the kind and press Add party.')
        except RuntimeError:
            pass

    def _similar(self, name):
        """جهات اسمها قريب من الاسم ده (عشان مايتعملش فولدرين لنفس الشركة)."""
        k = D.PT.key(D.en_company(name) or name)
        out = []
        for s in self._subs:
            for nm in [s.get('name', '')] + list(s.get('aliases') or []):
                ks = D.PT.key(nm)
                if k and ks and (k == ks or (min(len(k), len(ks)) >= 4 and (k in ks or ks in k))):
                    out.append(s); break
        return out

    def _use_party(self, s, read=''):
        if read: self._alias[D.PT.key(read)] = s['folder']
        self._pcache = {}
        self._filling = True
        try: self.cb_sub.setEditText(s['label'])
        finally: self._filling = False
        self.sub_changed()
        for iid, it in list(self.items.items()):           # باقي ورق نفس الشركة في القايمة
            sg = it.get('sugg') or []
            if not sg or sg[0]['code'] not in C.SUB_CODES or (it.get('sub') or {}).get('folder'): continue
            try: self.paint(iid)
            except Exception: LOG.exception('repaint')

    def add_party(self):
        code = self.cur_code(); read = self.cb_sub.currentText().strip()
        name = ' '.join(self.e_new.text().split())
        if code not in C.SUB_CODES: return
        if len(name) < 2: return self.msg('Type the company name first.', 'warn')
        self.sub_list()
        s = self._party(name)
        if s:
            self.info(f"Already in the list: {s['dir']} – used it."); return self._use_party(s, read)
        sim = self._similar(name)
        if sim:
            b = QMessageBox(self); b.setWindowTitle('DC'); b.setIcon(QMessageBox.Question)
            b.setText('A company with a similar name is already there:\n\n' + '\n'.join(f"   {x['cat']}\\{x['label']}" for x in sim[:4])
                      + f'\n\nUse it for «{name}», or create a new folder?')
            use = [b.addButton(f"Use  {x['dir']}", QMessageBox.AcceptRole) for x in sim[:3]]
            b_new = b.addButton('Create new', QMessageBox.DestructiveRole); b.addButton(QMessageBox.Cancel)
            b.exec()
            if b.clickedButton() in use: return self._use_party(sim[use.index(b.clickedButton())], read)
            if b.clickedButton() is not b_new: return
        cat = self.cur_cat()
        if not cat: return self.msg('Choose the kind first:  SUB  /  SUP  /  PLANT  /  LAB.', 'warn')
        en = D.en_company(name) or name
        if not D.PT.en_folder(en):
            return self.msg('Write the name in English letters – it becomes the folder name.', 'warn')
        rel_ = D.PT.new_rel(D.project_root(), en, cat)
        path = os.path.join(D.project_root(), rel_)
        if not self.ask(f'Create this folder?\n\n{rel_}'): return
        try:
            ar = next((t for t in (read, name) if re.search(r'[\u0600-\u06FF]', t)), '')
            D.PT.ensure(path, arabic=ar)
            extra = [t for t in (read, name) if t and not re.search(r'[\u0600-\u06FF]', t)
                     and D.PT.key(t) != D.PT.key(D.PT.en_folder(en))]
            if extra:                                      # الاسم زي ما اتقرا ← الورق الجاي يتعرف لوحده
                rd = os.path.join(path, 'README.txt')
                with open(rd, 'a', encoding='utf-8' if os.path.exists(rd) else 'utf-8-sig') as f_:
                    f_.write('ALIAS: ' + ' | '.join(dict.fromkeys(extra)) + '\r\n')
        except OSError as e:
            LOG.exception('add party'); return self.msg(f'Could not create the folder:\n{e}', 'err')
        LOG.info('Party added by user: %s', rel_)
        self.sub_list()
        k = os.path.normcase(os.path.normpath(rel_))
        s = next((x for x in self._subs if os.path.normcase(os.path.normpath(x['folder'])) == k), None)
        if not s: return self.msg('The folder was created but is not in the list – press Refresh list.', 'warn')
        self.info(f'Added: {rel_}')
        self._use_party(s, read)

    def cur_code(self):
        code = self.cb_type.currentData()
        return code if code in CODES else None

    def type_changed(self, keep_folder=False):
        if self._filling: return
        iid, it = self.sel(); code = self.cur_code()
        if not it or not code: return
        sg = next((x for x in it.get('sugg') or [] if x['code'] == code), None)
        f = dict(sg['fields']) if sg else C.fields(code, it.get('text', ''), it.get('text', '')[:600], it['name'])
        no = self.e_no.text().strip()
        if no.isdigit(): f['no'] = int(no)
        o = C.find_out(no)
        if o and code in C.OUT_NAMED:
            f['out'] = o
            if code == 'OUT_YA': f['no'] = int(o[-4:])
        if code in ('PAYROLL', 'GIS') and not f.get('ipc'): C.set_ipc(code, f, it['name'])
        f['status'] = {'Approved': 'approved', 'Approved with comments': 'approved_comments', 'Rejected': 'rejected'}.get(self.cb_stat.currentText())
        self.set_sub_state(code)
        if not keep_folder:
            if code in C.SUB_CODES:
                if not self.cb_sub.currentText() and it.get('sub'): self.cb_sub.setEditText(it['sub']['label'])
                self.sub_list(); self.sub_changed()
            else:
                self.e_folder.setText(C.type_folder(code, f))
        name = C.file_name(code, f, it['name'], it.get('text', ''))
        self.e_name.setText(name or '')
        if code == 'SITE_WIR':
            self.info(f'⚠ {C.wir_note(f)}' if not name else 'Name from the WIR log ✔')
        elif C.out_note(code, f):
            self.info('⚠ ' + C.out_note(code, f))
        it['fields'] = f
        self._sync_add()

    # ------------------------------------------------------------ approve
    def approve(self):
        iid, it = self.sel(); code = self.cur_code()
        if not it: return
        if self.busy: return self.msg('Wait until reading finishes.')
        if getattr(self, 'pv', None): self.pv.clear()   # نسيب الملف قبل ما ينتقل
        if 'sugg' not in it: return self.msg('Press Distribute first so the file is read.')
        folder, name = self.e_folder.text().strip(), self.e_name.text().strip()
        if not code or not folder or not name:
            return self.msg('Choose the type, company, folder and name first.', 'warn')
        if code in C.SUB_CODES:                           # فولدر الجهة لازم يكون موجود – التوزيع مابيعملش جهات (طلب المستخدم 21-11-2026)
            self.sub_list()
            pty = self._party(self.cb_sub.currentText())
            m_ = re.match(r'(13_PARTIES[\\/][^\\/]+[\\/][^\\/]+)', os.path.normpath(folder))
            if not pty or not m_ or not os.path.isdir(os.path.join(D.project_root(), m_.group(1))):
                self._sync_add()
                return self.msg('This company is not in your list.\n\nChoose it from the list, or add it with «＋ Add party» first.', 'warn')
        guessed = it['sugg'][0]['code'] if it.get('sugg') else None
        it.setdefault('fields', (it['sugg'][0]['fields'] if it.get('sugg') else {}))
        upd_msg, dn_data = '', None
        if code == 'UPD_CAL':                             # تقويم تحديث الشركة ← تراجع التواريخ (شباك Tk فوق البرنامج) والملف يتنقل بعد ما تحفظ
            import tkinter as tk
            root = tk.Tk(); root.withdraw()
            try:
                got = D.review_update_dates(root, it)
            except PermissionError:
                got = 'locked'
            except Exception as e:
                LOG.exception('Update dates failed'); got = ('err', e)
            finally:
                try: root.destroy()
                except Exception: pass
            if got == 'locked': return self.msg('12-LOG UPDATE DATES.xlsx is open in Excel.\nClose it and press Approve again.', 'warn')
            if isinstance(got, tuple) and got and got[0] == 'err': return self.msg(f'Could not save the update dates:\n{got[1]}', 'err')
            if not got: return                            # Cancel ← الملف يفضل مكانه
            upd_msg, month = got
            name = f'{month} Company Update Calendar' + (os.path.splitext(name)[1] or '.pdf'); self.e_name.setText(name)
        if code == 'DN':                                  # إذن التوريد ← الأصناف والـ MIR جنب صورته (شباك Qt من Materials)
            mp = os.path.join(os.path.dirname(HERE), 'mat')
            if mp not in sys.path: sys.path.insert(0, mp)
            work = None
            try:
                import mat_qt
                pics, work = D.page_pictures(it)
                dn_data = mat_qt.delivery_dialog(self, D.delivery_prefill(it), pics)
            except Exception as e:
                LOG.exception('Delivery window failed'); return self.msg(f'Could not open the delivery window:\n{e}', 'err')
            finally:
                if work: shutil.rmtree(work, ignore_errors=True)
            if not dn_data: return                        # Cancel ← الملف يفضل مكانه
            folder = C.TYPES['DN'][1]; self.e_folder.setText(folder)
            name = (dn_data.get('file') or 'Delivery Note') + '.pdf'; self.e_name.setText(name)
        # المكررات: نفس الملف بالظبط ← _duplicates من غير أسئلة
        try: same = same_file(it['path'], folder)
        except Exception: same = None; LOG.exception('Duplicate check failed')
        if same: return self.to_duplicates(iid, it, code, same)
        olds = []                                        # نفس الرقم بس ورقة مختلفة ← اسأل
        if code in C.OUT_NAMED and (it.get('fields') or {}).get('out'):
            try: olds = same_number(it['fields']['out'], folder)
            except Exception: LOG.exception('Same-number check failed')
        if olds:
            b = QMessageBox(self); b.setWindowTitle('DC'); b.setIcon(QMessageBox.Question)
            b.setText(f"{it['fields']['out']} is already filed:\n\n" + '\n'.join('•  ' + rel(p) for p in olds[:6])
                      + ('\n…' if len(olds) > 6 else '') + '\n\nIs this paper a newer version (e.g. the approved copy)?')
            if not STAGE.match(os.path.basename(os.path.normpath(folder))):
                b.setInformativeText('This folder has no 4_ARCHIVED stage – a newer version is kept next to the old one.')
            b_new = b.addButton('Newer version – file it', QMessageBox.AcceptRole)
            b_dup = b.addButton('Duplicate – put aside', QMessageBox.DestructiveRole)
            b_cancel = b.addButton('Cancel', QMessageBox.RejectRole)
            b.setDefaultButton(b_cancel); b.setEscapeButton(b_cancel)     # Enter بالغلط مايأرشفش حاجة
            b.exec()
            if b.clickedButton() is b_dup: return self.to_duplicates(iid, it, code, olds[0])
            if b.clickedButton() is not b_new: return
        ct = None
        if code == 'SUB_CONTRACT':
            try: D.contract_log_ready()
            except PermissionError: return self.msg('10-LOG SUB.xlsx is open in Excel.\nClose it and press Approve again.', 'warn')
            except Exception as e: return self.msg(f'10-LOG SUB cannot be updated:\n{e}', 'err')
            ct = self.review_contract(it)
            if not ct: return
            folder = D.sub_folder('SUB_CONTRACT', D.party_rel(self.cb_sub.currentText()) or '', ct['en']); self.e_folder.setText(folder)
            name = ct['name']; self.e_name.setText(name)
        folder = os.path.normpath(folder)
        dest = os.path.normpath(os.path.join(D.project_root(), folder))
        aside, arc = [], None                            # نسخ قديمة بنفس الرقم في نفس الفولدر ← 4_ARCHIVED (من غير (2))
        if olds and STAGE.match(os.path.basename(dest)):
            arc = os.path.join(os.path.dirname(dest), '4_ARCHIVED')
            for p in olds:
                if os.path.normcase(os.path.dirname(p)) == os.path.normcase(dest) and os.path.normcase(arc) != os.path.normcase(dest):
                    try:
                        os.makedirs(arc, exist_ok=True); q = D.free_name(arc, os.path.basename(p)); safe_move(p, q)
                        aside.append((p, q))
                    except OSError as e:
                        LOG.exception('Archive old copy failed: %s', p)
                        self._restore(aside)
                        return self.msg(f'Could not move the old copy to 4_ARCHIVED – nothing was filed.\n\n{e}', 'warn')
        try:
            dst = D.approve(it, code, folder, name, guessed)
        except Exception as e:
            self._restore(aside)                         # النقل وقع ← القديم يرجع مكانه
            LOG.exception('Move failed: %s', it.get('path')); return self.msg(f'Could not move the file:\n{e}', 'err')
        if aside:
            D._log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='archive', src=os.path.dirname(os.path.dirname(dst)), dst=arc,
                        files=[os.path.basename(p) for p, _ in aside],
                        back=[(os.path.basename(q), os.path.basename(dest)) for _, q in aside], out=it['fields'].get('out')))
        f_ = it.get('fields') or {}
        if code in C.OUT_NAMED and f_.get('out'):
            stem = os.path.splitext(os.path.basename(dst))[0]
            k = stem.upper().find(f_['out'].upper() + ' - ')
            if k >= 0:
                t = re.sub(r'\s*\(\d+\)$', '', stem[k + len(f_['out']) + 3:])
                try: D.save_title(f_['out'], t)
                except Exception: LOG.exception('Save title failed')
        try: self.log_sheet(SH.entry_from_move(code, it.get('fields') or {}, it.get('text', ''), dst))
        except Exception: LOG.exception('DC sheet entry failed')
        extra = ('   |   ' + upd_msg) if upd_msg else ''
        if code == 'DN' and dn_data:                      # سطور التوريدات + الـ MIR جاهز للتوقيع
            QGuiApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                ok, bad, mir_path = D.after_delivery(dn_data, dst)
            except Exception as e:
                LOG.exception('Delivery after move failed'); ok, bad, mir_path = [], [str(e)], None
            QGuiApplication.restoreOverrideCursor()
            if ok: extra += '   |   ' + ' · '.join(ok)
            if bad: self.msg('The delivery note was filed, but:\n\n' + '\n'.join(bad), 'warn')
            if mir_path and self.ask(f'The MIR is ready for signature:\n{mir_path}\n\nOpen it?'):
                try: os.startfile(mir_path)
                except Exception: pass
        if code == 'SUB_CONTRACT' and ct:
            try:
                miss = D.after_contract(ct['en'], ct, dst)
                extra += '   |   ' + (f"Contract {ct['no']} – missing: {', '.join(miss)}" if miss else f"Contract {ct['no']} complete ✔")
            except Exception as e:
                LOG.exception('Contract register failed')
                self.msg(f'The contract was filed, but 10-LOG SUB was not updated:\n{e}\n\nAdd the row in sheet SUB_CONTRACTS by hand.', 'warn')
            if ct.get('mksub'):
                try:
                    sid, n, new_ = D.contract_to_sub(self.cb_sub.currentText().strip() or ct['en'], ct, f_)
                    extra += f"   |   SUB: {sid} {'created' if new_ else 'updated'}" + (f', {n} items' if n else '')
                except PermissionError:
                    self.msg('The contract was filed, but 10.1-SUBCONTRACTORS.xlsx is open in Excel – the subcontractor was not added to SUB.\n'
                             'Close it and add him from SUB → + New Subcontractor.', 'warn')
                except Exception as e:
                    LOG.exception('SUB from contract failed')
                    self.msg(f'The contract was filed, but the subcontractor was not added to SUB:\n{e}', 'warn')
        if code in C.OUT_NAMED:
            try:
                old = D.archive_old_copies(it.get('fields') or {}, dst)
                if old: extra += f'   |   {len(old)} older copy(ies) moved to 4_ARCHIVED'
            except Exception:
                LOG.exception('Archive old copies failed')
        st_now = {'Approved': 'approved', 'Approved with comments': 'approved_comments', 'Rejected': 'rejected'}.get(self.cb_stat.currentText())
        if code == 'MAR' and f_.get('out') and st_now:      # MAR رجع من الاستشاري ← MAR LOG
            self.mar_log(f_, st_now)
        if code == 'MIR':
            no = self.e_no.text().strip()
            no = int(no) if no.isdigit() else (it.get('fields') or {}).get('no')
            if no and D.after_signed_mir(no, dst): extra += '   |   unsigned copy moved to 4_ARCHIVED'
        if code == 'IPC_INV':
            try:
                ok, m = D.certify_invoice(it.get('text', ''), ask=lambda q: self.ask(q))
                if ok: extra = '   |   ' + m
                else: self.msg(m, 'warn')
            except PermissionError:
                self.msg('11-LOG IPC is open in Excel.\nClose it – the IPC will be marked Certified when you open the IPC screen.', 'warn')
            except Exception as e:
                LOG.exception('IPC certify failed'); self.msg(f'The invoice was moved, but the IPC was not marked Certified:\n{e}', 'warn')
        if code == 'IPC_SIGNED':
            no = self.e_no.text().strip()
            no = int(no) if no.isdigit() else (it.get('fields') or {}).get('no')
            if not no:
                self.msg('The IPC number is not known – the IPC was not marked Certified.\nUse "Mark Certified…" in the IPC screen.', 'warn')
            else:
                try:
                    ok, m = D.certify_signed(int(no), ask=lambda q: self.ask(q))
                    if ok: extra = '   |   ' + m
                    elif 'already' not in m and 'not marked' not in m: self.msg(m, 'warn')
                except PermissionError:
                    self.msg('11-LOG IPC is open in Excel.\nClose it – the IPC screen will ask you again about the signed copy.', 'warn')
                except Exception as e:
                    LOG.exception('IPC certify (signed) failed'); self.msg(f'The signed copy was moved, but the IPC was not marked Certified:\n{e}', 'warn')
        D.drop_work(it)
        r = self._row(iid)
        self.tv.removeRow(r); self.items.pop(iid, None)
        self.info(f'Moved ✔  {os.path.relpath(dst, D.project_root())}' + ('' if guessed == code else '   (correction saved – it will learn)') + extra)
        if self.tv.rowCount(): self.select_row(min(r, self.tv.rowCount() - 1))
        else: self.clear_box()

    def mar_log(self, f, status):
        """تحديث سطر الـ MAR في MAR LOG (الحالة + تاريخ الختم) – في الخلفية عشان Excel بطيء."""
        reply = f.get('reply')
        if not reply:
            try: row = D.mar_row(f['out'])
            except Exception: row = None; LOG.exception('MAR LOG read failed')
            if row and not row.get('reply'):
                t, ok = QInputDialog.getText(self, 'MAR LOG', f"{row['no']} – {f['out']}\n\nDate on the consultant's stamp (DD/MM/YYYY)\n"
                                                              '(leave empty = status only):')
                t = (t or '').strip() if ok else ''
                m = re.fullmatch(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', t) or re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', t)
                if m:
                    a, b, c = m.groups()
                    try: reply = (dt.date(int(a), int(b), int(c)) if len(a) == 4 else dt.date(int(c), int(b), int(a))).isoformat()
                    except ValueError: reply = None
                elif t: self.msg(f'«{t}» is not a date – only the status is written.', 'warn')
        out, resub = f['out'], bool(f.get('resubmit'))
        self.info(f'Writing MAR LOG for {out}…')
        def done(r):
            if isinstance(r, Exception):
                LOG.exception('MAR LOG update failed', exc_info=r)
                return self.msg(f'The file was filed, but MAR LOG was not updated:\n{r}\n\nWrite the status by hand.', 'warn')
            if r: self.info(r)
        self.bg(lambda: D.mar_log_update(out, status, reply, resub), done)

    def _restore(self, aside):
        bad = []
        for p, q in aside:
            try: shutil.move(q, D.free_name(os.path.dirname(p), os.path.basename(p)))
            except OSError: bad.append(q); LOG.exception('Restore failed: %s', q)
        if bad: self.msg('These old copies stayed in 4_ARCHIVED – move them back by hand:\n\n' + '\n'.join(bad), 'warn')

    def to_duplicates(self, iid, it, code, same):
        """الورقة مكررة ← 00_INBOX\\_duplicates (Undo last move بيرجّعها)."""
        dd = os.path.join(D.inbox(), DUP)
        try:
            os.makedirs(dd, exist_ok=True)
            dst = D.free_name(dd, os.path.basename(it['path'])); safe_move(it['path'], dst)
        except PermissionError as e:
            return self.msg(str(e), 'warn')
        except Exception as e:
            LOG.exception('Duplicate move failed: %s', it.get('path')); return self.msg(f'Could not move the file:\n{e}', 'err')
        D._log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='move', src=it['path'], dst=dst, type='DUPLICATE',
                    guessed=code, same_as=same, fields=it.get('fields') or {}))
        try: D._chat_done(it['path'])
        except Exception: pass
        D.drop_work(it)
        r = self._row(iid)
        self.tv.removeRow(r); self.items.pop(iid, None)
        self.info(f'Duplicate ✔  moved to 00_INBOX\\{DUP}  –  same as {rel(same)}')
        if self.tv.rowCount(): self.select_row(min(r, self.tv.rowCount() - 1))
        else: self.clear_box()

    # ------------------------------------------------------------ dialogs
    def _dialog(self, title, w=520):
        d = QDialog(self); d.setWindowTitle(title); d.setStyleSheet(CSS + f'QDialog {{ background: {TT["card2"]}; }}')
        d.setMinimumWidth(w)
        return d

    def review_contract(self, it):
        """شباك العقد: اسم المقاول بالإنجليزي، رقم العقد وتاريخه، والنسخة دي فيها أنهي اعتماد من الـ 3."""
        f = it.get('fields') or {}
        label = self.cb_sub.currentText().strip()
        if not label: self.msg('Choose the subcontractor first.', 'warn'); return None
        en0 = D.sub_english_folder(label) or ''
        try: d0 = dt.date.fromisoformat(f.get('date', '')).strftime('%d-%m-%Y')
        except (TypeError, ValueError): d0 = ''
        appr0 = set(f.get('appr') or [])
        d = self._dialog('Subcontract – approvals')
        g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 16); g.setVerticalSpacing(8)
        t = QLabel(f'Subcontract – {label}'); t.setFont(HQ.font(16, True)); g.addWidget(t, 0, 0, 1, 2)
        e_en, e_no, e_d = QLineEdit(str(en0)), QLineEdit(str(f.get('ct') or '')), QLineEdit(d0)
        for i, (txt, e) in enumerate((('English name (folder & file):', e_en), ('Contract No. (as written):', e_no),
                                      ('Contract date (dd-mm-yyyy):', e_d)), 1):
            g.addWidget(QLabel(txt), i, 0); g.addWidget(e, i, 1)
        h = QLabel('This copy has:'); h.setFont(HQ.font(13, True)); g.addWidget(h, 4, 0, 1, 2)
        names = {'legal': 'Legal Affairs stamp', 'sub': "Subcontractor's own signature", 'chair': 'Chairman of the Board approval'}
        vs = {}
        for i, k in enumerate(('legal', 'sub', 'chair')):
            vs[k] = QCheckBox(names[k]); vs[k].setChecked(k in appr0); g.addWidget(vs[k], 5 + i, 0, 1, 2)
        l_reg, l_prev = QLabel(''), QLabel('')
        l_reg.setStyleSheet(f'color: {MUTED};'); l_prev.setStyleSheet(f'color: {NAVY}; font-weight: bold;')
        g.addWidget(l_reg, 8, 0, 1, 2); g.addWidget(l_prev, 9, 0, 1, 2)
        its, si = f.get('items') or [], f.get('sub_info') or {}
        tot = sum(C_num(x.get('rate')) for x in its)
        mk = QCheckBox('Create the subcontractor in SUB' + (f' – {len(its)} contract items read (total of rates {tot:,.0f} SAR)' if its else ' (no items read)')
                       + (f" · {C_num(si.get('duration_months')):g} months" if C_num(si.get('duration_months')) else '')
                       + (f" · Retention {C_num(si.get('retention_pct')):g}%" if C_num(si.get('retention_pct')) else ''))
        mk.setChecked(True); mk.setStyleSheet('color: #2C7536; font-weight: bold;')
        g.addWidget(mk, 10, 0, 1, 2)
        S = D._sub()
        try: reg = S.contracts()
        except Exception: reg = []
        out = {}

        def calc():
            en = ' '.join(e_en.text().split()); no = e_no.text().strip()
            appr = {k: v.isChecked() for k, v in vs.items()}
            old = next((c for c in reg if c['no'] == no and (S._same(c['sub'], en) or S._same(c['sub'], label))), None)
            if old:
                have = {k: appr[k] or old['have'][k] for k in appr}
                l_reg.setText('In the register: ' + '  '.join(('✔ ' if old['have'][k] else '✘ ') + s for k, _, s in S.APPROVALS))
            else:
                have = appr; l_reg.setText('New contract in the register')
            miss = S.ct_missing(have)
            try: dd = dt.datetime.strptime(e_d.text().strip(), '%d-%m-%Y').date()
            except ValueError: dd = None
            name = f"CT-{C.ct_no(no) or 'XXXX'} - {en or '?'} - {S.ct_stage(appr)}{f' ({dd:%Y-%m-%d})' if dd else ''}.pdf"
            l_prev.setText(f'{name}\nAfter this copy: ' + (f"missing {', '.join(miss)}" if miss else 'complete ✔'))
            return en, no, dd, appr, name
        for e in (e_en, e_no, e_d): e.textChanged.connect(lambda *_: calc())
        for c in vs.values(): c.toggled.connect(lambda *_: calc())
        calc()

        def save():
            en, no, dd, appr, name = calc()
            if not S.english_ok(en): return QMessageBox.warning(d, 'DC', 'English name: English letters, numbers and spaces.')
            if len(en) > S.EN_MAX: return QMessageBox.warning(d, 'DC', f'English name: {S.EN_MAX} letters at most.')
            if not C.ct_no(no): return QMessageBox.warning(d, 'DC', 'Write the contract number.')
            if e_d.text().strip() and not dd: return QMessageBox.warning(d, 'DC', 'Date format: dd-mm-yyyy')
            out.update(en=en, no=no, date=dd, appr=appr, name=name, mksub=mk.isChecked()); d.accept()
        bb = QHBoxLayout(); bb.addStretch(1)
        bb.addWidget(btn('Save & file', save, True)); bb.addWidget(btn('Cancel', d.reject))
        g.addLayout(bb, 11, 0, 1, 2)
        d.exec()
        return out or None

    # ------------------------------------------------------------ Claude
    def claude_state(self):
        k = CL.key()
        self.l_claude.setText('☁ ' + ('Claude: ON' if CL.enabled() else ('Claude: off' if k else 'Claude: no key')))

    def send_claude(self):
        iid, it = self.sel()
        if not it: return self.msg('Select the file first.')
        if self.busy: return self.msg('Wait until reading finishes.')
        if it.get('claude'): return self.msg('This file is already in the Claude folder.')
        if it['path'].lower().endswith('.zip'): return self.msg('Claude cannot read a zip – choose the type yourself.')
        try: D.send_to_claude(it['path'])
        except Exception as e:
            LOG.exception('Send to Claude failed: %s', it.get('path')); return self.msg(f'Could not move the file:\n{e}', 'err')
        m = 'شوف فولدر Claude'
        try: QGuiApplication.clipboard().setText(m)
        except Exception: pass
        self.scan()
        self.info(f'Sent to INBOX\\{D.CLAUDE_DIR} ✔  – hidden until Claude reads it  –  in the chat write: «{m}» (copied – just paste it)')

    def split_sel(self):
        iid, it = self.sel()
        if not it: return
        if self.busy: return self.msg('Wait until reading finishes.')
        if self.split_item(iid): QTimer.singleShot(300, self.distribute)

    def split_item(self, iid, ask=True):
        it = self.items.get(iid); parts = it.get('parts') if it else None
        if not parts:
            self.msg('Claude did not find several papers in this file.'); return False
        def title(d):
            t = C.wir_title(d['number']) if d.get('type') == 'SITE_WIR' and str(d.get('number') or '').isdigit() else None
            return t or f"{C.EN.get(d['type'], d['type'])} {d.get('number') or ''}  {d.get('title') or ''}"
        lines = '\n'.join(f"• pages {d['first_page']}–{d['last_page']}:  {title(d)}".rstrip()
                          for d in parts[:15]) + ('\n…' if len(parts) > 15 else '')
        if ask and not self.ask(f"{it['name']}\nhas {len(parts)} papers (Claude):\n\n{lines}\n\nSplit it into {len(parts)} files?"):
            return False
        try: new = D.split_bundle(it['path'], parts, dest=os.path.join(D.inbox(), D.REVIEW))
        except Exception as e:
            LOG.exception('Split failed: %s', it.get('path')); self.msg(f'Could not split the file:\n{e}', 'err'); return False
        self.info(f'Split ✔  {len(new)} files in INBOX\\{D.REVIEW} – the original is kept in INBOX\\{D.ORIGINALS}')
        return True

    def claude_settings(self):
        d = self._dialog('Claude', 560)
        g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 16); g.setVerticalSpacing(8)
        st, k = CL.settings(), CL.key()
        t = QLabel('Claude reads the paper when the DC is not sure, and finds the papers inside a long scan.'); t.setWordWrap(True)
        g.addWidget(t, 0, 0, 1, 3)
        g.addWidget(QLabel('API key:'), 1, 0)
        e_key = QLineEdit(); e_key.setEchoMode(QLineEdit.Password); g.addWidget(e_key, 1, 1, 1, 2)
        s = QLabel((f'saved key: ••••{k[-4:]}' if k else 'no key saved yet') + ('  (from ANTHROPIC_API_KEY)' if os.environ.get('ANTHROPIC_API_KEY') else ''))
        s.setStyleSheet(f'color: {MUTED};'); g.addWidget(s, 2, 1, 1, 2)
        g.addWidget(QLabel('Model:'), 3, 0)
        cb_m = QComboBox(); ms = [str(x) for x in CL.MODELS]
        if CL.model() and str(CL.model()) not in ms: ms.append(str(CL.model()))
        cb_m.addItems(ms); cb_m.setCurrentText(str(CL.model())); g.addWidget(cb_m, 3, 1)
        c_auto = QCheckBox('Send automatically when the DC is not sure (and long scans)'); c_auto.setChecked(bool(st.get('auto', True)))
        g.addWidget(c_auto, 4, 1, 1, 2)
        p = QLabel(f'The key stays on this PC only:  {CL.CFG}'); p.setStyleSheet(f'color: {MUTED};'); p.setWordWrap(True); g.addWidget(p, 5, 0, 1, 3)
        l_msg = QLabel(''); l_msg.setStyleSheet(f'color: {NAVY};'); g.addWidget(l_msg, 6, 0, 1, 3)

        def save(close=True):
            kw = dict(model=cb_m.currentText(), auto=bool(c_auto.isChecked()))
            if e_key.text().strip(): kw['key'] = e_key.text().strip()
            CL.save_settings(**kw); self.claude_state()
            if close: d.accept()

        def test():
            save(close=False); l_msg.setText('Testing…')
            self.bg(CL.test, lambda r: l_msg.setText(f'✖ {r}' if isinstance(r, Exception) else f'✔ Claude answered ({r})'))

        def remove():
            if QMessageBox.question(d, 'DC', 'Remove the saved API key from this PC?') == QMessageBox.Yes:
                CL.forget_key(); self.claude_state(); d.accept()
        bb = QHBoxLayout(); bb.addStretch(1)
        bb.addWidget(btn('Test', test)); bb.addWidget(btn('Remove key', remove))
        bb.addWidget(btn('Save', lambda: save(), True)); bb.addWidget(btn('Close', d.reject))
        g.addLayout(bb, 7, 0, 1, 3)
        d.exec()

    def undo(self):
        if self.busy: return self.msg('Wait until reading finishes.')
        if not self.ask('Move the last moved file back to INBOX?'): return
        try: e = D.undo_last()
        except Exception as ex: return self.msg(str(ex), 'err')
        if not e: return self.msg('Nothing to undo.')
        self.scan(); self.info(f"Undone ✔  {os.path.basename(e['dst'])} is back in INBOX")

    # ------------------------------------------------------------ شيت DC.xlsx
    def bg(self, fn, done):
        def run():
            try: r = fn()
            except Exception as e: r = e
            self.q.put(lambda: done(r))
        threading.Thread(target=run, daemon=True).start()

    def refresh_next(self):
        def got(r):
            self.l_next.setText('DC sheet: not found' if isinstance(r, Exception) else f'✉  Next YA letter: {r or "—"}')
        self.bg(lambda: SH.next_number('YA'), got)
        try: n = len(SH.load_pending())
        except Exception: n = 0
        self.b_pend.setText(f'Update DC sheet  ({n})' if n else 'Update DC sheet')

    def log_sheet(self, entry):
        if not entry: return
        def done(r):
            ok, m = r if isinstance(r, tuple) else (False, str(r))
            self.info(self.l_info.text() + ('   |  DC sheet ✔' if ok else '   |  DC sheet: waiting – press Update DC sheet'))
            self.refresh_next()
        self.bg(lambda: SH.write([entry]), done)

    def flush(self):
        def done(r):
            ok, m = r if isinstance(r, tuple) else (False, str(r))
            (QMessageBox.information if ok else QMessageBox.warning)(self, 'DC sheet', m or 'Done'); self.refresh_next()
        self.info('Writing to DC sheet…'); self.bg(SH.flush_pending, done)

    def register(self):
        """تضيف أي سطر للشيت بإيدك: خطاب صادر برقمه أو وارد."""
        d = self._dialog('DC Register', 900); d.resize(900, 620)
        v = QVBoxLayout(d); v.setContentsMargins(18, 14, 18, 14)
        g = QGridLayout(); g.setVerticalSpacing(6)
        kinds = list(SH.KINDS)
        cb = QComboBox(); cb.addItems([SH.LABEL[x] for x in kinds]); cb.setCurrentText(SH.LABEL['YA'])
        E = {k: QLineEdit() for k in ('no', 'date', 'party', 'desc', 'file')}
        E['date'].setText(f'{dt.date.today():%Y-%m-%d}')
        rows = [('Register:', None), ('Letter No.:', 'no'), ('Date (yyyy-mm-dd):', 'date'), ('To / From:', 'party'),
                ('Subject:', 'desc'), ('File:', 'file')]
        for i, (lbl, k) in enumerate(rows):
            l = QLabel(lbl); l.setAlignment(Qt.AlignRight | Qt.AlignVCenter); g.addWidget(l, i, 0)
            g.addWidget(cb if k is None else E[k], i, 1)
        g.addWidget(btn('Browse…', lambda: E['file'].setText(QFileDialog.getOpenFileName(d, 'File', D.project_root())[0] or E['file'].text())), 5, 2)
        g.setColumnStretch(1, 1)
        v.addLayout(g)
        tv = QTableWidget(0, 3); tv.setHorizontalHeaderLabels(['NO.', 'DATE', 'SUBJECT'])
        tv.verticalHeader().setVisible(False); tv.setShowGrid(False); tv.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tv.setColumnWidth(0, 200); tv.setColumnWidth(1, 110); tv.horizontalHeader().setStretchLastSection(True)
        tv.setStyleSheet(f'QTableWidget {{ border: 1px solid {SOFT}; border-radius: 10px; }}')
        v.addWidget(tv, 1)
        l_msg = QLabel(''); l_msg.setStyleSheet(f'color: {C_OK};'); v.addWidget(l_msg)

        def kind(): return kinds[cb.currentIndex()]

        def load(*_):
            k = kind(); l_msg.setText('Reading the DC sheet…')
            def got(r):
                tv.setRowCount(0)
                if isinstance(r, Exception): return l_msg.setText(f'Cannot read the sheet: {r}')
                rows_, nxt = r
                for x in [x for x in rows_ if not SH._empty(x.get('desc'))][-15:]:
                    dd = x.get('date'); dd = dd.strftime('%d/%m/%Y') if hasattr(dd, 'strftime') else (dd or '')
                    n = tv.rowCount(); tv.insertRow(n)
                    for c, val in enumerate((x.get('no') or '', dd, x.get('desc') or '')): tv.setItem(n, c, QTableWidgetItem(str(val)))
                E['no'].setText(str(nxt or '') if k != 'IN' else '')
                l_msg.setText(f'Next number: {nxt}' if nxt else '')
            self.bg(lambda: (SH.read(k), SH.next_number(k)), got)

        def add():
            k = kind()
            try: dd = dt.date.fromisoformat(E['date'].text().strip())
            except ValueError: return QMessageBox.warning(d, 'DC', 'Date must be yyyy-mm-dd')
            if not E['desc'].text().strip(): return QMessageBox.warning(d, 'DC', 'Write the subject.')
            e = SH.make_entry(k, E['no'].text(), dd, E['desc'].text().strip(), E['file'].text().strip(),
                              frm=E['party'].text().strip() if k == 'IN' else '', to=E['party'].text().strip() if k != 'IN' else '')
            l_msg.setText('Writing to the DC sheet through Excel…')
            def done(r):
                ok, m = r if isinstance(r, tuple) else (False, str(r))
                l_msg.setText('Added ✔' if ok else m); E['desc'].clear(); E['file'].clear(); load(); self.refresh_next()
            self.bg(lambda: SH.write([e]), done)
        bb = QHBoxLayout()
        bb.addWidget(btn('Add to DC sheet', add, True)); bb.addWidget(btn('Open DC sheet', lambda: self.open_path(SH.sheet_path())))
        bb.addStretch(1); bb.addWidget(btn('Close', d.reject))
        v.addLayout(bb)
        cb.currentIndexChanged.connect(load); load()
        d.exec()

    # ------------------------------------------------------------ helpers
    def open_path(self, p):
        try:
            if os.name == 'nt': os.startfile(p)
            else: subprocess.Popen(['xdg-open', p])
        except Exception as e:
            self.msg(str(e), 'err')

    def open_sel(self):
        iid, it = self.sel()
        if it: self.open_path(it['path'])

    def _text_win(self, title, txt):
        d = self._dialog(title, 820); d.resize(820, 560); d.setAttribute(Qt.WA_DeleteOnClose)
        v = QVBoxLayout(d); v.setContentsMargins(10, 10, 10, 10)
        t = QPlainTextEdit(); t.setReadOnly(True); t.setPlainText(txt); t.setFont(HQ.font(13))
        t.setStyleSheet(f'border: 1px solid {SOFT}; border-radius: 8px;')
        v.addWidget(t)
        d.show()                                          # مش modal زي القديم

    def show_text(self):
        iid, it = self.sel()
        if it: self._text_win(f'Text read – {it["name"]}', f"[{it.get('method') or 'not read'}]\n\n{it.get('text') or ''}")

    def show_log(self):
        L = D.load_moves()[-200:][::-1]
        rows = [f"{e['time']}  {'UNDONE  ' if e.get('undone') else ''}{e.get('act')}  {e.get('type') or ''}\n   {e['src']}\n → {e['dst']}" for e in L]
        self._text_win('Moves log (last 200)', '\n\n'.join(rows) or 'No moves yet.')

    def on_close(self):
        """main_qt بيسأل قبل ما يقفل/يغيّر الشاشة: لو لسه بيقرا ← استنى."""
        if self.busy:
            QMessageBox.information(self, 'DC', 'Wait until reading finishes.'); return False
        self.set_cursor_busy(False)
        return True
