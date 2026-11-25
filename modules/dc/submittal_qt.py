# -*- coding: utf-8 -*-
"""Submittal Builder – تجهيز أي submittal بكل متطلباته (طلب المستخدم 21-11-2026).

بيقرا المتطلبات من submittal_reqs.json (تقدر تعدّله بإيدك).
الخطوات: تختار النوع ← تملا البيانات ← ترفق المطلوب ← Build package (PDF واحد) ← يتحفظ في فولدره.
"""
import os, re, sys, json, shutil, threading, datetime as dt

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QHBoxLayout, QVBoxLayout, QGridLayout,
                               QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
                               QLineEdit, QComboBox, QMessageBox, QFileDialog, QApplication,
                               QPushButton)

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import dc_core as D
import supplier_files as SF                           # ورق المورد من فولدره في 13_PARTIES
import mar_xl as MX                                   # MAR ← سطر MAR LOG + الفورم من شيت MAR (الداتا في الإكسيل)
import dc_sheet as SH                                 # رقم الـ DC: اللي عليه الدور + حجزه في DC.xlsx
import dc_qt as DQ                                   # نفس الشكل الغامق والزراير
from dc_io import LOG
try:
    from dc_preview import PreviewPane
except Exception:
    PreviewPane = None

PLAT = os.path.dirname(os.path.dirname(HERE))
if PLAT not in sys.path: sys.path.insert(0, PLAT)
import home_qt as HQ

BG, CARD, INK, MUTED, SOFT = HQ.BG, HQ.CARD, HQ.INK, HQ.MUTED, HQ.SOFT_LINE
TT = HQ.T
GREEN = '#7FD98A' if HQ.DARK else '#2C7536'
AMBER = '#F0C860' if HQ.DARK else '#8A6100'
REDC = '#FF8A8A' if HQ.DARK else '#C00000'
ACCENT = '#7FB3E8' if HQ.DARK else '#1F5C99'
BTN_BG = TT.get('card2', CARD)

REQS_FILE = os.path.join(HERE, 'submittal_reqs.json')
COLS = [('ok', '', 30), ('n', 'NO.', 42), ('name', 'REQUIREMENT', 215),
        ('req', '', 104), ('file', 'FILE', 220), ('st', 'STATUS', 76)]
SF_COLS = [('FILE', 250), ('FOLDER', 150), ('USED IN', 80)]

CSS = f"""
QWidget#subm {{ background: {BG}; }}
QFrame#card {{ background: {CARD}; border: 1px solid {SOFT}; border-radius: 10px; }}
QLabel#k {{ color: {MUTED}; font: 10pt "Segoe UI"; }}
QLabel#sec {{ color: {ACCENT}; font: bold 10pt "Segoe UI"; letter-spacing: .5px; }}
QLineEdit, QComboBox {{ background: {BG}; color: {INK}; border: 1px solid {SOFT};
                        border-radius: 6px; padding: 5px 8px; font: 11pt "Segoe UI"; }}
QLineEdit:focus, QComboBox:focus {{ border-color: {ACCENT}; }}
QTableWidget {{ background: {CARD}; color: {INK}; border: none; gridline-color: transparent;
                font: 10.5pt "Segoe UI"; }}
QTableWidget::item {{ border-bottom: 1px solid {SOFT}; padding-left: 6px; }}
QHeaderView::section {{ background: {TT.get('card2', CARD)}; color: {MUTED}; border: none;
                        border-bottom: 1px solid {SOFT}; padding: 6px; font: bold 8.5pt "Segoe UI"; }}
QPushButton#t {{ background: {TT.get('card2', CARD)}; color: {INK}; border: 1px solid {SOFT};
                 border-radius: 8px; padding: 9px 12px; font: 11pt "Segoe UI"; text-align: left; }}
QPushButton#t:hover {{ border-color: {ACCENT}; }}
QPushButton#t:checked {{ background: {HQ.NAVY}; border-color: {ACCENT}; color: #FFFFFF; font-weight: bold; }}
"""


def _card():
    f = QFrame(); f.setObjectName('card'); return f


class SubmittalQt(QWidget):
    """شاشة تجهيز الـ submittals."""

    def __init__(self, open_section=None, run_old=None, log=None, embedded=False):
        """embedded=True ← جوه New Document: من غير عنوان ولا عمود الأنواع (النوع بيتختار من الصف اللي فوق)."""
        super().__init__()
        self.embedded = embedded
        self.setObjectName('subm')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(DQ.CSS + CSS)        # ستايل الزراير من Distribute + اللي فوق
        self.open_section = open_section or (lambda k: None)
        self.run_old, self.log = run_old, log
        self.types, self.cur, self.rows = [], None, []
        self.sfiles, self._parties, self.party_dir = [], {}, ''     # ورق المورد المختار
        self.root = D.project_root()
        self._load_reqs()
        self.l_info = QLabel(''); self.l_info.setFont(HQ.font(12.5, True))
        self.l_info.setStyleSheet(f'color: {GREEN};'); self.l_info.setWordWrap(True)
        self.l_prog = QLabel(''); self.l_prog.setObjectName('k'); self.l_prog.setWordWrap(True)
        self.tbtn = []
        self._build()
        if self.types and not embedded: self.pick(0)

    # ------------------------------------------------------------ المتطلبات
    def _load_reqs(self):
        try:
            with open(REQS_FILE, encoding='utf-8') as f:
                self.types = json.load(f).get('types', [])
        except Exception as e:
            LOG.exception('submittal_reqs.json'); self.types = []
            self._err = str(e)

    # ------------------------------------------------------------ الشكل
    def _build(self):
        v = QVBoxLayout(self); v.setSpacing(11)
        if self.embedded: v.setContentsMargins(0, 0, 0, 0)
        else:
            v.setContentsMargins(20, 14, 20, 16)
            v.addLayout(self._head())

        row = QHBoxLayout(); row.setSpacing(12)
        if not self.embedded: row.addWidget(self._left(), 0)
        row.addWidget(self._mid(), 1)
        if PreviewPane:
            self.pv = PreviewPane(self, width=330)
            row.addWidget(self.pv, 0)
        else:
            self.pv = None
        v.addLayout(row, 1)

    def _head(self):
        t = QHBoxLayout(); t.setSpacing(14)
        b = DQ.btn('←  Document Center', lambda: self.open_section('doc')); b.setObjectName('back')
        t.addWidget(b)
        box = QVBoxLayout(); box.setSpacing(0)
        c = QLabel('Technical Office Platform  ›  Document Center  ›  Submittal Builder')
        c.setFont(HQ.font(12.5, True)); c.setStyleSheet(f'color: {MUTED};'); box.addWidget(c)
        tt = QLabel(f'<span style="color:{GREEN}; font-size:20pt; font-weight:bold;">Submittal Builder</span>'
                    f'<span style="color:{MUTED}; font-size:11pt;">  ·  every submittal with all its requirements</span>')
        box.addWidget(tt)
        t.addLayout(box); t.addStretch(1)
        self.l_info.setMaximumWidth(520); self.l_info.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        t.addWidget(self.l_info)
        return t

    def _left(self):
        c = _card(); c.setFixedWidth(258)
        v = QVBoxLayout(c); v.setContentsMargins(11, 11, 11, 11); v.setSpacing(6)
        k = QLabel('1 · SUBMITTAL TYPE'); k.setObjectName('sec'); v.addWidget(k)
        for i, t in enumerate(self.types):
            b = QPushButton(f"{t['title']}\n{t.get('pmm', '')}")
            b.setObjectName('t'); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, j=i: self.pick(j))
            v.addWidget(b); self.tbtn.append(b)
        if not self.types:
            v.addWidget(QLabel('submittal_reqs.json could not be read'))
        v.addStretch(1)
        v.addWidget(self.l_prog)
        b = DQ.btn('Edit the requirements file', self.edit_reqs); v.addWidget(b)
        return c

    def _mid(self):
        c = _card()
        v = QVBoxLayout(c); v.setContentsMargins(13, 12, 13, 12); v.setSpacing(9)

        k = QLabel('2 · SUBMITTAL DATA'); k.setObjectName('sec'); v.addWidget(k)
        g = QGridLayout(); g.setHorizontalSpacing(9); g.setVerticalSpacing(7)

        def lab(t):
            l = QLabel(t); l.setObjectName('k')
            l.setAlignment(Qt.AlignRight | Qt.AlignVCenter); return l

        self.e_no = QLineEdit(); self.e_rev = QLineEdit('R0'); self.e_rev.setFixedWidth(64)
        self.e_date = QLineEdit(f'{dt.date.today():%d-%m-%Y}'); self.e_date.setFixedWidth(110)
        self.e_subj = QLineEdit(); self.e_party = QComboBox(); self.e_party.setEditable(True)
        # البند والمواصفة خانتين هنا، وفي MAR LOG عمود واحد (G) سطرين: الأول البند والتاني المواصفة (طلب المستخدم 21-11)
        self.e_boq = QLineEdit(); self.e_boq.setPlaceholderText('BOQ item no(s). – e.g. 1000001, 1000002')
        self.e_spec = QLineEdit(); self.e_spec.setPlaceholderText('Contract spec – e.g. Section 032000 · ASTM A615 Grade 60')
        # خانات سطر MAR LOG اللي الفورم بتتملى منها (بتظهر مع MAR بس)
        self.e_disc = QComboBox(); self.e_disc.addItems(list(MX.DISCIPLINES)); self.e_disc.setFixedWidth(150)
        self.e_factory = QLineEdit(); self.e_factory.setPlaceholderText('only if different from the supplier')
        self.e_prev = QLineEdit(); self.e_prev.setFixedWidth(190); self.e_prev.setPlaceholderText('for a resubmittal')
        self._log0 = None                                # قيم السطر زي ما اتقرت من السجل (عشان نعرف إيه اللي اتغيّر بإيدك)
        self.l_subj = lab('Subject'); self.l_party = lab('Party')

        # رقم التقديم = حاجتين جنب بعض (طلب المستخدم 21-11): رقم النوع (M 27) + الرقم المسحوب من الـ DC
        self.e_dc = QLineEdit(); self.e_dc.setPlaceholderText('CTR-YA-WA-26-####')
        self.e_dc.setStyleSheet(f'font: bold 11pt "Segoe UI"; color: {GREEN};')
        self.l_dc = lab('DC No.')
        self.b_dc = DQ.btn('Take next DC number', self.take_dc)
        self.e_no.setFixedWidth(170); self.dc_locked = False; self._dc_busy = False
        nb = QHBoxLayout(); nb.setSpacing(9); nb.setContentsMargins(0, 0, 0, 0)
        nb.addWidget(self.e_no); nb.addWidget(self.l_dc); nb.addWidget(self.e_dc, 1); nb.addWidget(self.b_dc)
        g.addWidget(lab('No.'), 0, 0);       g.addLayout(nb, 0, 1)
        g.addWidget(lab('Rev'), 0, 2);       g.addWidget(self.e_rev, 0, 3)
        g.addWidget(lab('Date'), 0, 4);      g.addWidget(self.e_date, 0, 5)
        g.addWidget(self.l_subj, 1, 0);      g.addWidget(self.e_subj, 1, 1, 1, 5)
        g.addWidget(self.l_party, 2, 0);     g.addWidget(self.e_party, 2, 1, 1, 5)
        g.addWidget(lab('BOQ item'), 3, 0);  g.addWidget(self.e_boq, 3, 1, 1, 5)
        g.addWidget(lab('Spec'), 5, 0);      g.addWidget(self.e_spec, 5, 1, 1, 3)
        self.l_folder = QLabel(''); self.l_folder.setObjectName('k')
        self.l_folder.setFont(HQ.font(10.5, fam=HQ.MONO))
        g.addWidget(self.l_folder, 5, 4, 1, 2)
        self.l_disc, l_fac, l_prev = lab('Discipline'), lab('Manufacturer'), lab('Previous No.')
        lb = QHBoxLayout(); lb.setSpacing(9); lb.setContentsMargins(0, 0, 0, 0)
        lb.addWidget(self.e_disc); lb.addWidget(l_fac); lb.addWidget(self.e_factory, 1); lb.addWidget(l_prev); lb.addWidget(self.e_prev)
        g.addWidget(self.l_disc, 6, 0);      g.addLayout(lb, 6, 1, 1, 5)
        self.log_w = (self.l_disc, self.e_disc, l_fac, self.e_factory, l_prev, self.e_prev)
        g.setColumnStretch(1, 1)
        v.addLayout(g)

        k2 = QLabel('3 · REQUIRED ATTACHMENTS'); k2.setObjectName('sec')
        self.tv = QTableWidget(0, len(COLS))
        self.tv.setHorizontalHeaderLabels([c2[1] for c2 in COLS])
        self.tv.verticalHeader().setVisible(False); self.tv.setShowGrid(False)
        self.tv.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tv.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tv.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tv.verticalHeader().setDefaultSectionSize(32)
        hh = self.tv.horizontalHeader(); hh.setHighlightSections(False)
        hh.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        for i, (_, _, w) in enumerate(COLS): self.tv.setColumnWidth(i, w)
        hh.setSectionResizeMode(len(COLS) - 2, QHeaderView.Stretch)
        self.tv.itemSelectionChanged.connect(self.show_sel)
        self.tv.itemDoubleClicked.connect(lambda *_: self.attach())
        tr = QHBoxLayout(); tr.setSpacing(12)
        lv = QVBoxLayout(); lv.setSpacing(7); lv.addWidget(k2); lv.addWidget(self.tv, 1); tr.addLayout(lv, 6)
        # ورق المورد (طلب المستخدم 21-11): تختار المورد اللي فاتحله فولدر ← ورقه يظهر هنا، والواضح يترفق لوحده
        self.sf_box = QWidget(); sv = QVBoxLayout(self.sf_box); sv.setContentsMargins(0, 0, 0, 0); sv.setSpacing(7)
        self.l_sf = QLabel('SUPPLIER FILES'); self.l_sf.setObjectName('sec'); sv.addWidget(self.l_sf)
        self.tf = QTableWidget(0, len(SF_COLS)); self.tf.setHorizontalHeaderLabels([c2[0] for c2 in SF_COLS])
        self.tf.verticalHeader().setVisible(False); self.tf.setShowGrid(False)
        self.tf.setSelectionBehavior(QAbstractItemView.SelectRows); self.tf.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tf.setEditTriggers(QAbstractItemView.NoEditTriggers); self.tf.setWordWrap(False)
        self.tf.verticalHeader().setDefaultSectionSize(32)
        fh = self.tf.horizontalHeader(); fh.setHighlightSections(False); fh.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        for i, (_, w) in enumerate(SF_COLS): self.tf.setColumnWidth(i, w)
        fh.setSectionResizeMode(0, QHeaderView.Stretch)
        self.tf.itemSelectionChanged.connect(self.show_sfile)
        self.tf.itemDoubleClicked.connect(lambda *_: self.attach_sfile())
        sv.addWidget(self.tf, 1)
        self.l_sfnote = QLabel('Click a file to see it  ·  double-click to attach it to the selected requirement')
        self.l_sfnote.setObjectName('k'); self.l_sfnote.setWordWrap(True); sv.addWidget(self.l_sfnote)
        tr.addWidget(self.sf_box, 5)
        v.addLayout(tr, 1)
        self.e_party.activated.connect(lambda *_: self.on_party(True))     # اختيار من القايمة ← يتطابق من جديد
        self.e_party.lineEdit().editingFinished.connect(self.on_party)

        bar = QHBoxLayout(); bar.setSpacing(8)
        for txt, fn, pri in (('+  Attach file', self.attach, False),
                             ('Remove', self.unattach, False),
                             ('Open folder', self.open_folder, False),
                             ('⬇  Build package (PDF)', self.build_pkg, True)):
            bar.addWidget(DQ.btn(txt, fn, pri))
        bar.addStretch(1)
        if self.embedded:                                # مفيش عمود شمال ← العدّاد وزرار الملف هنا
            self.l_prog.setWordWrap(False); bar.addWidget(self.l_prog)
            bar.addWidget(DQ.btn('Edit requirements', self.edit_reqs))
        v.addLayout(bar)
        if self.embedded:
            self.l_info.setWordWrap(False); v.addWidget(self.l_info)
        return c

    # ------------------------------------------------------------ النوع
    def pick(self, i):
        if not (0 <= i < len(self.types)): return
        self.cur = self.types[i]
        for j, b in enumerate(self.tbtn): b.setChecked(j == i)
        t = self.cur
        self.l_subj.setText(t.get('subject_label') or 'Subject')
        self.l_party.setText(t.get('party_label') or 'Party')
        self.l_folder.setText(t.get('folder', ''))
        self.dc_locked = False; self.e_dc.clear(); self.e_dc.setReadOnly(False)
        for w in (self.l_dc, self.e_dc, self.b_dc): w.setVisible(bool(t.get('dc')))
        self.b_dc.setEnabled(True); self.e_subj.clear()
        QApplication.setOverrideCursor(Qt.WaitCursor)    # MAR / MIR بيقروا السجل (ثواني)
        try:
            self.e_no.setText(self.next_no(t))
            self.party_list(t)
        finally:
            while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()
        self.rows = [dict(it, files=[], auto=False) for it in t.get('items', [])]
        if self.is_mar() and not any(r.get('form') for r in self.rows):      # ملف المتطلبات قديم ← الفورم برضه بتتعمل لوحدها
            for r in self.rows:
                if 'approval form' in str(r.get('name', '')).lower(): r['form'] = 'mar'; break
        for w in self.log_w: w.setVisible(self.is_mar())
        self._log0 = None; self.e_spec.clear(); self.e_boq.clear(); self.e_factory.clear(); self.e_prev.clear()
        self.e_disc.setCurrentIndex(0); self.e_rev.setText('R0'); self.e_date.setText(f'{dt.date.today():%d-%m-%Y}')
        self.sf_box.setVisible(bool(t.get('party_folder'))); self.sfiles, self.party_dir = [], ''; self.fill_sfiles()
        self.fill()
        if getattr(self, 'pv', None): self.pv.clear('Attach a file to see it here')
        self.info(f"{t['code']}  –  {len(self.rows)} requirement(s)")

    def pick_code(self, code):
        i = next((j for j, t in enumerate(self.types) if t.get('code') == code), -1)
        if i >= 0: self.pick(i)

    def prefill(self, no='', subject='', party='', dc=''):
        """«Edit» من القايمة: خانات السطر المختار (المرفقات بتتحط من جديد). dc = رقم الـ DC بتاعه – ثابت."""
        if no: self.e_no.setText(str(no))
        self.e_subj.setText(str(subject or '')); self.e_party.setEditText(str(party or ''))
        self.e_dc.setText(str(dc or ''))
        self.dc_locked = bool(dc)                        # رقم متسجّل قبل كده ← مايتغيرش ومايتكتبش في الشيت تاني
        self.e_dc.setReadOnly(self.dc_locked); self.b_dc.setEnabled(not self.dc_locked)
        self.on_party(True)                              # المورد ده ليه فولدر؟ ← ورقه يظهر ويترفق
        if self.is_mar() and no: self._from_log(str(no).strip(), bool(subject))

    def is_mar(self):
        """النوع ده بيتسجّل في MAR LOG والفورم بتاعته بتطلع من شيت MAR."""
        return bool(self.cur) and (self.cur.get('log') == 'mar' or self.cur.get('code') == 'MAR')

    def _from_log(self, no, keep_subject=True):
        """«Edit»: باقي خانات السطر من MAR LOG (البند، التصنيف، المصنع، التاريخ …) – زي ما هي في الإكسيل."""
        try:
            m = next((x for x in D._mat().load_marmir()['mar'] if x['no'] == no), None)
        except Exception:
            LOG.exception('read MAR LOG'); m = None
        if not m: return
        if not keep_subject or not self.e_subj.text().strip(): self.e_subj.setText(m.get('desc') or '')
        raw = str(m.get('spec') or '')
        boq, sp = MX.split_spec(raw)                     # عمود G: السطر الأول البند، والباقي المواصفة
        self.e_boq.setText(boq); self.e_spec.setText(sp)
        spec = MX.join_spec(boq, sp)
        if m.get('disc'): self.e_disc.setCurrentText(m['disc'])
        self.e_factory.setText(m.get('factory') or ''); self.e_prev.setText(m.get('prev') or '')
        if str(m.get('rev') or '').strip(): self.e_rev.setText(f"R{MX.rev_num(m['rev'])}")
        if m.get('sent'):
            try: self.e_date.setText(f"{m['sent']:%d-%m-%Y}")
            except Exception: pass
        self._log0 = dict(desc=_flat(self.e_subj.text()), spec=_flat(spec), disc=_flat(self.e_disc.currentText()),
                          factory=_flat(self.e_factory.text()), prev=_flat(self.e_prev.text()), rev=MX.rev_num(self.e_rev.text()))

    def _mar_log(self, no, dc, rev, subj, party, want_form):
        """سطر MAR LOG + الفورم PDF من شيت MAR. يرجّع (مسار الفورم أو ''، ملحوظة)؛ None = المستخدم لغى."""
        d = MX.parse_date(self.e_date.text())
        if not d: self.msg('Write the date as  dd-mm-yyyy.', 'warn'); return None
        spec = MX.join_spec(self.e_boq.text(), self.e_spec.text())
        mar = dict(no=no, ref=dc, rev=MX.rev_num(rev), desc=subj, supplier=MX.supplier_name(party),
                   disc=self.e_disc.currentText().strip(), spec=spec, date=d,
                   factory=self.e_factory.text().strip(), prev=self.e_prev.text().strip())
        edited = []
        if self._log0 is not None:                       # سطر موجود: اللي غيّرته بإيدك بس هو اللي يتكتب فوق القديم
            edited = [k for k in ('desc', 'spec', 'disc', 'factory', 'prev') if _flat(mar[k]) and _flat(mar[k]) != self._log0.get(k, '')]
            if mar['rev'] != self._log0.get('rev'): edited.append('rev')
        import tempfile
        fd, pdf = tempfile.mkstemp(suffix='.pdf', prefix='mar_form_'); os.close(fd)
        try: os.remove(pdf)
        except OSError: pass
        err = ''
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            M = D._mat()
            res = MX.make(M.P(M.MARMIR_REL), mar, pdf if want_form else None, edited)
            try: M._CACHE.pop(M.MARMIR_REL, None)
            except Exception: pass
        except MX.XlError as e:
            err = str(e)
        except Exception as e:
            LOG.exception('MAR LOG'); err = str(e)
        finally:
            while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()
        if err:
            if not self.ask(f'MAR LOG was not written and the form was not made:\n\n{err}\n\n'
                            'Build the package anyway (attachments only)?'): return None
            return '', 'MAR LOG: NOT written – ' + err.splitlines()[0][:120]
        self._log0 = dict(desc=_flat(subj), spec=_flat(' / '.join(spec.splitlines())), disc=_flat(mar['disc']),
                          factory=_flat(mar['factory']), prev=_flat(mar['prev']), rev=mar['rev'])
        note = f"MAR LOG: line {res.get('row')} {'added' if res.get('isnew') else 'updated'} ✔"
        form = pdf if want_form and os.path.isfile(pdf) else ''
        if want_form: note += '   ·   form made from sheet MAR ✔' if form else '   ·   the form PDF was NOT made – print it from sheet MAR'
        if '⚠' in str(res.get('warn') or ''): note += f"\n{res['warn']}"
        return form, note

    def uses_dc(self):
        return bool(self.cur and self.cur.get('dc'))

    def take_dc(self):
        """الرقم اللي عليه الدور في DC.xlsx (تبويب الاستشاري) – بيتحجز في الشيت وقت «Build package»."""
        if self._dc_busy or self.dc_locked: return
        self._dc_busy = True; self.b_dc.setEnabled(False); self.info('Reading the DC sheet…')
        res = {}; cur = self.cur

        def work():
            try: res['n'] = SH.next_number('YA')
            except Exception as e: res['e'] = e

        th = threading.Thread(target=work, daemon=True); th.start()

        def wait():
            if th.is_alive(): return QTimer.singleShot(250, wait)
            self._dc_busy = False
            try:
                self.b_dc.setEnabled(not self.dc_locked)
                if self.dc_locked or self.cur is not cur: return        # اتغيّر النوع أو فتح Edit وهو بيقرا ← الرقم ده مش ليه
                if res.get('n'):
                    self.e_dc.setText(str(res['n'])); self.info(f"DC number {res['n']} – it is written in the DC sheet when you build the package")
                else:
                    self.info('DC sheet not read – type the number' + (f" ({res['e']})" if res.get('e') else ''))
            except RuntimeError:
                pass                                     # الشاشة اتقفلت وهو بيقرا

        QTimer.singleShot(250, wait)

    def reserve_dc(self, dc, subject, out):
        """يكتب سطر الرقم في DC.xlsx (زي الخطابات): الرقم يتحجز ومايتاخدش تاني. يرجّع (تمام، رسالة)."""
        try:
            d = dt.datetime.strptime(self.e_date.text().strip(), '%d-%m-%Y').date()
        except ValueError:
            d = dt.date.today()
        try:
            r = SH.write([SH.make_entry('YA', dc, d, subject, out)])
            return r if isinstance(r, tuple) else (bool(r), '')
        except Exception as e:
            LOG.exception('reserve DC number'); return False, str(e)

    def party_list(self, t):
        """قايمة الجهات: الأنواع اللي ليها فولدر مورد (MAR / SUB) ← من 13_PARTIES. MIR ← شركات سجل المخزن."""
        self.e_party.blockSignals(True); self.e_party.clear()
        vals, self._parties = [], {}
        try:
            if t.get('party_folder'):
                root = D.project_root()
                for p in D.subs():
                    lab = f"{p['cat']} · {p['label']}"
                    vals.append(lab); self._parties[lab] = os.path.join(root, p['folder'])
            elif t['code'] == 'MIR':
                vals = D._mat().load_stock().get('companies', [])
        except Exception:
            LOG.exception('party list')
        self.e_party.addItems(vals); self.e_party.setEditText(''); self.e_party.blockSignals(False)

    # ------------------------------------------------------------ ورق المورد
    def _party_folder(self, text):
        """فولدر المورد من اللي مكتوب في الخانة: اختيار من القايمة، أو اسم (عربي / إنجليزي) معروف في 13_PARTIES."""
        text = (text or '').strip()
        if not text: return ''
        if text in self._parties: return self._parties[text]
        try:
            p = D.PT.find(D.project_root(), text.split('·')[-1].strip()) or D.PT.find(D.project_root(), text)
            return p['path'] if p else ''
        except Exception:
            return ''

    def on_party(self, force=False):
        """المورد اتغيّر ← ورقه يظهر، واللي اسمه واضح يترفق في بنده (البنود اللي اتحط فيها ملف بالإيد مابتتغيرش)."""
        if not (self.cur and self.cur.get('party_folder')): return
        d = self._party_folder(self.e_party.currentText())
        if d == self.party_dir and not (force and d): return        # نفس الفولدر (خروج من الخانة بس) ← مفيش حاجة تتغير
        self.party_dir = d
        for r in self.rows:                              # مرفقات المورد اللي فات (التلقائية بس) تتشال
            if r.get('auto'): r['files'], r['auto'] = [], False
        try: self.sfiles = SF.list_files(d) if d else []
        except Exception:
            LOG.exception('supplier files'); self.sfiles = []
        n = 0
        if self.sfiles:
            for key, path in SF.auto_match(self.sfiles, self.rows).items():
                for r in self.rows:
                    if r.get('n') == key and not r['files']:
                        r['files'], r['auto'] = [path], True; n += 1
        self.tf.clearSelection()
        self.fill(); self.fill_sfiles()
        if d: self.info(f'{len(self.sfiles)} file(s) in the supplier folder  ·  {n} attached automatically' if self.sfiles
                        else 'No papers found in this supplier folder (02_APPROVAL / 01_QUOTATIONS)')

    def fill_sfiles(self):
        used = {}
        for r in self.rows:
            for p in r.get('files', []): used.setdefault(p, []).append(str(r.get('n', '')))
        keep = self.tf.currentRow() if self.tf.selectionModel().hasSelection() else -1
        self.tf.blockSignals(True); self.tf.setRowCount(0)
        for f in self.sfiles:
            i = self.tf.rowCount(); self.tf.insertRow(i)
            u = ', '.join(used.get(f['path'], [])) or ('old package' if f['package'] else '—')
            for c, val in enumerate((f['name'], f['sub'], u)):
                it = QTableWidgetItem(str(val))
                if c == 1 or (c == 2 and f['path'] not in used): it.setForeground(_col(MUTED))
                if c == 2 and f['path'] in used: it.setForeground(_col(GREEN))
                self.tf.setItem(i, c, it)
        try: rel = os.path.relpath(self.party_dir, D.project_root()) if self.party_dir else ''
        except ValueError: rel = self.party_dir
        if 0 <= keep < self.tf.rowCount(): self.tf.selectRow(keep)
        self.tf.blockSignals(False)
        self.l_sf.setText(f'SUPPLIER FILES  ·  {rel}' if rel else 'SUPPLIER FILES  ·  choose the supplier')

    def sel_sfile(self):
        i = self.tf.currentRow()
        return self.sfiles[i] if 0 <= i < len(self.sfiles) and self.tf.selectionModel().hasSelection() else None

    def show_sfile(self):
        f = self.sel_sfile()
        if f and self.pv and os.path.isfile(f['path']): self.pv.load(f['path'])

    def attach_sfile(self):
        """دبل كليك على ملف المورد ← يتضاف للبند المختار في جدول المتطلبات."""
        f = self.sel_sfile()
        if not f: return
        i, r = self.sel()
        if r is None: return self.msg('Choose the requirement line on the left first, then double-click the file.', 'warn')
        if f['path'] in r['files']: return
        if r.get('auto'): r['files'], r['auto'] = [], False          # اختيارك بيبدّل التلقائي
        r['files'].append(f['path'])
        self.fill(); self.tv.selectRow(i); self.fill_sfiles()
        self.info(f"{f['name']}  →  {r['name']}")

    def next_no(self, t):
        src = t.get('number')
        try:
            if src in ('mar', 'mir'):
                mp = os.path.join(os.path.dirname(HERE), 'mat')
                if mp not in sys.path: sys.path.insert(0, mp)
                import mat_core as M
                mm = M.load_marmir()
                if src == 'mir': return M.next_mir(mm['mir'])
                return f"M{max([m['n'] for m in mm['mar']] or [0]) + 1}"
        except Exception:
            LOG.exception('next number')
        return ''

    # ------------------------------------------------------------ الجدول
    def fill(self):
        self.tv.setRowCount(0)
        for r in self.rows: self._add(r)
        self.progress()

    def _add(self, r):
        i = self.tv.rowCount(); self.tv.insertRow(i)
        fl = r.get('files') or []
        auto_form = bool(r.get('form')) and not fl       # الفورم: البرنامج بيعملها من السجل وقت الـ Build
        got = bool(fl) or auto_form
        shown = 'made by the program from MAR LOG  (at Build)' if auto_form else '—'
        if fl:
            shown = os.path.basename(fl[0]) + (f'  + {len(fl) - 1} more' if len(fl) > 1 else '') + ('   · auto' if r.get('auto') else '')
        vals = ['✔' if got else '', r.get('n', ''), r.get('name', ''),
                'MANDATORY' if r.get('req') else 'optional', shown,
                'Auto' if auto_form else 'Ready' if got else ('Missing' if r.get('req') else 'Optional')]
        for c, val in enumerate(vals):
            it = QTableWidgetItem(str(val))
            if c == 0: it.setForeground(_col(GREEN))
            if c == 3: it.setForeground(_col(AMBER if r.get('req') else MUTED))
            if c == 5: it.setForeground(_col(GREEN if got else (REDC if r.get('req') else MUTED)))
            if c == 4 and not fl: it.setForeground(_col(MUTED))
            if c == 4 and fl: it.setToolTip('\n'.join(os.path.basename(p) for p in fl))
            self.tv.setItem(i, c, it)

    def progress(self):
        need = [r for r in self.rows if r.get('req')]
        got = [r for r in need if r.get('files') or r.get('form')]
        n = sum(len(r.get('files') or []) for r in self.rows)
        sep = '   ·   ' if self.embedded else '\n'
        self.l_prog.setText(f'{len(got)} of {len(need)} mandatory ready{sep}{n} file(s) attached')

    def sel(self):
        i = self.tv.currentRow()
        if i < 0 or i >= len(self.rows) or not self.tv.selectionModel().hasSelection(): return None, None
        return i, self.rows[i]

    def show_sel(self):
        if not getattr(self, 'pv', None): return
        i, r = self.sel()
        fl = (r or {}).get('files') or []
        if fl and os.path.isfile(fl[0]): self.pv.load(fl[0])
        else: self.pv.clear('Attach a file to see it here')

    # ------------------------------------------------------------ أفعال
    def attach(self):
        """ملف (أو أكتر) من أي مكان ← يتضاف للبند المختار (البند ياخد أكتر من ملف: سجل + ضريبة + عنوان …)."""
        i, r = self.sel()
        if r is None: return self.msg('Choose a line from the list first.', 'warn')
        start = self.party_dir or self.root
        ps, _ = QFileDialog.getOpenFileNames(self, f"Attach – {r['name']}", start,
                                             'Papers (*.pdf *.png *.jpg *.jpeg *.tif *.tiff);;All files (*.*)')
        if not ps: return
        if r.get('auto'): r['files'], r['auto'] = [], False
        r['files'] += [p for p in ps if p not in r['files']]
        self.fill(); self.tv.selectRow(i); self.fill_sfiles()
        self.info(f"{len(ps)} file(s)  →  {r['name']}")

    def unattach(self):
        """يشيل آخر ملف من البند المختار."""
        i, r = self.sel()
        if r is None or not r.get('files'): return
        r['files'].pop(); r['auto'] = False
        self.fill(); self.tv.selectRow(i); self.fill_sfiles()
        if self.pv and not r['files']: self.pv.clear('Attach a file to see it here')

    def open_folder(self):
        if not self.cur: return
        p = os.path.join(D.project_root(), self.cur.get('folder', ''))
        try: os.makedirs(p, exist_ok=True)
        except Exception: pass
        self.open_path(p)

    def edit_reqs(self):
        self.open_path(REQS_FILE)

    def open_path(self, p):
        try:
            if os.name == 'nt': os.startfile(p)
            else:
                import subprocess; subprocess.Popen(['xdg-open', p])
        except Exception as e:
            self.msg(str(e), 'err')

    # ------------------------------------------------------------ الباكدج
    def build_pkg(self):
        if not self.cur: return
        miss = [r['name'] for r in self.rows if r.get('req') and not r.get('files') and not r.get('form')]
        if miss:
            if not self.ask('These mandatory papers are missing:\n\n  •  '
                            + '\n  •  '.join(miss[:8])
                            + '\n\nBuild the package anyway?'): return
        parts = [p for r in self.rows for p in (r.get('files') or []) if os.path.isfile(p)]
        want_form = self.is_mar() and any(r.get('form') and not r.get('files') for r in self.rows)
        if not parts and not want_form: return self.msg('Attach at least one file first.', 'warn')

        no = self.e_no.text().strip() or self.cur['code']
        rev = self.e_rev.text().strip()
        subj = self.e_subj.text().strip() or self.cur['code']
        party = self.e_party.currentText().strip()
        dc = self.e_dc.text().strip() if self.uses_dc() else ''
        if self.is_mar():                                # MAR بيتسجّل في MAR LOG ← لازم رقمه ومادته ومورده
            lack = [n for n, v in (('the M number', self.e_no.text().strip()), ('the DC number', dc),
                                   ('the material', self.e_subj.text().strip()), ('the supplier', party)) if not v]
            if lack: return self.msg('Fill these first – the MAR LOG line and the form need them:\n\n  •  ' + '\n  •  '.join(lack), 'warn')
        if self.uses_dc() and not dc and not self.ask('There is no DC number.\n\nBuild the package without it?'): return
        if dc and not self.dc_locked:                    # رقم جديد: لازم يكون فاضي في الشيت (مش واخده ورقة تانية)
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                hit = SH.find_row('YA', SH._tail(dc)); used = bool(hit) and not SH._empty(hit.get('desc')); err = ''
            except Exception as e:
                used, err, hit = False, str(e), None
            finally:
                while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()
            if used:
                return self.msg(f"{dc} is already used in the DC sheet:\n\n{hit.get('desc')}\n\nPress «Take next DC number».", 'warn')
            if err and not self.ask(f'Could not check {dc} in the DC sheet:\n{err}\n\nBuild anyway?'): return
        pname = MX.supplier_name(party) if self.cur.get('party_folder') else party      # «SUP · 18-NAME · عربي» ← الاسم بس
        base = ' - '.join(x for x in (no, dc, rev if rev and rev.upper() not in ('R0', '0') else '',
                                      subj, pname) if x)                         # M27 - CTR-YA-WA-26-0079 - المادة - المورد
        base = re.sub(r'\s+', ' ', re.sub(r'[\\/:*?"<>|]', '-', base)).strip(' .-')[:120]

        folder = os.path.join(D.project_root(), self.cur.get('folder', ''))
        try: os.makedirs(folder, exist_ok=True)
        except Exception as e: return self.msg(f'Could not make the folder:\n{e}', 'err')
        out = os.path.join(folder, base + '.pdf')
        n = 1
        while os.path.exists(out):
            n += 1; out = os.path.join(folder, f'{base} ({n}).pdf')

        form_pdf, log_note = '', ''
        if self.is_mar():                                # الأول السجل والفورم (الإكسيل هو الأصل)، وبعدين الملف
            got = self._mar_log(no, dc, rev, subj, party, want_form)
            if got is None: return
            form_pdf, log_note = got
            if form_pdf: parts = [form_pdf] + parts
            if not parts: return self.msg(log_note + '\n\nThere is nothing to put in the package – attach the papers and build again.', 'warn')
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            made, skipped = self._merge(parts, out)
        except Exception as e:
            LOG.exception('build package')
            return self.msg((log_note + '\n\n' if log_note else '') + f'Could not build the package:\n{e}', 'err')
        finally:
            while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()
            if form_pdf:
                try: os.remove(form_pdf)
                except OSError: pass

        note = f'{made} page(s) in:\n{out}'
        if log_note: note += '\n\n' + log_note
        if skipped: note += '\n\nNot added (not a PDF or could not be read):\n  •  ' + '\n  •  '.join(skipped[:6])
        if dc and not self.dc_locked:                    # رقم جديد ← يتحجز في DC.xlsx
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try: ok, m = self.reserve_dc(dc, ' - '.join(x for x in (subj, pname) if x), out)
            finally:
                while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()
            if ok and '(0 cells)' in str(m): ok = False       # السطر كان مليان ← ماتكتبش حاجة
            note += f'\n\nDC sheet: {dc} ' + ('written ✔' if ok else f'NOT written yet ({m}) – press «Update DC sheet» in Distribute')
            if ok or 'Saved for later' in str(m):            # اتكتب أو مستني في الانتظار ← مايتبعتش تاني
                self.dc_locked = True; self.e_dc.setReadOnly(True); self.b_dc.setEnabled(False)
        self.msg(note)
        self.info(f'Package ready ✔   {os.path.basename(out)}')
        if self.pv: self.pv.load(out)

    def _merge(self, parts, out):
        """يدمج الـ PDF والصور في ملف واحد. يرجّع (عدد الصفحات، اللي ما اتضمّش)."""
        import pypdf
        tmp = []
        w = pypdf.PdfWriter(); skipped = []
        try:
            for p in parts:
                ext = os.path.splitext(p)[1].lower()
                src = p
                if ext in ('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp'):
                    src = self._img_pdf(p)
                    if not src: skipped.append(os.path.basename(p)); continue
                    tmp.append(src)
                try:
                    for pg in pypdf.PdfReader(src).pages: w.add_page(pg)
                except Exception:
                    skipped.append(os.path.basename(p))
            n = len(w.pages)
            if not n: raise RuntimeError('nothing could be read')
            t = out + '.tmp'
            with open(t, 'wb') as fh: w.write(fh)
            os.replace(t, out)
            return n, skipped
        finally:
            for t2 in tmp:
                try: os.remove(t2)
                except Exception: pass

    def _img_pdf(self, img):
        """صورة ← صفحة PDF مؤقتة (بـ PyMuPDF لو موجود)."""
        try:
            try: import fitz
            except Exception: import pymupdf as fitz
            r = fitz.open(img)
            pdfbytes = r.convert_to_pdf(); r.close()
            import tempfile
            fd, path = tempfile.mkstemp(suffix='.pdf'); os.close(fd)
            with open(path, 'wb') as fh: fh.write(pdfbytes)
            return path
        except Exception:
            LOG.exception('image to pdf: %s', img); return None

    # ------------------------------------------------------------ رسايل
    def info(self, t): self.l_info.setText(t)

    def msg(self, t, kind='info'):
        {'info': QMessageBox.information, 'warn': QMessageBox.warning,
         'err': QMessageBox.critical}[kind](self, 'Submittal Builder', t)

    def ask(self, t):
        return QMessageBox.question(self, 'Submittal Builder', t,
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes

    def on_close(self): return True


def _flat(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def _col(hexs):
    return QColor(hexs)
