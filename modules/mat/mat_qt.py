# -*- coding: utf-8 -*-
"""MAR / MIR | Materials بـ Qt – نفس شاشة mat_app.py في الشغل (mat_core زي ما هو)، والشكل الجديد الغامق.
تابات: MIR – MAR – Stock – Deliveries – Issues to Site.
كليك على خانة Unifier = علامة «اترفع» (كليك تاني يشيلها).  دبل كليك على صف = يفتح ملفه."""
import os, sys, shutil, traceback, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
PLAT = os.path.dirname(os.path.dirname(HERE))
for p in (PLAT, os.path.join(PLAT, 'modules', 'dc'), os.path.join(PLAT, 'modules', 'qs')):
    if p not in sys.path: sys.path.insert(0, p)

from PySide6.QtCore import Qt, QTimer, QDate
from PySide6.QtGui import QColor, QGuiApplication, QPixmap
from PySide6.QtWidgets import (QStyledItemDelegate, QStyle, QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem,
                               QAbstractItemView, QComboBox, QMessageBox, QTabWidget, QDialog, QLineEdit, QDateEdit,
                               QCheckBox, QScrollArea, QGroupBox)
import home_qt as HQ
import dc_qt as DQ
import mat_core as M
import mat_defs as OLDM                     # نفس الأعمدة والتخمينات (best_company / best_item / guess_type / route_choices)

INK, MUTED, NAVY = DQ.INK, DQ.MUTED, DQ.NAVY
TT = HQ.T
CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#matq') + f"""
QTabWidget::pane {{ border: none; border-top: 2px solid {NAVY}; }}
QTabBar::tab {{ background: {TT['chip']}; color: {MUTED}; padding: 7px 16px; margin-right: 2px; border-top-left-radius: 8px;
               border-top-right-radius: 8px; font: 600 10pt "Segoe UI"; }}
QTabBar::tab:selected {{ background: {NAVY}; color: white; }}
QGroupBox {{ color: {INK}; font: bold 10pt "Segoe UI"; border: 1px solid {DQ.SOFT}; border-radius: 10px; margin-top: 12px; padding: 10px 8px 8px 8px; }}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px; }}
QDateEdit {{ background: {TT['card2']}; border: 1px solid {DQ.SOFT}; border-radius: 8px; padding: 5px 9px; color: {INK}; }}
"""
TAG_BG = ({'appr': '#1F3A24', 'sign': '#3A3018', 'wait': '#222E3A', 'none': None, 'resub': '#3E2A1E', 'nomir': '#3A3018', 'neg': None, 'rej': None}
          if HQ.DARK else {'appr': '#E2EFDA', 'sign': '#FFF2CC', 'wait': '#EEF3F8', 'none': None, 'resub': '#FCE4D6', 'nomir': '#FFF2CC',
                           'neg': None, 'rej': None})
RED = '#FF8A8A' if HQ.DARK else '#C00000'
LEFT = ('desc', 'name', 'file', 'dn', 'sup', 'status', 'route', 'note', 'ref', 'company', 'mat', 'dl')
fd = OLDM.fd


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


def table(cols, multi=False):
    t = QTableWidget(0, len(cols))
    t.setHorizontalHeaderLabels([c[1] for c in cols]); t.verticalHeader().setVisible(False)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.ExtendedSelection if multi else QAbstractItemView.SingleSelection)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers); t.setShowGrid(False); t.setWordWrap(False)
    t.verticalHeader().setDefaultSectionSize(28); t.horizontalHeader().setHighlightSections(False)
    for i, c in enumerate(cols): t.setColumnWidth(i, int(c[2] * 1.1))
    t.horizontalHeader().setStretchLastSection(True)
    t.setItemDelegate(BgDelegate(t))
    return t


def qd(d):
    d = d or dt.date.today()
    return QDate(d.year, d.month, d.day)


def pyd(q): return dt.date(q.year(), q.month(), q.day())


def date_edit(d=None):
    e = QDateEdit(qd(d)); e.setCalendarPopup(True); e.setDisplayFormat('dd-MM-yyyy'); e.setMinimumWidth(130)
    return e


def combo(values, text='', editable=True, w=None):
    c = QComboBox(); c.setEditable(editable); c.addItems([str(v) for v in values])
    if editable: c.setCurrentText(text or '')
    elif text: c.setCurrentIndex(max(c.findText(text), 0))
    if w: c.setMinimumWidth(w)
    return c


# ---------------------------------------------------------------- إذن التوريد / الـ MIR
class DeliveryDlg(QDialog):
    """نفس DeliveryDialog القديم: بيانات التوريد وأصنافه + الـ MIR بتاعه (جنب صورة الإذن لو موجودة).
    existing=True ← توريد متسجل قبل كده: الـ MIR بس."""

    def __init__(self, parent, stock, marmir, prefill=None, pictures=None, existing=False):
        super().__init__(parent, Qt.Window)
        self.setWindowTitle('Make MIR for this delivery' if existing else 'Delivery Note → Stock + MIR')
        self.setStyleSheet(CSS + f'QDialog {{ background: {HQ.BG}; }}')
        self.result_, self.stock, self.mm, self.existing = None, stock, marmir, existing
        pf = prefill or {}
        self.names = [it['name'] for it in stock['items']]
        self.units = sorted(set(M.UNITS) | {it['unit'] for it in stock['items'] if it['unit']})
        root = QHBoxLayout(self); root.setContentsMargins(16, 14, 16, 14); root.setSpacing(14)
        self.pics = [p for p in (pictures or []) if p and os.path.isfile(p)]
        if self.pics: root.addLayout(self._viewer())
        form = QVBoxLayout(); form.setSpacing(8); root.addLayout(form, 1)
        en = not existing
        # ---- التوريد
        g1 = QGroupBox('Delivery'); l1 = QGridLayout(g1)
        self.e_date = date_edit(pf.get('date')); self.e_date.setEnabled(en)
        self.e_dn = QLineEdit(pf.get('dn', '')); self.e_inv = QLineEdit(pf.get('inv', ''))
        self.e_comp = combo(stock['companies'], OLDM.best_company(pf.get('company', ''), stock['companies']) if pf.get('company') else '')
        texts = [x.get('name', '') for x in pf.get('items', [])]
        self.e_type = combo(stock['types'], pf.get('type') or OLDM.guess_type(texts, stock['types']))
        for w in (self.e_dn, self.e_inv, self.e_comp, self.e_type): w.setEnabled(en)
        self.l_file = QLabel(''); self.l_file.setStyleSheet(f'color: {MUTED};')
        for r, (a, wa, b, wb) in enumerate((('Date', self.e_date, 'DN No.', self.e_dn), ('Invoice No.', self.e_inv, 'Company', self.e_comp),
                                            ('Type', self.e_type, '', self.l_file))):
            l1.addWidget(QLabel(a), r, 0); l1.addWidget(wa, r, 1)
            if b: l1.addWidget(QLabel(b), r, 2)
            l1.addWidget(wb, r, 3)
        self.e_comp.currentTextChanged.connect(self._file); self.e_dn.textChanged.connect(self._file)
        form.addWidget(g1)
        # ---- الأصناف
        g2 = QGroupBox('Items (each line goes to the deliveries sheet)'); self.l2 = QGridLayout(g2)
        for j, h in enumerate(('Material (as in the stock list)', 'Code / Model', 'Qty', 'Unit', 'Status', 'Note')):
            x = QLabel(h); x.setStyleSheet(f'color: {MUTED}; font-size: 8pt; font-weight: bold;'); self.l2.addWidget(x, 0, j)
        self.rows = []
        for it in (pf.get('items') or [{}])[:8]: self._item_row(it, en)
        if en:
            self.b_add = DQ.btn('+ Add item', lambda: self._item_row({}, True)); self.l2.addWidget(self.b_add, 99, 0, Qt.AlignLeft)
        form.addWidget(g2)
        # ---- الـ MIR
        g3 = QGroupBox('MIR'); l3 = QGridLayout(g3)
        self.c_make = QCheckBox('Make the MIR now – PDF ready for signature in 05_Quality\\MIR\\1_WIP'); self.c_make.setChecked(True)
        if en: l3.addWidget(self.c_make, 0, 0, 1, 4)
        self.mar_list = sorted(marmir['mar'], key=lambda m: m['n'])
        self.e_mno = QLineEdit(M.next_mir(marmir['mir'])); self.e_rev = QLineEdit('1'); self.e_rev.setMaximumWidth(70)
        self.e_mdate = date_edit()
        self.e_mar = combo([M.mar_label(m) for m in self.mar_list], '', True, 380)
        self.e_mdesc, self.e_model, self.e_orig = QLineEdit(), QLineEdit(), QLineEdit(pf.get('origin', ''))
        self.e_serial, self.e_mqty = QLineEdit(), QLineEdit()
        self.e_munit = combo(self.units, '')
        rows3 = (('MIR No.', self.e_mno, 'Rev', self.e_rev), ('Date', self.e_mdate, 'MAR', self.e_mar), ('Description', self.e_mdesc, None, None),
                 ('Model No.', self.e_model, 'Origin', self.e_orig), ('Serial No.', self.e_serial, None, None), ('Quantity', self.e_mqty, 'Unit', self.e_munit))
        for i, (a, wa, b, wb) in enumerate(rows3, 1):
            l3.addWidget(QLabel(a), i, 0)
            if b: l3.addWidget(wa, i, 1); l3.addWidget(QLabel(b), i, 2); l3.addWidget(wb, i, 3)
            else: l3.addWidget(wa, i, 1, 1, 3)
        l3.addWidget(DQ.btn('↻ Fill from the delivery', lambda: self._mir_defaults()), 9, 0, 1, 2, Qt.AlignLeft)
        form.addWidget(g3)
        bar = QHBoxLayout()
        self.l_msg = QLabel(''); self.l_msg.setStyleSheet(f'color: {RED};'); bar.addWidget(self.l_msg, 1)
        bar.addWidget(DQ.btn('Cancel', self.reject)); bar.addWidget(DQ.btn('Make MIR' if existing else 'Save', self.ok, True))
        form.addLayout(bar); form.addStretch(1)
        self._file(); self._mir_defaults(first=True)
        self.resize(1300 if self.pics else 820, 760)

    def _viewer(self):
        v = QVBoxLayout()
        nav = QHBoxLayout()
        nav.addWidget(DQ.btn('◀', lambda: self.page(-1))); self.l_pg = QLabel(''); nav.addWidget(self.l_pg)
        nav.addWidget(DQ.btn('▶', lambda: self.page(1))); nav.addStretch(1); v.addLayout(nav)
        self.sc = QScrollArea(); self.sc.setMinimumWidth(560); self.sc.setWidgetResizable(False)
        self.img = QLabel(); self.sc.setWidget(self.img); v.addWidget(self.sc, 1)
        self.pi = 0; QTimer.singleShot(0, lambda: self.page(0))
        return v

    def page(self, k):
        self.pi = max(0, min(len(self.pics) - 1, self.pi + k))
        px = QPixmap(self.pics[self.pi])
        if px.isNull(): self.l_pg.setText('Picture not available'); return
        px = px.scaledToWidth(min(px.width(), 560), Qt.SmoothTransformation)
        self.img.setPixmap(px); self.img.resize(px.size())
        self.l_pg.setText(f'Page {self.pi + 1} / {len(self.pics)}')

    def _item_row(self, it, en):
        if len(self.rows) >= 12: return
        i = len(self.rows) + 1
        name = OLDM.best_item(it.get('name', ''), self.stock['items'], it.get('code', ''), self.stock['deliveries']) if it.get('name') else ''
        qty = M.num_text(it['qty']) if isinstance(it.get('qty'), (int, float)) else str(it.get('qty') or '')
        unit = it.get('unit', '')
        if not unit and name:
            hit = next((x for x in self.stock['items'] if x['name'] == name), None)
            if hit: unit = hit['unit']
        w = dict(name=combo(self.names, name, True, 330), code=QLineEdit(it.get('code', '')), qty=QLineEdit(qty),
                 unit=combo(self.units, unit), status=combo(M.IN_STATUS, it.get('status') or M.IN_STATUS[0], False),
                 note=QLineEdit(it.get('note', '')))
        w['qty'].setMaximumWidth(90); w['unit'].setMaximumWidth(110)
        w['name'].currentTextChanged.connect(lambda _t, w=w: (self._unit_from_stock(w), self._mark(w)))
        w['name'].setToolTip(f"On the delivery note:  {it.get('name', '')}" if it.get('name') else '')
        for j, k in enumerate(('name', 'code', 'qty', 'unit', 'status', 'note')):
            w[k].setEnabled(en); self.l2.addWidget(w[k], i, j)
        self.rows.append(w); self._mark(w)

    def _mark(self, w):
        """البند مش في قايمة المخزون ← أحمر (هيتضاف صنف جديد في المخزون لو اتحفظ كده)."""
        t = w['name'].currentText().strip()
        new = bool(t) and t not in self.names
        w['name'].setStyleSheet(f'QComboBox {{ border: 2px solid {RED}; }}' if new else '')
        w['name'].setToolTip(('NOT in the stock list – choose the stock item, or it will be added as a NEW item\n' if new else '')
                             + (w['name'].toolTip().split('\n')[-1] if w['name'].toolTip() else ''))

    def _unit_from_stock(self, w):
        hit = next((x for x in self.stock['items'] if x['name'] == w['name'].currentText()), None)
        if hit and hit['unit']: w['unit'].setCurrentText(hit['unit'])

    def _items(self):
        out = []
        for w in self.rows:
            name = w['name'].currentText().strip()
            if not name: continue
            out.append(dict(name=name, code=w['code'].text().strip(), qty=M._n(w['qty'].text().replace(',', '')),
                            unit=w['unit'].currentText().strip(), status=w['status'].currentText().strip() or M.IN_STATUS[0],
                            note=w['note'].text().strip()))
        return out

    def _file(self, *_):
        c, d = self.e_comp.currentText().strip(), self.e_dn.text().strip()
        self.file_base = M.dn_file_name(c, d, self.stock['deliveries']) if c and d else ''
        self.l_file.setText(f'File:  {self.file_base}.pdf' if self.file_base else 'File name: needs Company and DN No.')

    def _mir_defaults(self, first=False):
        its = self._items(); names = [x['name'] for x in its]
        best = M.mar_from_history(self.e_comp.currentText(), names, self.stock['deliveries'], self.mm['mir'], self.mar_list) \
            or M.match_mar(self.e_comp.currentText(), names, self.mar_list)[0]          # نفس الصنف اتعمله MIR قبل كده ← نفس الـ MAR
        if best and (first or not self.e_mar.currentText()): self.e_mar.setCurrentText(M.mar_label(best))
        stock_names = [n for n in names if n in self.names] or names       # الوصف بأسامي المخزن (زي MIR-001)
        desc = ' + '.join(stock_names)
        self.e_mdesc.setText(stock_names[0] if len(desc) > 80 and stock_names else desc)
        if its:
            self.e_model.setText(its[0]['code']); self.e_munit.setCurrentText(its[0]['unit'])
            self.e_mqty.setText(M.num_text(its[0]['qty']) if len(its) == 1 and its[0]['qty'] is not None else '')
        dn, inv = self.e_dn.text().strip(), self.e_inv.text().strip()
        inv = inv if inv.upper().startswith('INV') else (f'INV {inv}' if inv else '')
        self.e_serial.setText(' / '.join(x for x in (f'Delivery Note {dn}' if dn else '', inv) if x))

    def ok(self):
        msg = self.l_msg.setText; msg('')
        d, its = pyd(self.e_date.date()), self._items()
        if not self.existing:
            if not self.e_comp.currentText().strip(): return msg('Choose or type the company')
            if not self.e_dn.text().strip(): return msg('Type the delivery note number')
            if not its: return msg('Add at least one item')
            bad = [x['name'] for x in its if not x['qty'] or x['qty'] <= 0]
            if bad: return msg(f'Quantity missing: {bad[0][:40]}')
            new = [x['name'] for x in its if x['name'] not in self.names]
            if new and QMessageBox.question(self, 'Delivery Note', 'These are NOT in the stock list:\n\n  •  ' + '\n  •  '.join(n[:70] for n in new)
                                            + '\n\nAdd them to the stock as NEW items?\n(No = go back and choose the stock item for each red line)') != QMessageBox.Yes:
                return msg('Choose the stock item for the red line(s)')
        mir = None
        if self.existing or self.c_make.isChecked():
            no = self.e_mno.text().strip()
            if not no: return msg('MIR No. and date are needed')
            if not self.e_mdesc.text().strip(): return msg('MIR description is empty')
            if any(m['no'].strip().upper() == no.upper() for m in self.mm['mir']): return msg(f'{no} is already in the MIR LOG')
            lab = self.e_mar.currentText().strip()
            mar = next((m for m in self.mar_list if M.mar_label(m) == lab), None)
            mir = dict(no=no, rev=self.e_rev.text().strip() or '1', date=pyd(self.e_mdate.date()), desc=self.e_mdesc.text().strip(),
                       origin=self.e_orig.text().strip(), model=self.e_model.text().strip(), serial=self.e_serial.text().strip(),
                       qty=self.e_mqty.text().strip(), unit=self.e_munit.currentText().strip(), mat=(mar['ref'] if mar else lab))
        self.result_ = dict(date=d, dn=self.e_dn.text().strip(), inv=self.e_inv.text().strip(), company=self.e_comp.currentText().strip(),
                            type=self.e_type.currentText().strip(), items=its, mir=mir, file=self.file_base)
        self.accept()


def without_mir(st):
    """التوريدات اللي لسه ملهاش MIR (اللي وصل منها): [(company, dn, date, عدد البنود)] – الأحدث الأول."""
    covered = {(d['company'], d['dn']) for d in st['deliveries'] if d['mir']}
    out = {}
    for d in st['deliveries']:
        k = (d['company'], d['dn'])
        if k in covered or d['status'] == 'لم يصل': continue
        c, n, date = out.get(k, (d['company'], 0, d['date']))
        out[k] = (c, n + 1, date)
    return sorted(((c, dn, date, n) for (c, dn), (_, n, date) in out.items()), key=lambda t: (t[2] or dt.date.min), reverse=True)


def make_mir_for(parent, st, mm, company, dn):
    """الـ MIR لتوريد متسجل: شباك الـ MIR بداتا التوريد ← سطر MIR LOG + PDF (الإذن وراه) + رقمه قدام بنود التوريد.
    يرجّع مسار الـ PDF أو None."""
    same = [d for d in st['deliveries'] if d['dn'] == dn and d['company'] == company]
    if not same: QMessageBox.information(parent, 'MIR', f'{company} {dn} is not in the deliveries sheet.'); return None
    has = sorted({d['mir'] for d in same if d['mir']})
    if has: QMessageBox.information(parent, 'MIR', f"This delivery note already has {', '.join(has)}."); return None
    x = same[0]
    group = [d for d in same if d['status'] != 'لم يصل'] or same
    pf = dict(date=x['date'], dn=x['dn'], company=x['company'], type=x['type'],
              items=[dict(name=d['name'], code=d['note'].split(' – ')[0] if d['note'] else '', qty=d['qty'], unit=d['unit'], status=d['status']) for d in group])
    dn_pdf = M.find_dn_file(x['dl'] or M.dn_file_name(x['company'], x['dn'], st['deliveries']))
    pics, work = [], None
    if dn_pdf:                                           # صورة الإذن جنب الشباك
        try:
            import tempfile
            try: import pymupdf as fitz
            except ImportError: import fitz
            work = tempfile.mkdtemp(prefix='mir_dn_')
            doc = fitz.open(dn_pdf)
            for i, pg in enumerate(doc):
                if i >= 4: break
                f = os.path.join(work, f'p{i + 1}.png'); pg.get_pixmap(dpi=110).save(f); pics.append(f)
            doc.close()
        except Exception:
            pics = []
    try:
        w = DeliveryDlg(parent, st, mm, pf, pics, existing=True)
        if not w.exec() or not w.result_ or not w.result_['mir']: return None
    finally:
        if work: shutil.rmtree(work, ignore_errors=True)
    QGuiApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        path = M.mir_for_delivery([d['row'] for d in group], w.result_['mir'], dn_pdf)
    except Exception as e:
        QGuiApplication.restoreOverrideCursor()
        QMessageBox.critical(parent, 'MIR', f'The MIR was not made:\n{e}'); return None
    QGuiApplication.restoreOverrideCursor()
    if QMessageBox.question(parent, 'MIR', f"{w.result_['mir']['no']} is ready for signature:\n{path}\n\nOpen it?") == QMessageBox.Yes:
        try: os.startfile(path)
        except Exception: pass
    return path


def new_mir(parent):
    """«＋ New MIR» من Documents: اختار توريد لسه ملوش MIR ← نفس شباك الـ MIR."""
    from PySide6.QtWidgets import QInputDialog
    try:
        st, mm = M.load_stock(), M.load_marmir()
    except Exception as e:
        QMessageBox.critical(parent, 'New MIR', f'Could not read the registers:\n{e}'); return None
    todo = without_mir(st)
    if not todo:
        QMessageBox.information(parent, 'New MIR', 'Every delivery already has its MIR.\n\n'
                                'A new MIR starts from the delivery note: distribute it first (Distribute → Approve).')
        return None
    labels = [f"{c}   ·   {dn}   ·   {fd(date)}   ·   {n} item(s)" for c, dn, date, n in todo]
    pick, ok = QInputDialog.getItem(parent, 'New MIR', 'Delivery notes without MIR – choose one:', labels, 0, False)
    if not ok: return None
    c, dn, _, _ = todo[labels.index(pick)]
    return make_mir_for(parent, st, mm, c, dn)


def delivery_dialog(parent, prefill=None, pictures=None):
    """للـ DC (Qt): يفتح شباك إذن التوريد. يرجّع البيانات أو None."""
    try:
        stock, mm = M.load_stock(), M.load_marmir()
    except Exception as e:
        QMessageBox.critical(parent, 'Delivery Note', f'Could not read the registers:\n{e}'); return None
    w = DeliveryDlg(parent, stock, mm, prefill, pictures)
    return w.result_ if w.exec() else None


# ---------------------------------------------------------------- الصرف للموقع
class IssueDlg(QDialog):
    def __init__(self, parent, stock, routes):
        super().__init__(parent)
        self.setWindowTitle('Issue to Site'); self.setStyleSheet(CSS + f'QDialog {{ background: {HQ.BG}; }}')
        self.result_, self.stock = None, stock
        g = QGridLayout(self); g.setContentsMargins(18, 16, 18, 14); g.setVerticalSpacing(8)
        self.e_date = date_edit()
        self.e_name = combo([x['name'] for x in stock['items']], '', False, 420); self.e_name.setCurrentIndex(-1)
        self.e_name.currentIndexChanged.connect(self._bal)
        self.l_bal = QLabel(''); self.l_bal.setStyleSheet(f'color: {DQ.GREEN}; font-weight: bold;')
        self.e_qty = QLineEdit(); self.e_qty.setMaximumWidth(120)
        self.e_route = combo(routes, '', True, 420); self.e_who, self.e_note = QLineEdit(), QLineEdit()
        for i, (lab, w) in enumerate((('Date', self.e_date), ('Material', self.e_name), ('Balance in store', self.l_bal), ('Quantity', self.e_qty),
                                       ('Line / Location', self.e_route), ('Receiver', self.e_who), ('Notes', self.e_note))):
            l = QLabel(lab); l.setStyleSheet('font-weight: bold;'); g.addWidget(l, i, 0); g.addWidget(w, i, 1)
        self.l_msg = QLabel(''); self.l_msg.setStyleSheet(f'color: {RED};'); g.addWidget(self.l_msg, 8, 0, 1, 2)
        bar = QHBoxLayout(); bar.addStretch(1); bar.addWidget(DQ.btn('Cancel', self.reject)); bar.addWidget(DQ.btn('Save', self.ok, True))
        g.addLayout(bar, 9, 0, 1, 2)

    def _item(self): return next((x for x in self.stock['items'] if x['name'] == self.e_name.currentText()), None)

    def _bal(self, *_):
        it = self._item()
        self.l_bal.setText(f"{M.num_text(it['bal'] or 0)} {it['unit']}" if it else '')

    def ok(self):
        it, q = self._item(), M._n(self.e_qty.text().replace(',', ''))
        if not it: return self.l_msg.setText('Choose the material')
        if not q or q <= 0: return self.l_msg.setText('Type the quantity')
        if it['bal'] is not None and q > it['bal'] + 1e-9:
            if QMessageBox.question(self, 'Issue to Site', f"Only {M.num_text(it['bal'])} {it['unit']} in store.\nIssue {M.num_text(q)} anyway?") != QMessageBox.Yes:
                return
        self.result_ = dict(date=pyd(self.e_date.date()), name=it['name'], qty=q, route=self.e_route.currentText().strip(),
                            who=self.e_who.text().strip(), note=self.e_note.text().strip())
        self.accept()


# ---------------------------------------------------------------- الشاشة
# المخزن بس (طلب المستخدم 21-11-2026): الـ MAR والـ MIR اتنقلوا لتابات Documents في Document Center
TABS = [('stock', 'Stock', OLDM.STOCK_COLS, ['All', 'Negative balance', 'Still to deliver', 'In store']),
        ('del', 'Deliveries', OLDM.DEL_COLS, ['All', 'Without MIR', 'Not arrived']),
        ('iss', 'Issues to Site', OLDM.ISS_COLS, ['All'])]


class MatQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        self.setObjectName('matq'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.marks = M.load_marks()
        self.mm, self.st, self.fmir, self.fmar = {'mar': [], 'mir': []}, None, {}, {}
        self.views = {}
        self._build()
        QTimer.singleShot(100, self.refresh)

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 14, 22, 16); v.setSpacing(10)
        self.l_info = QLabel(''); self.l_info.setStyleSheet(f'color: {DQ.GREEN};')
        self.b_pend = DQ.btn('Write waiting rows', self.flush); self.b_pend.hide()
        v.addLayout(header('Store', 'Stock · Deliveries · Issues to site   –   MAR and MIR are in Document Center › Documents',
                           HQ.COL.get('mat', '#D2BC4A'), lambda: self.open_section('doc'),
                           [self.l_info, self.b_pend, DQ.btn('⟳  Refresh', self.refresh, True),
                            DQ.btn('Delivery notes', lambda: self.open_path(M.P(M.DELIVERY_REL))),
                            DQ.btn('Stock log', lambda: self.open_path(M.P(M.STOCK_REL)))]))
        self.l_sum = QLabel(''); self.l_sum.setWordWrap(True); self.l_sum.setFont(HQ.font(12.5, True))
        self.l_sum.setStyleSheet(f'color: {INK}; background: {TT["chip"]}; border-radius: 10px; padding: 8px 12px;')
        v.addWidget(self.l_sum)
        c = DQ.card(); cv = QVBoxLayout(c); cv.setContentsMargins(14, 12, 14, 12)
        self.tabs = QTabWidget(); self.tabs.setDocumentMode(True); cv.addWidget(self.tabs)
        acts = {'stock': [('Issue to Site…', self.issue, True)],
                'del': [('Open file', lambda: self.open_file('del'), False), ('Make MIR…', self.mir_from_delivery, True)],
                'iss': [('Issue to Site…', self.issue, True)]}
        for key, title, cols, filters in TABS:
            w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 8, 0, 0); wv.setSpacing(6)
            bar = QHBoxLayout(); l = QLabel('Show'); l.setStyleSheet(f'color: {MUTED};'); bar.addWidget(l)
            cb = QComboBox(); cb.addItems(filters); cb.setMinimumWidth(240); cb.currentIndexChanged.connect(lambda _i, k=key: self.fill(k))
            bar.addWidget(cb); bar.addStretch(1)
            for text, fn, pri in acts[key]: bar.addWidget(DQ.btn(text, fn, pri))
            wv.addLayout(bar)
            if key == 'stock':
                self.l_alert = QLabel(''); self.l_alert.setStyleSheet(f'color: {RED}; font-weight: bold;'); self.l_alert.setWordWrap(True)
                wv.addWidget(self.l_alert)
            t = table(cols)
            t.cellClicked.connect(lambda r, cc, k=key: self.on_click(k, r, cc))
            t.cellDoubleClicked.connect(lambda r, cc, k=key: self.open_file(k))
            wv.addWidget(t, 1)
            self.tabs.addTab(w, f'  {title}  ')
            self.views[key] = dict(tv=t, cb=cb, cols=cols, rows=[])
        h = QLabel('Double-click a delivery to open its delivery note.   A delivery without MIR is highlighted – «Make MIR…» makes it.')
        h.setStyleSheet(f'color: {MUTED};'); cv.addWidget(h)
        v.addWidget(c, 1)

    def open_path(self, p):
        try: os.startfile(p)
        except Exception as e: QMessageBox.critical(self, 'Materials', str(e))

    # ------------------------------------------------------------ data
    def refresh(self):
        errs = []
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            try: self.mm = M.load_marmir()
            except Exception as e: errs.append(f'02-LOG MAR+MIR: {e}'); traceback.print_exc()
            try: self.st = M.load_stock()
            except Exception as e: errs.append(f'04-LOG STOCK: {e}'); traceback.print_exc(); self.st = None
            self.fmir, self.fmar = M.folder_files(M.MIR_DIR_REL), M.folder_files(M.MAR_DIR_REL)
            self.marks = M.load_marks()
            for k in self.views: self.fill(k)
        finally:
            QGuiApplication.restoreOverrideCursor()
        self.summary()
        n = len(M.load_pending())
        self.b_pend.setText(f'Write waiting rows ({n})'); self.b_pend.setVisible(bool(n))
        self.l_info.setText(('  |  '.join(errs)) if errs else f'Refreshed {dt.datetime.now():%H:%M}')
        if errs: self.l_info.setStyleSheet(f'color: {RED};')

    def _without_mir(self):
        if not self.st: return set()
        covered = {(d['company'], d['dn']) for d in self.st['deliveries'] if d['mir']}
        return {(d['company'], d['dn']) for d in self.st['deliveries'] if d['status'] != 'لم يصل' and (d['company'], d['dn']) not in covered}

    def summary(self):
        t = ''
        alert = []
        if self.st:
            neg = [x for x in self.st['items'] if (x['bal'] or 0) < -1e-9]
            nomir = self._without_mir()
            t += f"Stock: {len(self.st['items'])} materials" + (f", {len(neg)} negative" if neg else '') + \
                 (f"     |     {len(nomir)} delivery note(s) without MIR" if nomir else '')
            if neg: alert.append('Negative balance: ' + ' · '.join(x['name'][:40] for x in neg[:4]))
            if nomir: alert.append('Delivered without MIR: ' + ' · '.join(f'{c} {d}' for c, d in sorted(nomir)[:4]))
        self.l_alert.setText('\n'.join(alert)); self.l_alert.setVisible(bool(alert))
        self.l_sum.setText(t)

    def _unf(self, kind, no):
        d = self.marks.get(M.mkey(kind, no))
        return f'✔ {d}' if d else '—'

    def fill(self, key):
        V = self.views[key]; t, flt = V['tv'], V['cb'].currentText()
        rows = []
        if key == 'mir':
            for m in sorted(self.mm['mir'], key=lambda m: -m['n']):
                stg = M.mir_stage(m, self.fmir); unf = self.marks.get(M.mkey('MIR', m['no']))
                if flt == 'Not on Unifier' and (unf or stg in ('Not prepared', 'Rejected')): continue
                if flt not in ('All', 'Not on Unifier') and stg != flt: continue
                bf = M.best_file(M.mir_files(m, self.fmir))
                vals = (m['no'], fd(m['date']), m['desc'], m['qty'], m['unit'], m['serial'], m['mat'], stg, self._unf('MIR', m['no']),
                        f'{bf[0]}\\{bf[1]}' if bf else '')
                rows.append((m, vals, {'Approved': 'appr', 'Signed': 'sign', 'Waiting signature': 'wait', 'Rejected': 'rej'}.get(stg, 'none'), (M.MIR_DIR_REL, bf)))
        elif key == 'mar':
            for m in sorted(self.mm['mar'], key=lambda m: -m['n']):
                stg = M.mar_stage(m, self.fmar); unf = self.marks.get(M.mkey('MAR', m['no']))
                if flt == 'Approved – not on Unifier' and (stg != 'Approved' or unf): continue
                if flt not in ('All', 'Approved – not on Unifier') and stg != flt: continue
                bf = M.best_file(M.mar_files(m, self.fmar))
                vals = (m['no'], m['ref'], m['desc'], m['supplier'], fd(m['sent']), fd(m['reply']), m['status'], stg,
                        self._unf('MAR', m['no']), f'{bf[0]}\\{bf[1]}' if bf else '')
                rows.append((m, vals, {'Approved': 'appr', 'Submitted': 'wait', 'Resubmit': 'resub', 'Rejected': 'rej'}.get(stg, 'none'), (M.MAR_DIR_REL, bf)))
        elif key == 'stock' and self.st:
            for x in self.st['items']:
                bal = x['bal'] or 0
                if flt == 'Negative balance' and bal >= -1e-9: continue
                if flt == 'Still to deliver' and not (x['left'] or 0) > 0: continue
                if flt == 'In store' and bal <= 1e-9: continue
                vals = (x['name'], x['code'], x['unit'], x['type'], *(M.num_text(x[k]) for k in ('need', 'got', 'wait', 'out', 'bal', 'left')), x['status'])
                rows.append((x, vals, 'neg' if bal < -1e-9 else ('appr' if (x['need'] and (x['got'] or 0) >= x['need']) else 'none'), None))
        elif key == 'del' and self.st:
            nomir = self._without_mir()
            for x in reversed(self.st['deliveries']):
                if flt == 'Without MIR' and (x['company'], x['dn']) not in nomir: continue
                if flt == 'Not arrived' and x['status'] != 'لم يصل': continue
                vals = (x['tr'], fd(x['date']), x['dn'], x['company'], x['type'], x['name'], M.num_text(x['qty']), x['unit'], x['status'], x['mir'], x['dl'])
                rows.append((x, vals, 'nomir' if (x['company'], x['dn']) in nomir else 'none', None))
        elif key == 'iss' and self.st:
            for x in reversed(self.st['issues']):
                rows.append((x, (fd(x['date']), x['name'], M.num_text(x['qty']), x['unit'], x['route'], x['who'], x['note']), 'none', None))
        t.setUpdatesEnabled(False); t.setRowCount(len(rows))
        for i, (_, vals, tag, _) in enumerate(rows):
            for j, (cid, _, _) in enumerate(V['cols']):
                val = vals[j] if j < len(vals) else ''
                it = QTableWidgetItem('' if val is None else (f'{val:g}' if isinstance(val, float) else str(val)))
                it.setTextAlignment(Qt.AlignVCenter | (Qt.AlignLeft if cid in LEFT else Qt.AlignCenter))
                if TAG_BG.get(tag): it.setBackground(QColor(TAG_BG[tag]))
                if tag in ('neg', 'rej'): it.setForeground(QColor(RED))
                if cid == 'unf': it.setForeground(QColor(DQ.GREEN if str(val).startswith('✔') else MUTED))
                t.setItem(i, j, it)
        t.setUpdatesEnabled(True)
        V['rows'] = rows

    def _sel(self, key):
        V = self.views[key]; r = V['tv'].currentRow()
        return V['rows'][r] if 0 <= r < len(V['rows']) else None

    # ------------------------------------------------------------ actions
    def on_click(self, key, row, col):
        if key not in ('mir', 'mar'): return
        V = self.views[key]
        if V['cols'][col][0] != 'unf' or row >= len(V['rows']): return
        m = V['rows'][row][0]
        k = M.mkey(key.upper(), m['no'])
        if self.marks.get(k): self.marks.pop(k)
        else: self.marks[k] = dt.date.today().strftime('%d-%m-%Y')
        try: M.save_marks(self.marks)
        except Exception as e:
            QMessageBox.critical(self, 'Materials', f'Could not save the mark:\n{e}'); return
        sv = V['tv'].verticalScrollBar().value(); self.fill(key); V['tv'].verticalScrollBar().setValue(sv); self.summary()
        self.l_info.setText(f"{m['no']}: " + (f"uploaded to Unifier ✔ {self.marks[k]}" if k in self.marks else 'Unifier mark removed'))

    def open_file(self, key):
        r = self._sel(key)
        if not r: return QMessageBox.information(self, 'Materials', 'Select a row first.')
        if key == 'del':
            p = M.find_dn_file(r[0]['dl'] or M.dn_file_name(r[0]['company'], r[0]['dn'], self.st['deliveries']))
            return self.open_path(p) if p else QMessageBox.information(self, 'Materials', f"No file named «{r[0]['dl']}» in {M.DELIVERY_REL}")
        if not r[3]: return
        rel, bf = r[3]
        if not bf: return QMessageBox.information(self, 'Materials', f"No file found for {r[0]['no']} in {rel}")
        self.open_path(os.path.join(M.P(rel), bf[0], bf[1]))

    def issue(self):
        if not self.st: return QMessageBox.critical(self, 'Materials', 'The stock log could not be read – press Refresh.')
        QGuiApplication.setOverrideCursor(Qt.WaitCursor)
        try: routes = OLDM.route_choices(self.st)
        finally: QGuiApplication.restoreOverrideCursor()
        w = IssueDlg(self, self.st, routes)
        if not w.exec() or not w.result_: return
        try:
            M.write_issue(w.result_)
            self.l_info.setText(f"Issued ✔ {M.num_text(w.result_['qty'])} – {w.result_['name'][:40]}")
        except Exception as e:
            if QMessageBox.question(self, 'Issue to Site', f'Could not write to the stock log:\n{e}\n\nKeep it and write it later?') == QMessageBox.Yes:
                M.add_pending('issue', w.result_); self.l_info.setText('Issue saved – waiting to be written')
        self.refresh()

    def mir_from_delivery(self):
        if not self.st: return QMessageBox.critical(self, 'Store', 'The stock log could not be read – press Refresh.')
        r = self._sel('del') if self.tabs.currentIndex() == 1 else None
        if not r:
            self.tabs.setCurrentIndex(1)
            return QMessageBox.information(self, 'Store', 'Select the delivery (any of its lines) in Deliveries, then press «Make MIR…».')
        path = make_mir_for(self, self.st, self.mm, r[0]['company'], r[0]['dn'])
        if path: self.l_info.setText(f'MIR ready ✔ {os.path.basename(path)}')
        self.refresh()

    def flush(self):
        n, err = M.flush_pending()
        if err: QMessageBox.warning(self, 'Materials', f'{n} written. Still waiting:\n{err}')
        else: self.l_info.setText(f'{n} waiting row(s) written ✔')
        self.refresh()
