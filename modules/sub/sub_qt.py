# -*- coding: utf-8 -*-
"""قسم SUB (مقاولين الباطن) – نسخة Qt (طلب المستخدم 17-11).

كل المقاولين اللي في 13_PARTIES\\SUB في شاشة واحدة: على الشمال القايمة، وعلى اليمين ملف المقاول كامل –
بيانات العقد والاعتمادات، والمستخلصات والمذكرات، والخطابات (05_LETTERS + خطابات DC في 02_Correspondence)،
وملف الاعتماد وعروض الأسعار. دبل كليك على أي ملف بيفتحه.
عمل المستخلص والمذكرات وتعديل البيانات والبنود لسه في قسم SUB القديم (بيفتح في شباك لوحده).
البيانات: 00_Registers\\10.1-SUBCONTRACTORS.xlsx (sub_core.Register) و00_Registers\\10-LOG SUB.xlsx (العقود ومذكرات الخصم).
"""
import os, re, sys, shutil, traceback, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
MODS = os.path.dirname(HERE)
PLAT = os.path.dirname(MODS)
for _p in (HERE, os.path.join(MODS, 'dc'), PLAT):
    if _p not in sys.path: sys.path.insert(0, _p)

import sub_core as S  # noqa: E402
import parties as PT  # noqa: E402

CORR_REL = '02_Correspondence'
SKIP = ('readme', 'desktop.ini', 'thumbs.db', '~$')
TABS = [('contract', 'Contract'), ('items', 'Items'), ('areas', 'Stations'), ('certs', 'Certificates'), ('deds', 'Deductions'), ('memos', 'Memos'),
        ('letters', 'Letters'), ('approval', 'Approval'), ('quotes', 'Quotations'), ('sign', 'Signatures')]


# ---------------------------------------------------------------- الداتا (من غير Qt – بتتجرب لوحدها)
def _files(folder, base=None):
    """كل الملفات جوه الفولدر (وجوه الفولدرات اللي جواه) ← [dict(path, name, where, date, size)] الأحدث فوق"""
    out = []
    if not folder or not os.path.isdir(folder): return out
    base = base or os.path.dirname(folder)
    for d, _, names in os.walk(folder):
        for n in names:
            if n.lower().startswith(SKIP) or n.lower().endswith('.claude.json'): continue
            p = os.path.join(d, n)
            try: st = os.stat(p)
            except OSError: continue
            out.append(dict(path=p, name=n, where=os.path.relpath(d, base), date=dt.datetime.fromtimestamp(st.st_mtime), size=st.st_size))
    out.sort(key=lambda f: f['date'], reverse=True)
    return out


def _keys(e):
    """أسامي المقاول للبحث في أسامي الملفات (الإنجليزي والعربي) – 4 حروف على الأقل"""
    ks = {PT.key(e['en'])} | {PT.key(a) for a in e.get('aliases') or []}
    if e.get('sub'): ks |= {PT.key(e['sub'].get('Name')), PT.key(e['sub'].get('Folder'))}
    ks |= {PT.key(re.sub(r'^(شركه|شركة|مؤسسه|مؤسسة)\s+', '', a)) for a in e.get('aliases') or []}
    return {k for k in ks if len(k) >= 4}


def corr_files(project):
    """كل ملفات 02_Correspondence مرة واحدة (بنفلتر منها لكل مقاول)"""
    root = os.path.join(project, CORR_REL)
    out = _files(root, project)
    for f in out: f['k'] = PT.key(os.path.splitext(f['name'])[0])
    return out


def letters(e, corr):
    """خطابات المقاول: فولدره 05_LETTERS + خطابات DC اللي اسم المقاول في اسم ملفها"""
    out = _files(os.path.join(e['path'], PT.LETTERS), e['path']) if e.get('path') else []
    ks = _keys(e)
    seen = {os.path.normcase(f['path']) for f in out}
    for f in corr:
        if os.path.normcase(f['path']) in seen: continue
        if any(k in f['k'] for k in ks): out.append(f)
    out.sort(key=lambda f: f['date'], reverse=True)
    return out


def all_subs(project=None):
    """المقاولين: فولدرات 13_PARTIES\\SUB + أي مقاول في سجل SUB ملوش فولدر ← [dict(no, en, ar, aliases, path, sub)]"""
    project = project or S.proj()
    try: reg = S.Register()
    except Exception: reg = None
    rsubs = list(reg.subs()) if reg else []
    out, used = [], set()
    for p in PT.all_parties(project):
        if p['cat'] != 'SUB': continue
        sub = next((s for s in rsubs if PT.key(s.get('Folder')) == PT.key(p['name'])
                    or PT.key(s.get('Name')) in {PT.key(a) for a in p['aliases']}), None)
        if sub: used.add(sub['ID'])
        ar = (sub or {}).get('Name') or next(iter(p['aliases']), '')
        if not re.search(r'[؀-ۿ]', str(ar or '')): ar = next((a for a in p['aliases'] if re.search(r'[؀-ۿ]', a)), '')
        out.append(dict(no=p['no'], en=p['name'], ar=ar, aliases=p['aliases'], path=p['path'], sub=sub))
    for s in rsubs:
        if s['ID'] in used: continue
        try: path = S.sub_dir(s)
        except Exception: path = ''
        out.append(dict(no=0, en=s.get('Folder') or s.get('Name'), ar=s.get('Name') or '', aliases=[], path=path if os.path.isdir(path or '') else '', sub=s))
    out.sort(key=lambda e: (-1 if e['sub'] else 0, e['no']))
    return out, reg


def details(e, reg, corr):
    """كل اللي يخص المقاول ده: العقود والاعتمادات، البنود، المستخلصات، المذكرات، والملفات في كل فولدر"""
    who = dict(Folder=e['en'], Name=e['ar'] or e['en'])
    if e.get('sub'): who = e['sub']
    d = dict(err=[])
    try: d['contracts'] = S.contracts(who)
    except Exception as x: d['contracts'] = []; d['err'].append(f'10-LOG SUB (contracts): {x}')
    try: d['memos'] = S.memos(who)
    except Exception as x: d['memos'] = []; d['err'].append(f'10-LOG SUB (memos): {x}')
    sid = (e.get('sub') or {}).get('ID')
    d['items'] = reg.items(sid) if (reg and sid) else []
    d['certs'] = reg.certs(sid) if (reg and sid) else []
    P = e.get('path') or ''
    d['f_contract'] = _files(os.path.join(P, PT.CONTRACT), P) if P else []
    d['f_certs'] = _files(os.path.join(P, PT.CERTS), P) if P else []
    d['f_memos'] = (_files(os.path.join(P, PT.DMEMOS), P) + _files(os.path.join(P, PT.AMEMOS), P)
                    + _files(os.path.join(P, PT.FINAL), P)) if P else []
    d['f_approval'] = _files(os.path.join(P, PT.APPROVAL), P) if P else []
    d['f_quotes'] = _files(os.path.join(P, PT.QUOTES), P) if P else []
    d['f_letters'] = letters(e, corr)
    ct = d['contracts'][-1] if d['contracts'] else None
    d['ct'] = ct
    last = d['certs'][-1] if d['certs'] else None
    d['certified'] = S.num(last.get('Works To Date')) if last else 0.0
    d['retention'] = sum(S.num(c.get('Retention')) or 0.0 for c in d['certs'])
    d['ded_total'] = sum(m['amount'] or 0.0 for m in d['memos'] if m['status'] != 'Cancelled')
    return d


def contract_chip(d_or_ct):
    """شارة العقد في القايمة: (الكلام, نوعها)"""
    ct = d_or_ct
    if not ct: return 'No contract', 'wait'
    if not ct['missing']: return 'Contract complete ✔', 'sure'
    return 'Contract: missing ' + ', '.join(ct['missing']), 'ask'


def fdate(v):
    if isinstance(v, dt.datetime): return v.strftime('%d-%m-%Y')
    if isinstance(v, dt.date): return v.strftime('%d-%m-%Y')
    return str(v or '')


def money(v):
    try: return f'{float(v):,.2f}'
    except (TypeError, ValueError): return ''


# ---------------------------------------------------------------- الحصر: كل كميات QS جوه الاستيشنات
TAKE = [('EXC1', 'Excavation 0–1.5 m', 'm3'), ('EXC2', 'Excavation 1.5–2.5 m', 'm3'), ('EXC3', 'Excavation 2.5–3.5 m', 'm3'),
        ('EXC4', 'Excavation deeper than 3.5 m', 'm3'), ('TRL', 'Trench length excavated', 'm'), ('BED', 'Bedding & initial backfill', 'm3'),
        ('PIPL', 'Pipe laid – full length', 'm'), ('PIP', 'Pipe – paid length after % (like the owner)', 'm'),
        ('BF', 'Backfill by layers', 'm3'), ('RB', 'Road base RB-1 + RB-2', 'm3'), ('ASP1', 'Asphalt – normal', 'm3'),
        ('ASP2', 'Asphalt – paver & milling', 'm3')]
BANDS = [(None, 1.5, '0–1.5'), (1.5, 2.5, '1.5–2.5'), (2.5, 3.5, '2.5–3.5'), (3.5, None, '> 3.5')]
DEPTHED = {'PIP', 'PIPL', 'TRL'}


def _dia(v):
    try: return int(float(v))
    except (TypeError, ValueError): return None


def takeoff(items, areas, sub_items=()):
    """كل كميات حصر QS جوه الاستيشنات: [dict(code, work, dia, depth, unit, qty, item)] – بالقطر، والمواسير والخندق بالعمق كمان.
    item = بند العقد اللي الكمية دي بتروحله (نفس طريقة الحساب والقطر والعمق) لو فيه."""
    out = []
    if not areas: return out
    for code, work, unit in TAKE:
        try: tot, rows = S.item_quantity(items, {'Measure': code}, areas)
        except Exception: continue
        if not rows or abs(tot or 0) < 1e-9: continue
        dias = sorted({_dia(r.get('dia')) for r in rows if _dia(r.get('dia'))})
        for dia in dias or [None]:
            bands = BANDS if code in DEPTHED else [(None, None, '')]
            for a, b, lab in bands:
                q, _ = S.item_quantity(items, {'Measure': code, 'Dia': dia, 'Depth From': a, 'Depth To': b}, areas)
                if abs(q or 0) < 1e-9: continue
                it = next((i for i in sub_items if str(i.get('Measure') or '') == code
                           and (_dia(i.get('Dia')) in (None, dia))
                           and (code not in DEPTHED or ((S.num(i.get('Depth From')) or 0) == (a or 0)
                                                        and S.num(i.get('Depth To')) == b))), None)
                out.append(dict(code=code, work=work, dia=dia or '', depth=lab, unit=unit, qty=round(q, 3),
                                item=str(it['Item No.']) if it else ''))
    return out


# ---------------------------------------------------------------- الشاشة
from PySide6.QtCore import Qt, QTimer, QDate  # noqa: E402
from PySide6.QtGui import QCursor, QGuiApplication, QColor, QFont  # noqa: E402
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QHBoxLayout, QVBoxLayout, QGridLayout, QTableWidget,  # noqa: E402
                               QTableWidgetItem, QHeaderView, QAbstractItemView, QListWidget, QListWidgetItem,
                               QTabWidget, QMessageBox, QSplitter, QDialog, QComboBox, QLineEdit, QDateEdit, QCheckBox,
                               QInputDialog, QFileDialog)
import home_qt as HQ  # noqa: E402
import dc_qt as DQ  # noqa: E402  (نفس الشكل الغامق والزراير والشارات بتاعة Distribute)

INK, MUTED, NAVY, SOFT = DQ.INK, DQ.MUTED, DQ.NAVY, DQ.SOFT
TT = HQ.T
CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#subq') + f"""
QListWidget {{ background: {TT['card2']}; border: none; outline: 0; }}
QListWidget::item {{ border-radius: 10px; margin: 3px 4px; }}
QListWidget::item:selected {{ background: #2B5F8E; }}
QListWidget::item:hover:!selected {{ background: {TT['pill']}; }}
QTabWidget::pane {{ border: none; border-top: 2px solid {NAVY}; }}
QTabBar::tab {{ background: {TT['chip']}; color: {MUTED}; padding: 7px 14px; margin-right: 2px; border-top-left-radius: 8px;
               border-top-right-radius: 8px; font: 600 10pt "Segoe UI"; }}
QTabBar::tab:selected {{ background: {NAVY}; color: white; }}
QFrame#kpi {{ background: {TT['chip']}; border-radius: 10px; }}
"""


def _table(cols):
    t = QTableWidget(0, len(cols))
    t.setHorizontalHeaderLabels([c[1] for c in cols])
    t.verticalHeader().setVisible(False)
    t.setSelectionBehavior(QAbstractItemView.SelectRows); t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers); t.setShowGrid(False); t.setWordWrap(False)
    t.verticalHeader().setDefaultSectionSize(30)
    h = t.horizontalHeader(); h.setHighlightSections(False)
    for i, c in enumerate(cols):
        t.setColumnWidth(i, c[2])
    h.setStretchLastSection(True)
    return t


def pdate(s):
    """dd-mm-yyyy (أو تاريخ جاهز) ← date"""
    if isinstance(s, (dt.date, dt.datetime)): return s if isinstance(s, dt.date) else s.date()
    s = str(s or '').strip().replace('/', '-')
    for fm in ('%d-%m-%Y', '%Y-%m-%d', '%d-%m-%y'):
        try: return dt.datetime.strptime(s, fm).date()
        except ValueError: pass
    return None


def form(parent, title, fields, values=None, width=380):
    """فورم بخانات (label, key, kind, options) – kind: entry / combo / rocombo / date. يرجّع dict أو None"""
    values = values or {}
    d = QDialog(parent); d.setWindowTitle(title); d.setStyleSheet(CSS + f'QDialog {{ background: {HQ.BG}; }}')
    g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 14); g.setVerticalSpacing(7)
    w = {}
    for i, (label, key, kind, opts) in enumerate(fields):
        l = QLabel(label); l.setStyleSheet('font-weight: bold;'); g.addWidget(l, i, 0)
        v = values.get(key)
        txt = '' if v is None else (fdate(v) if kind == 'date' and isinstance(v, (dt.date, dt.datetime)) else
                                     (f'{v:g}' if isinstance(v, float) else str(v)))
        if kind in ('combo', 'rocombo'):
            e = QComboBox(); e.setEditable(kind == 'combo'); e.addItems([str(o) for o in (opts or ())])
            if kind == 'combo': e.setCurrentText(txt)
            else: e.setCurrentIndex(max(e.findText(txt), 0))
        else:
            e = QLineEdit(txt)
            if kind == 'date': e.setPlaceholderText('dd-mm-yyyy')
        e.setMinimumWidth(width); g.addWidget(e, i, 1); w[key] = e
    bar = QHBoxLayout(); bar.addStretch(1); bar.addWidget(DQ.btn('Cancel', d.reject)); bar.addWidget(DQ.btn('Save', d.accept, True))
    g.addLayout(bar, len(fields), 0, 1, 2)
    if not d.exec(): return None
    return {k: (e.currentText() if isinstance(e, QComboBox) else e.text()).strip() for k, e in w.items()}


def _front_filter():
    """أي شباك أو رسالة بتفتح (QDialog / QMessageBox / اختيار ملف) تيجي قدام على طول – عشان ماتستخباش ورا شاشة تانية
    والبرنامج يبان واقف. بيتركّب مرة واحدة على البرنامج كله."""
    from PySide6.QtCore import QObject, QEvent
    from PySide6.QtWidgets import QApplication, QDialog

    class _Front(QObject):
        def eventFilter(self, obj, ev):
            if ev.type() == QEvent.Show and isinstance(obj, QDialog) and obj.isWindow():
                def up(o=obj):
                    try: o.raise_(); o.activateWindow()
                    except RuntimeError: pass
                QTimer.singleShot(0, up)
            return False
    app = QApplication.instance()
    if app is not None and not app.property('sub_front_filter'):
        app._sub_front = _Front(app)
        app.installEventFilter(app._sub_front)
        app.setProperty('sub_front_filter', True)


class SubQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        try: _front_filter()
        except Exception: pass
        self.setObjectName('subq'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.subs, self.reg, self.corr, self.cur, self.d = [], None, [], None, {}
        self._build()
        QTimer.singleShot(100, self.reload)

    # ------------------------------------------------------------ UI
    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 14, 22, 16); v.setSpacing(12)
        t = QHBoxLayout(); t.setSpacing(14)
        b = DQ.btn('←  Back', lambda: self.open_section('doc')); b.setObjectName('back'); t.addWidget(b)
        tv = QVBoxLayout(); tv.setSpacing(0)
        c = QLabel('Technical Office Platform  ›  Subcontractors'); c.setFont(HQ.font(12.5, True)); c.setStyleSheet(f'color: {MUTED};')
        tv.addWidget(c)
        tt = QLabel(f'<span style="color:{HQ.COL.get("sub", "#E58A55")}; font-size:20pt; font-weight:bold;">Subcontractors</span>'
                    f'<span style="color:{MUTED}; font-size:11pt;">   everything about each subcontractor in one place</span>')
        tv.addWidget(tt); t.addLayout(tv, 1)
        t.addWidget(DQ.btn('⟳  Refresh', self.reload))
        t.addWidget(DQ.btn('All deductions', self.all_deds))
        t.addWidget(DQ.btn('Edit data', self.edit_sub))
        v.addLayout(t)

        sp = QSplitter(Qt.Horizontal); sp.setHandleWidth(10); sp.setChildrenCollapsible(False)
        # الشمال: القايمة
        lc = DQ.card(); lv = QVBoxLayout(lc); lv.setContentsMargins(10, 12, 10, 10); lv.setSpacing(6)
        self.l_count = QLabel('Subcontractors'); self.l_count.setFont(HQ.font(12, True)); self.l_count.setStyleSheet(f'color: {MUTED};')
        lv.addWidget(self.l_count)
        self.lst = QListWidget(); self.lst.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.lst.currentRowChanged.connect(self.pick)
        lv.addWidget(self.lst, 1)
        nb = DQ.btn('+  New subcontractor', self.new_sub); lv.addWidget(nb)
        sp.addWidget(lc)
        # اليمين: ملف المقاول
        rc = DQ.card(); rv = QVBoxLayout(rc); rv.setContentsMargins(18, 14, 18, 14); rv.setSpacing(8)
        self.l_title = QLabel(''); self.l_title.setFont(HQ.font(20, True)); self.l_title.setStyleSheet(f'color: {INK};')
        self.l_title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        rv.addWidget(self.l_title)
        self.l_info = QLabel(''); self.l_info.setWordWrap(True); self.l_info.setFont(HQ.font(12.5))
        self.l_info.setStyleSheet(f'color: {MUTED};'); self.l_info.setTextInteractionFlags(Qt.TextSelectableByMouse)
        rv.addWidget(self.l_info)
        kr = QHBoxLayout(); kr.setSpacing(10); self.kpi = {}
        for k, cap in (('appr', 'Contract approvals'), ('cert', 'Certified to date (SAR)'), ('ded', 'Deduction memos'), ('ret', 'Retention held (SAR)')):
            f = QFrame(); f.setObjectName('kpi'); fv = QVBoxLayout(f); fv.setContentsMargins(12, 8, 12, 8); fv.setSpacing(2)
            a = QLabel(cap); a.setFont(HQ.font(11)); a.setStyleSheet(f'color: {MUTED};')
            b2 = QLabel('—'); b2.setFont(HQ.font(15, True)); b2.setStyleSheet(f'color: {INK};'); b2.setWordWrap(True)
            fv.addWidget(a); fv.addWidget(b2); kr.addWidget(f, 1); self.kpi[k] = b2
        rv.addLayout(kr)
        self.tabs = QTabWidget(); self.tabs.setDocumentMode(True)
        FILE_COLS = [('name', 'FILE', 520), ('where', 'FOLDER', 260), ('date', 'DATE', 110), ('size', 'SIZE', 80)]
        self.t = {}
        for k, cap in TABS:
            if k == 'items':
                tb = _table([('no', 'NO.', 60), ('desc', 'DESCRIPTION', 520), ('unit', 'UNIT', 60), ('rate', 'RATE', 90),
                             ('m', 'QUANTITY FROM', 230), ('dia', 'DIA', 60), ('dep', 'DEPTH', 90)])
            elif k == 'areas':
                tb = _table([('item', 'ITEM', 70), ('desc', 'DESCRIPTION', 460), ('line', 'LINE', 120), ('from', 'FROM', 100),
                             ('to', 'TO', 100), ('len', 'LENGTH (m)', 100), ('notes', 'NOTES', 200)])
            elif k == 'sign':
                tb = _table([('role', 'ROLE', 260), ('name', 'NAME (printed on the certificate)', 400)])
            elif k == 'deds':
                import ded_qt as DD
                tb = _table(DD.COLS)
            elif k == 'certs':
                tb = _table([('no', 'NO.', 70), ('to', 'PERIOD TO', 110), ('wtd', 'WORKS TO DATE', 140), ('cur', 'THIS CERT.', 130),
                             ('net', 'NET PAYABLE', 130), ('st', 'STATUS', 100), ('name', 'FILE', 400)])
            elif k == 'memos':
                tb = _table([('no', 'NO.', 80), ('date', 'DATE', 100), ('subj', 'SUBJECT', 380), ('amt', 'AMOUNT', 110),
                             ('cert', 'DEDUCTED IN', 100), ('st', 'STATUS', 100), ('name', 'FILE', 400)])
            else:
                tb = _table(FILE_COLS)
            if k == 'areas': tb.cellDoubleClicked.connect(lambda r, c: self.edit_area())
            elif k == 'items': tb.cellDoubleClicked.connect(lambda r, c: self.edit_item())
            elif k == 'sign': tb.cellDoubleClicked.connect(lambda r, c: self.edit_sign())
            elif k == 'certs': tb.cellDoubleClicked.connect(lambda r, c: self.view_cert(r))
            elif k == 'deds': tb.cellDoubleClicked.connect(lambda r, c: self.open_ded(r))
            else: tb.cellDoubleClicked.connect(lambda r, c, tb=tb: self.open_row(tb, r))
            self.t[k] = tb
            bar = {'areas': [('Edit', self.edit_area, False), ('Delete', self.del_area, False)],
                   'certs': [('+ New certificate', self.new_cert, True), ('Open', lambda: self.view_cert(self.t['certs'].currentRow()), False)],
                   'sign': [('Change name', self.edit_sign, False)],
                   'items': [('+ Add item', self.add_item, True), ('Edit', self.edit_item, False), ('Delete', self.del_item, False)],
                   'memos': [('+ Add memo', self.add_memo, True), ('Open', lambda: self.open_row(self.t['memos'], self.t['memos'].currentRow()), False),
                             ('Cancel memo', self.cancel_memo, False)],
                   'deds': [('+ Deduction', self.add_ded, True), ('Open', lambda: self.open_ded(self.t['deds'].currentRow()), False),
                            ('Memo PDF', self.memo_pdf, False)]}.get(k)
            if bar:
                w = QWidget(); wv = QVBoxLayout(w); wv.setContentsMargins(0, 6, 0, 0); wv.setSpacing(6)
                hb = QHBoxLayout(); hb.setSpacing(8)
                for txt, fn, pri in bar: hb.addWidget(DQ.btn(txt, fn, pri))
                tip = {'areas': 'The stations are added in each certificate (+ New certificate) – NOTES = the certificate they came with.',
                       'certs': 'Double-click a certificate to see it. A new one starts after the last one is Approved.',
                       'sign': 'Double-click a row to change the name.',
                       'items': 'Double-click an item to change it. «Quantity from» = how the program takes it from QS.',
                       'memos': 'Memos made from the deductions come here by themselves – «+ Add memo» files a memo you already have.',
                       'deds': 'Violations, company equipment, materials … – the open ones show in his next certificate.'}[k]
                lt = QLabel(tip); lt.setStyleSheet(f'color: {MUTED};'); hb.addWidget(lt, 1)
                wv.addLayout(hb); wv.addWidget(tb, 1)
                self.tabs.addTab(w, cap)
            else:
                self.tabs.addTab(tb, cap)
        rv.addWidget(self.tabs, 1)
        br = QHBoxLayout(); br.setSpacing(8)
        br.addWidget(DQ.btn('Show in folder', self.show_sel))
        br.addWidget(DQ.btn('Open contract', self.open_contract))
        self.l_hint = QLabel('Double-click a file to open it.'); self.l_hint.setStyleSheet(f'color: {MUTED};')
        br.addWidget(self.l_hint, 1)
        br.addWidget(DQ.btn('+ New certificate', self.new_cert, True))
        rv.addLayout(br)
        sp.addWidget(rc)
        sp.setSizes([320, 1100])
        v.addWidget(sp, 1)

    # ------------------------------------------------------------ الداتا
    def reload(self):
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        try:
            keep = self.cur['en'] if self.cur else None
            self.subs, self.reg = all_subs()
            try: self.corr = corr_files(S.proj())
            except Exception: self.corr = []
            try: cts = S.contracts()
            except Exception: cts = []
            self.lst.blockSignals(True); self.lst.clear()
            row = 0
            for i, e in enumerate(self.subs):
                ks = {PT.key(e['en'])} | ({PT.key(e['sub'].get('Folder'))} if e.get('sub') else set())
                ct = next((c for c in reversed(cts) if PT.key(c['sub']) in ks), None)
                chip, tag = contract_chip(ct)
                n = len(self.reg.certs(e['sub']['ID'])) if (self.reg and e.get('sub')) else 0
                it = QListWidgetItem(); it.setData(Qt.UserRole, i)
                w = self._row_widget(e, chip, tag, n)
                it.setSizeHint(w.sizeHint()); self.lst.addItem(it); self.lst.setItemWidget(it, w)
                if keep and e['en'] == keep: row = i
            if self.subs: self.lst.setCurrentRow(row)
            self.lst.blockSignals(False)
            self.l_count.setText(f'Subcontractors  ·  {len(self.subs)}')
            if self.subs: self.pick(row)
            else: self.l_title.setText('No subcontractors in 13_PARTIES\\SUB yet')
        except Exception as x:
            if self.log: self.log(f'SUB reload: {x}')
            QMessageBox.warning(self, 'Subcontractors', f'Could not read the subcontractors:\n{x}')
        finally:
            QGuiApplication.restoreOverrideCursor()

    def _row_widget(self, e, chip, tag, n):
        w = QWidget(); w.setAttribute(Qt.WA_TranslucentBackground, True)
        g = QVBoxLayout(w); g.setContentsMargins(10, 7, 10, 8); g.setSpacing(2)
        a = QLabel(f"{e['no']:02d} · {e['en'].replace('_', ' ')}" if e['no'] else e['en'].replace('_', ' '))
        a.setFont(HQ.font(13, True)); a.setStyleSheet(f'color: {INK}; background: transparent;'); g.addWidget(a)
        if e['ar']:
            b = QLabel(e['ar']); b.setFont(HQ.font(12)); b.setStyleSheet(f'color: {MUTED}; background: transparent;'); g.addWidget(b)
        bg, fg = DQ.PILL.get(tag, DQ.PILL['wait'])
        cs = (f'<span style="background-color:{bg}; color:{fg};">&nbsp;{chip}&nbsp;</span>'
              + (f'&nbsp;&nbsp;<span style="color:{MUTED};">{n} IPC</span>' if n else '')
              + ('' if e.get('sub') else f'&nbsp;&nbsp;<span style="color:{MUTED};">not in SUB yet</span>'))
        c = QLabel(cs); c.setFont(HQ.font(11, True)); c.setStyleSheet('background: transparent;'); g.addWidget(c)
        return w

    def pick(self, row):
        if row is None or row < 0 or row >= len(self.subs): return
        e = self.subs[row]; self.cur = e
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        try:
            d = self.d = details(e, self.reg, self.corr)
        except Exception as x:
            QGuiApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, 'Subcontractors', f'Could not read {e["en"]}:\n{x}')
        QGuiApplication.restoreOverrideCursor()
        s = e.get('sub') or {}
        self.l_title.setText(e['en'].replace('_', ' ') + (f'  ·  {e["ar"]}' if e['ar'] else ''))
        ct = d['ct']
        parts = []
        no = s.get('Contract No.') or (ct['no'] if ct else '')
        cd = s.get('Contract Date') or (ct['date'] if ct else '')
        if no: parts.append(f'Contract <b style="color:{INK}">{no}</b>' + (f' ({fdate(cd)})' if cd else ''))
        for k, cap, suf in (('Contract Type', 'Type', ''), ('Duration (months)', 'Duration', ' months'), ('Retention %', 'Retention', '%'),
                            ('VAT %', 'VAT', '%'), ('CR No.', 'CR', ''), ('Representative', 'Rep.', ''), ('Status', '', '')):
            val = s.get(k)
            if val not in (None, ''):
                if isinstance(val, float) and val.is_integer(): val = int(val)
                parts.append((f'{cap} ' if cap else '') + f'<b style="color:{INK}">{val}{suf}</b>')
        if not s: parts.append('Not in the SUB register yet – it is added when its contract is distributed, or from «+ New subcontractor»')
        self.l_info.setText('&nbsp;&nbsp;·&nbsp;&nbsp;'.join(parts))
        if ct:
            self.kpi['appr'].setText('  '.join(('✔ ' if ct['have'][k] else '✘ ') + sh for k, _, sh in S.APPROVALS))
        else:
            self.kpi['appr'].setText('No contract')
        self.kpi['cert'].setText(money(d['certified']) if d['certs'] else '—')
        live = [m for m in d['memos'] if m['status'] != 'Cancelled']
        self.kpi['ded'].setText(f"{len(live)}  ·  {money(d['ded_total'])}" if live else '—')
        self.kpi['ret'].setText(money(d['retention']) if d['certs'] else '—')
        self._fill(e, d)

    def _put(self, tb, rows, paths):
        tb.setRowCount(0)
        for i, (vals, path) in enumerate(zip(rows, paths)):
            tb.insertRow(i)
            for j, val in enumerate(vals):
                x = QTableWidgetItem(str(val if val is not None else ''))
                x.setData(Qt.UserRole, path)
                if path: x.setToolTip(path)
                tb.setItem(i, j, x)

    def _fill(self, e, d):
        def frows(fs): return [(f['name'], f['where'], fdate(f['date']), f'{f["size"] / 1048576:,.1f} MB' if f['size'] >= 1048576 else f'{f["size"] / 1024:,.0f} KB') for f in fs]
        for k, key in (('contract', 'f_contract'), ('letters', 'f_letters'), ('approval', 'f_approval'), ('quotes', 'f_quotes')):
            self._put(self.t[k], frows(d[key]), [f['path'] for f in d[key]])
        # البنود
        rows = []
        for it in d['items']:
            dep = ''
            if it.get('Depth From') is not None or it.get('Depth To') is not None:
                a, b = it.get('Depth From'), it.get('Depth To')
                dep = f"{a or 0:g}–{b:g}" if b is not None else f'> {a or 0:g}'
            rows.append((it['Item No.'], it['Description'] or '', it['Unit'] or '', money(it['Rate']),
                         S.MEASURE_NAME.get(it.get('Measure') or 'MAN', it.get('Measure') or ''), it.get('Dia') or '', dep))
        self._put(self.t['items'], rows, [None] * len(rows))
        # المستخلصات: اللي في السجل + ملفات 04_CERTIFICATES اللي مش متسجلة
        rows, paths, used = [], [], set()
        self._cert_recs = []
        for c in d['certs']:
            self._cert_recs.append(c)
            p = S.live_path(c.get('Folder')) if c.get('Folder') else ''
            used.add(os.path.normcase(p or ''))
            rows.append((S.vlabel(c['Cert No.']), fdate(c.get('Period To')), money(c.get('Works To Date')), money(c.get('Current')),
                         money(c.get('Net Payable')), c.get('Status') or '', os.path.basename(p or '')))
            paths.append(p)
        for f in d['f_certs']:
            if any(os.path.normcase(f['path']).startswith(u) for u in used if u): continue
            rows.append(('', '', '', '', '', '', os.path.join(f['where'], f['name']))); paths.append(f['path'])
            self._cert_recs.append(None)
        self._put(self.t['certs'], rows, paths)
        # المذكرات: اللي في 10-LOG SUB + الملفات
        rows, paths, names = [], [], set()
        for m in d['memos']:
            p = next((f['path'] for f in d['f_memos'] if f['name'] == m['file'] or os.path.splitext(f['name'])[0] == os.path.splitext(m['file'])[0]), '')
            names.add(os.path.splitext(m['file'])[0].lower())
            rows.append((m['no'], fdate(m['date']), m['subject'], money(m['amount']), m['cert'] or '', m['status'], m['file'])); paths.append(p)
        for f in d['f_memos']:
            if os.path.splitext(f['name'])[0].lower() in names: continue
            rows.append(('', fdate(f['date']), '', '', '', '', os.path.join(f['where'], f['name']))); paths.append(f['path'])
        self._put(self.t['memos'], rows, paths)
        # الاستيشنات
        sid = (e.get('sub') or {}).get('ID')
        self._areas = sorted(self.reg.areas(sid), key=lambda a: (S.item_key(a['Item No.']), str(a['Line']), S.parse_sta(a['From']) or 0)) if (self.reg and sid) else []
        desc = {str(i['Item No.']): i['Description'] or '' for i in d['items']}
        rows = []
        for a in self._areas:
            x0, x1 = S.parse_sta(a['From']), S.parse_sta(a['To'])
            rows.append((a['Item No.'], desc.get(str(a['Item No.']), ''), a['Line'], a['From'], a['To'],
                         f'{abs(x1 - x0):,.2f}' if x0 is not None and x1 is not None else '', a.get('Notes') or ''))
        self._put(self.t['areas'], rows, [None] * len(rows))
        sg = self.reg.signs() if self.reg else []
        self._put(self.t['sign'], [(r, n) for r, n in sg], [None] * len(sg))
        import ded_qt as DD
        self._deds = DD.ded_rows(self.reg, sid) if (self.reg and sid) else []
        DD.fill_table(self.t['deds'], self.reg, self._deds) if self.reg else None
        counts = dict(contract=len(d['f_contract']), items=len(d['items']), areas=len(self._areas), sign=0, deds=len(self._deds), certs=self.t['certs'].rowCount(),
                      memos=self.t['memos'].rowCount(), letters=len(d['f_letters']), approval=len(d['f_approval']), quotes=len(d['f_quotes']))
        for i, (k, cap) in enumerate(TABS):
            self.tabs.setTabText(i, f'{cap} ({counts[k]})' if counts[k] else cap)
        if d['err']: self.l_hint.setText('⚠ ' + ' | '.join(d['err']))
        else: self.l_hint.setText('Double-click a file to open it.')

    # ------------------------------------------------------------ الأزرار
    def _open(self, p):
        if not p or not os.path.exists(p):
            return QMessageBox.information(self, 'Subcontractors', 'The file was not found.')
        try:
            if os.name == 'nt': os.startfile(p)
        except OSError as x:
            QMessageBox.warning(self, 'Subcontractors', f'Could not open it:\n{x}')

    def open_row(self, tb, r):
        x = tb.item(r, 0)
        p = x.data(Qt.UserRole) if x else None
        if p: self._open(p)

    def show_sel(self):
        """الملف (أو الفولدر) بتاع السطر المعلّم عليه في التاب المفتوح ← يفتح فولدره والملف متعلّم عليه. مفيش اختيار ← فولدر المقاول"""
        k = TABS[self.tabs.currentIndex()][0]
        tb = self.t.get(k); r = tb.currentRow() if tb is not None else -1
        p = None
        if r >= 0:
            if k == 'deds':
                rows = getattr(self, '_deds', [])
                p = S.ded_dir(rows[r]) if r < len(rows) else None
            elif k in ('contract', 'letters', 'approval', 'quotes', 'memos', 'certs'):
                x = tb.item(r, 0) or tb.item(r, tb.columnCount() - 1)
                p = x.data(Qt.UserRole) if x else None
        if p and os.path.exists(p): return reveal(p)
        if r >= 0 and k not in ('items', 'areas', 'sign'):
            QMessageBox.information(self, 'Subcontractors', 'This row has no file or folder yet – opening the subcontractor folder.')
        self.open_folder()

    def open_folder(self):
        if self.cur and self.cur.get('path'): self._open(self.cur['path'])
        else: QMessageBox.information(self, 'Subcontractors', 'This subcontractor has no folder in 13_PARTIES\\SUB yet.')

    def open_contract(self):
        fs = (self.d or {}).get('f_contract') or []
        if not fs: return QMessageBox.information(self, 'Subcontractors', 'No contract in 03_CONTRACT yet.')
        self._open(fs[0]['path'])

    # ------------------------------------------------------------ تعديل السجل (الاستيشنات والتوقيعات)
    def _sid(self):
        s = (self.cur or {}).get('sub')
        if not s:
            QMessageBox.information(self, 'Subcontractors', 'This subcontractor is not in the SUB register yet – '
                                    'distribute his contract, or add him from the old window.')
            return None
        return s['ID']

    def _save_reg(self):
        try:
            self.reg.save(); return True
        except PermissionError as x:
            QMessageBox.warning(self, 'Subcontractors', f'{x}\n\nNothing was saved.')
        except Exception as x:
            QMessageBox.warning(self, 'Subcontractors', f'Could not save:\n{x}')
        self.reg = S.Register()
        return False

    def lines(self):
        if getattr(self, '_lines', None) is None:
            QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
            try: self._lines = S.lines_in_log()
            except Exception: self._lines = []
            finally: QGuiApplication.restoreOverrideCursor()
        return self._lines

    def _area_dlg(self, title, v, allow_all):
        its = [i for i in self.reg.items(self.cur['sub']['ID']) if (i['Measure'] or 'MAN') != 'MAN']
        if not its:
            QMessageBox.information(self, 'Subcontractors', 'No items calculated from the requests (all are «Manual»).'); return None
        d = QDialog(self); d.setWindowTitle(title); d.setStyleSheet(CSS); d.setMinimumWidth(620)
        g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 16); g.setVerticalSpacing(8)
        h = QLabel(title); h.setFont(HQ.font(16, True)); g.addWidget(h, 0, 0, 1, 2)
        cb = QComboBox()
        if allow_all: cb.addItem(f'All {len(its)} items calculated from QS (pipes, asphalt …)', '*')
        for i in its: cb.addItem(f"{i['Item No.']}  –  {str(i['Description'] or '')[:70]}", str(i['Item No.']))
        if v.get('Item No.'):
            k = cb.findData(str(v['Item No.']))
            if k >= 0: cb.setCurrentIndex(k)
        ln = QComboBox(); ln.setEditable(True); ln.addItems(self.lines()); ln.setCurrentText(str(v.get('Line') or ''))
        e0, e1, en = QLineEdit(str(v.get('From') or '')), QLineEdit(str(v.get('To') or '')), QLineEdit(str(v.get('Notes') or ''))
        e0.setPlaceholderText('0+350'); e1.setPlaceholderText('0+800')
        for r, (txt, w) in enumerate((('Item', cb), ('Line (as in the WIR log)', ln), ('From station', e0), ('To station', e1),
                                      ('Notes', en)), 1):
            g.addWidget(QLabel(txt), r, 0); g.addWidget(w, r, 1)
        out = {}

        def ok():
            x0, x1 = S.parse_sta(e0.text()), S.parse_sta(e1.text())
            line = ln.currentText().strip()
            if not line or x0 is None or x1 is None:
                return QMessageBox.warning(d, 'Subcontractors', 'Line, From and To are needed (stations like 0+350).')
            if abs(x1 - x0) < 0.001: return QMessageBox.warning(d, 'Subcontractors', 'From = To')
            if self.lines() and line not in self.lines() and QMessageBox.question(
                    d, 'Subcontractors', f'Line «{line}» is not in the WIR log – save it anyway?') != QMessageBox.Yes: return
            x0, x1 = sorted((x0, x1))
            out.update(item=cb.currentData(), Line=line, From=S.sta(x0), To=S.sta(x1), Notes=en.text().strip())
            d.accept()
        bb = QHBoxLayout(); bb.addStretch(1); bb.addWidget(DQ.btn('Save', ok, True)); bb.addWidget(DQ.btn('Cancel', d.reject))
        g.addLayout(bb, 6, 0, 1, 2)
        d.exec()
        return out or None

    def _area_sel(self):
        r = self.t['areas'].currentRow()
        return self._areas[r] if 0 <= r < len(getattr(self, '_areas', [])) else None

    def add_area(self):
        sid = self._sid()
        if not sid: return
        last = self._areas[-1] if getattr(self, '_areas', None) else {}
        r = self._area_dlg('Add stations', {'Item No.': '*', 'Line': last.get('Line', '')}, True)
        if not r: return
        its = [str(i['Item No.']) for i in self.reg.items(sid) if (i['Measure'] or 'MAN') != 'MAN'] if r['item'] == '*' else [r['item']]
        have = {(str(a['Item No.']), a['Line'], a['From'], a['To']) for a in self.reg.areas(sid)}
        n = 0
        for k in its:
            if (k, r['Line'], r['From'], r['To']) in have: continue
            self.reg.data['AREAS'].append({'Sub ID': sid, 'Item No.': k, 'Line': r['Line'], 'From': r['From'], 'To': r['To'], 'Notes': r['Notes']})
            n += 1
        if not n: return QMessageBox.information(self, 'Subcontractors', 'These stations are already there.')
        if self._save_reg(): self.pick(self.lst.currentRow())

    def edit_area(self):
        a = self._area_sel()
        if not a: return QMessageBox.information(self, 'Subcontractors', 'Select a stations row first.')
        r = self._area_dlg('Edit stations', a, False)
        if not r: return
        a.update({'Item No.': r['item'], 'Line': r['Line'], 'From': r['From'], 'To': r['To'], 'Notes': r['Notes']})
        if self._save_reg(): self.pick(self.lst.currentRow())

    def del_area(self):
        a = self._area_sel()
        if not a: return QMessageBox.information(self, 'Subcontractors', 'Select a stations row first.')
        same = [x for x in self.reg.areas(a['Sub ID']) if (x['Line'], x['From'], x['To']) == (a['Line'], a['From'], a['To'])]
        b = QMessageBox(self); b.setWindowTitle('Subcontractors'); b.setIcon(QMessageBox.Question)
        b.setText(f"Delete {a['Line']}  {a['From']} → {a['To']}?")
        b_one = b.addButton(f"Item {a['Item No.']} only", QMessageBox.AcceptRole)
        b_all = b.addButton(f'All {len(same)} items', QMessageBox.DestructiveRole) if len(same) > 1 else None
        b_no = b.addButton('Cancel', QMessageBox.RejectRole); b.setDefaultButton(b_no)
        b.exec()
        c = b.clickedButton()
        if c is b_one: gone = [a]
        elif b_all is not None and c is b_all: gone = same
        else: return
        ids = {id(x) for x in gone}
        self.reg.data['AREAS'] = [x for x in self.reg.data['AREAS'] if id(x) not in ids]
        if self._save_reg(): self.pick(self.lst.currentRow())

    def edit_sign(self):
        r = self.t['sign'].currentRow()
        if not self.reg or r < 0 or r >= len(self.reg.data['SIGN']): return
        x = next((s for s in self.reg.data['SIGN'] if s['Role'] == self.t['sign'].item(r, 0).text()), None)
        if not x: return
        name, ok = QInputDialog.getText(self, 'Signatures', x['Role'], text=str(x.get('Name') or ''))
        if not ok: return
        x['Name'] = name.strip() or None
        if self._save_reg(): self.pick(self.lst.currentRow())

    # ------------------------------------------------------------ الخصومات
    def add_ded(self):
        sid = self._sid()
        if not sid: return
        import ded_qt as DD
        DD.DedDlg(self, self.reg, None, sid, CSS, DQ.btn).exec()
        self.pick(self.lst.currentRow())

    def open_ded(self, r):
        rows = getattr(self, '_deds', [])
        if r is None or r < 0 or r >= len(rows): return
        import ded_qt as DD
        DD.DedDlg(self, self.reg, rows[r], None, CSS, DQ.btn).exec()
        self.pick(self.lst.currentRow())

    def memo_pdf(self):
        """مذكرة الخصم بتاعة الخصم المعلّم عليه ← تتعمل تاني PDF (بالأسعار والصور) مكان القديمة وتتفتح"""
        rows = getattr(self, '_deds', [])
        r = self.t['deds'].currentRow()
        if r < 0 or r >= len(rows): return QMessageBox.information(self, 'Subcontractors', 'Select a deduction that is in a memo.')
        no = rows[r].get('Memo')
        if not no: return QMessageBox.information(self, 'Subcontractors', 'This deduction is not in a memo yet.')
        import ded_qt as DD
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        try:
            p, tot = DD.reissue_memo(self.reg, self.cur['sub'], no)
        except PermissionError:
            QGuiApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, 'Subcontractors', f'The {no} PDF is open in another program – close it and press Memo PDF again.')
        except Exception as x:
            QGuiApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, 'Subcontractors', f'{no} was not made:\n{x}')
        QGuiApplication.restoreOverrideCursor()
        self._open(p)
        self.pick(self.lst.currentRow())

    def all_deds(self):
        if not self.reg: return
        import ded_qt as DD
        DD.DedList(self, self.reg, CSS, _table, DQ.btn).exec()
        self.reload()

    # ------------------------------------------------------------ المستخلص
    def new_cert(self):
        sid = self._sid()
        if not sid: return
        if not self.reg.items(sid): return QMessageBox.information(self, 'Subcontractors', 'Add the contract items first.')
        cs = self.reg.certs(sid)
        if cs and str(cs[-1].get('Status') or 'Draft') not in ('Approved', 'Paid'):
            return QMessageBox.information(self, 'Subcontractors',
                                           f"V-{int(cs[-1]['Cert No.']):02d} is «{cs[-1].get('Status') or 'Draft'}» – a new certificate "
                                           'starts after it is Approved.\n\nOpen it (double-click) to edit it or change its status.')
        self._open_cert_dlg(sid)

    def _open_cert_dlg(self, sid, edit=None):
        dlg = getattr(self, '_cert_dlg', None)
        try:
            if dlg is not None and dlg.isVisible(): dlg.raise_(); dlg.activateWindow(); return
        except RuntimeError:                              # الشباك اتقفل واتمسح
            pass
        self._cert_dlg = CertDlg(self, self.reg, sid, edit)
        self._cert_dlg.show()

    def view_cert(self, r):
        recs = getattr(self, '_cert_recs', [])
        if r is None or r < 0 or r >= len(recs): return
        c = recs[r]
        if c is None: return self.open_row(self.t['certs'], r)          # ملف في 04_CERTIFICATES مش متسجل
        CertView(self, self.reg, self.cur['sub'], c).exec()

    def open_cert(self):
        r = self.t['certs'].currentRow()
        x = self.t['certs'].item(r, 0) if r >= 0 else None
        p = x.data(Qt.UserRole) if x else None
        if p: self._open(p if os.path.isdir(p) else os.path.dirname(p))
        else: self._open(os.path.join(self.cur['path'], PT.CERTS) if self.cur and self.cur.get('path') else '')


    # ------------------------------------------------------------ المقاول نفسه (كان في الشباك القديم)
    SUB_FIELDS = [('Name (as in the contract)', 'Name', 'combo', None),
                  ('English name (folder & files)', 'Folder', 'entry', None), ('CR No.', 'CR No.', 'entry', None),
                  ('VAT No.', 'VAT No.', 'entry', None), ('IBAN', 'IBAN', 'entry', None),
                  ('Representative', 'Representative', 'entry', None), ('Contract No.', 'Contract No.', 'entry', None),
                  ('Contract Date (dd-mm-yyyy)', 'Contract Date', 'date', None), ('Contract Type', 'Contract Type', 'entry', None),
                  ('Site Handover (dd-mm-yyyy)', 'Site Handover', 'date', None), ('Duration (months)', 'Duration (months)', 'entry', None),
                  ('Retention %', 'Retention %', 'entry', None), ('VAT %', 'VAT %', 'entry', None),
                  ('Advance %', 'Advance %', 'entry', None), ('Status', 'Status', 'rocombo', ('Active', 'Stopped', 'Withdrawn', 'Closed')),
                  ('Notes', 'Notes', 'entry', None)]

    def _sub_form(self, title, values):
        try: names = S.submittal_names()
        except Exception: names = []
        fields = [(l, k, kind, (names if k == 'Name' else o)) for l, k, kind, o in self.SUB_FIELDS]
        while True:
            r = form(self, title, fields, values)
            if not r: return None
            values = r
            err = self._sub_check(r)
            if not err: return r
            QMessageBox.warning(self, 'Subcontractors', err)

    def _sub_check(self, r):
        if not r['Name']: return 'Write the subcontractor name'
        folder = re.sub(r'\s+', ' ', r.get('Folder') or '').strip()
        if not folder and S.english_ok(r['Name']) and len(r['Name'].strip()) <= S.EN_MAX: folder = r['Name'].strip()
        if not S.english_ok(folder):
            return 'English name: letters, numbers and spaces only – it is the folder name and it goes in the file names'
        if len(folder) > S.EN_MAX: return f'English name: {S.EN_MAX} letters at most (Windows file paths are limited)'
        for k in ('Contract Date', 'Site Handover'):
            if r[k] and not pdate(r[k]): return f'{k}: invalid date (dd-mm-yyyy)'
        for k in ('Duration (months)', 'Retention %', 'VAT %', 'Advance %'):
            if r[k] and S.num(r[k]) is None: return f'{k}: not a number'
        r['Folder'] = folder
        for k in ('Contract Date', 'Site Handover'): r[k] = pdate(r[k]) if r[k] else None
        for k in ('Duration (months)', 'Retention %', 'VAT %', 'Advance %'): r[k] = S.num(r[k])
        return None

    def new_sub(self):
        if self.reg is None: self.reg = S.Register()
        r = self._sub_form('New subcontractor', {'Retention %': 15, 'VAT %': 15, 'Advance %': 0, 'Status': 'Active'})
        if not r: return
        if any(x['Name'] == r['Name'] for x in self.reg.subs()):
            return QMessageBox.warning(self, 'Subcontractors', 'This name is already there')
        if any(str(x.get('Folder') or '').lower() == r['Folder'].lower() for x in self.reg.subs()):
            return QMessageBox.warning(self, 'Subcontractors', 'Another subcontractor has this English name')
        r['ID'] = self.reg.new_id()
        self.reg.data['SUBS'].append(r)
        if self._save_reg():
            try: S.sub_folder(r)
            except OSError as x: QMessageBox.warning(self, 'Subcontractors', f'Folder not created:\n{x}')
            self.reload()

    def edit_sub(self):
        sid = self._sid()
        if not sid: return
        sub = self.reg.sub(sid)
        r = self._sub_form('Edit subcontractor', sub)
        if not r: return
        if r['Name'] != sub['Name'] and any(x['Name'] == r['Name'] for x in self.reg.subs()):
            return QMessageBox.warning(self, 'Subcontractors', 'This name is already there')
        old = sub.get('Folder') or S.safe_name(sub['Name'])
        if r['Folder'] != old:                              # الاسم الإنجليزي اتغير ← الفولدر ياخد الاسم الجديد
            if any(x is not sub and str(x.get('Folder') or '').lower() == r['Folder'].lower() for x in self.reg.subs()):
                return QMessageBox.warning(self, 'Subcontractors', 'Another subcontractor has this English name')
            try: has = S.memos(sub)
            except Exception: has = []
            if has:
                return QMessageBox.warning(self, 'Subcontractors', 'This subcontractor has deduction memos with the old English name '
                                           'in their file names – the English name cannot change')
            paths = {id(c): c.get('Folder') for c in self.reg.data['CERTS']}
            try:
                self.reg.writable()
                po, pn = S.rename_sub_dir(self.reg, sub['ID'], old, r['Folder'])
            except Exception as x:
                traceback.print_exc(); return QMessageBox.critical(self, 'Subcontractors', f'The folder name did not change:\n{x}')
            keep = dict(sub); sub.update(r)
            if self._save_reg(): return self.reload()
            try:                                            # السجل محفظش ← الفولدر يرجع زي ما كان
                if os.path.isdir(pn) and not os.path.exists(po): os.rename(pn, po)
            except OSError: traceback.print_exc()
            return
        sub.update(r)
        if self._save_reg(): self.reload()

    # ------------------------------------------------------------ البنود
    UNITS = ('م3', 'م.ط', 'م2', 'عدد', 'مقطوعية', 'طن')

    def _item_form(self, title, values, sid):
        mnames = [n for _, n in S.MEASURES]
        v = dict(values); v['Measure'] = S.MEASURE_NAME.get(v.get('Measure') or 'MAN', '')
        fields = [('Item No.', 'Item No.', 'entry', None),
                  ('Section (أ / ب / ج …)', 'Section', 'combo', sorted({it['Section'] for it in self.reg.items(sid) if it['Section']})),
                  ('Description', 'Description', 'entry', None), ('Unit', 'Unit', 'combo', self.UNITS),
                  ('Contract Qty', 'Contract Qty', 'entry', None), ('Rate (SAR)', 'Rate', 'entry', None),
                  ('Quantity from', 'Measure', 'rocombo', mnames), ('Dia (only this dia – empty = all)', 'Dia', 'combo', ('', '300', '400', '600')),
                  ('Depth from (m)  – depth items', 'Depth From', 'entry', None), ('Depth to (m)', 'Depth To', 'entry', None),
                  ('Notes', 'Notes', 'entry', None)]
        r = form(self, title, fields, v, width=520)
        if not r: return None
        if not r['Item No.'] or not r['Description']:
            QMessageBox.warning(self, 'Subcontractors', 'Item No. and Description are needed'); return None
        r['Measure'] = next((c for c, n in S.MEASURES if n == r['Measure']), 'MAN')
        for k in ('Contract Qty', 'Rate', 'Dia', 'Depth From', 'Depth To'):
            if r[k] and S.num(r[k]) is None:
                QMessageBox.warning(self, 'Subcontractors', f'{k}: not a number'); return None
            r[k] = S.num(r[k]) if r[k] else None
        if r['Dia']: r['Dia'] = int(r['Dia'])
        return r

    def _item_sel(self):
        r = self.t['items'].currentRow()
        its = (self.d or {}).get('items') or []
        return its[r] if 0 <= r < len(its) else None

    def add_item(self):
        sid = self._sid()
        if not sid: return
        its = self.reg.items(sid)
        nxt = max([int(S.num(i['Item No.']) or 0) for i in its] or [0]) + 1
        r = self._item_form('New item', {'Item No.': nxt, 'Section': its[-1]['Section'] if its else '', 'Unit': 'م3', 'Measure': 'MAN'}, sid)
        if not r: return
        if any(str(i['Item No.']) == str(r['Item No.']) for i in its):
            return QMessageBox.warning(self, 'Subcontractors', 'This item number is already there')
        r['Sub ID'] = sid
        self.reg.data['ITEMS'].append(r)
        if self._save_reg(): self.reload()

    def edit_item(self):
        sid = self._sid(); sel = self._item_sel()
        if not sid or not sel: return
        k = str(sel['Item No.'])
        it = next((i for i in self.reg.items(sid) if str(i['Item No.']) == k), None)
        if not it: return
        r = self._item_form(f'Item {k}', it, sid)
        if not r: return
        if str(r['Item No.']) != k:
            if any(str(i['Item No.']) == str(r['Item No.']) for i in self.reg.items(sid)):
                return QMessageBox.warning(self, 'Subcontractors', 'This item number is already there')
            if self.reg.certs(sid):
                return QMessageBox.warning(self, 'Subcontractors', 'Certificates were saved with this item number – it cannot be changed')
            for a in self.reg.areas(sid, k): a['Item No.'] = r['Item No.']
        it.update(r)
        if self._save_reg(): self.reload()

    def del_item(self):
        sid = self._sid(); sel = self._item_sel()
        if not sid or not sel: return
        k = str(sel['Item No.'])
        if any(str(q['Item No.']) == k and q['Sub ID'] == sid for q in self.reg.data['CERT_QTY']):
            return QMessageBox.warning(self, 'Subcontractors', 'This item is in a saved certificate – it cannot be deleted')
        if QMessageBox.question(self, 'Subcontractors', f'Delete item {k} and its stations?') != QMessageBox.Yes: return
        self.reg.data['ITEMS'] = [i for i in self.reg.data['ITEMS'] if not (i['Sub ID'] == sid and str(i['Item No.']) == k)]
        self.reg.data['AREAS'] = [a for a in self.reg.data['AREAS'] if not (a['Sub ID'] == sid and str(a['Item No.']) == k)]
        if self._save_reg(): self.reload()

    # ------------------------------------------------------------ مذكرات الخصم (ملف جاهز)
    def _memo_sel(self):
        r = self.t['memos'].currentRow()
        ms = (self.d or {}).get('memos') or []
        return ms[r] if 0 <= r < len(ms) else None

    def add_memo(self):
        sid = self._sid()
        if not sid: return
        sub = self.reg.sub(sid)
        if not S.english_ok(sub.get('Folder')):
            return QMessageBox.warning(self, 'Subcontractors', 'Set the English name of this subcontractor first (Edit data) – '
                                       'it is the folder name and it goes in the memo file names')
        try: S.log_writable()
        except Exception as x: return QMessageBox.critical(self, 'Subcontractors', str(x))
        no = S.next_memo_no(sub)
        r = form(self, f"Deduction memo DM-{no:03d}  –  {sub['Folder']}",
                 [('Subject (in English – it goes in the file name)', 'Subject', 'entry', None),
                  ('Memo date (dd-mm-yyyy)', 'Date', 'date', None), ('Amount (SAR)', 'Amount', 'entry', None)], {'Date': dt.date.today()})
        if not r: return
        if not S.english_ok(r['Subject']):
            return QMessageBox.warning(self, 'Subcontractors', 'Write the subject in English (letters, numbers and spaces)')
        if len(r['Subject'].strip()) > S.SUBJ_MAX:
            return QMessageBox.warning(self, 'Subcontractors', f'The subject is too long for a file name ({S.SUBJ_MAX} letters at most)')
        amt = S.num(r['Amount'])
        if amt is None or amt <= 0: return QMessageBox.warning(self, 'Subcontractors', 'Amount: a number more than 0')
        d = pdate(r['Date']) if r['Date'] else dt.date.today()
        if not d: return QMessageBox.warning(self, 'Subcontractors', 'Memo date: invalid')
        ib = S.inbox_dir()
        src, _ = QFileDialog.getOpenFileName(self, 'Choose the memo file (PDF, Word or scan)', ib if os.path.isdir(ib) else '',
                                             'Memo (*.pdf *.docx *.doc *.xlsx *.jpg *.jpeg *.png);;All files (*.*)')
        if not src: return
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        try:
            memo_no, dst, left = S.add_memo(sub, r['Subject'], d, amt, src)
        except Exception as x:
            QGuiApplication.restoreOverrideCursor(); traceback.print_exc(); return QMessageBox.critical(self, 'Subcontractors', str(x))
        QGuiApplication.restoreOverrideCursor()
        self.reload()
        QMessageBox.information(self, 'Subcontractors', f'{memo_no} saved as\n{os.path.basename(dst)}\n\nin DEDUCTION_MEMOS, and added to '
                                '00_Registers\\10-LOG SUB.xlsx (sheet DEDUCTION_MEMOS).')
        if left:
            QMessageBox.warning(self, 'Subcontractors', f'The original is still in 00_INBOX (it is open in another program):\n'
                                f'{os.path.basename(left)}\n\nClose it and delete it from 00_INBOX.')

    def cancel_memo(self):
        sid = self._sid(); m = self._memo_sel()
        if not sid: return
        if not m: return QMessageBox.information(self, 'Subcontractors', 'Select a memo first')
        sub = self.reg.sub(sid)
        if m['status'] == 'Deducted':
            return QMessageBox.warning(self, 'Subcontractors', f"{m['no']} is deducted in {m['cert']} – it cannot be cancelled")
        if m['status'] == 'Cancelled': return
        if QMessageBox.question(self, 'Subcontractors', f"Cancel {m['no']} – {m['subject']}?\n\nIt stays in the register as "
                                'Cancelled, and its number is not used again.') != QMessageBox.Yes: return
        try: S.cancel_memo(sub, m['no'])
        except Exception as x: traceback.print_exc(); return QMessageBox.critical(self, 'Subcontractors', str(x))
        self.reload()

    def old(self):
        """قسم SUB القديم في شباك لوحده (احتياطي)"""
        if self.run_old:
            try: self.run_old('sub')
            except Exception as x: QMessageBox.warning(self, 'Subcontractors', f'Could not open it:\n{x}')

    def on_close(self): return True


# ---------------------------------------------------------------- تفاصيل مستخلص متسجل (دبل كليك)
class CertView(QDialog):
    """المستخلص جوه البرنامج: الفلوس والبنود والاستيشنات والخصومات – والإكسل والـ PDF بزراير بس"""

    def __init__(self, app, reg, sub, c):
        super().__init__(app, Qt.Window)
        self.app, self.reg, self.sub, self.c = app, reg, sub, c
        self.no = int(c['Cert No.']); self.lab = f'V-{self.no:02d}'
        self.setWindowTitle(f"{self.lab}  ·  {(sub.get('Folder') or sub['Name']).replace('_', ' ')}")
        self.setStyleSheet(CSS); self.setObjectName('subq'); self.setAttribute(Qt.WA_StyledBackground, True)
        self.resize(1250, 780)
        v = QVBoxLayout(self); v.setContentsMargins(18, 14, 18, 14); v.setSpacing(10)
        self.h = QLabel(''); self.h.setFont(HQ.font(18, True)); v.addWidget(self.h)
        self.l_info = QLabel(''); self.l_info.setStyleSheet(f'color: {MUTED};'); self.l_info.setWordWrap(True); v.addWidget(self.l_info)
        kr = QHBoxLayout(); kr.setSpacing(8); self.k = {}
        for k, cap in (('gross', 'Works to date'), ('prev', 'Previous'), ('cur', 'This certificate'), ('vat', 'VAT'),
                       ('ret', 'Retention'), ('adddel', 'Additions / deductions'), ('net', 'NET PAYABLE (SAR)')):
            f = QFrame(); f.setObjectName('kpi'); fv = QVBoxLayout(f); fv.setContentsMargins(10, 6, 10, 6); fv.setSpacing(1)
            a = QLabel(cap); a.setStyleSheet(f'color: {MUTED};'); b = QLabel('—'); b.setFont(HQ.font(15, True))
            fv.addWidget(a); fv.addWidget(b); kr.addWidget(f, 1); self.k[k] = b
        v.addLayout(kr)
        tabs = QTabWidget(); tabs.setDocumentMode(True)
        self.t_it = _table([('no', 'ITEM', 60), ('desc', 'DESCRIPTION', 520), ('unit', 'UNIT', 60), ('rate', 'RATE', 90),
                            ('prev', 'PREVIOUS', 105), ('cur', 'THIS CERT.', 105), ('tot', 'TO DATE', 105), ('val', 'VALUE TO DATE', 130)])
        self.t_st = _table([('line', 'LINE', 120), ('from', 'FROM', 110), ('to', 'TO', 110), ('len', 'LENGTH (m)', 120)])
        self.t_dm = _table([('no', 'MEMO', 90), ('date', 'DATE', 100), ('subj', 'SUBJECT', 460), ('amt', 'AMOUNT', 120)])
        tabs.addTab(self.t_it, 'Items'); tabs.addTab(self.t_st, 'Stations'); tabs.addTab(self.t_dm, 'Deductions')
        v.addWidget(tabs, 1)
        br = QHBoxLayout(); br.setSpacing(8)
        br.addWidget(DQ.btn('Edit certificate', self.edit, True))
        br.addWidget(DQ.btn('Status / invoice', self.status))
        br.addStretch(1)
        br.addWidget(DQ.btn('Excel', lambda: self.open_file('.xlsx')))
        br.addWidget(DQ.btn('PDF', lambda: self.open_file(' - IPC.pdf')))
        br.addWidget(DQ.btn('Follow-up', lambda: self.open_file(' - Follow-up.pdf')))
        br.addWidget(DQ.btn('Open folder', lambda: self.app._open(self.folder())))
        br.addWidget(DQ.btn('Close', self.reject))
        v.addLayout(br)
        self.fill()

    def folder(self): return S.cert_dir(self.sub, self.c)

    def open_file(self, suf):
        d = self.folder()
        try: names = sorted(os.listdir(d)) if d and os.path.isdir(d) else []
        except OSError: names = []
        p = next((os.path.join(d, n) for n in names if n.startswith(self.lab) and n.endswith(suf)), None)
        if not p and suf == ' - IPC.pdf':               # مستخلص قديم اتطلع قبل الـ PDF الواحد ← ورقة الـ V بتاعته
            p = next((os.path.join(d, n) for n in names if n.startswith(self.lab) and n.endswith(' - V.pdf')), None)
        if not p: return QMessageBox.information(self, 'Subcontractors', f'No «{suf.strip(" -")}» file in the certificate folder.')
        self.app._open(p)

    def fill(self):
        c = self.c
        st = c.get('Status') or 'Draft'
        bg, fg = DQ.PILL.get({'Approved': 'sure', 'Paid': 'sure', 'Submitted': 'ask'}.get(st, 'wait'))
        self.h.setText(f"{'Final certificate' if c.get('Final') else 'Certificate'} {self.lab}   "
                       f"<span style='background-color:{bg}; color:{fg}; font-size:12pt;'>&nbsp;{st}&nbsp;</span>")
        inv = (f"   ·   invoice {c.get('Invoice No.')} ({fdate(c.get('Invoice Date'))}) {money(c.get('Invoice Amount'))}"
               if c.get('Invoice No.') else '')
        self.l_info.setText(f"Period {fdate(c.get('Period From'))} → {fdate(c.get('Period To'))}   ·   basis {c.get('Basis') or ''}"
                            f"   ·   saved {fdate(c.get('Saved At'))}"
                            + (f"   ·   in owner IPC-{int(c['Owner IPC']):02d}" if S.num(c.get('Owner IPC')) else '') + inv)
        for k, key in (('gross', 'Works To Date'), ('prev', 'Previous'), ('cur', 'Current'), ('vat', 'VAT'), ('net', 'Net Payable')):
            self.k[k].setText(money(c.get(key)) or '0.00')
        self.k['ret'].setText('-' + (money(c.get('Retention')) or '0.00'))
        self.k['adddel'].setText(f"+{money(c.get('Additions')) or '0.00'} / -{money(c.get('Deductions')) or '0.00'}")
        sid = c['Sub ID']
        its = {str(i['Item No.']): i for i in self.reg.items(sid)}
        now, prev = self.reg.cert_qty(sid, self.no), (self.reg.cert_qty(sid, self.no - 1) if self.no > 1 else {})
        rows, sec = [], None
        bold = QFont(); bold.setBold(True)
        self.t_it.setRowCount(0)
        for k, i in its.items():
            tot = S.num(now.get(k, {}).get('Qty To Date')) or 0.0
            pv = S.num(prev.get(k, {}).get('Qty To Date')) or 0.0
            if abs(tot) < 1e-9 and abs(pv) < 1e-9: continue
            r = self.t_it.rowCount()
            if i.get('Section') and i['Section'] != sec:
                sec = i['Section']; self.t_it.insertRow(r)
                x = QTableWidgetItem(sec); x.setFont(bold); x.setForeground(QColor('#9CC8F0')); self.t_it.setItem(r, 1, x); r += 1
            rate = S.num(now.get(k, {}).get('Rate')) or S.num(i.get('Rate')) or 0.0
            self.t_it.insertRow(r)
            for cc, val in enumerate((k, i['Description'] or '', i['Unit'] or '', money(rate), qty(pv), qty(tot - pv), qty(tot), money(tot * rate))):
                x = QTableWidgetItem(str(val))
                if cc >= 3: x.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.t_it.setItem(r, cc, x)
        if self.t_it.rowCount() == 0:
            self.t_it.insertRow(0); self.t_it.setItem(0, 1, QTableWidgetItem('No quantities in this certificate.'))
        seen, st_rows = set(), []
        for a in self.reg.areas(sid):
            if str(a.get('Notes') or '') != self.lab: continue
            key = (a['Line'], a['From'], a['To'])
            if key in seen: continue
            seen.add(key)
            x0, x1 = S.parse_sta(a['From']), S.parse_sta(a['To'])
            st_rows.append((a['Line'], a['From'], a['To'], f'{abs(x1 - x0):,.2f}' if x0 is not None and x1 is not None else ''))
        self.t_st.setRowCount(len(st_rows))
        for r, vals in enumerate(st_rows):
            for cc, val in enumerate(vals): self.t_st.setItem(r, cc, QTableWidgetItem(str(val)))
        try: dms = [m for m in S.memos(self.sub) if m['cert'] == self.lab and m['status'] == 'Deducted']
        except Exception: dms = []
        self.t_dm.setRowCount(len(dms))
        for r, m in enumerate(dms):
            for cc, val in enumerate((m['no'], fdate(m['date']), m['subject'], money(m['amount']))): self.t_dm.setItem(r, cc, QTableWidgetItem(str(val)))

    def edit(self):
        st = self.c.get('Status') or 'Draft'
        if st in ('Approved', 'Paid') and QMessageBox.question(
                self, 'Subcontractors', f'{self.lab} is «{st}». Edit it anyway?') != QMessageBox.Yes: return
        later = [x for x in self.reg.certs(self.c['Sub ID']) if int(x['Cert No.']) > self.no]
        if later and QMessageBox.question(
                self, 'Subcontractors', f"{', '.join(S.vlabel(x['Cert No.']) for x in later)} came after {self.lab} – "
                'their «previous» will not change by themselves. Edit anyway?') != QMessageBox.Yes: return
        self.accept()
        self.app._open_cert_dlg(self.c['Sub ID'], self.c)

    def status(self):
        c = self.c
        d = QDialog(self); d.setWindowTitle(f'{self.lab} – status / invoice'); d.setStyleSheet(CSS); d.setMinimumWidth(460)
        g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 16); g.setVerticalSpacing(8)
        cb = QComboBox(); cb.addItems(list(S.STATUSES)); cb.setCurrentText(c.get('Status') or 'Draft')
        e_no, e_dt, e_amt = QLineEdit(str(c.get('Invoice No.') or '')), QLineEdit(fdate(c.get('Invoice Date'))), QLineEdit(
            f"{S.num(c['Invoice Amount']):.2f}" if S.num(c.get('Invoice Amount')) else '')
        e_dt.setPlaceholderText('dd-mm-yyyy')
        for r, (txt, w) in enumerate((('Status', cb), ("Subcontractor's invoice No.", e_no), ('Invoice date', e_dt),
                                      ('Invoice amount (SAR)', e_amt))):
            g.addWidget(QLabel(txt), r, 0); g.addWidget(w, r, 1)

        def ok():
            amt = S.num(e_amt.text()) if e_amt.text().strip() else None
            if e_amt.text().strip() and amt is None: return QMessageBox.warning(d, 'Subcontractors', 'Invoice amount: not a number')
            try: idt = dt.datetime.strptime(e_dt.text().strip(), '%d-%m-%Y').date() if e_dt.text().strip() else None
            except ValueError: return QMessageBox.warning(d, 'Subcontractors', 'Invoice date: dd-mm-yyyy')
            c.update({'Status': cb.currentText(), 'Invoice No.': e_no.text().strip() or None, 'Invoice Date': idt, 'Invoice Amount': amt})
            if not self.app._save_reg(): return
            if amt is not None:
                net, tot = S.num(c.get('Net Payable')) or 0.0, S.num(c.get('Total incl. VAT')) or 0.0
                if abs(amt - net) > 1 and abs(amt - tot) > 1:
                    QMessageBox.warning(d, 'Subcontractors', f'The invoice ({amt:,.2f}) is not the certificate net ({net:,.2f}) '
                                                             f'nor the total incl. VAT ({tot:,.2f}) – check it before paying.')
            d.accept()
        bb = QHBoxLayout(); bb.addStretch(1); bb.addWidget(DQ.btn('Save', ok, True)); bb.addWidget(DQ.btn('Cancel', d.reject))
        g.addLayout(bb, 4, 0, 1, 2)
        if d.exec():
            self.fill()
            try: self.app.pick(self.app.lst.currentRow())
            except Exception: pass


# ---------------------------------------------------------------- شباك المستخلص (نفس حسابات وورقة القديم – sub_core)

# ---------------------------------------------------------------- PDF واحد للمستخلص: كفر ← متابعة ← V ← Q ← مذكرة الخصم
def _m(v): return f'{S.num(v) or 0:,.2f}'


def _cover_follow_pdf(path, cm):
    """بيان المتابعة: صفحة A4 بإطار ورأس فيه اللوجو، جداول من اليمين للشمال أبيض وأسود، والتوقيعات تحت الجدول على طول"""
    from PySide6.QtCore import QMarginsF, QRectF
    from PySide6.QtGui import (QPdfWriter, QPageSize, QPageLayout, QPainter, QTextDocument, QImage, QPen, QTextOption,
                               QAbstractTextDocumentLayout, QPalette)
    sub = cm['sub']
    sig = dict(cm.get('signs') or [])
    ct = str(sub.get('Contract No.') or '').strip()
    ct_line = f"{ct} – {fdate(sub.get('Contract Date'))}" if ct else ''
    nm = S.ar_name(sub)
    proj = 'مشروع شبكات المياه (مثال) – أمر شراء 000000000'
    contract = S.num(cm.get('contract')) or 0
    dur = S.num(sub.get('Duration (months)'))
    TB = "border='1' cellspacing='0' cellpadding='7' width='100%' style='border-collapse:collapse; border-color:#000; color:#000;"

    def kv(rows, size=10.5):
        # الجدول بيترسم من الشمال لليمين ← القيمة الأول (شمال) والعنوان آخر (يمين)
        return (f"<table {TB} font-size:{size}pt;'>"
                + ''.join(f"<tr><td align='right'>{v}</td><td align='right' width='34%'><b>{k}</b></td></tr>" for k, v in rows)
                + '</table>')
    inc = cm.get('owner_ipc')
    prog = f"{S.num(cm['progress']) * 100:.1f} %" if contract else '—'
    h = lambda t: f"<p align='right' style='font-size:12pt; font-weight:bold; color:#000; margin-top:12px; margin-bottom:4px;'>{t}</p>"
    ck = lambda b: '√' if b else ''
    cond = [('تم إدراج الأعمال بمستخلصات الجهة المالكة', bool(inc), not inc, inc or ''),
            ('أ- تم تسليمها للجهة المشرفة بدون ملاحظات', (not inc) and cm.get('all_approved'), False, ''),
            ('ب- تم تسليمها بملاحظات توقف الصرف بنسبة', False, False, '')]
    ctab = (f"<table {TB} font-size:10.5pt;'>"
            "<tr><th width='13%'>نسبة الصرف</th><th width='15%'>مستخلص رقم</th><th width='10%'>لا</th>"
            "<th width='10%'>نعم</th><th align='right'>البيان</th></tr>"
            + ''.join(f"<tr><td></td><td align='center'>{no}</td><td align='center'>{ck(n)}</td>"
                      f"<td align='center'>{ck(y)}</td><td align='right'>{t}</td></tr>" for t, y, n, no in cond) + '</table>')
    body = (f"<p align='center' style='font-size:17pt; font-weight:bold; color:#000;'>بيان المتابعة الشهري لأعمال مقاولي الباطن</p>"
            f"<p align='center' style='font-size:11pt; color:#000;'>{S.cert_title(cm)} – حتى {fdate(cm['cutoff'])}</p>"
            + h('أولاً: بيانات المشروع') + kv([('المشروع', proj), ('مقاول الباطن', nm)])
            + h('ثانياً: مدة التنفيذ') + kv([('رقم العقد / تاريخه', ct_line), ('تاريخ تسليم الموقع', fdate(sub.get('Site Handover'))),
                                             ('مدة العقد', f'{dur:g} أشهر' if dur else ''),
                                             ('مدة التأخير', f"{cm['delay']} يوم" if cm.get('delay') else 'لا يوجد')])
            + h('ثالثاً: الموقف المالي') + kv([('قيمة الأعمال بالمستخلص', _m(cm['net']) + ' ريال'), ('نسبة الإنجاز', prog),
                                              ('نسبة التأخير', '—'),
                                              ('الإجراءات بحق المقاول', f"{cm['ded_ref']} مرفقة" if cm.get('ded_ref') else 'لا يوجد')])
            + h('رابعاً: شروط الصرف حسب العقد') + ctab)
    roles = ['مدير المشروع', 'مدير إدارة المتابعة', 'مدير قطاع المقاولات']
    w = QPdfWriter(path)
    w.setPageSize(QPageSize(QPageSize.A4)); w.setResolution(150)
    w.setPageMargins(QMarginsF(10, 10, 10, 10), QPageLayout.Millimeter)
    w.setTitle('بيان المتابعة – ' + S.cert_title(cm))
    p = QPainter(w)
    W, H = w.width(), w.height()
    black = QColor('black')
    p.setPen(QPen(black, 4)); p.drawRect(QRectF(3, 3, W - 6, H - 6))
    p.setPen(QPen(black, 1.5)); p.drawRect(QRectF(13, 13, W - 26, H - 26))
    L, R, hh = 50, W - 50, 120
    logo = QImage(S.LOGO) if os.path.exists(S.LOGO) else QImage()
    if not logo.isNull(): p.drawImage(L, 40, logo.scaledToHeight(hh, Qt.SmoothTransformation))
    p.setPen(black)
    p.setFont(QFont('Arial', 16, QFont.Bold)); p.drawText(QRectF(L, 50, R - L, 50), int(Qt.AlignRight | Qt.AlignVCenter), S.COMPANY)
    p.setFont(QFont('Arial', 10)); p.drawText(QRectF(L, 100, R - L, 36), int(Qt.AlignRight | Qt.AlignVCenter),
                                              f'عقد رقم {ct} – {nm}' if ct else nm)
    p.setPen(QPen(black, 3)); p.drawLine(L, 40 + hh + 14, R, 40 + hh + 14)
    top = 40 + hh + 30
    doc = QTextDocument(); doc.documentLayout().setPaintDevice(w)
    opt = QTextOption(); opt.setTextDirection(Qt.RightToLeft); doc.setDefaultTextOption(opt)
    doc.setDefaultFont(QFont('Arial', 11)); doc.setDocumentMargin(0)
    doc.setDefaultStyleSheet('body, p, td, th { color: #000000; } th { font-weight: bold; }')
    doc.setHtml(f"<html><body dir='rtl' style='font-family:Arial; color:#000;'>{body}</body></html>")
    doc.setTextWidth(R - L)
    ctx = QAbstractTextDocumentLayout.PaintContext()       # الكلام أسود حتى لو البرنامج غامق
    pal = QPalette(); pal.setColor(QPalette.Text, black); pal.setColor(QPalette.WindowText, black); ctx.palette = pal
    ctx.clip = QRectF(0, 0, R - L, H - top - 60)
    p.save(); p.translate(L, top); doc.documentLayout().draw(p, ctx); p.restore()
    # التوقيعات تحت الجدول على طول: المسمى وتحته الاسم، من غير خط
    y = int(top + doc.size().height() + 60)
    cw = (R - L) / len(roles)
    for k, role in enumerate(roles):
        x = R - (k + 1) * cw                       # من اليمين للشمال
        p.setPen(black); p.setFont(QFont('Arial', 11, QFont.Bold))
        p.drawText(QRectF(x, y, cw, 36), Qt.AlignCenter, role)
        p.setFont(QFont('Arial', 10)); p.drawText(QRectF(x, y + 40, cw, 36), Qt.AlignCenter, sig.get(role, ''))
    p.end()


def memo_files(sub, ref):
    """ملفات مذكرات الخصم اللي في المستخلص (DM-001 …) من 06_DEDUCTION_MEMOS"""
    d = S.memo_dir(sub)
    try: names = sorted(os.listdir(d)) if d and os.path.isdir(d) else []
    except OSError: names = []
    out = []
    for no in re.findall(r'DM-\d+', str(ref or ''), re.I):
        f = next((n for n in names if n.upper().startswith(no.upper() + ' ') and n.lower().endswith('.pdf')), None)
        if f: out.append(os.path.join(d, f))
    return out


def make_ipc_pdf(xlsx, parts, base, cm):
    """ملفين PDF: «IPC» = V ← Q ← مذكرة الخصم بصورها، و«Follow-up» = بيان المتابعة لوحده. الـ V والـ Q من إكسل زي ما هما.
    الـ PDFات القديمة بتاعة نفس المستخلص (V / Q / Cover) تتنقل _OLD. يرجّع (المسار، المذكرات اللي ماتلاقتش)"""
    import ipc_core, pypdf
    folder = os.path.dirname(base)
    tmp = {k: base + f' - ~{k}.tmp.pdf' for k in ('cf', 'v', 'q')}
    jobs = [(parts['v'], tmp['v'])] + ([(parts['q'], tmp['q'])] if parts.get('q') else [])
    out, fol = base + ' - IPC.pdf', base + ' - Follow-up.pdf'
    try:
        ipc_core.export_pdfs(xlsx, jobs)
        _cover_follow_pdf(tmp['cf'], cm)
        os.replace(tmp['cf'], fol)
        mf = memo_files(cm['sub'], cm.get('ded_ref'))
        wr = pypdf.PdfWriter()
        for f in [tmp['v']] + ([tmp['q']] if parts.get('q') else []) + mf:
            for pg in pypdf.PdfReader(f).pages: wr.add_page(pg)
        with open(out + '.part', 'wb') as fh: wr.write(fh)
        os.replace(out + '.part', out)
    finally:
        for f in tmp.values():
            try: os.remove(f)
            except OSError: pass
    lab = os.path.basename(base)
    old = [n for n in os.listdir(folder) if n.startswith(lab) and n.endswith((' - V.pdf', ' - Q.pdf', ' - Cover & Follow-up.pdf'))]
    if old:
        arc = os.path.join(os.path.dirname(folder), '_OLD'); os.makedirs(arc, exist_ok=True)
        for n in old:
            try: shutil.move(os.path.join(folder, n), os.path.join(arc, n[:-4] + f' - replaced {dt.datetime.now():%Y-%m-%d %H%M}.pdf'))
            except OSError: pass
    missing = [no for no in re.findall(r'DM-\d+', str(cm.get('ded_ref') or ''), re.I)
               if not any(os.path.basename(f).upper().startswith(no.upper() + ' ') for f in memo_files(cm['sub'], no))]
    return out, missing


def _cno(notes):
    """رقم المستخلص من Notes بتاعة الاستيشنات («V-01» أو «V-01 *» = استثناء) أو None"""
    m = re.match(r'V-(\d+)', str(notes or ''))
    return int(m.group(1)) if m else None


def reveal(p):
    """فولدر ← يتفتح، ملف ← فولدره يتفتح والملف متعلّم عليه (Explorer)"""
    if not p or not os.path.exists(p) or os.name != 'nt': return
    try:
        if os.path.isdir(p): os.startfile(p)  # noqa
        else:
            import subprocess
            subprocess.Popen(f'explorer /select,"{os.path.normpath(p)}"')
    except Exception:
        pass


def qty(v):
    try: return f'{float(v):,.3f}'
    except (TypeError, ValueError): return ''


def _qd(d):
    return QDate(d.year, d.month, d.day) if isinstance(d, (dt.date, dt.datetime)) else QDate.currentDate()


def _pyd(qd): return dt.date(qd.year(), qd.month(), qd.day())


class CertDlg(QDialog):
    """مستخلص جديد: الفترة والأساس والإضافات والخصومات والبنود اليدوي ← Calculate ← Save & Export (Excel + PDF)"""

    def __init__(self, app, reg, sid, edit=None):
        super().__init__(app, Qt.Window)
        self.app, self.reg, self.sid = app, reg, sid
        self.sub = reg.sub(sid); self.cm = None
        self.edit = edit                                  # مستخلص متسجل بيتعدّل (تعديلات المالية)
        certs = reg.certs(sid)
        last = certs[-1] if certs else None
        if edit:
            self.no = int(edit['Cert No.'])
            last = next((c for c in certs if int(c['Cert No.']) == self.no - 1), None)
        else:
            self.no = (int(last['Cert No.']) + 1) if last else 1
        what = f'Edit certificate V-{self.no:02d}' if edit else 'New certificate'
        self.setWindowTitle(f"{what}  ·  {self.sub.get('Folder') or self.sub['Name']}")
        self.setStyleSheet(CSS); self.setObjectName('subq'); self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.resize(1300, 820)
        v = QVBoxLayout(self); v.setContentsMargins(18, 14, 18, 14); v.setSpacing(10)
        h = QLabel(f"{what}  ·  {(self.sub.get('Folder') or self.sub['Name']).replace('_', ' ')}"); h.setFont(HQ.font(18, True))
        v.addWidget(h)
        # الصف الأول: الرقم والأساس والفترة
        top = QHBoxLayout(); top.setSpacing(10)
        self.cb_no = QComboBox(); self.cb_no.addItem(f'V-{self.no:02d}', self.no)
        self.cb_basis = QComboBox(); self.cb_basis.addItems(['Approved', 'Requested'])
        p0 = (S.todate(last['Period To']) + dt.timedelta(days=1)) if last and S.todate(last['Period To']) else S.todate(self.sub.get('Site Handover'))
        self.d0, self.d1 = QDateEdit(_qd(p0)), QDateEdit(QDate.currentDate())
        for de in (self.d0, self.d1): de.setCalendarPopup(True); de.setDisplayFormat('dd-MM-yyyy')
        self.ck_final = QCheckBox('Final (الختامي)')
        if edit:                                          # نفس اللي اتكتب فيه قبل كده
            self.cb_basis.setCurrentText('Requested' if str(edit.get('Basis') or '').lower().startswith('req') else 'Approved')
            if S.todate(edit.get('Period From')): self.d0.setDate(_qd(S.todate(edit['Period From'])))
            if S.todate(edit.get('Period To')): self.d1.setDate(_qd(S.todate(edit['Period To'])))
            self.ck_final.setChecked(bool(edit.get('Final')))
        for txt, w in (('Certificate', self.cb_no), ('Basis', self.cb_basis), ('Period from', self.d0), ('to (cut-off)', self.d1)):
            l = QLabel(txt); l.setStyleSheet(f'color: {MUTED};'); top.addWidget(l); top.addWidget(w)
        top.addWidget(self.ck_final); top.addStretch(1)
        v.addLayout(top)
        # الإضافات والخصومات
        mid = QHBoxLayout(); mid.setSpacing(10)
        self.e_add, self.e_addref = QLineEdit(), QLineEdit('مذكرة إضافة رقم 1')
        self.e_add.setPlaceholderText('0.00'); self.e_add.setMaximumWidth(140)
        if edit and S.num(edit.get('Additions')):
            self.e_add.setText(f"{S.num(edit['Additions']):.2f}"); self.e_addref.setText(str(edit.get('Additions Ref') or ''))
        for txt, w in (('Additions (SAR, incl. VAT)', self.e_add), ('Additions memo', self.e_addref)):
            l = QLabel(txt); l.setStyleSheet(f'color: {MUTED};'); mid.addWidget(l); mid.addWidget(w)
        l = QLabel('Deductions'); l.setStyleSheet(f'color: {MUTED};'); mid.addWidget(l)
        mid.addWidget(DQ.btn('Choose memos…', self.pick_memos))
        self.l_ded = QLabel(''); mid.addWidget(self.l_ded, 1)
        v.addLayout(mid)
        # الأزرار فوق التابات، وكل جزء في تاب بالشاشة كلها (طلب المستخدم 17-11)
        p1 = QWidget(); v1 = QVBoxLayout(p1); v1.setContentsMargins(0, 8, 0, 0); v1.setSpacing(8)
        p2 = QWidget(); v2 = QVBoxLayout(p2); v2.setContentsMargins(0, 8, 0, 0); v2.setSpacing(8)
        # الاستيشنات اللي اشتغل فيها في المستخلص ده (+ اللي قبله بالرمادي – الإجمالي لحد تاريخه)
        l = QLabel('Stations he worked in this certificate  –  the program takes everything finished inside them from QS')
        l.setFont(HQ.font(12.5, True)); v1.addWidget(l)
        sh = QHBoxLayout(); sh.setSpacing(10)
        self.t_st = _table([('line', 'LINE', 140), ('from', 'FROM', 110), ('to', 'TO', 110), ('len', 'LENGTH (m)', 110), ('cert', 'CERTIFICATE', 160)])
        sh.addWidget(self.t_st, 1)
        sb = QVBoxLayout(); sb.setSpacing(6)
        sb.addWidget(DQ.btn('+ Add stations', self.add_st, True)); sb.addWidget(DQ.btn('Delete', self.del_st)); sb.addStretch(1)
        sh.addLayout(sb)
        v1.addLayout(sh, 3)
        # الاستثناءات: بند معين على استيشنات لوحده، أو كمية تزيد / تنقص – بالسبب
        l = QLabel('Exceptions  –  one item on its own stations, or a quantity + / − (with the reason)'); l.setFont(HQ.font(12.5, True))
        v1.addWidget(l)
        eh = QHBoxLayout(); eh.setSpacing(10)
        self.t_ex = _table([('item', 'ITEM', 70), ('desc', 'DESCRIPTION', 330), ('line', 'LINE', 90), ('from', 'FROM', 90), ('to', 'TO', 90),
                            ('q', 'QTY + / −', 110), ('why', 'REASON', 300)])
        eh.addWidget(self.t_ex, 1)
        eb = QVBoxLayout(); eb.setSpacing(6)
        eb.addWidget(DQ.btn('+ Exception', self.add_ex)); eb.addWidget(DQ.btn('Delete', self.del_ex)); eb.addStretch(1)
        eh.addLayout(eb)
        v1.addLayout(eh, 2)
        self.memos_err = ''
        try: self.memos_all = S.memos(self.sub)
        except Exception as x: self.memos_all, self.memos_err = [], str(x)
        self.ded_nos = self._linked(); self._show_ded()
        # البنود اليدوي
        self.man = [i for i in reg.items(sid) if (i['Measure'] or 'MAN') == 'MAN']
        self.man_vals = {}
        self.t_man = None
        if self.man:
            l = QLabel('Items you type – quantity TO DATE (double-click the number to change it)'); l.setFont(HQ.font(12.5, True)); v2.addWidget(l)
            prevq = reg.cert_qty(sid, self.no - 1) if self.no > 1 else {}
            nowq = reg.cert_qty(sid, self.no) if edit else {}
            self.t_man = _table([('no', 'ITEM', 60), ('desc', 'DESCRIPTION', 640), ('unit', 'UNIT', 70), ('prev', 'PREVIOUS', 110),
                                 ('tot', 'TO DATE ✎', 120)])
            self.t_man.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
            self.t_man.setRowCount(len(self.man))
            for r, i in enumerate(self.man):
                k = str(i['Item No.'])
                pv = S.num(prevq.get(k, {}).get('Qty To Date')) or 0.0
                tv = S.num(nowq.get(k, {}).get('Qty To Date')) if k in nowq else pv
                self.man_vals[k] = tv or 0.0
                for c, val in enumerate((k, i['Description'] or '', i['Unit'] or '', qty(pv), qty(tv))):
                    x = QTableWidgetItem(val); x.setData(Qt.UserRole, k)
                    if c != 4: x.setFlags(x.flags() & ~Qt.ItemIsEditable)
                    else: x.setBackground(QColor('#1E3448'))
                    self.t_man.setItem(r, c, x)
            self.t_man.itemChanged.connect(self.man_changed)
            v2.addWidget(self.t_man, 1)
        # الأزرار
        bar = QHBoxLayout(); bar.setSpacing(8)
        bar.addWidget(DQ.btn('Takeoff && Calculate', self.calc, True))
        self.b_take = DQ.btn('Takeoff rows (request by request)', self.takeoff); self.b_take.setEnabled(False); bar.addWidget(self.b_take)
        self.b_save = DQ.btn('Save && Export (Excel + PDF)', self.save_export, True); self.b_save.setEnabled(False); bar.addWidget(self.b_save)
        self.l_msg = QLabel('Choose the period and basis, then press Calculate.'); self.l_msg.setStyleSheet(f'color: {MUTED};')
        bar.addWidget(self.l_msg, 1)
        v.addLayout(bar)
        # النتيجة
        self.rt = QTabWidget(); self.rt.setDocumentMode(True)
        self.rt.addTab(p1, '1 · Stations && exceptions')
        if self.man: self.rt.addTab(p2, f'2 · Typed items ({len(self.man)})')
        else: v2.addWidget(QLabel('No typed items in this contract.')); self.rt.addTab(p2, '2 · Typed items')
        self.t_take = _table([('work', 'WORK (from QS)', 330), ('dia', 'DIA', 70), ('depth', 'DEPTH (m)', 100), ('unit', 'UNIT', 60),
                              ('q', 'QUANTITY', 130), ('item', '→ CONTRACT ITEM', 140)])
        self.rt.addTab(self.t_take, '3 · Takeoff')
        self.tv = _table([('no', 'ITEM', 60), ('desc', 'DESCRIPTION', 470), ('unit', 'UNIT', 60), ('rate', 'RATE', 90), ('prev', 'PREVIOUS', 105),
                          ('cur', 'THIS CERT.', 105), ('tot', 'TO DATE', 105), ('val', 'VALUE TO DATE', 130)])
        self.rt.addTab(self.tv, '4 · Certificate items')
        v.addWidget(self.rt, 1)
        kr = QHBoxLayout(); kr.setSpacing(8); self.k = {}
        for k, cap in (('gross', 'Works to date'), ('cur', 'This certificate'), ('vat', 'VAT'), ('ret', 'Retention'),
                       ('adddel', 'Additions / deductions'), ('net', 'NET PAYABLE (SAR)')):
            f = QFrame(); f.setObjectName('kpi'); fv = QVBoxLayout(f); fv.setContentsMargins(10, 6, 10, 6); fv.setSpacing(1)
            a = QLabel(cap); a.setStyleSheet(f'color: {MUTED};'); b = QLabel('—'); b.setFont(HQ.font(15, True))
            fv.addWidget(a); fv.addWidget(b); kr.addWidget(f, 1); self.k[k] = b
        v.addLayout(kr)
        self.l_info = QLabel(''); self.l_info.setStyleSheet(f'color: {MUTED};'); v.addWidget(self.l_info)
        for w in (self.cb_no, self.cb_basis): w.currentIndexChanged.connect(lambda *_: self.dirty())
        self.cb_no.currentIndexChanged.connect(lambda *_: self._no_changed())
        for w in (self.d0, self.d1): w.dateChanged.connect(lambda *_: self.dirty())
        for w in (self.e_add, self.e_addref): w.textChanged.connect(lambda *_: self.dirty())
        self.ck_final.toggled.connect(lambda *_: self.dirty())
        self._load_st()

    # ---- مذكرات الخصم
    def _label(self): return f'V-{int(self.cb_no.currentData()):02d}'

    def _linked(self):
        lab = self._label()
        return [m['no'] for m in self.memos_all if m['status'] == 'Deducted' and m['cert'] == lab]

    def _no_changed(self):
        self.ded_nos = self._linked(); self._show_ded(); self._load_st()

    # ---- الاستيشنات
    def _qs_items(self): return [str(i['Item No.']) for i in self.reg.items(self.sid) if (i['Measure'] or 'MAN') != 'MAN']

    def _load_st(self):
        """القديمة (مستخلصات قبل كده) والجديدة (بتاعة رقم المستخلص ده) – من غير تكرار لكل بند"""
        lab, seen = f'V-{self.no:02d}', set()
        self.prev_st, self.new_st = [], []
        for a in self.reg.areas(self.sid):
            k = (str(a['Line']), str(a['From']), str(a['To']), str(a.get('Notes') or ''))
            if k in seen: continue
            seen.add(k)
            n = _cno(k[3])
            if n is not None and n > self.no: continue          # استيشنات مستخلص بعده
            if n == self.no and k[3] != lab: continue           # استثناء المستخلص ده (بيتقري من EXCEPTIONS)
            (self.new_st if k[3] == lab else self.prev_st).append(dict(Line=k[0], From=k[1], To=k[2], Notes=k[3]))
        self.exc = [dict(e) for e in self.reg.data.get('EXCEPTIONS', [])
                    if e['Sub ID'] == self.sid and int(S.num(e['Cert No.']) or 0) == self.no]
        self._show_st()
        if hasattr(self, 't_ex'): self._show_ex()

    # ---- الاستثناءات
    def _show_ex(self):
        desc = {str(i['Item No.']): i['Description'] or '' for i in self.reg.items(self.sid)}
        self.t_ex.setRowCount(len(self.exc))
        for r, e in enumerate(self.exc):
            qv = S.num(e.get('Qty'))
            vals = (e['Item No.'], desc.get(str(e['Item No.']), ''), e.get('Line') or '', e.get('From') or '', e.get('To') or '',
                    (f'{qv:+,.3f}' if qv else ''), e.get('Reason') or '')
            for c, val in enumerate(vals):
                it = QTableWidgetItem(str(val)); it.setBackground(QColor('#3A3018')); self.t_ex.setItem(r, c, it)

    def add_ex(self):
        its = self.reg.items(self.sid)
        d = QDialog(self); d.setWindowTitle('Exception'); d.setStyleSheet(CSS); d.setMinimumWidth(620)
        g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 16); g.setVerticalSpacing(8)
        cb = QComboBox()
        for i in its: cb.addItem(f"{i['Item No.']}  –  {str(i['Description'] or '')[:70]}", str(i['Item No.']))
        kind = QComboBox(); kind.addItems(['On its own stations (from QS)', 'Quantity + / − (typed)'])
        ln = QComboBox(); ln.setEditable(True); ln.addItems(self.app.lines())
        last = (self.new_st or self.prev_st or [{}])[-1]; ln.setCurrentText(str(last.get('Line') or ''))
        e0, e1, eq, ew = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        e0.setPlaceholderText('0+130'); e1.setPlaceholderText('0+200'); eq.setPlaceholderText('e.g. 12.5 or -12.5')
        ew.setPlaceholderText('why – it is printed in the takeoff')
        lab_q = QLabel('Quantity + / −')
        rows = (('Item', cb), ('Type', kind), ('Line', ln), ('From station', e0), ('To station', e1), ('Quantity + / −', eq), ('Reason', ew))
        labels = {}
        for r, (txt, w) in enumerate(rows):
            labels[txt] = QLabel(txt); g.addWidget(labels[txt], r, 0); g.addWidget(w, r, 1)

        def mode():
            st = kind.currentIndex() == 0
            for t, w in (('Line', ln), ('From station', e0), ('To station', e1)): w.setVisible(st); labels[t].setVisible(st)
            eq.setVisible(not st); labels['Quantity + / −'].setVisible(not st)
        kind.currentIndexChanged.connect(lambda *_: mode()); mode()

        def ok():
            k = cb.currentData(); why = ew.text().strip()
            if not why: return QMessageBox.warning(d, 'Subcontractors', 'Write the reason.')
            it = next((i for i in its if str(i['Item No.']) == k), {})
            if kind.currentIndex() == 0:
                if (it.get('Measure') or 'MAN') == 'MAN':
                    return QMessageBox.warning(d, 'Subcontractors', 'This item is typed by hand (Manual) – use «Quantity + / −».')
                x0, x1 = S.parse_sta(e0.text()), S.parse_sta(e1.text()); line = ln.currentText().strip()
                if not line or x0 is None or x1 is None or abs(x1 - x0) < 0.001:
                    return QMessageBox.warning(d, 'Subcontractors', 'Line, From and To are needed (stations like 0+130).')
                x0, x1 = sorted((x0, x1))
                self.exc.append({'Sub ID': self.sid, 'Cert No.': self.no, 'Item No.': k, 'Line': line, 'From': S.sta(x0),
                                 'To': S.sta(x1), 'Qty': None, 'Reason': why})
            else:
                if (it.get('Measure') or 'MAN') == 'MAN':
                    return QMessageBox.warning(d, 'Subcontractors', 'This item is typed by hand – change its TO DATE in the Manual items table.')
                qv = S.num(eq.text())
                if not qv: return QMessageBox.warning(d, 'Subcontractors', 'Write the quantity (+ to add, − to take off).')
                self.exc.append({'Sub ID': self.sid, 'Cert No.': self.no, 'Item No.': k, 'Line': None, 'From': None, 'To': None,
                                 'Qty': qv, 'Reason': why})
            d.accept()
        bb = QHBoxLayout(); bb.addStretch(1); bb.addWidget(DQ.btn('Add', ok, True)); bb.addWidget(DQ.btn('Cancel', d.reject))
        g.addLayout(bb, len(rows), 0, 1, 2)
        if d.exec(): self._show_ex(); self.dirty()

    def del_ex(self):
        r = self.t_ex.currentRow()
        if r < 0 or r >= len(self.exc): return QMessageBox.information(self, 'Subcontractors', 'Select an exception first.')
        self.exc.pop(r); self._show_ex(); self.dirty()

    def _show_st(self):
        rows = [(x, False) for x in self.prev_st] + [(x, True) for x in self.new_st]
        self.t_st.setRowCount(len(rows))
        for r, (x, new) in enumerate(rows):
            x0, x1 = S.parse_sta(x['From']), S.parse_sta(x['To'])
            vals = (x['Line'], x['From'], x['To'], f'{abs(x1 - x0):,.2f}' if x0 is not None and x1 is not None else '',
                    (self._label() + '  (this one)') if new else (x['Notes'] or 'before'))
            for c, val in enumerate(vals):
                it = QTableWidgetItem(str(val))
                if new: it.setBackground(QColor('#1E3448'))
                else: it.setForeground(QColor('#7C8896'))
                self.t_st.setItem(r, c, it)
        tot = sum(abs((S.parse_sta(x['To']) or 0) - (S.parse_sta(x['From']) or 0)) for x in self.new_st)
        self.l_msg.setText(f'{len(self.new_st)} new station range(s) – {tot:,.2f} m. Press Calculate.' if self.new_st
                           else 'Add the stations he worked in, then press Calculate.')

    def add_st(self):
        d = QDialog(self); d.setWindowTitle('Add stations'); d.setStyleSheet(CSS); d.setMinimumWidth(460)
        g = QGridLayout(d); g.setContentsMargins(18, 16, 18, 16); g.setVerticalSpacing(8)
        last = (self.new_st or self.prev_st or [{}])[-1]
        ln = QComboBox(); ln.setEditable(True); ln.addItems(self.app.lines()); ln.setCurrentText(str(last.get('Line') or ''))
        e0, e1 = QLineEdit(), QLineEdit(); e0.setPlaceholderText('0+000'); e1.setPlaceholderText('0+450')
        for r, (txt, w) in enumerate((('Line (as in the WIR log)', ln), ('From station', e0), ('To station', e1))):
            g.addWidget(QLabel(txt), r, 0); g.addWidget(w, r, 1)

        def ok():
            x0, x1 = S.parse_sta(e0.text()), S.parse_sta(e1.text())
            line = ln.currentText().strip()
            if not line or x0 is None or x1 is None:
                return QMessageBox.warning(d, 'Subcontractors', 'Line, From and To are needed (stations like 0+350).')
            if abs(x1 - x0) < 0.001: return QMessageBox.warning(d, 'Subcontractors', 'From = To')
            if self.app.lines() and line not in self.app.lines() and QMessageBox.question(
                    d, 'Subcontractors', f'Line «{line}» is not in the WIR log – add it anyway?') != QMessageBox.Yes: return
            x0, x1 = sorted((x0, x1))
            for x in self.prev_st + self.new_st:                  # متداخل مع استيشنات قبل كده؟ نبّه بس (الحساب بيدمجهم من غير تكرار)
                a0, a1 = S.parse_sta(x['From']), S.parse_sta(x['To'])
                if x['Line'] == line and a0 is not None and a1 is not None and min(x1, max(a0, a1)) - max(x0, min(a0, a1)) > 0.001:
                    if QMessageBox.question(d, 'Subcontractors', f"{line} {S.sta(x0)} → {S.sta(x1)} overlaps {x['From']} → {x['To']} "
                                            f"({x['Notes'] or 'before'}).\nThe overlap is counted once. Add it anyway?") != QMessageBox.Yes:
                        return
                    break
            self.new_st.append(dict(Line=line, From=S.sta(x0), To=S.sta(x1), Notes=self._label()))
            d.accept()
        bb = QHBoxLayout(); bb.addStretch(1); bb.addWidget(DQ.btn('Add', ok, True)); bb.addWidget(DQ.btn('Cancel', d.reject))
        g.addLayout(bb, 3, 0, 1, 2)
        if d.exec():
            self._show_st(); self.dirty()

    def del_st(self):
        r = self.t_st.currentRow() - len(self.prev_st)
        if r < 0 or r >= len(self.new_st):
            return QMessageBox.information(self, 'Subcontractors', 'Select one of the stations of this certificate (the blue rows).')
        self.new_st.pop(r); self._show_st(); self.dirty()

    def _areas(self):
        """AREAS بعد المستخلص ده: كل اللي قبله + استيشناته لكل بنود QS"""
        lab = self._label()
        keep = [a for a in self.reg.data['AREAS'] if not (a['Sub ID'] == self.sid and _cno(a.get('Notes')) == self.no)]
        self._later = [a for a in keep if a['Sub ID'] == self.sid and (_cno(a.get('Notes')) or 0) > self.no]
        own = [e for e in getattr(self, 'exc', []) if e.get('Line') and e.get('From') and e.get('To')]
        own_items = {str(e['Item No.']) for e in own}
        add = [{'Sub ID': self.sid, 'Item No.': k, 'Line': x['Line'], 'From': x['From'], 'To': x['To'], 'Notes': lab}
               for x in self.new_st for k in self._qs_items() if k not in own_items]
        add += [{'Sub ID': self.sid, 'Item No.': str(e['Item No.']), 'Line': e['Line'], 'From': e['From'], 'To': e['To'],
                 'Notes': lab + ' *'} for e in own]
        return keep + add

    def _apply_ex(self, cm):
        """الكميات + / − (اللي في المستخلصات اللي قبله كمان – الإجمالي لحد تاريخه) وسطور الاستثناء في الحصر"""
        allx = [e for e in self.reg.data.get('EXCEPTIONS', []) if e['Sub ID'] == self.sid and int(S.num(e['Cert No.']) or 0) < self.no]
        allx += self.exc
        delta = {}
        for e in allx:
            if S.num(e.get('Qty')): delta[str(e['Item No.'])] = delta.get(str(e['Item No.']), 0.0) + S.num(e['Qty'])
        if delta:
            for r in cm['rows']:
                dq = delta.get(r['item'])
                if not dq: continue
                r['tot'] = round(r['tot'] + dq, 3); r['cur'] = round(r['tot'] - r['prev'], 3); r['val'] = r['tot'] * r['rate']
                r['over'] = (r['tot'] > r['cqty'] + 1e-6) if r['cqty'] else False
            cm['gross'] = round(sum(r['val'] for r in cm['rows']), 2)
            cm['cur'] = round(cm['gross'] - cm['prev'], 2)
            cm['vat'] = round(cm['cur'] * cm['vat_p'], 2); cm['ret'] = round((cm['cur'] + cm['vat']) * cm['ret_p'], 2)
            cm['net'] = round(cm['cur'] + cm['vat'] - cm['ret'] - cm['deductions'] + cm['additions'], 2)
            cm['progress'] = cm['gross'] / cm['contract'] if cm['contract'] else 0.0
        for e in allx:                                    # سطر في ورقة الحصر بالسبب
            k = str(e['Item No.']); qv = S.num(e.get('Qty'))
            x0, x1 = S.parse_sta(e.get('From')), S.parse_sta(e.get('To'))
            cm['qrows'].setdefault(k, []).append(dict(
                wir='استثناء', unf='', dist=S.vlabel(e['Cert No.']), street=str(e.get('Reason') or ''), line=e.get('Line') or '',
                x0=x0, x1=x1, L=abs(x1 - x0) if x0 is not None and x1 is not None else None, dia='', avg=None, pct=None,
                qty=qv if qv else None, appdate=None, reqdate=None))

    def _ded_open(self):
        """خصومات المقاول المفتوحة (مخالفات، معدات، مواد …) اللي لسه مادخلتش مذكرة"""
        return [d for d in self.reg.data.get('DEDUCTIONS', []) if d.get('Sub ID') == self.sid and (d.get('Status') or 'Open') == 'Open']

    def _ded(self):
        ch = [m for m in self.memos_all if m['no'] in self.ded_nos]
        items = [d for d in self._ded_open() if d['ID'] in getattr(self, 'ded_ids', [])]
        nos = [m['no'] for m in ch]
        if items:
            try: nos.append(f'DM-{S.next_memo_no(self.sub):03d}')     # المذكرة الجديدة اللي هتتعمل وقت الحفظ
            except Exception: nos.append('DM-new')
        return round(sum(m['amount'] for m in ch) + sum(S.ded_amount(d) for d in items), 2), S.memo_ref(nos)

    def _show_ded(self):
        if self.memos_err: self.l_ded.setText('⚠ 10-LOG SUB cannot be read – ' + self.memos_err[:120]); return
        amt, _ = self._ded()
        what = list(self.ded_nos) + list(getattr(self, 'ded_ids', []))
        self.l_ded.setText(f"{amt:,.2f} SAR   ({', '.join(what)})" if what else 'none')

    def pick_memos(self):
        if self.memos_err: return QMessageBox.warning(self, 'Subcontractors', f'10-LOG SUB cannot be read:\n{self.memos_err}')
        import ded_qt as DD
        lab = self._label()
        av = [m for m in self.memos_all if S.can_deduct(m, lab)]
        r = DD.pick_dialog(self, CSS, DQ.btn, av, self._ded_open(), self.ded_nos, getattr(self, 'ded_ids', []), lab)
        if r is None: return
        self.ded_nos, self.ded_ids = r
        self._show_ded(); self.dirty()

    def man_changed(self, x):
        if x.column() != 4: return
        k = x.data(Qt.UserRole)
        val = S.num(x.text())
        self.t_man.blockSignals(True)
        if val is None:
            QMessageBox.warning(self, 'Subcontractors', 'Not a number'); x.setText(qty(self.man_vals.get(k, 0)))
        else:
            self.man_vals[k] = val; x.setText(qty(val))
        self.t_man.blockSignals(False)
        self.dirty()

    def dirty(self):
        self.cm = None
        self.b_save.setEnabled(False); self.b_take.setEnabled(False)

    # ---- الحساب
    def calc(self):
        no = int(self.cb_no.currentData())
        d0, d1 = _pyd(self.d0.date()), _pyd(self.d1.date())
        if d0 > d1: return QMessageBox.warning(self, 'Subcontractors', 'Period from is after period to.')
        add = S.num(self.e_add.text() or 0)
        if add is None: return QMessageBox.warning(self, 'Subcontractors', 'Additions: not a number.')
        ded, ded_ref = self._ded()
        if not self.new_st and not self.prev_st and self._qs_items():
            return QMessageBox.information(self, 'Subcontractors', 'Add the stations he worked in first (+ Add stations).')
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        orig, orig_pi = self.reg.data['AREAS'], S.project_items
        try:
            ar = self._areas()
            self.reg.data['AREAS'] = [a for a in ar if not any(a is x for x in self._later)]
            res = S.project_items(d1, self.cb_basis.currentText().lower())          # QS مرة واحدة للحصر والمستخلص
            S.project_items = lambda *a, **k: res
            st = [dict(Line=x['Line'], From=x['From'], To=x['To']) for x in self.prev_st + self.new_st]
            self.take = takeoff(res[1], st, self.reg.items(self.sid))           # من الطلبات زي حسابات المالك
            self.cm = S.compute(self.reg, self.sid, no, d1, self.cb_basis.currentText().lower(), manual=self.man_vals,
                                additions=add, deductions=ded, add_ref=self.e_addref.text().strip() if add else '',
                                ded_ref=ded_ref if ded else '', period_from=d0, final=self.ck_final.isChecked())
            self._apply_ex(self.cm)
        except Exception as x:
            self.cm = None
            QGuiApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, 'Subcontractors', f'Could not calculate:\n{x}')
        finally:
            S.project_items = orig_pi
            self.reg.data['AREAS'] = orig                        # لسه ماتحفظش – بيتحفظ مع المستخلص
            QGuiApplication.restoreOverrideCursor()
        cm = self.cm
        self.t_take.setRowCount(0)
        bold = QFont(); bold.setBold(True)
        for r, x in enumerate(self.take):
            self.t_take.insertRow(r)
            for c, val in enumerate((x['work'], x['dia'], x['depth'], x['unit'], qty(x['qty']), x['item'] or '—')):
                it = QTableWidgetItem(str(val))
                if c == 4: it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter); it.setFont(bold)
                if c == 5: it.setForeground(QColor(DQ.GREEN if x['item'] else '#7C8896'))
                self.t_take.setItem(r, c, it)
        if not self.take:
            self.t_take.insertRow(0); self.t_take.setItem(0, 0, QTableWidgetItem('No requests inside these stations up to the cut-off date.'))
        self.rt.setCurrentIndex(2)                        # بعد الحساب ← تاب الحصر
        self.tv.setRowCount(0)
        sec, r = None, 0
        shown = [x for x in cm['rows'] if abs(x['tot']) > 1e-9 or abs(x['prev']) > 1e-9]
        for x in shown:
            if x['section'] and x['section'] != sec:
                sec = x['section']; self.tv.insertRow(r)
                it = QTableWidgetItem(sec); it.setFont(bold); it.setForeground(QColor('#9CC8F0'))
                self.tv.setItem(r, 1, it); r += 1
            self.tv.insertRow(r)
            vals = (x['item'], x['desc'], x['unit'], money(x['rate']), qty(x['prev']), qty(x['cur']), qty(x['tot']), money(x['val']))
            bad = x['over'] or x['cur'] < -1e-6
            for c, val in enumerate(vals):
                it = QTableWidgetItem(str(val))
                if c >= 3: it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if bad: it.setForeground(QColor('#FF8A8A'))
                elif c == 6 and x['measure'] != 'MAN' and x['tot']: it.setForeground(QColor(DQ.GREEN))
                self.tv.setItem(r, c, it)
            r += 1
        if not shown:
            self.tv.insertRow(0); self.tv.setItem(0, 1, QTableWidgetItem('Nothing finished inside these stations up to the cut-off date.'))
        self.k['gross'].setText(money(cm['gross'])); self.k['cur'].setText(money(cm['cur']))
        self.k['vat'].setText(money(cm['vat'])); self.k['ret'].setText('-' + money(cm['ret']))
        self.k['adddel'].setText(f"+{money(cm['additions'])} / -{money(cm['deductions'])}")
        self.k['net'].setText(money(cm['net']))
        ipc = (f"in owner IPC-{cm['owner_ipc']:02d}" if cm['owner_ipc'] else
               ('approved – not in an owner IPC yet' if cm['all_approved'] else 'not all approved'))
        self.l_info.setText(f"Previous {money(cm['prev'])}   ·   progress {cm['progress'] * 100:.2f}%   ·   requests: {ipc}"
                            + (f"   ·   delay {cm['delay']} days" if cm['delay'] else ''))
        msgs = []
        over = [x['item'] for x in cm['rows'] if x['over']]
        neg = [x['item'] for x in cm['rows'] if x['cur'] < -1e-6]
        if over: msgs.append(f"over the contract qty: {', '.join(map(str, over))}")
        if neg: msgs.append(f"less than the previous certificate: {', '.join(map(str, neg))}")
        if not any(cm['qrows'].values()) and any(x['measure'] != 'MAN' for x in cm['rows']):
            msgs.append('no requests inside the stations (check the line, the stations and the basis Approved / Requested)')
        self.l_msg.setText(('⚠ ' + ' | '.join(msgs)) if msgs else 'Check the numbers, then Save & Export.')
        self.l_msg.setStyleSheet(f"color: {'#F0C860' if msgs else DQ.GREEN};")
        self.b_save.setEnabled(True); self.b_take.setEnabled(True)

    def takeoff(self):
        """صفوف الحصر لكل بند: الطلب والخط والاستيشنات والطول والقطر ومتوسط العمق ونسبة التنفيذ والكمية (زي ورقة Q)"""
        cm = self.cm
        if not cm: return
        d = QDialog(self, Qt.Window); d.setWindowTitle('Takeoff – request by request'); d.setStyleSheet(CSS); d.resize(1250, 680)
        dv = QVBoxLayout(d)
        tb = _table([('item', 'ITEM', 60), ('wir', 'REQUEST', 80), ('dist', 'DISTRICT', 110), ('line', 'LINE', 70), ('from', 'FROM', 85),
                     ('to', 'TO', 85), ('l', 'LENGTH', 80), ('dia', 'DIA', 60), ('avg', 'AVG DEPTH', 90), ('pct', 'DONE %', 75),
                     ('q', 'QTY', 110), ('unit', 'UNIT', 60)])
        units = {x['item']: x['unit'] for x in cm['rows']}
        rows = []
        for k, rs in cm['qrows'].items():
            for q in rs:
                rows.append((k, q.get('wir', ''), q.get('dist', ''), q.get('line', ''), S.sta(q['x0']), S.sta(q['x1']),
                             f"{q['L']:,.2f}" if q.get('L') is not None else '',
                             q.get('dia') or '', f"{q['avg']:.2f}" if q.get('avg') is not None else '',
                             f"{q['pct']:g}" if q.get('pct') is not None else '', qty(q['qty']), units.get(k, '')))
        tb.setRowCount(len(rows))
        for r, vals in enumerate(rows):
            for c, val in enumerate(vals): tb.setItem(r, c, QTableWidgetItem(str(val)))
        dv.addWidget(tb, 1)
        dv.addWidget(QLabel(f'{len(rows)} rows'))
        d.show()
        self._take = d

    # ---- الحفظ (نفس خطوات القديم)
    def save_export(self):
        cm = self.cm
        if not cm: return
        Y = QMessageBox.Yes
        if not self.edit and any(int(c['Cert No.']) == cm['no'] for c in self.reg.certs(self.sid)) and QMessageBox.question(
                self, 'Subcontractors', f"Certificate {cm['no']} is already saved – replace it?") != Y: return
        label = f"V-{cm['no']:02d}"
        if self.memos_err:
            return QMessageBox.warning(self, 'Subcontractors', f'10-LOG SUB cannot be read:\n{self.memos_err}\n\nNothing was saved.')
        items = [d for d in self._ded_open() if d['ID'] in getattr(self, 'ded_ids', [])]
        if items:                                         # الخصومات المختارة ← مذكرة خصم واحدة الأول (بتتخصم مع باقي المذكرات)
            import ded_qt as DD
            QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
            try:
                no_new, tot = DD.memo_from_deductions(self.reg, self.sub, items, label, cm['cutoff'])
                self.memos_all = S.memos(self.sub)
            except Exception as x:
                QGuiApplication.restoreOverrideCursor()
                return QMessageBox.warning(self, 'Subcontractors', f'The deduction memo was not made:\n{x}\n\nNothing was saved.')
            QGuiApplication.restoreOverrideCursor()
            for d in items: d['Cert'] = label
            self.ded_ids = []
            self.ded_nos = list(self.ded_nos) + [no_new]
            if abs(tot - sum(S.ded_amount(d) for d in items)) > 0.01:
                QMessageBox.warning(self, 'Subcontractors', f'{no_new} = {tot:,.2f} – press Takeoff & Calculate again before saving.')
                self._show_ded(); self.dirty(); return
        nos = [m['no'] for m in self.memos_all if m['no'] in self.ded_nos]
        touch = bool(nos) or any(m['cert'] == label and m['status'] == 'Deducted' for m in self.memos_all)
        back = []
        if touch:                                       # 10-LOG SUB لازم يبقى مقفول قبل ما نحفظ أي حاجة
            try:
                S.log_writable(); fresh = S.memos(self.sub)
            except Exception as x:
                return QMessageBox.warning(self, 'Subcontractors', f'{x}\n\nNothing was saved.')
            byno = {f['no'].upper(): f for f in fresh}
            old = {m['no'].upper(): m for m in self.memos_all}
            changed = [n for n in nos if not byno.get(n.upper()) or not S.can_deduct(byno[n.upper()], label)
                       or byno[n.upper()]['amount'] != old[n.upper()]['amount']]
            if changed:
                self.memos_all = fresh
                self.ded_nos = [n for n in self.ded_nos if n not in changed]
                self._show_ded(); self.dirty()
                return QMessageBox.warning(self, 'Subcontractors', f"{', '.join(changed)} changed in 10-LOG SUB after you chose them – "
                                           'choose the memos again, then Calculate.\n\nNothing was saved.')
            back = [f['no'] for f in fresh if f['status'] == 'Deducted' and f['cert'] == label and f['no'] not in nos]
        folder = S.cert_folder(cm)
        old_dir = S.cert_dir(self.sub, self.edit) if self.edit else None
        base = os.path.join(folder, f"V-{cm['no']:02d} ({cm['cutoff']:%d-%m-%Y}) - {S.safe_name(self.sub.get('Folder') or self.sub['Name'])}")
        xlsx = base + '.xlsx'
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        orig = self.reg.data['AREAS']
        try:
            parts = S.build_workbook(cm, xlsx)
            self.reg.data['AREAS'] = self._areas()               # استيشنات المستخلص تتحفظ معاه (Notes = رقمه)
            orig_ex = self.reg.data.get('EXCEPTIONS', [])
            self.reg.data['EXCEPTIONS'] = [e for e in orig_ex if not (e['Sub ID'] == self.sid and int(S.num(e['Cert No.']) or 0) == self.no)] \
                + [dict(e) for e in self.exc]
            S.save_cert(self.reg, cm, folder)
        except PermissionError as x:
            self.reg.data['AREAS'] = orig
            if 'orig_ex' in locals(): self.reg.data['EXCEPTIONS'] = orig_ex
            QGuiApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, 'Subcontractors', f'{x}\n\n(close the file if it is open in Excel)')
        except Exception as x:
            self.reg.data['AREAS'] = orig
            if 'orig_ex' in locals(): self.reg.data['EXCEPTIONS'] = orig_ex
            QGuiApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, 'Subcontractors', f'Could not save:\n{x}')
        QGuiApplication.restoreOverrideCursor()
        while touch:                                    # المذكرات المختارة ← Deducted برقم المستخلص في 10-LOG SUB
            try:
                S.mark_deducted(self.sub, nos, label); break
            except Exception as x:
                if QMessageBox.question(self, 'Subcontractors', f'Certificate {label} is saved, but 10-LOG SUB was not updated:\n{x}\n\n'
                                        'Close 10-LOG SUB in Excel, then press Yes to try again.') == Y: continue
                todo = (f"• Status Deducted, Deducted in {label}: {', '.join(nos)}" if nos else '') + \
                       (f"\n• Status Open, Deducted in empty: {', '.join(back)}" if back else '')
                QMessageBox.warning(self, 'Subcontractors', f'Fix 10-LOG SUB (sheet DEDUCTION_MEMOS) by hand:\n{todo.strip()}')
                break
        self.l_msg.setText('Making the PDF files with Excel…'); QGuiApplication.processEvents()
        QGuiApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        try:
            pdf, miss = make_ipc_pdf(xlsx, parts, base, cm)
            pdf_msg = 'Excel + IPC.pdf (V, Q' + (', deduction memo' if cm.get('ded_ref') else '') + ') + Follow-up.pdf'
            if miss: pdf_msg += f"\nMemo file not found in 06_DEDUCTION_MEMOS: {', '.join(miss)}"
        except Exception as x:
            pdf_msg = f'Excel saved – PDF not made:\n{x}'
        finally:
            QGuiApplication.restoreOverrideCursor()
            try: self.raise_(); self.activateWindow()       # إكسل خلص ← شباك المستخلص يرجع قدام قبل ما يسأل
            except RuntimeError: pass
        if old_dir and os.path.isdir(old_dir) and os.path.normcase(os.path.normpath(old_dir)) != os.path.normcase(os.path.normpath(folder)):
            try:                                          # التاريخ اتغير ← الورق القديم يتنقل _OLD جنبه (مايتمسحش)
                arc = os.path.join(os.path.dirname(old_dir), '_OLD')
                os.makedirs(arc, exist_ok=True)
                shutil.move(old_dir, os.path.join(arc, os.path.basename(old_dir) + f' - replaced {dt.datetime.now():%Y-%m-%d %H%M}'))
            except Exception as x:
                QMessageBox.warning(self, 'Subcontractors', f'The old files of this certificate were not moved to _OLD:\n{x}')
        try: self.app.reload()
        except Exception: pass
        if QMessageBox.question(self, 'Subcontractors', f"Certificate {cm['no']} saved.\n{pdf_msg}\n\nOpen the folder?") == Y:
            try: os.startfile(folder)  # noqa
            except Exception: pass
        self.close()
