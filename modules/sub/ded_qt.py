# -*- coding: utf-8 -*-
"""الخصومات على المقاولين (طلب المستخدم 18-11): مخالفات الجهات (بنية تحتية، مياه، بلدية، مرور…) + معدات الشركة + مواد + غيره.

- كل خصم: سطر في شيت DEDUCTIONS بتاع 00_Registers\\10.1-SUBCONTRACTORS.xlsx + فولدر في 10_Site_Records\\DEDUCTIONS فيه ورقته وصوره.
- بيتشال على مقاول (أو يفضل على الشركة)، وفي المستخلص بيظهر مع مذكرات المقاول: اللي يتعلّم عليه يتخصم دلوقتي والباقي يستنى.
- عند حفظ المستخلص: كل اللي اتعلّم عليه يطلع في مذكرة خصم واحدة (PDF بشكل DM-001: جدول + صور مترقّمة + ورق الـ PDF في الآخر)
  تتسجل زي أي مذكرة (06_DEDUCTION_MEMOS + 10-LOG SUB) وتتخصم في المستخلص.
"""
import os, sys, tempfile, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import sub_core as S  # noqa: E402

from PySide6.QtCore import Qt, QDate, QMarginsF, QRectF, QSizeF, QSize  # noqa: E402
from PySide6.QtGui import (QPdfWriter, QPageSize, QPageLayout, QPainter, QTextDocument, QImage, QFont, QPen, QColor,  # noqa: E402
                           QIcon, QPixmap, QTextOption, QCursor, QGuiApplication)
from PySide6.QtWidgets import (QDialog, QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QComboBox, QLineEdit,  # noqa: E402
                               QDateEdit, QMessageBox, QFileDialog, QListWidget, QListWidgetItem, QAbstractItemView, QTableWidget,
                               QTableWidgetItem)

TYPE_EN = {t: t for t in S.DED_TYPES}


def _num(v): return S.num(v)


def _money(v):
    try: return f'{float(v):,.2f}'
    except (TypeError, ValueError): return ''


def _fd(v):
    d = S.todate(v)
    return d.strftime('%d-%m-%Y') if d else ''


def _qd(v):
    d = S.todate(v) or dt.date.today()
    return QDate(d.year, d.month, d.day)


def _open(p):
    if p and os.path.exists(p) and os.name == 'nt':
        try: os.startfile(p)  # noqa
        except OSError: pass


# ---------------------------------------------------------------- مذكرة الخصم الواحدة (PDF)
def make_memo_pdf(path, sub, items, memo_no, cert_label, cutoff, signs, memo_date=None):
    """مذكرة خصم بشكل DM-001: صفحة الجدول (كل الخصومات) ← صفحات الصور (4 في الصفحة، مترقّمة) ← ورق الـ PDF في الآخر.
    items = سطور DEDUCTIONS. يرجّع الإجمالي."""
    # الصور بالترتيب: ورقة الخصم لو صورة، وبعدين صوره – والـ PDF يتلم للآخر
    photos, pdfs, rng, seen = [], [], {}, {}
    import hashlib
    for i, d in enumerate(items, 1):
        paper, ph = S.ded_files(d)
        for f in ([paper] if paper else []) + ph:
            if f.lower().endswith('.pdf'): pdfs.append((i, d, f)); continue
            try:
                with open(f, 'rb') as fh: h = hashlib.md5(fh.read()).hexdigest()
            except OSError:
                h = f
            if h in seen:                                 # نفس الصورة على أكتر من بند (زي اللودر والتريلا) ← مرة واحدة
                k = seen[h]
                if i not in photos[k][0]: photos[k][0].append(i)
                continue
            seen[h] = len(photos)
            photos.append(([i], d, f, f == paper))
    for i in range(1, len(items) + 1):
        ns = [n for n, x in enumerate(photos, 1) if i in x[0]]
        rng[i] = (f'{ns[0]}' if len(ns) == 1 else f'{ns[0]} – {ns[-1]}' if ns == list(range(ns[0], ns[-1] + 1))
                  else '، '.join(map(str, ns))) if ns else ''
        if any(x[0] == i for x in pdfs): rng[i] = (rng[i] + ' + ' if rng[i] else '') + 'مرفق PDF'
    total = round(sum(S.ded_amount(d) for d in items), 2)
    sig = dict(signs or [])
    ct = str(sub.get('Contract No.') or '').strip()
    rows = ''
    for i, d in enumerate(items, 1):
        q = _num(d.get('Qty'))
        cnt = (f"{q:g} {d.get('Unit') or ''}".strip() if q else '')
        rows += (f"<tr><td>{i}</td><td>{S.ded_type_ar(d)}</td><td>{d.get('Ref No.') or '—'}</td>"
                 f"<td align='right'>{d.get('Description') or ''}</td><td>{d.get('Place') or ''}</td><td>{_fd(d.get('Date'))}</td>"
                 f"<td>{cnt}</td><td>{_money(d.get('Rate')) if _num(d.get('Rate')) else ''}</td>"
                 f"<td>{_money(S.ded_amount(d))}</td><td>{rng[i]}</td></tr>")
    th = ''.join(f"<th bgcolor='#DCE6F2'>{h}</th>" for h in
                 ('م', 'النوع', 'رقم / مرجع', 'البيان', 'المكان', 'التاريخ', 'العدد', 'سعر الوحدة', 'القيمة (ريال)', 'الصور'))
    npdf = len(pdfs)
    md = S.todate(memo_date) or dt.date.today()
    when = (f"مرفقة مع المستخلص الجاري رقم {int(cert_label.split('-')[-1])} (حتى {cutoff:%d-%m-%Y})" if cert_label
            else f"التاريخ: {md:%d-%m-%Y} – تُخصم من مستحقات المقاول في أول مستخلص")
    html = f"""<html><body dir='rtl' style="font-family:Arial; font-size:10.5pt;">
<p align='center' style="font-size:16pt; font-weight:bold;"><u>مذكرة خصم رقم ({int(memo_no.split('-')[-1])})</u></p>
<p align='center' style="font-weight:bold;">مشروع شبكات المياه (مثال) – أمر شراء 000000000</p>
<p align='center' style="font-weight:bold;">عقد مقاول الباطن رقم ({ct}) – {S.ar_name(sub)}</p>
<p align='center'>{when}</p>
<p style="font-weight:bold;">يتم خصم البنود التالية من مستحقات المقاول:</p>
<table dir='rtl' border='1' cellspacing='0' cellpadding='4' width='100%' style="border-collapse:collapse; font-size:9.5pt;">
<tr>{th}</tr>{rows}
<tr><td colspan='8' align='center' bgcolor='#F2F2F2'><b>الإجمالي</b></td><td bgcolor='#F2F2F2'><b>{_money(total)}</b></td><td bgcolor='#F2F2F2'></td></tr>
</table>
<p style="font-size:12pt; font-weight:bold;">إجمالي قيمة الخصم = {_money(total)} ريال</p>
<p>يتم خصمها من مستحقات المقاول – {S.ar_name(sub)}.</p>
<p><b>المرفقات:</b> ورق وصور الخصومات ({len(photos)} صورة{f' + {npdf} ملف PDF' if npdf else ''})، موضّح مع كل صورة رقم البند الخاص بها.</p>
<br><table width='100%' dir='rtl' cellpadding='3'>
<tr><td width='34%'><b>مندوب المقاول</b></td><td width='33%' align='center'><b>مهندس مكتب فني</b></td><td width='33%' align='center'><b>مدير المشروع</b></td></tr>
<tr><td>الاسم : {sig.get('مندوب مقاول الباطن', '')}</td><td align='center'>{sig.get('مهندس مكتب فني', '')}</td><td align='center'>{sig.get('مدير المشروع', '')}</td></tr>
</table>
</body></html>"""
    tmp = path + '.part.pdf'
    w = QPdfWriter(tmp)
    w.setPageSize(QPageSize(QPageSize.A4)); w.setResolution(150)
    w.setPageMargins(QMarginsF(12, 12, 12, 12), QPageLayout.Millimeter)
    w.setTitle(f'{memo_no} – {sub.get("Folder") or ""}')
    p = QPainter(w)
    W, H = w.width(), w.height()
    p.setLayoutDirection(Qt.RightToLeft)
    navy = QColor('#1F3A7A')

    def frame():
        p.setPen(QPen(navy, 3)); p.drawRect(QRectF(2, 2, W - 4, H - 4))
        p.setPen(QPen(navy, 1)); p.drawRect(QRectF(10, 10, W - 20, H - 20))

    # صفحة (صفحات) الجدول
    doc = QTextDocument()
    doc.documentLayout().setPaintDevice(w)
    opt = QTextOption(); opt.setTextDirection(Qt.RightToLeft); doc.setDefaultTextOption(opt)
    doc.setDefaultFont(QFont('Arial', 10))
    doc.setDocumentMargin(40)
    doc.setHtml(html)
    doc.setPageSize(QSizeF(W, H))
    for i in range(doc.pageCount()):
        if i: w.newPage()
        frame()
        p.save(); p.translate(0, -i * H)
        doc.drawContents(p, QRectF(0, i * H, W, H))
        p.restore()
    # صفحات الصور: 4 في الصفحة
    cw, chh = (W - 60) / 2, (H - 140) / 2
    for k in range(0, len(photos), 4):
        w.newPage(); frame()
        p.setPen(QColor('black')); p.setFont(QFont('Arial', 13, QFont.Bold))
        p.drawText(QRectF(0, 20, W, 50), Qt.AlignCenter, 'مرفق – ورق وصور الخصومات' + (' (تابع)' if k else ''))
        for j, (ii, d, f, is_paper) in enumerate(photos[k:k + 4]):
            n = k + j + 1
            col, row = (1 - j % 2), j // 2                     # من اليمين للشمال
            x, y = 30 + col * cw, 80 + row * chh
            p.setPen(QPen(QColor('black'), 1)); p.drawRect(QRectF(x, y, cw, chh))
            p.setPen(navy); p.setFont(QFont('Arial', 10, QFont.Bold))
            p.drawText(QRectF(x, y + 4, cw, 28), Qt.AlignCenter, f"صورة ({n})  –  البند {' و '.join(map(str, ii))}")
            img = QImage(f)
            box = QRectF(x + 10, y + 36, cw - 20, chh - 110)
            if not img.isNull():
                sc = img.scaled(int(box.width()), int(box.height()), Qt.KeepAspectRatio, Qt.SmoothTransformation)
                p.drawImage(QRectF(box.x() + (box.width() - sc.width()) / 2, box.y() + (box.height() - sc.height()) / 2,
                                   sc.width(), sc.height()), sc)
            else:
                p.setPen(QColor('grey')); p.drawText(box, Qt.AlignCenter, os.path.basename(f))
            p.setPen(QColor('black')); p.setFont(QFont('Arial', 9))
            cap = (('ورقة ' if is_paper else '') + S.ded_type_ar(d) + (f"  {d.get('Ref No.')}" if d.get('Ref No.') else '')
                   + '\n' + ' – '.join(x for x in (str(d.get('Place') or ''), _fd(d.get('Date'))) if x))
            p.drawText(QRectF(x + 6, y + chh - 70, cw - 12, 66), int(Qt.AlignCenter) | int(Qt.TextWordWrap), cap)
    p.end()
    # ورق الـ PDF في الآخر
    if pdfs:
        import pypdf
        out = pypdf.PdfWriter()
        for f in [tmp] + [f for _, _, f in pdfs]:
            try:
                for pg in pypdf.PdfReader(f).pages: out.add_page(pg)
            except Exception:
                if f == tmp: raise
        with open(path, 'wb') as fh: out.write(fh)
        os.remove(tmp)
    else:
        os.replace(tmp, path)
    return total


def reissue_memo(reg, sub, memo_no):
    """زرار Memo PDF: مذكرة متطلّعة قبل كده تتعمل تاني من خصوماتها (بالأسعار والصور اللي اتكتبت بعدها)
    ← نفس اسم الملف في 06_DEDUCTION_MEMOS (القديم يتحفظ .bak-YYYY-MM-DD) ← يرجّع (المسار، الإجمالي)"""
    items = sorted([d for d in reg.data.get('DEDUCTIONS', []) if d.get('Sub ID') == sub['ID']
                    and str(d.get('Memo') or '').upper() == str(memo_no).upper() and d.get('Status') != 'Cancelled'],
                   key=lambda d: str(d.get('ID')))
    if not items: raise ValueError(f'No deductions are linked to {memo_no}.')
    m = next((x for x in S.memos(sub) if x['no'].upper() == str(memo_no).upper()), None)
    dd = S.memo_dir(sub); os.makedirs(dd, exist_ok=True)
    old = sorted(f for f in os.listdir(dd) if f.upper().startswith(str(memo_no).upper() + ' ') and f.lower().endswith('.pdf'))
    name = old[0] if old else (m['file'] if m and str(m.get('file') or '').lower().endswith('.pdf') else
                               f"{str(memo_no).upper()} - {sub.get('Folder')} - Deductions ({dt.date.today():%Y-%m-%d}).pdf")
    md = (m or {}).get('date')
    if not md:
        import re as _re
        g = _re.search(r'\((\d{4}-\d{2}-\d{2})\)', name)
        md = S.todate(g.group(1)) if g else dt.date.today()
    tmp = os.path.join(tempfile.gettempdir(), name)
    cert = (m or {}).get('cert') if m and m.get('status') == 'Deducted' else None
    total = make_memo_pdf(tmp, sub, items, str(memo_no).upper(), cert, S.todate(md), reg.signs(), memo_date=md)
    dst = os.path.join(dd, name)
    if os.path.exists(dst):
        base, n = dst + f'.bak-{dt.date.today():%Y-%m-%d}', 0
        bak = base
        while os.path.exists(bak):
            n += 1; bak = base + chr(ord('a') + n)
        os.replace(dst, bak)
    import shutil
    shutil.move(tmp, dst)
    if (m is None or m['status'] == 'Open') and total > 0:
        try: S.set_memo_amount(sub, memo_no, total)
        except Exception: pass
    return dst, total


def memo_from_deductions(reg, sub, items, cert_label, cutoff):
    """الخصومات اللي اتعلّم عليها ← مذكرة خصم واحدة متسجلة (06_DEDUCTION_MEMOS + 10-LOG SUB) ← حالتها «In memo».
    يرجّع (رقم المذكرة، الإجمالي)."""
    S.log_writable()
    memo_no = f'DM-{S.next_memo_no(sub):03d}'
    tmp = os.path.join(tempfile.gettempdir(), f'{memo_no} - deductions {cert_label}.pdf')
    total = make_memo_pdf(tmp, sub, items, memo_no, cert_label, cutoff, reg.signs())
    if total <= 0: raise ValueError('The ticked deductions have no amount – write the amount of each one first.')
    no, dst, _ = S.add_memo(sub, f'Deductions {cert_label}', dt.date.today(), total, tmp)
    for d in items: d.update({'Status': 'In memo', 'Memo': no})
    reg.save()
    return no, total


# ---------------------------------------------------------------- شاشة الخصم (بيانات + ورقة + صور)
class DedDlg(QDialog):
    def __init__(self, parent, reg, d=None, sub_id=None, css='', btn=None):
        super().__init__(parent, Qt.Window)
        self.reg, self.new = reg, d is None
        self.d = d if d is not None else {'ID': S.ded_next_id(reg), 'Date': dt.date.today(), 'Type': 'Violation', 'Source': None,
                                          'Status': 'Open' if sub_id else 'Company', 'Sub ID': sub_id}
        self.btn = btn
        self.setWindowTitle(f"Deduction {self.d['ID']}"); self.setStyleSheet(css); self.setObjectName('subq')
        self.setAttribute(Qt.WA_StyledBackground, True); self.resize(1000, 760)
        v = QVBoxLayout(self); v.setContentsMargins(18, 14, 18, 14); v.setSpacing(8)
        self.h = QLabel(f"Deduction {self.d['ID']}"); f = QFont(); f.setPointSize(16); f.setBold(True); self.h.setFont(f); v.addWidget(self.h)
        g = QGridLayout(); g.setHorizontalSpacing(12); g.setVerticalSpacing(6)
        self.cb_type = QComboBox(); self.cb_type.addItems(list(S.DED_TYPES)); self.cb_type.setCurrentText(self.d.get('Type') or 'Violation')
        self.cb_src = QComboBox(); self.cb_src.addItems([''] + list(S.DED_SOURCES)); self.cb_src.setCurrentText(self.d.get('Source') or '')
        self.e_ref, self.e_desc, self.e_place = (QLineEdit(str(self.d.get(k) or '')) for k in ('Ref No.', 'Description', 'Place'))
        self.e_desc.setPlaceholderText('in Arabic – it is printed in the memo')
        self.de = QDateEdit(_qd(self.d.get('Date'))); self.de.setCalendarPopup(True); self.de.setDisplayFormat('dd-MM-yyyy')
        self.e_qty = QLineEdit('' if self.d.get('Qty') is None else f"{_num(self.d['Qty']):g}")
        self.e_unit = QLineEdit(str(self.d.get('Unit') or ''))
        self.e_rate = QLineEdit('' if self.d.get('Rate') is None else f"{_num(self.d['Rate']):.2f}")
        self.e_amt = QLineEdit('' if self.d.get('Amount') is None else f"{_num(self.d['Amount']):.2f}")
        self.e_amt.setPlaceholderText('or Qty × Rate')
        self.cb_sub = QComboBox(); self.cb_sub.addItem('— the company (not charged)', None)
        for s in reg.subs(): self.cb_sub.addItem(f"{s['ID']} · {str(s.get('Folder') or s['Name']).replace('_', ' ')}", s['ID'])
        k = self.cb_sub.findData(self.d.get('Sub ID')); self.cb_sub.setCurrentIndex(max(k, 0))
        self.e_notes = QLineEdit(str(self.d.get('Notes') or ''))
        cells = [('Type', self.cb_type), ('Source (violations)', self.cb_src), ('Ref / violation No.', self.e_ref), ('Date', self.de),
                 ('What (Arabic)', self.e_desc), ('Place', self.e_place), ('Qty', self.e_qty), ('Unit', self.e_unit),
                 ('Rate (SAR)', self.e_rate), ('Amount (SAR)', self.e_amt), ('Charged to', self.cb_sub), ('Notes', self.e_notes)]
        for n, (txt, wdg) in enumerate(cells):
            r, c = divmod(n, 2)
            g.addWidget(QLabel(txt), r, c * 2); g.addWidget(wdg, r, c * 2 + 1)
        v.addLayout(g)
        st = self.d.get('Status') or 'Open'
        self.l_st = QLabel(f"Status: {st}" + (f"  ·  {self.d.get('Memo')}" if self.d.get('Memo') else '')
                           + (f"  ·  deducted in {self.d.get('Cert')}" if self.d.get('Cert') else ''))
        v.addWidget(self.l_st)
        # الورقة
        pr = QHBoxLayout(); pr.setSpacing(8)
        self.l_paper = QLabel(''); pr.addWidget(QLabel('Paper:')); pr.addWidget(self.l_paper, 1)
        pr.addWidget(self._b('Open', self.open_paper)); pr.addWidget(self._b('Choose / change…', self.pick_paper))
        v.addLayout(pr)
        # الصور
        v.addWidget(QLabel('Photos (pictures or PDF) – double-click to open'))
        self.ph = QListWidget(); self.ph.setViewMode(QListWidget.IconMode); self.ph.setIconSize(QSize(180, 130))
        self.ph.setResizeMode(QListWidget.Adjust); self.ph.setMovement(QListWidget.Static); self.ph.setSpacing(8)
        self.ph.itemDoubleClicked.connect(lambda it: _open(it.data(Qt.UserRole)))
        v.addWidget(self.ph, 1)
        br = QHBoxLayout(); br.setSpacing(8)
        br.addWidget(self._b('+ Photos', self.add_photos, True)); br.addWidget(self._b('Delete photo', self.del_photo))
        br.addWidget(self._b('Open folder', lambda: _open(S.ded_dir(self.d)))); br.addStretch(1)
        self.b_cancel_d = self._b('Cancel this deduction', self.cancel_ded); br.addWidget(self.b_cancel_d)
        br.addWidget(self._b('Save', self.save, True)); br.addWidget(self._b('Close', self.reject))
        v.addLayout(br)
        # دخل مذكرة لسه ماتخصمتش ← السعر والعدد والمبلغ بس يتكتبوا (ومبلغ المذكرة يتحدّث)، الباقي مقفول
        self.memo_open = False
        if st == 'In memo' and self.d.get('Memo') and self.d.get('Sub ID'):
            try:
                self.memo_open = not any(m['no'].upper() == str(self.d['Memo']).upper() and m['status'] in ('Deducted', 'Cancelled')
                                         for m in S.memos(self.reg.sub(self.d['Sub ID'])))
            except Exception:
                self.memo_open = False
        locked = st in ('In memo', 'Cancelled')
        for wdg in (self.cb_type, self.cb_src, self.e_ref, self.de, self.e_desc, self.e_place, self.cb_sub):
            wdg.setEnabled(not locked)
        for wdg in (self.e_qty, self.e_unit, self.e_rate, self.e_amt):
            wdg.setEnabled(not locked or self.memo_open)
        if self.memo_open:
            self.l_st.setText(self.l_st.text() + f"   –   {self.d['Memo']} is not deducted yet: Qty, Rate and Amount can still change")
        self.b_cancel_d.setEnabled(not locked and not self.new)
        self.fill_files()

    def _b(self, txt, fn, pri=False):
        if self.btn: return self.btn(txt, fn, pri)
        from PySide6.QtWidgets import QPushButton
        b = QPushButton(txt); b.clicked.connect(fn); return b

    def fill_files(self):
        paper, photos = S.ded_files(self.d)
        self.paper = paper
        self.l_paper.setText(os.path.basename(paper) if paper else '— none yet')
        self.ph.clear()
        for f in photos:
            it = QListWidgetItem(os.path.basename(f)); it.setData(Qt.UserRole, f)
            if not f.lower().endswith('.pdf'):
                img = QImage(f)
                if not img.isNull(): it.setIcon(QIcon(QPixmap.fromImage(img.scaled(180, 130, Qt.KeepAspectRatio, Qt.SmoothTransformation))))
            else:
                it.setText('📄 ' + os.path.basename(f))
            self.ph.addItem(it)

    def _start(self):
        ib = S.inbox_dir()
        return ib if os.path.isdir(ib) else ''

    def _ensure(self):
        """قبل ما نحط ملفات: البيانات الأساسية لازم تتحفظ (عشان اسم الفولدر)"""
        if self.new: return self.save(close=False)
        return True

    def pick_paper(self):
        if not self._ensure(): return
        f, _ = QFileDialog.getOpenFileName(self, 'The deduction paper (picture or PDF)', self._start(),
                                           'Paper (*.pdf *.jpg *.jpeg *.png *.bmp *.webp)')
        if not f: return
        try: S.ded_add_paper(self.d, f)
        except Exception as x: return QMessageBox.warning(self, 'Deductions', f'Could not add it:\n{x}')
        self._save_folder(); self.fill_files()

    def open_paper(self):
        if self.paper: _open(self.paper)

    def add_photos(self):
        if not self._ensure(): return
        fs, _ = QFileDialog.getOpenFileNames(self, 'Photos (pictures or PDF)', self._start(), 'Photos (*.jpg *.jpeg *.png *.bmp *.webp *.pdf)')
        if not fs: return
        try: S.ded_add_photos(self.d, fs)
        except Exception as x: QMessageBox.warning(self, 'Deductions', f'Could not add them all:\n{x}')
        self._save_folder(); self.fill_files()

    def del_photo(self):
        it = self.ph.currentItem()
        if not it: return QMessageBox.information(self, 'Deductions', 'Select a photo first.')
        if QMessageBox.question(self, 'Deductions', f'Take out {it.text()}? (it moves to _OLD in the folder)') != QMessageBox.Yes: return
        try: S.ded_del_photo(self.d, it.data(Qt.UserRole))
        except Exception as x: QMessageBox.warning(self, 'Deductions', str(x))
        self.fill_files()

    def _save_folder(self):
        try: self.reg.save()
        except Exception as x: QMessageBox.warning(self, 'Deductions', f'The register was not saved:\n{x}')

    def _read(self):
        qv = _num(self.e_qty.text()) if self.e_qty.text().strip() else None
        rate = _num(self.e_rate.text()) if self.e_rate.text().strip() else None
        amt = _num(self.e_amt.text()) if self.e_amt.text().strip() else None
        for txt, val, w in (('Qty', qv, self.e_qty), ('Rate', rate, self.e_rate), ('Amount', amt, self.e_amt)):
            if w.text().strip() and val is None: raise ValueError(f'{txt}: not a number')
        if not self.e_desc.text().strip(): raise ValueError('Write what it is (in Arabic).')
        if self.cb_type.currentText() == 'Violation' and not self.cb_src.currentText(): raise ValueError('Choose the source of the violation.')
        qd = self.de.date()
        sid = self.cb_sub.currentData()
        st = self.d.get('Status') or 'Open'
        if st in ('Open', 'Company'): st = 'Open' if sid else 'Company'
        return {'Type': self.cb_type.currentText(), 'Source': self.cb_src.currentText() or None, 'Ref No.': self.e_ref.text().strip() or None,
                'Date': dt.date(qd.year(), qd.month(), qd.day()), 'Description': self.e_desc.text().strip(),
                'Place': self.e_place.text().strip() or None, 'Qty': qv, 'Unit': self.e_unit.text().strip() or None, 'Rate': rate,
                'Amount': amt if amt is not None else (round(qv * rate, 2) if qv is not None and rate is not None else None),
                'Sub ID': sid, 'Status': st, 'Notes': self.e_notes.text().strip() or None}

    def save(self, close=True):
        try: vals = self._read()
        except ValueError as x:
            QMessageBox.warning(self, 'Deductions', str(x)); return False
        old = dict(self.d)
        self.d.update(vals)
        if self.new: self.reg.data.setdefault('DEDUCTIONS', []).append(self.d)
        try:
            S.ded_dir(self.d, make=True)
            self.reg.save()
        except Exception as x:
            if self.new: self.reg.data['DEDUCTIONS'].remove(self.d)
            self.d.clear(); self.d.update(old)
            QMessageBox.warning(self, 'Deductions', f'Not saved:\n{x}\n\n(close 10.1-SUBCONTRACTORS.xlsx if it is open in Excel)')
            return False
        self.new = False
        self.b_cancel_d.setEnabled(True)
        if getattr(self, 'memo_open', False):              # مبلغ المذكرة = إجمالي خصوماتها
            no = self.d['Memo']
            tot = round(sum(S.ded_amount(x) for x in self.reg.data.get('DEDUCTIONS', [])
                            if x.get('Sub ID') == self.d.get('Sub ID') and str(x.get('Memo') or '').upper() == str(no).upper()
                            and x.get('Status') != 'Cancelled'), 2)
            if tot > 0:
                try: S.set_memo_amount(self.reg.sub(self.d['Sub ID']), no, tot)
                except Exception as x:
                    QMessageBox.warning(self, 'Deductions', f'Saved – but the amount of {no} in 10-LOG SUB was not updated:\n{x}\n\n'
                                                            f'(close 10-LOG SUB if it is open in Excel, then press Save again)')
        if close: self.accept()
        return True

    def cancel_ded(self):
        if QMessageBox.question(self, 'Deductions', f"Cancel {self.d['ID']}? It stays in the register as «Cancelled».") != QMessageBox.Yes: return
        self.d['Status'] = 'Cancelled'
        try: self.reg.save()
        except Exception as x: return QMessageBox.warning(self, 'Deductions', str(x))
        self.accept()


# ---------------------------------------------------------------- جدول الخصومات (لكل المقاولين أو لمقاول واحد)
COLS = [('id', 'NO.', 70), ('date', 'DATE', 95), ('type', 'TYPE', 170), ('ref', 'REF', 110), ('desc', 'WHAT', 330), ('place', 'PLACE', 140),
        ('amt', 'AMOUNT', 110), ('sub', 'CHARGED TO', 170), ('st', 'STATUS', 160), ('ph', 'FILES', 60)]


def ded_rows(reg, sid=None):
    rows = [d for d in reg.data.get('DEDUCTIONS', []) if sid is None or d.get('Sub ID') == sid]
    return sorted(rows, key=lambda d: (S.todate(d.get('Date')) or dt.date.min, str(d.get('ID'))), reverse=True)


def fill_table(tb, reg, rows):
    names = {s['ID']: str(s.get('Folder') or s['Name']).replace('_', ' ') for s in reg.subs()}
    tb.setRowCount(len(rows))
    colors = {'Open': '#F0C860', 'In memo': '#7FD98A', 'Company': '#9CC8F0', 'Cancelled': '#7C8896'}
    for r, d in enumerate(rows):
        paper, ph = S.ded_files(d)
        st = d.get('Status') or 'Open'
        st_txt = st + (f"  {d['Memo']}" if d.get('Memo') else '') + (f"  → {d['Cert']}" if d.get('Cert') else '')
        t = (d.get('Type') or '') + (f" · {d['Source']}" if d.get('Source') else '')
        vals = (d.get('ID'), _fd(d.get('Date')), t, d.get('Ref No.') or '', d.get('Description') or '', d.get('Place') or '',
                _money(S.ded_amount(d)), names.get(d.get('Sub ID'), '— company'), st_txt, str(len(ph) + (1 if paper else 0)))
        for c, val in enumerate(vals):
            it = QTableWidgetItem(str(val if val is not None else '')); it.setData(Qt.UserRole, d.get('ID'))
            if c == 8: it.setForeground(QColor(colors.get(st, '#E3EAF1')))
            if c in (6,): it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            tb.setItem(r, c, it)


class DedList(QDialog):
    """كل الخصومات (كل الجهات وكل المقاولين وحتى اللي على الشركة)"""

    def __init__(self, parent, reg, css, table_fn, btn):
        super().__init__(parent, Qt.Window)
        self.reg, self.css, self.btn = reg, css, btn
        self.setWindowTitle('Deductions  ·  violations, company equipment, materials …'); self.setStyleSheet(css); self.setObjectName('subq')
        self.setAttribute(Qt.WA_StyledBackground, True); self.resize(1400, 780)
        v = QVBoxLayout(self); v.setContentsMargins(18, 14, 18, 14); v.setSpacing(8)
        top = QHBoxLayout(); top.setSpacing(8)
        t = QLabel('Deductions'); f = QFont(); f.setPointSize(18); f.setBold(True); t.setFont(f); top.addWidget(t)
        self.f_sub = QComboBox(); self.f_sub.addItem('All subcontractors', '*'); self.f_sub.addItem('— the company', None)
        for s in reg.subs(): self.f_sub.addItem(str(s.get('Folder') or s['Name']).replace('_', ' '), s['ID'])
        self.f_type = QComboBox(); self.f_type.addItems(['All types'] + list(S.DED_TYPES))
        self.f_st = QComboBox(); self.f_st.addItems(['All statuses'] + list(S.DED_STATUSES))
        for w in (self.f_sub, self.f_type, self.f_st): w.currentIndexChanged.connect(lambda *_: self.fill()); top.addWidget(w)
        top.addStretch(1)
        top.addWidget(btn('+ Deduction', self.add, True)); top.addWidget(btn('Open', self.open_sel))
        top.addWidget(btn('Show in folder', self.show_sel))
        v.addLayout(top)
        self.tb = table_fn(COLS); self.tb.cellDoubleClicked.connect(lambda r, c: self.open_sel())
        v.addWidget(self.tb, 1)
        self.l_tot = QLabel(''); v.addWidget(self.l_tot)
        self.fill()

    def rows(self):
        rs = ded_rows(self.reg)
        fs = self.f_sub.currentData()
        if fs != '*': rs = [d for d in rs if d.get('Sub ID') == fs]
        if self.f_type.currentIndex() > 0: rs = [d for d in rs if d.get('Type') == self.f_type.currentText()]
        if self.f_st.currentIndex() > 0: rs = [d for d in rs if (d.get('Status') or 'Open') == self.f_st.currentText()]
        return rs

    def fill(self):
        self._rows = self.rows()
        fill_table(self.tb, self.reg, self._rows)
        op = [d for d in self._rows if (d.get('Status') or 'Open') == 'Open']
        self.l_tot.setText(f"{len(self._rows)} deductions  ·  open (not deducted yet): {len(op)}  ·  {_money(sum(S.ded_amount(d) for d in op))} SAR")

    def add(self):
        fs = self.f_sub.currentData()
        if DedDlg(self, self.reg, None, fs if fs not in ('*', None) else None, self.css, self.btn).exec() or True: self.fill()

    def show_sel(self):
        r = self.tb.currentRow()
        if r < 0 or r >= len(self._rows): return QMessageBox.information(self, 'Deductions', 'Select a deduction first.')
        p = S.ded_dir(self._rows[r])
        if not p or not os.path.isdir(p): return QMessageBox.information(self, 'Deductions', 'This deduction has no paper or photos yet.')
        _open(p)

    def open_sel(self):
        r = self.tb.currentRow()
        if r < 0 or r >= len(self._rows): return
        DedDlg(self, self.reg, self._rows[r], None, self.css, self.btn).exec()
        self.fill()


# ---------------------------------------------------------------- اختيار الخصومات في المستخلص
def pick_dialog(parent, css, btn, memos_av, ded_av, memo_sel, ded_sel, label):
    """مذكرات المقاول المفتوحة + خصوماته المفتوحة في قايمة واحدة – اللي يتعلّم عليه يتخصم في المستخلص ده.
    يرجّع (memo_nos, ded_ids) أو None"""
    d = QDialog(parent); d.setWindowTitle(f'Deductions in {label}'); d.setStyleSheet(css); d.setMinimumSize(900, 520)
    v = QVBoxLayout(d)
    v.addWidget(QLabel('Tick what is deducted in this certificate – the rest waits for the next one.\n'
                       'The ticked deductions (violations, equipment, materials …) go in ONE new deduction memo when you save.'))
    lw = QListWidget()
    for m in memos_av:
        it = QListWidgetItem(f"Memo  {m['no']}    {_fd(m['date'])}    {m['subject'][:60]}    {m['amount']:,.2f} SAR")
        it.setData(Qt.UserRole, ('m', m['no'])); it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
        it.setCheckState(Qt.Checked if m['no'] in memo_sel else Qt.Unchecked); lw.addItem(it)
    for x in ded_av:
        it = QListWidgetItem(f"{x['ID']}    {_fd(x.get('Date'))}    {S.ded_type_ar(x)}    {x.get('Ref No.') or ''}    "
                             f"{str(x.get('Description') or '')[:50]}    {_money(S.ded_amount(x))} SAR")
        it.setData(Qt.UserRole, ('d', x['ID'])); it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
        it.setCheckState(Qt.Checked if x['ID'] in ded_sel else Qt.Unchecked); lw.addItem(it)
    if not lw.count(): lw.addItem('No open deductions or memos for this subcontractor.')
    v.addWidget(lw, 1)
    out = {}

    def ok():
        ms, ds = [], []
        for i in range(lw.count()):
            it = lw.item(i); k = it.data(Qt.UserRole)
            if k and it.checkState() == Qt.Checked: (ms if k[0] == 'm' else ds).append(k[1])
        out['r'] = (ms, ds); d.accept()
    hb = QHBoxLayout(); hb.addStretch(1); hb.addWidget(btn('OK', ok, True)); hb.addWidget(btn('Cancel', d.reject)); v.addLayout(hb)
    d.exec()
    return out.get('r')
