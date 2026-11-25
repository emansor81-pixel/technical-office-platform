# -*- coding: utf-8 -*-
"""Subcontractors | مقاولين الباطن – المقاول وعقده وبنوده والاستيشنات اللي اشتغل فيها، والمستخلصات.
الكميات من حصر الريكويستات (نفس حسابات QS) جوه استيشنات المقاول لكل بند، وبأسعار عقده.
المستخلص على شكل مستخلصات مقاولين مشروع الصرف (<SAMPLE_FOLDER>):
  غلاف + بيان المتابعة + المستخلص (V) + الحصر (Q1 حفر وفرشة ومواسير وردم وقاعدة طريق، Q2 أسفلت).
السجل: 00_Registers\\10.1-SUBCONTRACTORS.xlsx (جنب 10-LOG SUB) – وفولدر لكل مقاول باسمه الإنجليزي في 08_Cost_Payments\\SUBCONTRACTORS.
مذكرات الخصم: DM-001، DM-002 … لكل مقاول لوحده – ملفها في فولدره\\DEDUCTION_MEMOS وسطرها في 00_Registers\\10-LOG SUB.xlsx
(شيت DEDUCTION_MEMOS)، عشان لو البرنامج اتشال اللي بعدك يلاقي كل حاجة من الريجسترات."""
import os, re, sys, glob, math, shutil, tempfile, zipfile, calendar, datetime as dt
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from copy import copy
from urllib.parse import quote

HERE = os.path.dirname(os.path.abspath(__file__))
for _d in ('qs', 'ipc'):
    _p = os.path.join(os.path.dirname(HERE), _d)
    if _p not in sys.path: sys.path.insert(0, _p)
import qs_core as Q  # noqa: E402
_PLAT = os.path.dirname(os.path.dirname(HERE))                       # <PLATFORM> (فيه parties.py)
if _PLAT not in sys.path: sys.path.insert(0, _PLAT)
import parties as PT  # noqa: E402


def proj():
    """فولدر المشروع (نفس اللي qs_core بيحسب منه المسارات)."""
    return os.path.dirname(Q.project_path('_'))

ROOT_REL = PT.ROOT_REL                                                 # 13_PARTIES\SUB\NN-NAME (طلب المستخدم 12-11)
OLD_ROOT_REL = r'08_Cost_Payments\SUBCONTRACTORS'                      # مكان فولدرات المقاولين قبل كده
REG_REL = r'00_Registers\10.1-SUBCONTRACTORS.xlsx'                     # بيانات البرنامج للمقاولين – في الريجسترات جنب 10-LOG SUB
OLD_REG_REL = r'08_Cost_Payments\SUBCONTRACTORS\SUBCONTRACTORS.xlsx'   # مكانها القديم – بتتنقل لوحدها
LOG_REL = r'00_Registers\10-LOG SUB.xlsx'                             # سجل مقاولي الباطن – فيه شيت مذكرات الخصم
DM_SHEET = 'DEDUCTION_MEMOS'
INBOX_REL = '00_INBOX'
SUB_SUBMITTALS = r'04_Submittals\SUB'
OWNER_IPC_REL = r'08_Cost_Payments\CLIENT_IPC'
SUB_DIRS = (PT.CONTRACT, PT.CERTS, PT.DMEMOS, PT.AMEMOS, PT.FINAL)      # 03_CONTRACT، 04_CERTIFICATES، 06_… 07_… 08_…
OLD_DIRS = ('CONTRACT', 'CERTIFICATES', 'DEDUCTION_MEMOS', 'ADDITION_MEMOS', 'FINAL_ACCOUNT')   # الأسامي القديمة – بتتغير لوحدها
OLDER_DIRS = ('1-Contract', '2-Certificates', '3-Deduction Memos', '4-Addition Memos', '5-Final Account')
LOGO = os.path.join(HERE, 'assets', 'logo.jpeg')
LOGO_PX = (380, 243)
COMPANY = 'شركة المقاول'
PROJECT = 'مشروع شبكات المياه (مثال) ( أمر شراء 000000000 )'

SHEETS = {
    'SUBS': ['ID', 'Name', 'CR No.', 'VAT No.', 'IBAN', 'Representative', 'Contract No.', 'Contract Date', 'Contract Type',
             'Site Handover', 'Duration (months)', 'Retention %', 'VAT %', 'Advance %', 'Status', 'Folder', 'Notes'],
    'ITEMS': ['Sub ID', 'Item No.', 'Section', 'Description', 'Unit', 'Contract Qty', 'Rate', 'Measure', 'Dia',
              'Depth From', 'Depth To', 'Notes'],
    'AREAS': ['Sub ID', 'Item No.', 'Line', 'From', 'To', 'Notes'],
    'CERTS': ['Sub ID', 'Cert No.', 'Final', 'Period From', 'Period To', 'Basis', 'Works To Date', 'Previous', 'Current',
              'VAT', 'Total incl. VAT', 'Retention', 'Additions', 'Additions Ref', 'Deductions', 'Deductions Ref',
              'Net Payable', 'Progress %', 'Owner IPC', 'Invoice No.', 'Invoice Date', 'Invoice Amount', 'Status',
              'Saved At', 'Folder'],
    'CERT_QTY': ['Sub ID', 'Cert No.', 'Item No.', 'Qty To Date', 'Rate', 'Value To Date'],
    'SIGN': ['Role', 'Name'],
    # استثناءات المستخلص (طلب المستخدم 17-11): بند معين على استيشنات غير باقي البنود، أو كمية تزيد/تنقص – بالسبب
    'EXCEPTIONS': ['Sub ID', 'Cert No.', 'Item No.', 'Line', 'From', 'To', 'Qty', 'Reason'],
    # الخصومات على المقاولين (طلب المستخدم 18-11): مخالفات الجهات + معدات الشركة + مواد + غيره – بالورقة والصور
    'DEDUCTIONS': ['ID', 'Date', 'Type', 'Source', 'Ref No.', 'Description', 'Place', 'Qty', 'Unit', 'Rate', 'Amount',
                   'Sub ID', 'Status', 'Memo', 'Cert', 'Folder', 'Notes'],
}
DATE_COLS = {'Contract Date', 'Site Handover', 'Period From', 'Period To', 'Invoice Date', 'Saved At'}
SIGN_DEFAULT = [('مندوب مقاول الباطن', ''), ('مهندس مكتب فني', 'مهندس المكتب الفني'), ('مدير المشروع', 'م. (اسم المهندس)'),
                ('مدير الإدارة الفنية', 'م. (اسم المهندس)'), ('الإدارة المعنية', ''), ('مدير إدارة المتابعة', 'أ. (الاسم)'),
                ('مدير قطاع المقاولات', 'م. (اسم المهندس)')]
STATUSES = ('Draft', 'Submitted', 'Approved', 'Paid')

# طريقة حساب كمية البند – من صفوف حصر QS جوه استيشنات المقاول
MEASURES = [
    ('EXC1', 'Excavation 0–1.5 m (m3)'), ('EXC2', 'Excavation 1.5–2.5 m (m3)'), ('EXC3', 'Excavation 2.5–3.5 m (m3)'),
    ('EXC4', 'Excavation deeper than 3.5 m (m3)'), ('EXC', 'Excavation – all depths (m3)'),
    ('TRL', 'Trench length excavated (m)'), ('BED', 'Bedding & initial backfill (m3)'),
    ('PIP', 'Pipe laying – paid length after % (m)'), ('PIPL', 'Pipe laying – full laid length (m)'),
    ('BF', 'Backfill by approved layers (m3)'), ('RB', 'Road base RB-1 + RB-2 (m3)'),
    ('ASP1', 'Asphalt – normal (m3)'), ('ASP2', 'Asphalt – paver & milling (m3)'), ('MAN', 'Manual – typed in each certificate'),
]
MEASURE_NAME = dict(MEASURES)
ASPHALT = {'ASP1', 'ASP2'}
DEPTH_FILTER = {'EXC', 'TRL', 'PIPL', 'PIP'}          # فلتر العمق (متوسط عمق الخندق) – البنود اللي بالعمق
DIA_FILTER = {'EXC1', 'EXC2', 'EXC3', 'EXC4', 'EXC', 'TRL', 'BED', 'PIP', 'PIPL', 'BF', 'RB', 'ASP1', 'ASP2'}


# ---------------------------------------------------------------- مساعدات
def num(v): return Q.num(v)


def todate(v): return Q.todate(v)


def fdate(d): return d.strftime('%d-%m-%Y') if isinstance(d, (dt.date, dt.datetime)) else (d or '')


def parse_sta(s):
    """«0+350» أو «350» أو «1+025.5» ← متر"""
    if s is None: return None
    if isinstance(s, (int, float)): return float(s)
    t = str(s).strip().replace(' ', '')
    if not t: return None
    m = re.fullmatch(r'(\d+)\+(\d+(?:\.\d+)?)', t)
    if m: return int(m.group(1)) * 1000 + float(m.group(2))
    try: return float(t)
    except ValueError: return None


def sta(x): return Q.sta(x)


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]+', ' ', str(s or '')).strip().rstrip('.')


def add_months(d, n):
    if not d or not n: return None
    m = d.month - 1 + int(n); y = d.year + m // 12; m = m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def root_dir(): return Q.project_path(ROOT_REL)


def reg_path():
    """بيانات البرنامج للمقاولين في 00_Registers. لو لسه في المكان القديم بتتنقل لوحدها (ولو الملف مفتوح بنقرا من القديم)."""
    new, old = Q.project_path(REG_REL), Q.project_path(OLD_REG_REL)
    if not os.path.exists(new) and os.path.isfile(old):
        try:
            os.makedirs(os.path.dirname(new), exist_ok=True)
            os.replace(old, new)
        except OSError:
            return old
    return new


def submittal_names():
    """أسماء الجهات اللي ليها فولدر في 13_PARTIES (للاختيار عند إضافة مقاول) – الإنجليزي وبعده العربي"""
    out = []
    for p in PT.all_parties(proj()):
        out += [p['name']] + p['aliases']
    if out: return list(dict.fromkeys(out))
    root = Q.project_path(SUB_SUBMITTALS)
    for st in ('1_WIP', '2_SHARED', '3_PUBLISHED', '4_ARCHIVED'):
        d = os.path.join(root, st)
        if not os.path.isdir(d): continue
        for n in sorted(os.listdir(d)):
            if os.path.isdir(os.path.join(d, n)):
                m = re.match(r'\s*\d+\s*-\s*(.*)', n)
                out.append((m.group(1) if m else n).strip(' -'))
    return list(dict.fromkeys(out))


def owner_ipcs():
    """مستخلصات المالك اللي اتعملت: [(رقم, نهاية الفترة)] من أسماء الملفات IPC-NN (dd-mm-yyyy)"""
    out = {}
    base = Q.project_path(OWNER_IPC_REL)
    for p in glob.glob(os.path.join(base, '**', 'IPC-*'), recursive=True):
        m = re.match(r'IPC-(\d+)\s*\((\d{2}-\d{2}-\d{4})\)', os.path.basename(p))
        if m: out[int(m.group(1))] = todate(m.group(2))
    return sorted(out.items())


# ---------------------------------------------------------------- السجل (إكسل)
class Register:
    """كل بيانات المقاولين في ملف واحد – كل شيت جدول بعناوين في الصف الأول."""

    def __init__(self, path=None):
        self.path = path or reg_path()
        self.data = {k: [] for k in SHEETS}
        if os.path.exists(self.path): self.load()
        if not self.data['SIGN']: self.data['SIGN'] = [dict(Role=r, Name=n) for r, n in SIGN_DEFAULT]

    def load(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        for sh, hdr in SHEETS.items():
            if sh not in wb.sheetnames: continue
            ws = wb[sh]
            rows = list(ws.iter_rows(values_only=True))
            if not rows: continue
            head = [str(h).strip() if h is not None else '' for h in rows[0]]
            for r in rows[1:]:
                if not r or all(v in (None, '') for v in r): continue
                d = {h: (r[i] if i < len(r) else None) for i, h in enumerate(head) if h}
                for h in hdr: d.setdefault(h, None)
                for h in DATE_COLS & set(d):
                    if d[h] is not None: d[h] = todate(d[h]) or d[h]
                self.data[sh].append(d)
        wb.close()

    def save(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        hf, hfill = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='1F4E78')
        for sh, hdr in SHEETS.items():
            ws = wb.create_sheet(sh)
            for j, h in enumerate(hdr, 1):
                c = ws.cell(1, j, h); c.font = hf; c.fill = hfill
                ws.column_dimensions[c.column_letter].width = max(12, min(40, len(h) + 4))
            for i, d in enumerate(self.data[sh], 2):
                for j, h in enumerate(hdr, 1):
                    v = d.get(h)
                    c = ws.cell(i, j, v)
                    if h in DATE_COLS and isinstance(v, (dt.date, dt.datetime)): c.number_format = 'dd-mm-yyyy'
            ws.freeze_panes = 'A2'
        fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(self.path)); os.close(fd)
        try:
            wb.save(tmp)
            try:
                os.replace(tmp, self.path)
            except PermissionError:
                raise PermissionError(f'{os.path.basename(self.path)} is open in Excel – close it and try again')
        finally:
            if os.path.exists(tmp):
                try: os.remove(tmp)
                except OSError: pass

    def writable(self):
        """قبل ما نغيّر حاجة برّه السجل (زي اسم فولدر): السجل نفسه مش مفتوح في إكسل؟"""
        if not os.path.exists(self.path): return True
        try:
            with open(self.path, 'r+b'): return True
        except PermissionError:
            raise PermissionError(f'{os.path.basename(self.path)} is open in Excel – close it and try again')

    # ---------- مقاولين
    def subs(self): return self.data['SUBS']

    def sub(self, sid): return next((s for s in self.data['SUBS'] if s['ID'] == sid), None)

    def new_id(self):
        n = max([int(re.sub(r'\D', '', str(s['ID'])) or 0) for s in self.data['SUBS']] or [0]) + 1
        return f'S{n:02d}'

    def items(self, sid):
        its = [d for d in self.data['ITEMS'] if d['Sub ID'] == sid]
        return sorted(its, key=lambda d: item_key(d['Item No.']))

    def areas(self, sid, item=None):
        return [d for d in self.data['AREAS'] if d['Sub ID'] == sid and (item is None or str(d['Item No.']) == str(item))]

    def certs(self, sid):
        return sorted([d for d in self.data['CERTS'] if d['Sub ID'] == sid], key=lambda d: int(d['Cert No.'] or 0))

    def cert_qty(self, sid, no):
        return {str(d['Item No.']): d for d in self.data['CERT_QTY'] if d['Sub ID'] == sid and int(d['Cert No.'] or 0) == int(no)}

    def signs(self): return [(d['Role'], d['Name'] or '') for d in self.data['SIGN']]

    def sign(self, role): return next((n for r, n in self.signs() if r == role), '')

    def delete_sub(self, sid):
        for sh in ('SUBS', 'ITEMS', 'AREAS', 'CERTS', 'CERT_QTY'):
            k = 'ID' if sh == 'SUBS' else 'Sub ID'
            self.data[sh] = [d for d in self.data[sh] if d[k] != sid]


def from_contract(name, en, no, date, info=None, items=None):
    """عقد مقاول اتوزع (Distribute ← Approve) ← المقاول في سجل SUB ببيانات العقد وبنوده.
    مقاول موجود (بالاسم الإنجليزي أو العربي) ← نكمّل الناقص بس. البنود تتحط لو المقاول ملوش بنود (مابنغيّرش بنود اتعدلت).
    يرجّع (ID, عدد البنود اللي اتضافت, اتعمل جديد؟)"""
    info = info or {}
    reg = Register()
    reg.writable()
    nm = str(info.get('name_ar') or name or '').strip()
    nm = re.sub(r'\s*\([^)]*\)\s*$', '', nm).strip() or str(en)
    sub = next((s for s in reg.subs() if _same(s.get('Folder'), en) or _same(s.get('Name'), nm)), None)
    created = sub is None
    if created:
        sub = {h: None for h in SHEETS['SUBS']}
        sub.update({'ID': reg.new_id(), 'Name': nm, 'Folder': en, 'Status': 'Active', 'Retention %': 15, 'VAT %': 15, 'Advance %': 0})
        reg.data['SUBS'].append(sub)
    vals = {'Contract No.': str(no or '').strip() or None, 'Contract Date': todate(date) if date else None,
            'CR No.': info.get('cr'), 'VAT No.': info.get('vat_no'), 'Representative': info.get('representative'),
            'Contract Type': info.get('contract_type'), 'Duration (months)': num(info.get('duration_months')),
            'Retention %': num(info.get('retention_pct')), 'VAT %': num(info.get('vat_pct')), 'Advance %': num(info.get('advance_pct'))}
    for k, v in vals.items():
        if v in (None, ''): continue
        if created or sub.get(k) in (None, ''): sub[k] = v
        elif k == 'Contract No.' and not _same(sub.get(k), v): sub['Notes'] = ' | '.join(x for x in (sub.get('Notes'), f'Contract {v}') if x)
    added = 0
    if items and not reg.items(sub['ID']):
        for it in items:
            if not it.get('no') or not it.get('description'): continue
            m = str(it.get('measure') or 'MAN').upper()
            reg.data['ITEMS'].append({'Sub ID': sub['ID'], 'Item No.': str(it['no']).strip(), 'Section': it.get('section'),
                                      'Description': it['description'], 'Unit': it.get('unit'), 'Contract Qty': num(it.get('qty')),
                                      'Rate': num(it.get('rate')), 'Measure': m if m in MEASURE_NAME else 'MAN',
                                      'Dia': int(num(it['dia'])) if num(it.get('dia')) else None,
                                      'Depth From': num(it.get('depth_from')), 'Depth To': num(it.get('depth_to')),
                                      'Notes': it.get('notes')})
            added += 1
    reg.save()
    try: sub_folder(sub)
    except OSError: pass
    return sub['ID'], added, created


def item_key(no):
    """ترتيب أرقام البنود: 1، 2، 10، 2.1 ..."""
    parts = re.findall(r'\d+|[^\d.]+', str(no or ''))
    return [(0, int(p)) if p.isdigit() else (1, p) for p in parts]


def sub_dir(sub):
    """فولدر المقاول في 13_PARTIES باسمه الإنجليزي (عمود Folder) أو العربي، ولو مش موجود ← SUB\\NN-NAME الجديد"""
    pr = proj()
    for n in (sub.get('Folder'), sub.get('Name')):
        p = PT.find(pr, n) if n else None
        if p: return p['path']
    return os.path.join(pr, PT.new_rel(pr, sub.get('Folder') or sub['Name'], 'SUB'))


def sub_folder(sub):
    """فولدر المقاول + الفولدرات الفرعية (CONTRACT، CERTIFICATES، DEDUCTION_MEMOS …). الأسامي القديمة (1-Contract …) بتتغير لوحدها."""
    d = sub_dir(sub)
    PT.ensure(d, arabic=sub['Name'] if re.search(r'[\u0600-\u06FF]', str(sub.get('Name') or '')) else '')
    for old, new in zip(OLDER_DIRS, SUB_DIRS):
        po, pn = os.path.join(d, old), os.path.join(d, new)
        if not os.path.isdir(po): continue
        try:
            if not os.path.exists(pn): os.rename(po, pn)
            else: _merge_dir(po, pn)
        except OSError:
            pass                                          # ملف جواه مفتوح – بنكمّل المرة الجاية
    for old, new in zip(OLD_DIRS, SUB_DIRS):
        os.makedirs(os.path.join(d, new), exist_ok=True)
    return d


def _merge_dir(po, pn):
    """القديم والجديد الاتنين موجودين: كل اللي في القديم يتنقل للجديد (اللي اسمه موجود هناك يفضل مكانه)، ولو فضي يتشال"""
    for n in os.listdir(po):
        a, b = os.path.join(po, n), os.path.join(pn, n)
        if not os.path.exists(b):
            try: os.rename(a, b)
            except OSError: pass
    try: os.rmdir(po)
    except OSError: pass


def live_path(p):
    """مسار اتسجل قبل تغيير أسامي الفولدرات (1-Contract …) ← مكانه دلوقتي"""
    if not p or os.path.exists(p): return p
    q = str(p)
    m = re.search(r'[\\/]08_Cost_Payments[\\/]SUBCONTRACTORS[\\/]([^\\/]+)', q)    # المكان القديم ← فولدره في 13_PARTIES
    if m:
        pp = PT.find(proj(), m.group(1))
        if pp: q = pp['path'] + q[m.end():]
    for old, new in list(zip(OLD_DIRS, SUB_DIRS)) + list(zip(OLDER_DIRS, SUB_DIRS)):
        for sep in ('\\', '/'):
            q = q.replace(sep + old + sep, sep + new + sep)
            if q.endswith(sep + old): q = q[:-len(old)] + new
    return q


def rename_sub_dir(reg, sid, old, new):
    """الاسم الإنجليزي اتغير: فولدر المقاول ياخد الاسم الجديد (بنفس رقمه ونوعه)، ومسارات مستخلصاته المتسجلة تتظبط.
    بيرجّع (القديم، الجديد)."""
    pp = PT.find(proj(), old)
    if pp:
        po = pp['path']; pn = os.path.join(os.path.dirname(po), f"{pp['no']:02d}-{PT.en_folder(new) or new}")
    else:
        po = pn = os.path.join(proj(), PT.new_rel(proj(), new, 'SUB'))
    if os.path.isdir(po):
        if os.path.exists(pn) and os.path.normcase(os.path.abspath(po)) != os.path.normcase(os.path.abspath(pn)):
            raise FileExistsError(f'There is already a folder named «{new}»')
        os.rename(po, pn)
    for c in reg.data['CERTS']:
        f = str(c.get('Folder') or '')
        if c['Sub ID'] == sid and (f == po or f.startswith(po + os.sep)):
            c['Folder'] = pn + f[len(po):]
    return po, pn


# ---------------------------------------------------------------- الكميات
def measure_rows(items, code, dia=None, d0=None, d1=None):
    """صفوف حصر QS اللي بتدخل في البند ده: [(صف, الكمية الكاملة للصف)] – قبل القص على الاستيشنات"""
    def okd(r):
        if not dia: return True
        v = r.get('dia') if r.get('dia') is not None else r['q'].get('dia')
        try: return int(float(v)) == int(float(dia))
        except (TypeError, ValueError): return False

    def okdepth(avg):
        if d0 is None and d1 is None: return True
        if avg is None: return False
        return (d0 is None or avg > d0 + 1e-9) and (d1 is None or avg <= d1 + 1e-9)
    exc_ok = [r for r in items['exc'] if r['status'] == 'Approved']
    if code in ('EXC1', 'EXC2', 'EXC3', 'EXC4'):
        k = int(code[-1]) - 1
        return [(r, r['v'][k] or 0.0) for r in exc_ok if okd(r)]
    if code == 'EXC': return [(r, r['total'] or 0.0) for r in exc_ok if okd(r) and okdepth(r['avg'])]
    if code == 'TRL': return [(r, r['L']) for r in exc_ok if okd(r) and okdepth(r['avg'])]
    if code == 'BED': return [(r, r['total'] or 0.0) for r in items['bed'] if r['status'] == 'Approved' and okd(r)]
    if code in ('PIP', 'PIPL'):
        avg = _avg_depth(items)
        f = 'paid' if code == 'PIP' else 'L'
        return [(r, r[f] or 0.0) for r in items['pip'] if r['status'] == 'Approved' and okd(r)
                and okdepth(avg(r['q']['key'], r['x0'], r['x1']))]
    if code == 'BF': return [(r, r['paid'] or 0.0) for r in items.get('bf', []) if r.get('paid') and okd(r)]
    if code == 'RB': return [(r, r['qty'] or 0.0) for r in items.get('rb', []) if r.get('qty') and okd(r)]
    if code == 'ASP1': return [(r, r['q1'] or 0.0) for r in items.get('asp', []) if r.get('q1') and okd(r)]
    if code == 'ASP2': return [(r, r['q2'] or 0.0) for r in items.get('asp', []) if r.get('q2') and okd(r)]
    return []


def _avg_depth(items):
    """متوسط عمق الخندق عند أي جزء (من صفوف الحفر) – لفلتر العمق في بنود المواسير"""
    by = {}
    for r in items['exc']:
        if r['avg'] is not None: by.setdefault(r['q']['key'], []).append((r['x0'], r['x1'], r['avg']))

    def f(key, x0, x1):
        tot = w = 0.0
        for a, b, d in by.get(key, []):
            o = Q.overlap(a, b, x0, x1)
            if o > Q.EPS: tot += o * d; w += o
        return tot / w if w else None
    return f


def merged_areas(areas):
    """{الخط: [(من, إلى)]} – المدى المتكرر/المتداخل بيتحسب مرة واحدة"""
    by = {}
    for a in areas:
        x0, x1 = parse_sta(a['From']), parse_sta(a['To'])
        if x0 is None or x1 is None: continue
        x0, x1 = sorted((x0, x1))
        by.setdefault(Q.kstr(a['Line']), []).append((x0, x1))
    out = {}
    for ln, ivs in by.items():
        ivs.sort(); m = [list(ivs[0])]
        for a, b in ivs[1:]:
            if a <= m[-1][1] + Q.EPS: m[-1][1] = max(m[-1][1], b)
            else: m.append([a, b])
        out[ln] = [tuple(x) for x in m]
    return out


def item_quantity(items, it, areas):
    """كمية البند جوه استيشنات المقاول: (الكمية, صفوف الحصر) – كل صف بالجزء اللي جوه المدى بس (بالتناسب مع الطول)"""
    code = str(it.get('Measure') or 'MAN')
    if code == 'MAN': return None, []
    ar = merged_areas(areas)
    rows, total = [], 0.0
    avgf = None
    for r, qf in measure_rows(items, code, num(it.get('Dia')), num(it.get('Depth From')), num(it.get('Depth To'))):
        q = r['q']; L = r['x1'] - r['x0']
        if L <= Q.EPS: continue
        for a, b in ar.get(q['key'][0], []):
            o = Q.overlap(r['x0'], r['x1'], a, b)
            if o <= Q.EPS: continue
            qty = qf * o / L
            total += qty
            x0, x1 = max(r['x0'], a), min(r['x1'], b)
            avg = r.get('avg')
            if avg is None:                               # المواسير والفرشة: متوسط عمق الخندق من صفوف الحفر (المناسيب)
                avgf = avgf or _avg_depth(items)
                avg = avgf(q['key'], x0, x1)
            if code == 'PIP': pct = num(r.get('pct'))                                   # تركيب 85 / اختبار 90 / غسيل 100
            elif code == 'BF': pct = (num(r.get('pct')) or 0.0) * 100 if r.get('pct') is not None else None
            else: pct = 100.0
            rows.append(dict(wir=q['no'], unf=q.get('unf', ''), dist=q['dist'], street=q['street'], line=q['line'],
                             x0=x0, x1=x1, L=o, dia=q.get('dia_raw') or q.get('dia'),
                             avg=avg, pct=pct, qty=qty, appdate=q.get('appdate'), reqdate=q.get('reqdate')))
    rows.sort(key=lambda d: (Q.kstr(d['dist']), Q.kstr(d['line']), d['x0']))
    return total, rows


def project_items(cutoff, basis):
    """حصر QS بالتاريخ والأساس (approved / requested) – نفس اللي في برنامج الكميات"""
    p = Q.Project(Q.DEFAULT_INPUT)
    p.load_requests()
    p.mode = 'requested' if str(basis).lower().startswith('req') else 'approved'
    items, notes = p.build(cutoff)
    return p, items, notes


def lines_in_log():
    """أرقام الخطوط اللي في سجل الـ WIR (لاختيار الخط في الاستيشنات)"""
    p = Q.Project(Q.DEFAULT_INPUT)
    p.load_requests()
    return sorted({Q.kstr(q['line']) for q in p.reqs_all if Q.kstr(q['line'])}, key=item_key)


# ---------------------------------------------------------------- المستخلص
def compute(reg, sid, no, cutoff, basis, manual=None, additions=0.0, deductions=0.0, add_ref='', ded_ref='',
            period_from=None, final=False):
    """المستخلص رقم no: لكل بند السابق (من المستخلص اللي قبله) والإجمالي حتى cutoff والحالي، والخصومات بالترتيب."""
    sub = reg.sub(sid)
    its = reg.items(sid)
    p, items, notes = project_items(cutoff, basis)
    prev_q = reg.cert_qty(sid, no - 1) if no > 1 else {}
    prev_c = next((c for c in reg.certs(sid) if int(c['Cert No.']) == no - 1), None)
    manual = manual or {}
    rows, qrows = [], {}
    for it in its:
        k = str(it['Item No.'])
        rate = num(it['Rate']) or 0.0
        cq = num(it['Contract Qty']) or 0.0
        if str(it.get('Measure') or 'MAN') == 'MAN':
            tot = num(manual.get(k))
            if tot is None: tot = num(prev_q.get(k, {}).get('Qty To Date')) or 0.0
            qr = []
        else:
            tot, qr = item_quantity(items, it, reg.areas(sid, k))
        tot = round(tot or 0.0, 3)
        pq = round(num(prev_q.get(k, {}).get('Qty To Date')) or 0.0, 3)
        rows.append(dict(item=k, section=it.get('Section') or '', desc=it['Description'] or '', unit=it['Unit'] or '',
                         cqty=cq, rate=rate, cval=cq * rate, prev=pq, cur=round(tot - pq, 3), tot=tot, val=tot * rate,
                         measure=it.get('Measure') or 'MAN', over=(tot > cq + 1e-6) if cq else False))
        if qr: qrows[k] = qr
    vat_p = (num(sub.get('VAT %')) if sub.get('VAT %') is not None else 15.0) / 100
    ret_p = (num(sub.get('Retention %')) if sub.get('Retention %') is not None else 15.0) / 100
    gross = round(sum(r['val'] for r in rows), 2)
    prev = round(num(prev_c['Works To Date']) if prev_c else 0.0, 2)
    cur = round(gross - prev, 2)
    vat = round(cur * vat_p, 2)
    ret = round((cur + vat) * ret_p, 2)                # ضمان الأعمال على الأعمال الحالية بعد إضافة الضريبة
    additions, deductions = round(num(additions) or 0.0, 2), round(num(deductions) or 0.0, 2)
    net = round(cur + vat - ret - deductions + additions, 2)
    contract = round(sum(r['cval'] for r in rows), 2)
    # مستخلص المالك: كل الريكويستات اللي فيه اتعمدت ودخلت مستخلص مالك؟
    owner = owner_ipcs()
    apps = [d.get('appdate') for qr in qrows.values() for d in qr]
    inc_no = None
    if apps and all(apps):
        last = max(apps)
        inc_no = next((n for n, d in owner if d and d >= last), None)
    all_app = bool(apps) and all(apps)
    end = add_months(todate(sub.get('Site Handover')), num(sub.get('Duration (months)')))
    delay = max(0, (cutoff - end).days) if end else 0
    return dict(sub=sub, no=no, final=final, period_from=period_from, cutoff=cutoff, basis=basis, rows=rows, qrows=qrows,
                notes=notes, gross=gross, prev=prev, cur=cur, vat=vat, ret=ret, vat_p=vat_p, ret_p=ret_p,
                additions=additions, deductions=deductions, add_ref=add_ref, ded_ref=ded_ref, net=net, contract=contract,
                progress=(gross / contract if contract else 0.0), owner_ipc=inc_no, all_approved=all_app, delay=delay,
                end=end, signs=reg.signs())


def save_cert(reg, cm, folder):
    """المستخلص في السجل (يتبدل لو نفس الرقم) + كميات بنوده (السابق للمستخلص الجاي)"""
    sid, no = cm['sub']['ID'], cm['no']
    old = next((c for c in reg.data['CERTS'] if c['Sub ID'] == sid and int(c['Cert No.'] or 0) == no), None)
    reg.data['CERTS'] = [c for c in reg.data['CERTS'] if not (c['Sub ID'] == sid and int(c['Cert No.'] or 0) == no)]
    reg.data['CERT_QTY'] = [c for c in reg.data['CERT_QTY'] if not (c['Sub ID'] == sid and int(c['Cert No.'] or 0) == no)]
    reg.data['CERTS'].append({'Sub ID': sid, 'Cert No.': no, 'Final': 'Yes' if cm['final'] else '', 'Period From': cm['period_from'],
                              'Period To': cm['cutoff'], 'Basis': cm['basis'], 'Works To Date': cm['gross'], 'Previous': cm['prev'],
                              'Current': cm['cur'], 'VAT': cm['vat'], 'Total incl. VAT': round(cm['cur'] + cm['vat'], 2),
                              'Retention': cm['ret'], 'Additions': cm['additions'], 'Additions Ref': cm['add_ref'],
                              'Deductions': cm['deductions'], 'Deductions Ref': cm['ded_ref'], 'Net Payable': cm['net'],
                              'Progress %': round(cm['progress'] * 100, 2), 'Owner IPC': cm['owner_ipc'], 'Invoice No.': None,
                              'Invoice Date': None, 'Invoice Amount': None, 'Status': 'Draft', 'Saved At': dt.date.today(),
                              'Folder': folder})
    if old:                                               # تعديل مستخلص اتحفظ قبل كده ← حالته وفاتورته يفضلوا
        reg.data['CERTS'][-1].update({k: old.get(k) for k in ('Status', 'Invoice No.', 'Invoice Date', 'Invoice Amount')})
        reg.data['CERTS'][-1]['Status'] = reg.data['CERTS'][-1]['Status'] or 'Draft'
    for r in cm['rows']:
        reg.data['CERT_QTY'].append({'Sub ID': sid, 'Cert No.': no, 'Item No.': r['item'], 'Qty To Date': r['tot'],
                                     'Rate': r['rate'], 'Value To Date': round(r['val'], 2)})
    reg.save()


def cert_title(cm):
    return f"المستخلص الختامي رقم {cm['no']}" if cm['final'] else f"مستخلص جاري رقم {cm['no']}"


def cert_folder(cm):
    sub = cm['sub']
    return os.path.join(sub_folder(sub), PT.FINAL if cm['final'] else PT.CERTS,
                        f"V-{cm['no']:02d} ({cm['cutoff']:%d-%m-%Y})")


def cert_dir(sub, c):
    """فولدر مستخلص متسجل: المسار المتخزن لو موجود، وإلا مكانه في فولدر المقاول (لو الفولدرات اتغيرت أو المشروع اتنقل)"""
    p = live_path(c.get('Folder'))
    if p and os.path.isdir(p): return p
    d1 = todate(c.get('Period To'))
    if not d1 or not c.get('Cert No.'): return p
    return os.path.join(sub_dir(sub), PT.FINAL if c.get('Final') else PT.CERTS, f"V-{int(c['Cert No.']):02d} ({d1:%d-%m-%Y})")


# ---------------------------------------------------------------- مذكرات الخصم (00_Registers\10-LOG SUB.xlsx ← شيت DEDUCTION_MEMOS)
DM_HDR = ['Sr', 'Memo No.', 'Date', 'Subcontractor', 'Subject', 'Amount (SAR)', 'Deducted in (Sub IPC)', 'Status', 'File']
DM_WIDTH = (5, 12, 12, 24, 40, 14, 16, 12, 70)
DM_STATUSES = ('Open', 'Deducted', 'Cancelled')
DM_HEAD, DM_LAST = 11, 311                  # زي باقي شيتات السجل: العناوين في صف 11 والجدول لحد صف 311
DM_KEYS = {'no': 'Memo No.', 'date': 'Date', 'sub': 'Subcontractor', 'subject': 'Subject', 'amount': 'Amount (SAR)',
           'cert': 'Deducted in (Sub IPC)', 'status': 'Status', 'file': 'File'}
EN_NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9 .,&()'+_-]*")
EN_MAX, SUBJ_MAX, PATH_MAX = 30, 60, 240     # أطول اسم إنجليزي للمقاول، وأطول موضوع، وأطول مسار (ويندوز بيقف عند 260)
RISKY_PARTS = ('xl/media/', 'xl/charts/', 'xl/vbaProject', 'xl/pivot', 'xl/slicer', 'xl/timeline', 'xl/ctrlProps/',
               'xl/embeddings/', 'xl/threadedComments/', 'xl/activeX/', 'xl/metadata.xml', 'xl/richData/')


def log_path(): return Q.project_path(LOG_REL)


def inbox_dir(): return Q.project_path(INBOX_REL)


# ---------------------------------------------------------------- الخصومات (مخالفات + معدات + مواد + غيره)
# كل خصم: سطر في شيت DEDUCTIONS (10.1-SUBCONTRACTORS.xlsx) + فولدر في 10_Site_Records\DEDUCTIONS فيه الورقة والصور.
# Status: Open (مستني مستخلص) ← In memo (دخل مذكرة خصم DM-xxx) ← (المذكرة نفسها بتتخصم في المستخلص) / Company (على الشركة) / Cancelled
DED_REL = r'10_Site_Records\DEDUCTIONS'
DED_TYPES = ('Violation', 'Company equipment', 'Materials', 'Other')
DED_TYPE_AR = {'Violation': 'مخالفة', 'Company equipment': 'معدات الشركة', 'Materials': 'مواد', 'Other': 'خصم آخر'}
DED_SOURCES = ('Infrastructure', 'CLIENT', 'Municipality', 'Traffic', 'Other')
DED_SOURCE_AR = {'Infrastructure': 'البنية التحتية', 'CLIENT': 'الجهة المالكة', 'Municipality': 'البلدية',
                 'Traffic': 'المرور', 'Other': 'جهة أخرى'}
DED_STATUSES = ('Open', 'In memo', 'Company', 'Cancelled')
IMG_EXT = ('.jpg', '.jpeg', '.png', '.bmp', '.webp')


def ded_type_ar(d):
    t = DED_TYPE_AR.get(d.get('Type'), str(d.get('Type') or ''))
    if d.get('Type') == 'Violation' and d.get('Source'): t += ' – ' + DED_SOURCE_AR.get(d['Source'], str(d['Source']))
    return t


def ded_next_id(reg):
    nos = [int(m.group(1)) for d in reg.data.get('DEDUCTIONS', []) for m in [re.match(r'DD-(\d+)', str(d.get('ID') or ''))] if m]
    root = Q.project_path(DED_REL)
    if os.path.isdir(root):
        nos += [int(m.group(1)) for f in os.listdir(root) for m in [re.match(r'DD-(\d+)', f)] if m]
    return f'DD-{max(nos or [0]) + 1:03d}'


def ded_dir(d, make=False):
    """فولدر الخصم (المسار متسجل نسبي لفولدر المشروع)"""
    if d.get('Folder'):
        p = Q.project_path(d['Folder'])
        if make: os.makedirs(p, exist_ok=True)
        return p
    if not make: return ''
    dd = todate(d.get('Date')) or dt.date.today()
    what = d.get('Source') if d.get('Type') == 'Violation' and d.get('Source') else d.get('Type') or 'Other'
    ref = re.sub(r'[^A-Za-z0-9 ._()+-]+', ' ', str(d.get('Ref No.') or '')).strip()
    name = safe_name(f"{d['ID']} - {what}" + (f' - {ref}' if ref else '') + f' ({dd:%Y-%m-%d})')
    rel = os.path.join(DED_REL, name)
    os.makedirs(Q.project_path(rel), exist_ok=True)
    d['Folder'] = rel
    return Q.project_path(rel)


def ded_files(d):
    """(الورقة، [الصور والملفات المرفقة]) من فولدر الخصم – الورقة اسمها بيبدأ بـ «00 - »"""
    p = ded_dir(d)
    if not p or not os.path.isdir(p): return None, []
    names = sorted(n for n in os.listdir(p) if not n.lower().startswith(('desktop.ini', 'thumbs.db', '~$')) and os.path.isfile(os.path.join(p, n)))
    paper = next((os.path.join(p, n) for n in names if n.startswith('00 - ')), None)
    photos = [os.path.join(p, n) for n in names if not n.startswith('00 - ') and n.lower().endswith(IMG_EXT + ('.pdf',))]
    return paper, photos


def _bring(src, dst):
    """من 00_INBOX ← نقل، من أي مكان تاني ← نسخ"""
    ib = os.path.normcase(os.path.abspath(inbox_dir()))
    if os.path.normcase(os.path.abspath(src)).startswith(ib + os.sep):
        try: shutil.move(src, dst); return dst
        except OSError: pass
    shutil.copy2(src, dst)
    return dst


def ded_add_paper(d, src):
    p = ded_dir(d, make=True)
    old, _ = ded_files(d)
    if old:
        arc = os.path.join(p, '_OLD'); os.makedirs(arc, exist_ok=True)
        shutil.move(old, os.path.join(arc, f'{dt.datetime.now():%Y%m%d-%H%M%S} - ' + os.path.basename(old)))
    return _bring(src, os.path.join(p, '00 - ' + safe_name(os.path.basename(src))))


def ded_add_photos(d, srcs):
    p = ded_dir(d, make=True)
    _, have = ded_files(d)
    n = max([int(m.group(1)) for f in have for m in [re.match(r'(\d+)', os.path.basename(f))] if m] or [0])
    out = []
    for src in srcs:
        n += 1
        out.append(_bring(src, os.path.join(p, f'{n:02d}{os.path.splitext(src)[1].lower()}')))
    return out


def ded_del_photo(d, path):
    """الصورة ماتتمسحش – بتتنقل _OLD جوه فولدر الخصم"""
    arc = os.path.join(ded_dir(d), '_OLD'); os.makedirs(arc, exist_ok=True)
    shutil.move(path, os.path.join(arc, f'{dt.datetime.now():%Y%m%d-%H%M%S} - ' + os.path.basename(path)))


def ded_amount(d):
    a = num(d.get('Amount'))
    if a is None and num(d.get('Qty')) is not None and num(d.get('Rate')) is not None: a = num(d['Qty']) * num(d['Rate'])
    return round(a or 0.0, 2)


def english_ok(s):
    """اسم بالإنجليزي ينفع للفولدر والملف: حروف وأرقام ومسافات و . , & ( ) ' + _ -"""
    s = str(s or '').strip()
    return bool(s) and bool(EN_NAME.fullmatch(s)) and not s.endswith('.')


def memo_dir(sub): return os.path.join(sub_folder(sub), PT.DMEMOS)


def memo_ref(nos):
    """للمستخلص المطبوع: «مذكرة خصم DM-001» أو «مذكرات خصم DM-001 و DM-002»"""
    nos = [str(n) for n in nos if n]
    if not nos: return ''
    return ('مذكرة خصم ' if len(nos) == 1 else 'مذكرات خصم ') + ' و '.join(nos)


def _same(a, b): return str(a or '').strip().lower() == str(b or '').strip().lower()


def _owner(sub): return sub.get('Folder') or sub.get('Name')


def vlabel(v):
    """رقم مستخلص المقاول بشكل واحد: «V-2» أو «v 02» أو «2» ← «V-02»"""
    m = re.fullmatch(r'\s*(?:V\s*-?\s*)?0*(\d+)\s*', str(v if v is not None else ''), re.I)
    return f'V-{int(m.group(1)):02d}' if m else str(v or '').strip()


def _status(st, cert):
    """الحالة زي ما اتكتبت باليد (deducted / DEDUCTED …) ← Open / Deducted / Cancelled.
    فاضية ← Deducted لو فيه رقم مستخلص، وإلا Open. حاجة غريبة بتفضل زي ما هي (ومش بتتختار في المستخلص)."""
    t = str(st or '').strip()
    for k in DM_STATUSES:
        if t.lower() == k.lower(): return k
    if not t: return 'Deducted' if cert else 'Open'
    return t


def _amount(v):
    """القيمة: رقم، أو كلام زي «1,500.00» – المعادلة اللي ملهاش قيمة متحفوظة أو أي حاجة تانية ← None"""
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    t = str(v or '').replace(',', '').replace('SAR', '').strip()
    if not t or t.startswith('='): return None
    try: return float(t)
    except ValueError: return None


def _dm_map(values_rows):
    """صف العناوين وعمود كل عنوان بالاسم (لو حد زوّد عمود في النص يفضل شغال) – عنوان ناقص ← رسالة واضحة"""
    for i, row in enumerate(values_rows, 1):
        vals = {}
        for j, v in enumerate(row, 1):
            t = str(v or '').strip()
            if t and t not in vals: vals[t] = j
        if 'Memo No.' in vals:
            miss = [h for h in DM_KEYS.values() if h not in vals]
            if miss:
                raise RuntimeError(f"10-LOG SUB, sheet {DM_SHEET}: «{'», «'.join(miss)}» not found in the titles row – "
                                   'put the title back as it was')
            return i, {k: vals[h] for k, h in DM_KEYS.items()}
        if i >= 20: break
    raise RuntimeError(f'10-LOG SUB, sheet {DM_SHEET}: the titles row (Memo No. …) was not found')


def _dm_cols(ws):
    return _dm_map([[ws.cell(r, c).value for c in range(1, min(ws.max_column, 60) + 1)] for r in range(1, 21)])


def memos(sub=None, path=None):
    """مذكرات الخصم اللي في 10-LOG SUB (كلها أو بتاعة مقاول واحد) – كل واحدة dict فيها row = صفها في الشيت.
    amount = None لو القيمة مش رقم (اتكتبت غلط في إكسل) – المذكرة دي مش بتتختار في المستخلص لحد ما تتصلح."""
    p = path or log_path()
    if not p or not os.path.isfile(p): return []
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    try:
        if DM_SHEET not in wb.sheetnames: return []
        rows = [list(r) for r in wb[DM_SHEET].iter_rows(values_only=True)]
    finally:
        wb.close()
    h, col = _dm_map(rows[:20])
    out = []
    for i, r in enumerate(rows[h:], h + 1):
        g = lambda k: r[col[k] - 1] if col[k] - 1 < len(r) else None
        no = str(g('no') or '').strip()
        if not no: continue
        cert = vlabel(g('cert'))
        d = dict(row=i, no=no, date=todate(g('date')) or g('date'), sub=str(g('sub') or '').strip(),
                 subject=str(g('subject') or '').strip(), amount=_amount(g('amount')), cert=cert,
                 status=_status(g('status'), cert), file=str(g('file') or '').strip())
        if sub is None or _same(d['sub'], _owner(sub)): out.append(d)
    return out


def can_deduct(m, cert):
    """تتختار في المستخلص ده؟ مفتوحة ومتخصمتش، أو متخصمة في نفس المستخلص – وقيمتها رقم"""
    if m.get('amount') is None: return False
    return (m['status'] == 'Open' and not m['cert']) or (m['status'] == 'Deducted' and m['cert'] == cert)


def next_memo_no(sub, rows=None):
    """الرقم اللي عليه الدور للمقاول ده: الأكبر في السجل وفي فولدره + 1 (الملغية رقمها مش بيتكرر)"""
    rows = memos(sub) if rows is None else rows
    nos = [int(m.group(1)) for d in rows for m in [re.match(r'DM-(\d+)', d['no'], re.I)] if m]
    dd = os.path.join(sub_dir(sub), PT.DMEMOS)
    if os.path.isdir(dd):
        nos += [int(m.group(1)) for f in os.listdir(dd) for m in [re.match(r'DM-(\d+)', f, re.I)] if m]
    return max(nos or [0]) + 1


def _log_check(p):
    """الكتابة بـ openpyxl بتضيّع الصور والرسوم والماكرو وقوايم إكسل 2010 – لو الملف فيه حاجة منهم منكتبش فيه"""
    bad = []
    with zipfile.ZipFile(p) as z:
        for n in z.namelist():
            if n.startswith(RISKY_PARTS): bad.append(n.split('/')[1].replace('.xml', ''))
            elif n.startswith('xl/drawings/') and not re.fullmatch(r'xl/drawings/(vmlDrawing|commentsDrawing)\d*\.vml', n):
                bad.append('drawings')                  # الـ VML لوحده = ملاحظات الخلايا، و openpyxl بيحافظ عليها
            elif n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
                x = z.read(n)
                if b'<extLst' in x and any(t in x for t in (b'x14:dataValidation', b'x14:conditionalFormatting', b'x14:sparkline')):
                    bad.append('Excel 2010 lists or colours')
    if bad:
        raise RuntimeError(f"{os.path.basename(p)} has {', '.join(sorted(set(bad)))} that the program cannot keep – "
                           'write the memo row in it by hand')


def log_writable():
    """10-LOG SUB موجود ومش مفتوح في إكسل ومفيهوش حاجة هتضيع؟ – قبل ما نحرك أي ملف أو نحفظ مستخلص"""
    p = log_path()
    if not p or not os.path.isfile(p): raise FileNotFoundError(f'Not found: {p}')
    try:
        with open(p, 'r+b'): pass
    except PermissionError:
        raise PermissionError(f'{os.path.basename(p)} is open in Excel – close it and try again')
    _log_check(p)
    return p


def dm_sheet(wb):
    """شيت DEDUCTION_MEMOS – لو مش موجود بيتعمل بنفس شكل باقي شيتات السجل (البيانات فوق، العناوين في صف 11)"""
    if DM_SHEET in wb.sheetnames: return wb[DM_SHEET]
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import FormulaRule
    ref = next((wb[n] for n in ('Contracts & Invoices', 'Sub Approvals') if n in wb.sheetnames), None)
    ws = wb.create_sheet(DM_SHEET)
    navy, grey = Font(bold=True, color='1F4E78'), Font(color='808080')
    for r in range(2, 6):                       # المشروع والمالك والاستشاري والمقاول – منقولة من الشيت اللي جنبه
        src = ref.cell(r, 1) if ref is not None else None
        if src is not None and src.value:
            x = ws.cell(r, 1, src.value); x.font = copy(src.font)
    ws.cell(7, 1, 'Deduction Memos – Subcontractors').font = Font(size=14, bold=True, color='1F4E78')
    ws.cell(8, 1, 'Updated: ').font = grey
    ws.cell(8, 2, '=TODAY()').number_format = 'dd-mm-yyyy'
    ws.cell(9, 1, 'Memo No.: DM-001, DM-002 … for every subcontractor on its own   |   File: '
                  '13_PARTIES\\SUB\\<NN-SUB_NAME>\\06_DEDUCTION_MEMOS\\   |   Added by the SUB section of the program').font = grey
    st_rng, am_rng = f'$H${DM_HEAD + 1}:$H${DM_LAST}', f'$F${DM_HEAD + 1}:$F${DM_LAST}'
    for c, h in ((8, 'Status'), (10, 'Count'), (11, 'Amount (SAR)')): ws.cell(1, c, h).font = navy
    for i, st in enumerate(DM_STATUSES, 2):
        ws.cell(i, 8, st).font = Font(bold=True)
        ws.cell(i, 10, f'=COUNTIF({st_rng},"{st}")').font = Font(bold=True)
        x = ws.cell(i, 11, f'=SUMIF({st_rng},"{st}",{am_rng})'); x.font = Font(bold=True); x.number_format = '#,##0.00'
    ws.cell(5, 8, 'Total').font = Font(bold=True)
    ws.cell(5, 10, '=SUM(J2:J4)').font = Font(bold=True)
    x = ws.cell(5, 11, '=SUM(K2:K4)'); x.font = Font(bold=True); x.number_format = '#,##0.00'
    hs = ref.cell(DM_HEAD, 1) if ref is not None and ref.cell(DM_HEAD, 1).value else None
    h_font = copy(hs.font) if hs is not None else Font(bold=True, color='FFFFFF')
    h_fill = copy(hs.fill) if hs is not None else PatternFill('solid', fgColor='1F4E78')
    h_al = copy(hs.alignment) if hs is not None else Alignment(horizontal='center', vertical='center', wrap_text=True)
    h_bd = copy(hs.border) if hs is not None else BOX
    for j, (h, w) in enumerate(zip(DM_HDR, DM_WIDTH), 1):
        x = ws.cell(DM_HEAD, j, h); x.font, x.fill, x.alignment, x.border = h_font, h_fill, h_al, h_bd
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.row_dimensions[DM_HEAD].height = 32
    bs = ref.cell(DM_HEAD + 1, 2) if ref is not None else None
    bd = copy(bs.border) if bs is not None and bs.border.left.style else BOX
    cen = Alignment(horizontal='center')
    for r in range(DM_HEAD + 1, DM_LAST + 1):
        ws.cell(r, 1, f'=IF(B{r}="","",ROW()-{DM_HEAD})')
        for j in range(1, len(DM_HDR) + 1):
            x = ws.cell(r, j); x.border = bd
            if j in (1, 2, 3, 7, 8): x.alignment = cen
            if j == 3: x.number_format = 'dd-mm-yyyy'
            elif j == 6: x.number_format = '#,##0.00'
    dv = DataValidation(type='list', formula1='"' + ','.join(DM_STATUSES) + '"', allow_blank=True)
    dv.add(f'H{DM_HEAD + 1}:H{DM_LAST}'); ws.add_data_validation(dv)
    for st, fill, fc in (('Deducted', 'E2F2E6', '2C7536'), ('Open', 'FFF1D6', '7A5200'), ('Cancelled', 'EDEDED', '808080')):
        ws.conditional_formatting.add(f'A{DM_HEAD + 1}:I{DM_LAST}', FormulaRule(
            formula=[f'$H{DM_HEAD + 1}="{st}"'], fill=PatternFill(start_color=fill, end_color=fill, fill_type='solid'),
            font=Font(color=fc)))
    ws.auto_filter.ref = f'A{DM_HEAD}:I{DM_LAST}'
    ws.freeze_panes = f'C{DM_HEAD + 1}'
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f'{DM_HEAD}:{DM_HEAD}'
    return ws


def _log_edit(fn, sheet=None, name=DM_SHEET):
    """يفتح 10-LOG SUB، يعمل التغيير (fn(الشيت) ← عدد الصفوف اللي اتغيرت)، ويحفظ بأمان (ملف مؤقت ثم replace).
    sheet = دالة الشيت (dm_sheet لمذكرات الخصم، ct_sheet للعقود) – بتعمله لو مش موجود."""
    p = log_writable()
    try: wb = openpyxl.load_workbook(p, rich_text=True)  # الكلام الملوّن جوه الخلية يفضل زي ما هو
    except TypeError: wb = openpyxl.load_workbook(p)
    created = name not in wb.sheetnames
    n = fn((sheet or dm_sheet)(wb))
    if n or created:
        fd, tmp = tempfile.mkstemp(suffix='.xlsx', dir=os.path.dirname(p)); os.close(fd)
        try:
            wb.save(tmp)
            try:
                os.replace(tmp, p)
            except PermissionError:
                raise PermissionError(f'{os.path.basename(p)} is open in Excel – close it and try again')
        finally:
            if os.path.exists(tmp):
                try: os.remove(tmp)
                except OSError: pass
    return n


def ensure_log_sheet():
    """يعمل شيت DEDUCTION_MEMOS في 10-LOG SUB لو مش موجود (من غير ما يغيّر باقي الشيتات)"""
    return _log_edit(lambda ws: 0)


def add_memo(sub, subject, date, amount, src):
    """مذكرة خصم جديدة: الرقم اللي عليه الدور ← الملف في فولدر المقاول\\DEDUCTION_MEMOS باسم
    «DM-001 - <Sub Name> - <Subject> (YYYY-MM-DD)» ← سطر في 10-LOG SUB. الملف اللي في 00_INBOX بيتنقل، غير كده بيتنسخ.
    يرجّع (رقم المذكرة، مسار الملف)."""
    name = str(sub.get('Folder') or '').strip()
    if not english_ok(name): raise ValueError('Set the English name of this subcontractor first (Edit) – it goes in the folder and file names')
    subject = re.sub(r'\s+', ' ', str(subject or '')).strip()
    if not english_ok(subject): raise ValueError('Write the subject in English (letters, numbers and spaces)')
    if len(subject) > SUBJ_MAX: raise ValueError(f'The subject is too long for a file name ({SUBJ_MAX} letters at most)')
    amount = num(amount)
    if amount is None or amount <= 0: raise ValueError('Amount: a number more than 0')
    date = todate(date) or dt.date.today()
    if not src or not os.path.isfile(src): raise ValueError('Choose the memo file')
    p = log_writable()                                  # قبل ما نحرك أي ملف
    memo_no = f'DM-{next_memo_no(sub):03d}'
    ext = os.path.splitext(src)[1].lower() or '.pdf'
    dst = os.path.join(memo_dir(sub), f'{memo_no} - {name} - {subject} ({date:%Y-%m-%d}){ext}')
    if len(os.path.abspath(dst)) > PATH_MAX:
        raise ValueError('The file path is too long for Windows – make the subject (or the English name) shorter')
    if os.path.exists(dst): raise FileExistsError(f'Already there: {dst}')
    ib = os.path.normcase(os.path.abspath(inbox_dir()))
    from_inbox = os.path.normcase(os.path.abspath(src)).startswith(ib + os.sep)
    try:
        rel = os.path.relpath(dst, os.path.dirname(p)).replace(os.sep, '/')
    except ValueError:                                  # درايف تاني
        rel = 'file:///' + dst.replace(os.sep, '/')
    link = quote(rel, safe="/:._-()&,+'")

    def fn(ws):
        h, col = _dm_cols(ws)
        r = h + 1
        while ws.cell(r, col['no']).value not in (None, ''): r += 1
        if r > DM_LAST and r > h + 1:                   # الجدول خلص: الصف الجديد بنفس شكل اللي قبله
            for j in range(1, ws.max_column + 1): ws.cell(r, j)._style = copy(ws.cell(r - 1, j)._style)
        vals = dict(no=memo_no, date=date, sub=name, subject=subject, amount=round(amount, 2), cert=None, status='Open',
                    file=os.path.basename(dst))
        for k, v in vals.items(): ws.cell(r, col[k], v)
        ws.cell(r, col['date']).number_format = 'dd-mm-yyyy'; ws.cell(r, col['amount']).number_format = '#,##0.00'
        c = ws.cell(r, col['file']); c.hyperlink = link; c.font = Font(color='0563C1', underline='single')
        return 1
    try:
        shutil.copy2(src, dst)
        _log_edit(fn)
    except Exception:
        try:                                            # السطر متكتبش ← النسخة تتشال والأصل زي ما هو
            if os.path.exists(dst): os.remove(dst)
        except OSError:
            pass
        raise
    left = None
    if from_inbox:                                      # من 00_INBOX: الأصل يتشال (لو مفتوح في برنامج تاني بيفضل ونقول)
        try: os.remove(src)
        except OSError: left = src
    return memo_no, dst, left


def mark_deducted(sub, nos, cert):
    """حفظ مستخلص المقاول: المذكرات المختارة ← Deducted + رقم المستخلص (V-02). اللي كانت في نفس المستخلص ومبقتش مختارة ← ترجع Open."""
    who, cert = _owner(sub), vlabel(cert)
    want = {str(n).strip().upper() for n in nos if n}
    if not want and not any(m['cert'] == cert and m['status'] == 'Deducted' for m in memos(sub)): return 0

    def fn(ws):
        h, col = _dm_cols(ws)
        n, bad, seen = 0, [], set()
        for r in range(h + 1, ws.max_row + 1):
            no = str(ws.cell(r, col['no']).value or '').strip().upper()
            if not no or not _same(ws.cell(r, col['sub']).value, who): continue
            seen.add(no)
            cur = vlabel(ws.cell(r, col['cert']).value)
            st = _status(ws.cell(r, col['status']).value, cur)
            if no in want:
                if st == 'Deducted' and cur == cert: continue
                if st != 'Open' or cur:                 # اتلغت أو اتخصمت في مستخلص تاني من بعد ما اخترتها
                    bad.append(f"{no} ({st}{' in ' + cur if cur else ''})"); continue
                ws.cell(r, col['cert']).value, ws.cell(r, col['status']).value = cert, 'Deducted'; n += 1
            elif st == 'Deducted' and cur == cert:
                ws.cell(r, col['cert']).value, ws.cell(r, col['status']).value = None, 'Open'; n += 1
        bad += [f'{no} (not in the register)' for no in sorted(want - seen)]
        if bad: raise RuntimeError('These memos changed in 10-LOG SUB: ' + ', '.join(bad))
        return n
    return _log_edit(fn)


def set_memo_amount(sub, no, amount):
    """مبلغ مذكرة لسه Open في 10-LOG SUB (لما أسعار خصوماتها تتكتب بعد ما اتطلعت) – المخصومة ماتتغيرش.
    مذكرة ملفها في 06_DEDUCTION_MEMOS ومش متسجلة (زي DM-001 اللي اتعملت بإيد) ← تتسجل Open بالمبلغ."""
    who = _owner(sub)
    dd = memo_dir(sub)
    files = sorted(f for f in (os.listdir(dd) if os.path.isdir(dd) else []) if f.upper().startswith(str(no).upper() + ' '))
    f = next((x for x in files if x.lower().endswith('.pdf')), files[0] if files else None)

    def fn(ws):
        h, col = _dm_cols(ws)
        for r in range(h + 1, ws.max_row + 1):
            if str(ws.cell(r, col['no']).value or '').strip().upper() == str(no).upper() and _same(ws.cell(r, col['sub']).value, who):
                cur = vlabel(ws.cell(r, col['cert']).value)
                st = _status(ws.cell(r, col['status']).value, cur)
                if st != 'Open': raise ValueError(f'{no} is {st} – its amount cannot change')
                ws.cell(r, col['amount']).value = round(float(amount), 2)
                ws.cell(r, col['amount']).number_format = '#,##0.00'
                return 1
        if not f: raise ValueError(f'{no} is not in 10-LOG SUB and its file is not in {PT.DMEMOS}')
        m = re.match(r'DM-\d+ - .*? - (.*) \((\d{4}-\d{2}-\d{2})\)', os.path.splitext(f)[0])
        r = h + 1
        while ws.cell(r, col['no']).value not in (None, ''): r += 1
        if r > DM_LAST and r > h + 1:
            for j in range(1, ws.max_column + 1): ws.cell(r, j)._style = copy(ws.cell(r - 1, j)._style)
        vals = dict(no=str(no).upper(), date=todate(m.group(2)) if m else dt.date.today(), sub=str(sub.get('Folder') or who),
                    subject=m.group(1) if m else 'Deductions', amount=round(float(amount), 2), cert=None, status='Open', file=f)
        for k, v in vals.items(): ws.cell(r, col[k], v)
        ws.cell(r, col['date']).number_format = 'dd-mm-yyyy'; ws.cell(r, col['amount']).number_format = '#,##0.00'
        try:
            rel = os.path.relpath(os.path.join(dd, f), os.path.dirname(log_path())).replace(os.sep, '/')
            c = ws.cell(r, col['file']); c.hyperlink = quote(rel, safe="/:._-()&,+'"); c.font = Font(color='0563C1', underline='single')
        except ValueError:
            pass
        return 1
    return _log_edit(fn)


def cancel_memo(sub, no):
    """المذكرة تفضل في السجل بحالة Cancelled، ورقمها مش بيتكرر"""
    who = _owner(sub)

    def fn(ws):
        h, col = _dm_cols(ws)
        for r in range(h + 1, ws.max_row + 1):
            if str(ws.cell(r, col['no']).value or '').strip().upper() == str(no).upper() and _same(ws.cell(r, col['sub']).value, who):
                cur = vlabel(ws.cell(r, col['cert']).value)
                st = _status(ws.cell(r, col['status']).value, cur)
                if st == 'Deducted': raise ValueError(f'{no} is deducted in {cur or "a certificate"} – it cannot be cancelled')
                if st == 'Cancelled': return 0
                ws.cell(r, col['status']).value = 'Cancelled'
                return 1
        raise ValueError(f'{no} is not in 10-LOG SUB')
    return _log_edit(fn)


# ---------------------------------------------------------------- ملف المستخلص (إكسل)
AR = 'Arial'
THIN = Side(style='thin', color='000000')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill('solid', fgColor='D9E1F2')
SEC = PatternFill('solid', fgColor='EDEDED')
TOT = PatternFill('solid', fgColor='FFF2CC')
C_RIGHT = Alignment(horizontal='right', vertical='center', wrap_text=True, readingOrder=2)
C_MID = Alignment(horizontal='center', vertical='center', wrap_text=True, readingOrder=2)


def _f(size=10, bold=False, color='000000'): return Font(name=AR, size=size, bold=bold, color=color)


def _cell(ws, r, c, v=None, size=10, bold=False, al=C_MID, fill=None, fmt=None, border=True):
    x = ws.cell(r, c, v)
    x.font = _f(size, bold); x.alignment = al
    if fill: x.fill = fill
    if fmt: x.number_format = fmt
    if border: x.border = BOX
    return x


def _merge(ws, r0, c0, r1, c1, v=None, size=10, bold=False, al=C_MID, fill=None, border=True, fmt=None):
    if (r0, c0) != (r1, c1): ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
    x = _cell(ws, r0, c0, v, size, bold, al, fill, fmt, border)
    if border:
        for rr in range(r0, r1 + 1):
            for cc in range(c0, c1 + 1): ws.cell(rr, cc).border = BOX
    return x


def _page(ws, landscape=True, scale=None, fit=False, fitw=False):
    """fit = الشيت كله في صفحة، fitw = على عرض الصفحة (والطول صفحات)، scale = مقاس ثابت (لما فيه فواصل صفحات بإيدينا)"""
    ws.sheet_view.rightToLeft = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_margins = PageMargins(left=0.35, right=0.35, top=0.4, bottom=0.45, header=0.2, footer=0.25)
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = 'Page &P of &N'; ws.oddFooter.center.size = 8
    if fit or fitw:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1 if fit else 0
    elif scale:
        ws.page_setup.scale = scale


def _signs(ws, r, roles, signs, ncol):
    """خانات التوقيع: الدور ثم الاسم/التاريخ/التوقيع – موزعة على عرض الصفحة"""
    names = dict(signs)
    n = len(roles); w = ncol / n
    for i, role in enumerate(roles):
        c0 = int(round(i * w)) + 1; c1 = int(round((i + 1) * w))
        for k, (txt, b) in enumerate(((role, True), (f"الاسم / {names.get(role, '')}", False), ('التاريخ /', False), ('التوقيع /', False))):
            _merge(ws, r + k, c0, r + k, c1, txt, 9, b, Alignment(horizontal='right', vertical='center', readingOrder=2), border=False)
    return r + 4


def build_workbook(cm, out):
    """غلاف + بيان المتابعة + المستخلص (V) + الحصر (Q1/Q2) في ملف واحد. بيرجّع أسماء الشيتات لكل PDF."""
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    logos = {}
    _cover(wb.create_sheet('غلاف'), cm, logos)
    _follow(wb.create_sheet('بيان المتابعة'), cm)
    _vsheet(wb.create_sheet('المستخلص'), cm, logos)
    qsheets = []
    its = {r['item']: r for r in cm['rows']}
    groups = [('Q1', 'حصر أعمال الحفر والفرشة والتمديد والردم وطبقة قاعدة الطريق',
               [k for k, r in its.items() if k in cm['qrows'] and r['measure'] not in ASPHALT]),
              ('Q2', 'حصر أعمال الأسفلت', [k for k, r in its.items() if k in cm['qrows'] and r['measure'] in ASPHALT])]
    for name, title, keys in groups:
        if keys:
            _qsheet(wb.create_sheet(name), cm, title, keys, logos); qsheets.append(name)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    add_logos(out, logos)
    return dict(cover=['غلاف', 'بيان المتابعة'], v=['المستخلص'], q=qsheets)


def ar_name(sub):
    """اسم المقاول في الورق المطبوع بالعربي: الاسم لو عربي، وإلا الاسم العربي اللي في README فولدره في 13_PARTIES"""
    nm = str(sub.get('Name') or '').strip()
    if re.search(r'[\u0600-\u06FF]', nm): return nm
    try:
        p = PT.find(proj(), sub.get('Folder') or nm)
        ar = next((a for a in (p or {}).get('aliases') or [] if re.search(r'[\u0600-\u06FF]', a)), '')
        if ar: return ar if ar.startswith(('شركة', 'مؤسسة')) else 'شركة ' + ar
    except Exception:
        pass
    return nm.replace('_', ' ')


def _head(ws, r, cm, ncol, title=None, logos=None, logo_col=None):
    """رأس الصفحة: اللوجو شمال + الشركة والمشروع والمقاول في النص"""
    c_txt0, c_txt1 = 2, ncol - 2
    _merge(ws, r, c_txt0, r, c_txt1, COMPANY, 12, True, border=False)
    _merge(ws, r + 1, c_txt0, r + 1, c_txt1, PROJECT, 11, True, border=False)
    ct = str(cm['sub'].get('Contract No.') or '').strip()
    _merge(ws, r + 2, c_txt0, r + 2, c_txt1, ar_name(cm['sub']) + (f'   –   عقد رقم  {ct}' if ct else ''), 11, True, border=False)
    n = 4 if title else 3
    if title: _merge(ws, r + 3, c_txt0, r + 3, c_txt1, title, 12, True, border=False)
    for k in range(n): ws.row_dimensions[r + k].height = 20
    if logos is not None: logos.setdefault(ws.title, []).append((r - 1, (logo_col or ncol) - 1))
    return r + n


def _period_row(ws, r, cm, ncol):
    _merge(ws, r, 1, r, 3, cert_title(cm), 10, True, fill=HEAD)
    _merge(ws, r, 4, r, ncol // 2 + 1, f"عن الأعمال من  {fdate(cm['period_from'])}", 10, True, fill=HEAD)
    _merge(ws, r, ncol // 2 + 2, r, ncol, f"إلى  {fdate(cm['cutoff'])}", 10, True, fill=HEAD)
    ws.row_dimensions[r].height = 20
    return r + 1


V_COLS = [('م', 5), ('البند', 76), ('الوحدة', 8), ('الكمية', 11), ('السعر', 11), ('القيمة', 15), ('السابق', 11),
          ('الحالي', 11), ('الإجمالي', 11), ('قيمة الأعمال الإجمالية', 16)]
V_ROLES1 = ['مندوب مقاول الباطن', 'مهندس مكتب فني', 'مدير المشروع']
V_ROLES2 = ['مدير الإدارة الفنية', 'الإدارة المعنية', 'مدير إدارة المتابعة', 'مدير قطاع المقاولات']
V_SCALE = 80
PAGE_H = 534 / (V_SCALE / 100) - 4           # ارتفاع الصفحة بنقط الشيت (A4 بالعرض بعد الهوامش والمقاس)
FIXED_H = 15 + 60 + 36 + 20 + 15 + 8 + 2 * (4 * 13 + 8)     # رأس + عناوين + سطر الفترة + إجمالي الصفحة + التوقيعات
TOTALS_H = 10 * 15 + 8
SEC_H = 18


def _line_h(kind, v): return SEC_H if kind == 'sec' else (30 if len(str(v['desc'])) > 95 else 20)


def _paginate(lines):
    """تقسيم البنود على الصفحات بالارتفاع – وآخر صفحة لازم يبقى فيها مكان الخصومات (وإلا صفحة لوحدها للخصومات)"""
    pages, cur, used = [], [], 0.0
    for i, (kind, v) in enumerate(lines):
        h = _line_h(kind, v)
        cap = PAGE_H - FIXED_H - (15 if pages else 0)
        if cur and used + h > cap:
            pages.append(cur); cur, used = [], 0.0
        cur.append((kind, v)); used += h
    pages.append(cur)
    cap_last = PAGE_H - FIXED_H - (15 if len(pages) > 1 else 0) - TOTALS_H
    if used > cap_last: pages.append([])
    return pages


def _vsheet(ws, cm, logos):
    """المستخلص زي مستخلصات مقاولين الصرف: كل صفحة برأسها وتوقيعاتها، وإجمالي الصفحة / ما قبله، والخصومات في الآخر"""
    nc = len(V_COLS)
    _page(ws, landscape=True, scale=V_SCALE)
    for j, (_, w) in enumerate(V_COLS, 1): ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    # البنود مع عناوين الأقسام (أ، ب، ج…)
    lines, sec = [], None
    for r in cm['rows']:
        if (r['section'] or '') != (sec or '') and r['section']:
            lines.append(('sec', r['section'])); sec = r['section']
        lines.append(('item', r))
    pages = _paginate(lines)
    row, carry, first_rows = 2, None, []
    for pi, chunk in enumerate(pages):
        row = _head(ws, row, cm, nc, logos=logos)
        # عناوين الأعمدة (صفين – كميات البند فوق السابق/الحالي/الإجمالي)
        for j, (h, _) in enumerate(V_COLS, 1):
            if j in (7, 8, 9): continue
            _merge(ws, row, j, row + 1, j, h, 10, True, fill=HEAD)
        _merge(ws, row, 7, row, 9, 'كميات البند', 10, True, fill=HEAD)
        for j in (7, 8, 9): _cell(ws, row + 1, j, V_COLS[j - 1][0], 10, True, fill=HEAD)
        ws.row_dimensions[row].height = ws.row_dimensions[row + 1].height = 18
        row += 2
        row = _period_row(ws, row, cm, nc)
        if carry:
            _merge(ws, row, 1, row, 9, 'ما قبله', 10, True, C_RIGHT, fill=TOT)
            _cell(ws, row, 10, f'={carry}', 10, True, fill=TOT, fmt='#,##0.00'); row += 1
        r0 = row
        for kind, v in chunk:
            if kind == 'sec':
                _merge(ws, row, 1, row, nc, v, 10, True, C_RIGHT, fill=SEC); ws.row_dimensions[row].height = 18
            else:
                _cell(ws, row, 1, v['item'])
                _cell(ws, row, 2, v['desc'], al=C_RIGHT)
                _cell(ws, row, 3, v['unit'])
                _cell(ws, row, 4, v['cqty'], fmt='#,##0.000')
                _cell(ws, row, 5, v['rate'], fmt='#,##0.00')
                _cell(ws, row, 6, f'=D{row}*E{row}', fmt='#,##0.00')
                _cell(ws, row, 7, v['prev'], fmt='#,##0.000')
                _cell(ws, row, 8, f'=I{row}-G{row}', fmt='#,##0.000')
                _cell(ws, row, 9, v['tot'], fmt='#,##0.000')
                _cell(ws, row, 10, f'=I{row}*E{row}', fmt='#,##0.00')
                first_rows.append(row)
                ws.row_dimensions[row].height = 30 if len(str(v['desc'])) > 95 else 20
            row += 1
        pt = f'=SUM(J{r0}:J{row - 1})' + (f'+{carry}' if carry else '') if row > r0 else (f'={carry}' if carry else 0)
        _merge(ws, row, 1, row, 9, 'إجمالي الصفحة', 10, True, C_RIGHT, fill=TOT)
        _cell(ws, row, 10, pt, 10, True, fill=TOT, fmt='#,##0.00')
        carry = f'J{row}'; row += 1
        if pi == len(pages) - 1:
            ws.row_dimensions[row].height = 8
            r_tot = row + 1
            row = _vtotals(ws, r_tot, cm, carry, first_rows)
            for rr in range(r_tot, row): ws.row_dimensions[rr].height = max(15, ws.row_dimensions[rr].height or 0)
        for roles in (V_ROLES1, V_ROLES2):
            ws.row_dimensions[row].height = 8
            r_s = row + 1
            row = _signs(ws, r_s, roles, cm['signs'], nc)
            for rr in range(r_s, row): ws.row_dimensions[rr].height = 13
        if pi < len(pages) - 1:
            ws.row_breaks.append(Break(id=row)); row += 1
    ws.print_area = f'A1:J{row}'


def _vtotals(ws, r, cm, gross_cell, item_rows):
    """الخصومات بنفس ترتيب مستخلصات المقاولين: الإجمالي ← الحالي ← الضريبة ← بعد الضريبة ← الضمان ← الإضافات ← الحسميات ← الصافي"""
    no = cm['no']

    def jline(label, val, fmt='#,##0.00', strong=False):
        nonlocal r
        _merge(ws, r, 7, r, 9, label, 10, True, C_RIGHT, fill=TOT if strong else None)
        _cell(ws, r, 10, val, 10, True, fill=TOT if strong else None, fmt=fmt)
        if len(label) > 30: ws.row_dimensions[r].height = 15 * math.ceil(len(label) / 30)
        r += 1
        return r - 1
    # نفس الصف: إجمالي العقد (تحت عمود القيمة) + إجمالي أعمال المقاول
    _merge(ws, r, 1, r, 5, 'إجمالي قيمة الأعمال تعاقديًا', 10, True, C_RIGHT, fill=TOT)
    _cell(ws, r, 6, ('=SUM(' + ','.join(f'F{x}' for x in item_rows) + ')') if item_rows else 0, 10, True, fill=TOT, fmt='#,##0.00')
    g = jline('إجمالي قيمة أعمال المقاول', f'={gross_cell}', strong=True)
    p = jline(f'قيمة الأعمال حتى مستخلص رقم {no - 1}' if no > 1 else 'قيمة الأعمال بالمستخلصات السابقة', cm['prev'])
    c = jline(f'قيمة الأعمال الحالية بالمستخلص رقم {no}', f'=J{g}-J{p}')
    v = jline(f"ضريبة القيمة المضافة {cm['vat_p'] * 100:g} %", f"=ROUND(J{c}*{cm['vat_p']:g},2)")
    t = jline('إجمالي قيمة الأعمال بعد إضافة القيمة المضافة', f'=J{c}+J{v}')
    rt = jline(f"حسم {cm['ret_p'] * 100:g} % ضمان الأعمال من الإجمالي بعد الضريبة", f"=ROUND(J{t}*{cm['ret_p']:g},2)")
    a = jline(f"إضافات ( {cm['add_ref'] or 'مذكرة إضافة'} )", cm['additions']) if cm['additions'] else None
    d = jline(f"حسميات أخرى ( {cm['ded_ref'] or 'مذكرة خصم'} )", cm['deductions']) if cm['deductions'] else None
    jline('صافي المستحق صرفه للمقاول', f'=J{t}-J{rt}' + (f'-J{d}' if d else '') + (f'+J{a}' if a else ''), strong=True)
    return r


Q_COLS = [('م', 5), ('رقم الطلب', 9), ('الحي', 14), ('الشارع', 20), ('رقم الخط', 12), ('من', 10), ('إلى', 10),
          ('الطول (م)', 9), ('القطر', 7), ('متوسط العمق', 9), ('الكمية', 12), ('الوحدة', 7), ('نسبة التنفيذ %', 10)]
Q_ROLES = ['مندوب مقاول الباطن', 'مهندس مكتب فني']


def _qsheet(ws, cm, title, keys, logos):
    """الحصر: لكل بند جزء بعنوانه وصفوف الريكويستات اللي جوه استيشنات المقاول وإجمالي البند"""
    nc = len(Q_COLS)
    _page(ws, landscape=True, fitw=True)
    for j, (_, w) in enumerate(Q_COLS, 1): ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    row = _head(ws, 2, cm, nc, title, logos, logo_col=nc - 1)
    row = _period_row(ws, row, cm, nc)
    hdr_row = row
    for j, (h, _) in enumerate(Q_COLS, 1): _cell(ws, row, j, h, 10, True, fill=HEAD)
    ws.row_dimensions[row].height = 30
    row += 1
    ws.print_title_rows = f'{hdr_row}:{hdr_row}'
    its = {r['item']: r for r in cm['rows']}
    for k in keys:
        it = its[k]
        _merge(ws, row, 1, row, nc, f"البند {k}  –  {it['desc']}", 10, True, C_RIGHT, fill=SEC); row += 1
        r0 = row
        for i, d in enumerate(cm['qrows'][k], 1):
            vals = [i, d['wir'], d['dist'], d['street'], d['line'], sta(d['x0']), sta(d['x1']), d['L'], d['dia'],
                    d.get('avg'), d['qty'], it['unit'], (d['pct'] / 100 if d.get('pct') is not None else None)]
            for j, v in enumerate(vals, 1):
                fmt = '#,##0.00' if j in (8, 10) else '#,##0.000' if j == 11 else '0%' if j == 13 else None
                _cell(ws, row, j, round(v, 4) if isinstance(v, float) else v, 9, fmt=fmt,
                      al=C_RIGHT if j in (3, 4) else C_MID)
            row += 1
        _merge(ws, row, 1, row, 10, 'إجمالي البند', 10, True, C_RIGHT, fill=TOT)
        _cell(ws, row, 11, f'=SUM(K{r0}:K{row - 1})' if row > r0 else 0, 10, True, fill=TOT, fmt='#,##0.000')
        _cell(ws, row, 12, it['unit'], 10, True, fill=TOT)
        _cell(ws, row, 13, None, 10, True, fill=TOT)
        row += 2
    row = _signs(ws, row, Q_ROLES, cm['signs'], nc)
    ws.print_area = f'A1:{openpyxl.utils.get_column_letter(nc)}{row}'


def _cover(ws, cm, logos):
    _page(ws, landscape=False, fit=True)
    for j, w in enumerate((4, 22, 22, 22, 4), 1): ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    sub = cm['sub']
    lines = [(9, 'المملكة العربية السعودية', 14), (11, COMPANY, 18), (13, PROJECT, 14), (16, cert_title(cm), 22),
             (18, f"عن الأعمال من  {fdate(cm['period_from'])}   إلى  {fdate(cm['cutoff'])}", 14),
             (21, 'مقاول الباطن', 13), (22, ar_name(cm['sub']), 18),
             (24, f"عقد رقم  {sub.get('Contract No.') or ''}   بتاريخ  {fdate(sub.get('Contract Date'))}", 13),
             (26, f"تاريخ تسليم الموقع  {fdate(sub.get('Site Handover'))}", 13)]
    for r, t, s in lines:
        _merge(ws, r, 2, r, 4, t, s, True, border=False)
        ws.row_dimensions[r].height = s * 2
    logos.setdefault(ws.title, []).append(('cover', 0))
    ws.print_area = 'A1:E30'


def _follow(ws, cm):
    """بيان المتابعة الشهري لأعمال مقاولي الباطن – بيتملا من العقد والمستخلص ومستخلصات المالك"""
    _page(ws, landscape=False, fit=True)
    for j, w in enumerate((3, 30, 9, 9, 13, 13, 3), 1): ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    sub = cm['sub']
    R = Alignment(horizontal='right', vertical='center', readingOrder=2, wrap_text=True)
    r = 2
    _merge(ws, r, 2, r, 6, 'بيان المتابعة الشهري لأعمال مقاولي الباطن', 14, True, border=False); r += 1
    _merge(ws, r, 2, r, 6, f"بيان متابعة الأعمال  {fdate(cm['cutoff'])}", 12, True, border=False); r += 2

    def line(t, bold=False, size=11):
        nonlocal r
        _merge(ws, r, 2, r, 6, t, size, bold, R, border=False); ws.row_dimensions[r].height = 20; r += 1
    actions = cm['ded_ref'] and f"مرفق {cm['ded_ref']}" or 'لا يوجد'
    line('أولاً:-', True)
    line(f'1- اسم المشروع الذي تؤدى عليه العملية:  {PROJECT}')
    line(f"2- اسم مقاول الباطن:  {ar_name(sub)}")
    line('ثانياً:- (مدة التنفيذ)', True)
    line(f"1- عقد رقم  {sub.get('Contract No.') or ''}   المبرم بتاريخ  {fdate(sub.get('Contract Date'))}")
    line(f"2- تاريخ تسليم الموقع:  {fdate(sub.get('Site Handover'))}")
    dur = num(sub.get('Duration (months)'))
    line(f"3- المدة المحددة بموجب العقد:  {f'{dur:g} أشهر' if dur else ''}")
    line(f"4- مدة التأخير إن وجدت:  {str(cm['delay']) + ' يوم' if cm['delay'] else 'لا يوجد'}")
    line('ثالثاً:-', True)
    line(f"1- قيمة الأعمال بموجب المستخلص المرفق:  {cm['net']:,.2f}  ريال")
    line(f"2- نسبة الإنجاز الحالي:  {cm['progress'] * 100:.2f}%")
    line('3- نسبة التأخير إن وجدت:  ')
    line(f'4- الإجراءات المتخذة بحق المقاول:  {actions}')
    r += 1
    line('رابعاً:- شروط الصرف حسب العقد المبرم مع المقاول', True)
    line(f"حالة الأعمال المنفذة بمستخلص المقاول {'الختامي' if cm['final'] else 'الجاري'} رقم  {cm['no']}")
    hdr = ['البيان', 'لا', 'نعم', 'مستخلص رقم', 'نسبة الصرف']

    def table(rows):
        nonlocal r
        for j, h in enumerate(hdr, 2): _cell(ws, r, j, h, 10, True, fill=HEAD)
        r += 1
        for t, no, yes, ipc in rows:
            _cell(ws, r, 2, t, 10, al=R); _cell(ws, r, 3, '√' if no else ''); _cell(ws, r, 4, '√' if yes else '')
            _cell(ws, r, 5, ipc or ''); _cell(ws, r, 6, ''); ws.row_dimensions[r].height = 22; r += 1
    inc = cm['owner_ipc']
    table([('هل تم إدراجها بمستخلصات الجهة المالكة', not inc, bool(inc), inc)])
    r += 1
    line('في حالة عدم إدراجها بمستخلصات الجهة المالكة:-', True)
    table([('أ- تم تسليمها إلى الجهة المشرفة بدون ملاحظات', False, (not inc) and cm['all_approved'], None),
           ('ب- تم تسليمها إلى الجهة المشرفة بملاحظات توقف الصرف بنسبة', False, False, None)])
    r += 2
    names = dict(cm['signs'])
    for j, role in zip((2, 4, 6), ('مدير المشروع', 'مدير إدارة المتابعة', 'مدير قطاع المقاولات')):
        span = (j, j + 1) if j < 6 else (j, j)
        _merge(ws, r, span[0], r, span[1], role, 11, True, border=False)
        _merge(ws, r + 1, span[0], r + 1, span[1], names.get(role, ''), 10, False, border=False)
    ws.print_area = f'A1:G{r + 3}'


# ---------------------------------------------------------------- اللوجو (openpyxl من غير Pillow مش بيحط صور)
def add_logos(xlsx, logos):
    """logos = {اسم الشيت: [(صف, عمود) ...]} – اللوجو بيتحط فوق شمال كل صفحة (والغلاف في النص)"""
    if not logos or not os.path.exists(LOGO): return
    img = open(LOGO, 'rb').read()
    z = zipfile.ZipFile(xlsx)
    names = z.namelist()
    files = {n: z.read(n) for n in names}
    z.close()
    wbx = files['xl/workbook.xml'].decode('utf8')
    rels = files['xl/_rels/workbook.xml.rels'].decode('utf8')
    sheet_part = {}
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        t = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(2), rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % m.group(2), rels)
        if t: sheet_part[_unxml(m.group(1))] = 'xl/' + t.group(1).lstrip('/').replace('xl/', '', 1)
    files['xl/media/sub_logo.jpeg'] = img
    ct = files['[Content_Types].xml'].decode('utf8')
    if 'Extension="jpeg"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ', 1)
    emu = 9525
    for i, (sheet, spots) in enumerate(logos.items(), 1):
        part = sheet_part.get(sheet)
        if not part or part not in files: continue
        anchors = []
        for k, (r, c) in enumerate(spots, 1):
            if r == 'cover':
                w, h = 260, int(260 * LOGO_PX[1] / LOGO_PX[0]); r, c, xo = 0, 1, 40 * emu
            else:
                w, h = 105, int(105 * LOGO_PX[1] / LOGO_PX[0]); xo = 0
            anchors.append(
                f'<xdr:oneCellAnchor><xdr:from><xdr:col>{c}</xdr:col><xdr:colOff>{xo}</xdr:colOff><xdr:row>{r}</xdr:row>'
                f'<xdr:rowOff>{4 * emu}</xdr:rowOff></xdr:from><xdr:ext cx="{w * emu}" cy="{h * emu}"/>'
                f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{k + 1}" name="Logo {k}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/>'
                f'</xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch>'
                f'</xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * emu}" cy="{h * emu}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>')
        dname = f'xl/drawings/sublogo{i}.xml'
        files[dname] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                        + ''.join(anchors) + '</xdr:wsDr>').encode('utf8')
        files[f'xl/drawings/_rels/sublogo{i}.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            'Target="../media/sub_logo.jpeg"/></Relationships>').encode('utf8')
        ct = ct.replace('</Types>', f'<Override PartName="/{dname}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
        # علاقة الشيت بالرسمة
        srel = part.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if srel in files:
            sr = files[srel].decode('utf8')
            rid = 'rIdLogo1'
            sr = sr.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                                                f'Target="../drawings/sublogo{i}.xml"/></Relationships>')
        else:
            rid = 'rIdLogo1'
            sr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                  f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                  f'Target="../drawings/sublogo{i}.xml"/></Relationships>')
        files[srel] = sr.encode('utf8')
        sx = files[part].decode('utf8')
        if 'xmlns:r=' not in sx[:600]:
            sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        tag = f'<drawing r:id="{rid}"/>'
        # مكان <drawing> في ترتيب الشيت: بعد آخر واحد من printOptions/pageMargins/pageSetup/headerFooter/rowBreaks/colBreaks
        end = None
        for t in ('printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks'):
            for m in re.finditer(r'<%s\b[^>]*/>|<%s\b[^>]*>.*?</%s>' % (t, t, t), sx, re.S):
                end = max(end or 0, m.end())
        sx = sx[:end] + tag + sx[end:] if end else sx.replace('</worksheet>', tag + '</worksheet>')
        files[part] = sx.encode('utf8')
    files['[Content_Types].xml'] = ct.encode('utf8')
    tmp = xlsx + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, data in files.items(): zo.writestr(n, data)
    os.replace(tmp, xlsx)


def _unxml(s): return s.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'")


def export_pdfs(xlsx, parts, base):
    """3 ملفات PDF زي ورق مستخلصاتكم: الغلاف وبيان المتابعة، المستخلص (V)، الحصر (Q) – بإكسل (ويندوز)"""
    import ipc_core
    jobs = [(parts['cover'], base + ' - Cover & Follow-up.pdf'), (parts['v'], base + ' - V.pdf')]
    if parts['q']: jobs.append((parts['q'], base + ' - Q.pdf'))
    ipc_core.export_pdfs(xlsx, jobs)
    return [p for _, p in jobs]


# ---------------------------------------------------------------- عقود مقاولي الباطن واعتماداتها الـ 3
# العقد بيخلص لما ياخد 3 اعتمادات: ختم الشؤون القانونية + توقيع المقاول نفسه + اعتماد رئيس مجلس الإدارة.
# كل نسخة بتيجي في الـ DC بتتسأل عليها، وملفها في فولدر المقاول\CONTRACT وسطرها في 10-LOG SUB (شيت SUB_CONTRACTS)
CT_SHEET = 'SUB_CONTRACTS'
APPROVALS = (('legal', 'Legal Affairs', 'Legal'), ('sub', 'Subcontractor Signed', 'Sub Signature'),
             ('chair', 'Chairman Approved', 'Chairman'))          # (المفتاح، العنوان في السجل، الاسم المختصر)
CT_KEYS = {'no': 'Contract No.', 'date': 'Date', 'sub': 'Subcontractor', 'legal': 'Legal Affairs',
           'subsig': 'Subcontractor Signed', 'chair': 'Chairman Approved', 'missing': 'Missing', 'file': 'File'}
CT_HDR = ['Sr'] + list(CT_KEYS.values())
CT_WIDTH = (5, 16, 12, 26, 14, 14, 14, 36, 70)
_CT_K = {'legal': 'legal', 'sub': 'subsig', 'chair': 'chair'}


def ct_stage(appr):
    """اسم المرحلة في اسم الملف: Unsigned / Legal / Legal + Sub Signature / Fully Approved"""
    done = [short for k, _, short in APPROVALS if appr.get(k)]
    if len(done) == len(APPROVALS): return 'Fully Approved'
    return ' + '.join(done) if done else 'Unsigned'


def ct_missing(appr):
    return [short for k, _, short in APPROVALS if not appr.get(k)]


def ct_file_no(no):
    """رقم العقد في اسم الملف: 2610/0424 ← 2610-0424"""
    return re.sub(r'[\\/:*?"<>|\s]+', '-', str(no or '').strip()).strip('-')


def _ct_cols(ws):
    for i in range(1, 21):
        vals = {}
        for j in range(1, min(ws.max_column or 60, 60) + 1):
            t = str(ws.cell(i, j).value or '').strip()
            if t and t not in vals: vals[t] = j
        if 'Contract No.' in vals and 'Chairman Approved' in vals:
            miss = [h for h in CT_KEYS.values() if h not in vals]
            if miss:
                raise RuntimeError(f"10-LOG SUB, sheet {CT_SHEET}: «{'», «'.join(miss)}» not found in the titles row – "
                                   'put the title back as it was')
            return i, {k: vals[h] for k, h in CT_KEYS.items()}
    raise RuntimeError(f'10-LOG SUB, sheet {CT_SHEET}: the titles row (Contract No. …) was not found')


def _ct_formula(r, col=None):
    """عمود Missing: الناقص من الـ 3 اعتمادات (من غير TEXTJOIN عشان يشتغل على أي إكسل)"""
    L = lambda k, d: openpyxl.utils.get_column_letter(col[k]) if col else d
    b, e, f, g = L('no', 'B'), L('legal', 'E'), L('subsig', 'F'), L('chair', 'G')
    return (f'=IF({b}{r}="","",IF(AND({e}{r}<>"",{f}{r}<>"",{g}{r}<>""),"Complete ✔",'
            f'MID(IF({e}{r}="",", Legal Affairs","")&IF({f}{r}="",", Sub Signature","")&IF({g}{r}="",", Chairman",""),3,200)))')


def ct_sheet(wb):
    """شيت SUB_CONTRACTS – لو مش موجود بيتعمل بنفس شكل شيت مذكرات الخصم (العناوين في صف 11)"""
    if CT_SHEET in wb.sheetnames: return wb[CT_SHEET]
    from openpyxl.formatting.rule import FormulaRule
    ref = next((wb[n] for n in ('Contracts & Invoices', 'Sub Approvals') if n in wb.sheetnames), None)
    ws = wb.create_sheet(CT_SHEET)
    navy, grey = Font(bold=True, color='1F4E78'), Font(color='808080')
    for r in range(2, 6):
        src = ref.cell(r, 1) if ref is not None else None
        if src is not None and src.value:
            x = ws.cell(r, 1, src.value); x.font = copy(src.font)
    ws.cell(7, 1, 'Subcontracts – the 3 approvals').font = Font(size=14, bold=True, color='1F4E78')
    ws.cell(8, 1, 'Updated: ').font = grey
    ws.cell(8, 2, '=TODAY()').number_format = 'dd-mm-yyyy'
    ws.cell(9, 1, 'A subcontract is complete with: Legal Affairs stamp + the subcontractor\'s signature + the Chairman\'s approval '
                  '(date = when the program saw it)   |   File: 13_PARTIES\\SUB\\<NN-SUB_NAME>\\03_CONTRACT\\   |   '
                  'Added by DC (Document Control)').font = grey
    h, last = DM_HEAD, DM_LAST
    ws.cell(1, 8, 'Contracts').font = navy
    ws.cell(2, 8, 'Complete').font = Font(bold=True)
    ws.cell(2, 9, f'=COUNTIF($H${h + 1}:$H${last},"Complete ✔")').font = Font(bold=True)
    ws.cell(3, 8, 'Missing approvals').font = Font(bold=True)
    ws.cell(3, 9, f'=COUNTA($B${h + 1}:$B${last})-I2').font = Font(bold=True)
    hs = ref.cell(h, 1) if ref is not None and ref.cell(h, 1).value else None
    h_font = copy(hs.font) if hs is not None else Font(bold=True, color='FFFFFF')
    h_fill = copy(hs.fill) if hs is not None else PatternFill('solid', fgColor='1F4E78')
    h_al = copy(hs.alignment) if hs is not None else Alignment(horizontal='center', vertical='center', wrap_text=True)
    h_bd = copy(hs.border) if hs is not None else BOX
    for j, (t, w) in enumerate(zip(CT_HDR, CT_WIDTH), 1):
        x = ws.cell(h, j, t); x.font, x.fill, x.alignment, x.border = h_font, h_fill, h_al, h_bd
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.row_dimensions[h].height = 32
    bs = ref.cell(h + 1, 2) if ref is not None else None
    bd = copy(bs.border) if bs is not None and bs.border.left.style else BOX
    cen = Alignment(horizontal='center')
    for r in range(h + 1, last + 1):
        ws.cell(r, 1, f'=IF(B{r}="","",ROW()-{h})')
        ws.cell(r, 8, _ct_formula(r))
        for j in range(1, len(CT_HDR) + 1):
            x = ws.cell(r, j); x.border = bd
            if j in (1, 2, 3, 5, 6, 7): x.alignment = cen
            if j in (3, 5, 6, 7): x.number_format = 'dd-mm-yyyy'
    for f, fill, fc in ((f'$H{h + 1}="Complete ✔"', 'E2F2E6', '2C7536'), (f'AND($B{h + 1}<>"",$H{h + 1}<>"Complete ✔")', 'FFF1D6', '7A5200')):
        ws.conditional_formatting.add(f'H{h + 1}:H{last}', FormulaRule(
            formula=[f], fill=PatternFill(start_color=fill, end_color=fill, fill_type='solid'), font=Font(bold=True, color=fc)))
    ws.auto_filter.ref = f'A{h}:I{last}'
    ws.freeze_panes = f'C{h + 1}'
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f'{h}:{h}'
    return ws


def _link(dst):
    """لينك الملف من مكان 10-LOG SUB (نسبي، عشان يفضل شغال لو الدرايف اتغير)"""
    try:
        rel = os.path.relpath(dst, os.path.dirname(log_path())).replace(os.sep, '/')
    except ValueError:
        rel = 'file:///' + dst.replace(os.sep, '/')
    return quote(rel, safe="/:._-()&,+'")


def contract_record(sub, no, date, appr, dst, today=None):
    """نسخة عقد اتوزعت ← سطر العقد في SUB_CONTRACTS (رقم العقد + المقاول): الاعتماد الجديد ياخد تاريخ النهارده،
    اللي اتعلّم قبل كده بيفضل (مفيش نسخة بتشيل اعتماد)، ولينك الملف = آخر نسخة. يرجّع الناقص."""
    today = today or dt.date.today()
    no, sub = str(no).strip(), str(sub).strip()
    date = todate(date)
    out = {}

    def fn(ws):
        h, col = _ct_cols(ws)
        r, free = None, None
        for i in range(h + 1, max(ws.max_row, DM_LAST) + 1):
            v = ws.cell(i, col['no']).value
            if v in (None, ''):
                if free is None: free = i
                continue
            if _same(v, no) and _same(ws.cell(i, col['sub']).value, sub): r = i; break
        if r is None:
            r = free or max(ws.max_row, h) + 1
            if r > DM_LAST and r > h + 1:
                for j in range(1, ws.max_column + 1): ws.cell(r, j)._style = copy(ws.cell(r - 1, j)._style)
                ws.cell(r, col['missing'], _ct_formula(r, col))
            ws.cell(r, col['no'], no); ws.cell(r, col['sub'], sub)
        if date and not ws.cell(r, col['date']).value:
            ws.cell(r, col['date'], date).number_format = 'dd-mm-yyyy'
        have = {}
        for k, _, _ in APPROVALS:
            c = ws.cell(r, col[_CT_K[k]])
            if appr.get(k) and c.value in (None, ''):
                c.value = today; c.number_format = 'dd-mm-yyyy'
            have[k] = c.value not in (None, '')
        c = ws.cell(r, col['file'], os.path.basename(dst))
        c.hyperlink = _link(dst); c.font = Font(color='0563C1', underline='single')
        out['missing'] = ct_missing(have)
        return 1
    _log_edit(fn, ct_sheet, CT_SHEET)
    return out['missing']


def contracts(sub=None, path=None):
    """العقود اللي في SUB_CONTRACTS (كلها أو بتاعة مقاول واحد) ← [{no, date, sub, legal, sub_signed, chair, missing, file}]"""
    p = path or log_path()
    if not p or not os.path.isfile(p): return []
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    try:
        if CT_SHEET not in wb.sheetnames: return []
        rows = [list(r) for r in wb[CT_SHEET].iter_rows(values_only=True)]
    finally:
        wb.close()
    h = col = None
    for i, row in enumerate(rows[:20], 1):
        vals = {str(v or '').strip(): j for j, v in reversed(list(enumerate(row, 1))) if v}
        if 'Contract No.' in vals and 'Chairman Approved' in vals and all(t in vals for t in CT_KEYS.values()):
            h, col = i, {k: vals[t] for k, t in CT_KEYS.items()}; break
    if not h: return []
    out = []
    for i, r in enumerate(rows[h:], h + 1):
        g = lambda k: r[col[k] - 1] if col[k] - 1 < len(r) else None
        no = str(g('no') or '').strip()
        if not no: continue
        have = {k: g(_CT_K[k]) not in (None, '') for k, _, _ in APPROVALS}
        d = dict(row=i, no=no, date=todate(g('date')) or g('date'), sub=str(g('sub') or '').strip(), have=have,
                 dates={k: g(_CT_K[k]) for k, _, _ in APPROVALS}, missing=ct_missing(have), file=str(g('file') or '').strip())
        if sub is None or _same(d['sub'], _owner(sub)) or _same(d['sub'], sub.get('Name')): out.append(d)
    return out
