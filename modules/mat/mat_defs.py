# -*- coding: utf-8 -*-
"""MAR / MIR – الثوابت والدوال المشتركة (الأعمدة والتخمينات).

اتشالت من mat_app.py (الشاشة القديمة) عشان mat_qt.py ما يبقاش محتاجها. 21-11-2026
مفيش tkinter هنا خالص.
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import mat_core as M

TYPE_WORDS = (('مواسير', ('ماسور', 'مواسير', 'pipe')), ('محابس', ('محبس', 'صمام', 'valve')),
              ('قطع (Fittings)', ('كوع', 'مشترك', 'مسلوب', 'وصله', 'وصلة', 'fitting', 'tee', 'bend')),
              ('مسامير وجوانات', ('جوان', 'مسامير', 'gasket', 'bolt')), ('ملحقات', ('غلاف', 'شحم', 'sleeve', 'lubricant')))

def fd(d): return d.strftime('%d-%m-%Y') if d else ''

def guess_type(texts, types):
    t = ' '.join(texts).lower()
    for name, words in TYPE_WORDS:
        if any(w in t for w in words) and name in types: return name
    return 'مواد عامة' if 'مواد عامة' in types else (types[-1] if types else '')

def best_company(name, companies):
    """اسم الشركة زي ما هو في قايمة الشركات (SUPPLIER_A، Sample Supplier…) – ولو جديدة: الاسم من غير Trading Company/شركة."""
    nt = M._toks(name)
    for c in companies:
        ct = M._toks(c)
        if ct and ct <= nt | M._toks(name.replace('-', ' ')): return c
    import re
    short = re.sub(r'(?i)\b(trading|company|co\.?|ltd\.?|llc|est\.?|contracting|industrial|for|and)\b', ' ', name or '')
    short = re.sub(r'(شركة|مؤسسة|للتجارة|للمقاولات|المحدودة)', ' ', short)
    return ' '.join(short.split()).strip(' .,-') or (name or '').strip()

# كلام إذن التوريد الإنجليزي ← كلمة المخزن العربي (عشان البند يتطابق مع صنفه في المخزون)
EN_AR = {'pipe': 'مواسير', 'pipes': 'مواسير', 'ductile': 'دكتايل', 'gasket': 'جوان', 'gaskets': 'جوان', 'sleeve': 'غلاف',
         'slv': 'غلاف', 'polyethylene': 'غلاف', 'lubricant': 'شحم', 'grease': 'شحم', 'valve': 'محبس', 'gate': 'بوابه',
         'butterfly': 'فراشه', 'tape': 'شريط', 'warning': 'تحذير', 'elbow': 'كوع', 'bend': 'كوع', 'tee': 'مشترك',
         'reducer': 'مسلوب', 'coupling': 'وصله', 'hdpe': 'hdpe'}


def _item_toks(s):
    """كلمات الوصف + الأرقام (300، C40) + ترجمة الكلام الإنجليزي."""
    import re
    s = (s or '').lower()
    t = M._toks(s)
    t |= {EN_AR[w] for w in re.findall(r'[a-z]+', s) if w in EN_AR}
    t |= {'d.i'.replace('.', '') if w in ('d.i', 'di') else w for w in re.findall(r'\b(?:c\d{2}|d\.?i|\d{2,4})\b', s)}
    if 'di' in t: t.add('دكتايل')
    return {w.translate(M._AR) for w in t}


def best_item(desc, items, code='', deliveries=()):
    """المادة زي ما هي في شيت المخزون (عشان الرصيد يتجمع صح) – وإلا الوصف زي ما هو.
    الأول: نفس كود المورد اتورد قبل كده (التوريدات، أول الملاحظة) ← نفس الصنف. بعدها: كلام + أرقام الوصف."""
    names = {it['name'] for it in items}
    c = str(code or '').strip()
    if c:
        for d in deliveries:
            if d.get('name') in names and str(d.get('note') or '').split(' – ')[0].strip() == c: return d['name']
    dt_ = _item_toks(desc)
    best, sc = None, 0
    for it in items:
        s = len(dt_ & _item_toks(it['name']))
        if s > sc: best, sc = it, s
    return best['name'] if best and sc >= 2 else (desc or '').strip()

def route_choices(stock):
    """الخطوط من سجل الطلبات (الشارع – الخط) + الأماكن اللي اتصرف لها قبل كده."""
    out = []
    try:
        import openpyxl, warnings, tempfile, shutil
        src = M.P(r'00_Registers\01-LOG WIR.xlsm')
        fd_, tmp = tempfile.mkstemp(suffix='.xlsm'); os.close(fd_); shutil.copy2(src, tmp)
        try:
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
            for r in wb['SW'].iter_rows(min_row=4, max_col=5, values_only=True):
                try: int(r[1])
                except (TypeError, ValueError): continue
                street, line = M._s(r[3]), M._s(r[4])
                t = ' – '.join(x for x in (street, f'Line {line}' if line else '') if x)
                if t and t not in out: out.append(t)
            wb.close()
        finally:
            os.remove(tmp)
    except Exception:
        pass
    for x in stock['issues']:
        if x['route'] and x['route'] not in out: out.append(x['route'])
    return out

MIR_COLS = [('no', 'MIR No.', 130), ('date', 'Date', 85), ('desc', 'Material', 300), ('qty', 'Qty', 60), ('unit', 'Unit', 90),
            ('dn', 'Delivery Note', 210), ('mat', 'MAR', 150), ('stage', 'Stage', 125), ('unf', 'Unifier', 105), ('file', 'File', 300)]

MAR_COLS = [('no', 'No.', 50), ('ref', 'Submittal', 150), ('desc', 'Material', 330), ('sup', 'Supplier', 220), ('sent', 'Sent', 85),
            ('reply', 'Reply', 85), ('status', 'Status', 230), ('stage', 'Stage', 95), ('unf', 'Unifier', 105), ('file', 'File', 300)]

STOCK_COLS = [('name', 'Material', 300), ('code', 'Item code', 95), ('unit', 'Unit', 55), ('type', 'Type', 110), ('need', 'Required', 75),
              ('got', 'Received', 75), ('wait', 'Pending', 70), ('out', 'Issued', 70), ('bal', 'Balance', 75), ('left', 'To deliver', 80),
              ('status', 'Status', 140)]

DEL_COLS = [('tr', 'TR', 65), ('date', 'Date', 85), ('dn', 'DN No.', 135), ('company', 'Company', 130), ('type', 'Type', 105),
            ('name', 'Material', 270), ('qty', 'Qty', 65), ('unit', 'Unit', 55), ('status', 'Status', 65), ('mir', 'MIR', 130), ('dl', 'File', 230)]

ISS_COLS = [('date', 'Date', 85), ('name', 'Material', 300), ('qty', 'Qty', 70), ('unit', 'Unit', 55), ('route', 'Line / Location', 220),
            ('who', 'Receiver', 150), ('note', 'Notes', 220)]
