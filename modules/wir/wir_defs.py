# -*- coding: utf-8 -*-
"""WIR – الثوابت والدوال المشتركة (الأعمدة والعلامات).

اتشالت من wir_app.py (الشاشة القديمة) عشان wir_qt.py ما يبقاش محتاجها. 21-11-2026
مفيش tkinter هنا خالص.
"""
import os, json

HERE = os.path.dirname(os.path.abspath(__file__))

NAVY = '#1F4E78'

STATUS_FILE = os.path.join(HERE, 'wir_status.json')

COLS = [('seq', 'م', 40), ('no', 'رقم الطلب', 70), ('dist', 'الحي', 90), ('street', 'الشارع', 120), ('line', 'رقم الخط', 65),
        ('desc', 'الوصف', 160), ('a', 'FROM', 70), ('b', 'TO', 70), ('Ls', 'L', 70), ('dia', 'القطر', 55), ('rev', 'REV', 45),
        ('req', 'تاريخ الطلب', 90), ('app', 'تاريخ التسليم', 95), ('unf', 'رقم اليونيفاير', 100), ('code', 'الكود', 80), ('up', 'اترفع', 55), ('upd', 'تاريخ الرفع', 90),
        ('title', 'الوصف الكامل', 430), ('qty', 'الكمية', 95), ('unit', 'الوحدة', 55), ('item', 'البند', 110), ('note', 'ملاحظات', 320)]

COLS_T = {c: t for c, t, _ in COLS}

def ukey(r): return '|'.join(str(r.get(k) or '').strip() for k in ('line', 'code', 'a_raw', 'b_raw', 'dist', 'street'))

def load_marks():
    try:
        import json
        with open(STATUS_FILE, encoding='utf-8') as f: return json.load(f)
    except Exception: return {}

def save_marks(m):
    import json
    tmp = STATUS_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(m, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATUS_FILE)

def f3(v): return '' if v is None else f'{v:,.3f}'

def fd(d): return d.strftime('%d-%m-%Y') if d else ''

def s_(v):
    if v is None: return ''
    if isinstance(v, float): return f'{v:g}'
    return str(v)

def cells(r):
    """{col: (shown text, copied text, excel value)}"""
    q = r['qty']
    d = dict(seq=r['seq'], no=r['no'], dist=r['dist'], street=r['street'], line=r['line'], desc=r['desc'], a=r['a_raw'], b=r['b_raw'],
             Ls=r['L_raw'], dia=r['dia'], rev=r['rev'], unf=r['unf'], code=r['code'], title=r['title'], unit=r['unit'],
             item=r['item'], note=r['note'])
    out = {k: (s_(v), s_(v), v) for k, v in d.items()}
    out['req'] = (fd(r['reqdate']), fd(r['reqdate']), r['reqdate'])
    out['app'] = (fd(r['appdate']), fd(r['appdate']), r['appdate'])
    up = r.get('upd')
    out['up'] = ('☑' if up else '☐', '', None if not up else 'Yes')
    out['upd'] = (up or '', up or '', up or None)
    out['qty'] = (f3(q), '' if q is None else f'{q:.3f}', None if q is None else round(q, 3))
    return out
