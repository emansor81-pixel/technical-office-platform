# -*- coding: utf-8 -*-
"""ورق المورد من فولدره في 13_PARTIES + مطابقته ببنود الـ submittal (طلب المستخدم 21-11-2026). من غير Qt – قراية بس.

الفكرة: المستخدم بيفتح للمورد فولدر وبيحط فيه ورقه (Company Profile / Company Documents / Previous Approval …).
وهو بيعمل MAR يختار المورد، ورقه يظهر، واللي اسمه واضح يترفق في بنده لوحده.
"""
import os, re

ATTACH_EXT = ('.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp')
SUBDIRS = ('02_APPROVAL', '01_QUOTATIONS')           # اللي بيتعرض من فولدر المورد (الباقي عقود ومستخلصات)
# (كلمات في اسم البند، كلمات في اسم الملف) – الترتيب مهم: الأخص الأول
RULES = [
    (('profile',), ('profile', 'بروفايل', 'بروفيل')),
    (('data sheet', 'datasheet'), ('data sheet', 'datasheet', 'data shet', 'tds', 'technical submit')),
    (('specification', 'spec'), ('spec', 'مواصف')),
    (('boq', 'po items'), ('boq', 'أمر الشراء', 'امر الشراء', 'بنود')),
    (('test', 'mill'), ('mill', 'test', 'اختبار')),
    (('previous',), ('previous', 'اعتماد', 'approval')),
    (('cr ', 'cr+', 'certificates', 'zakat', 'vat'), ('company documents', 'سجل', 'commercial', 'ضريب', 'vat', 'zakat', 'زكا',
                                                      'certificate', 'شهاد', 'iso', 'عنوان وطني', 'مستندات')),
    (('catalogue', 'drawings'), ('catalog', 'كتالوج', 'drawing')),
]


def _skip(n):
    n = n.lower()
    return (n.startswith(('.', '~', '_')) or '.bak-' in n or n in ('readme.txt', 'desktop.ini', 'thumbs.db'))


def is_package(name):
    """ملف باكدج اتعمل قبل كده (اسمه بيبدأ برقم الـ DC أو رقم M) – مايترفقش تلقائي."""
    return bool(re.match(r'^\s*(CTR-|M\s*\d+\s*-)', name, re.I))


def list_files(party_dir, depth=3):
    """ورق المورد اللي ينفع يترفق: [dict(name, sub, path, mtime, package)] – الأحدث في 1_WIP الأول."""
    out = []

    def walk(d, rel, lvl):
        try: names = sorted(os.listdir(d))
        except OSError: return
        for n in names:
            if _skip(n): continue
            p = os.path.join(d, n)
            if os.path.isdir(p):
                if lvl < depth: walk(p, os.path.join(rel, n), lvl + 1)
            elif n.lower().endswith(ATTACH_EXT):
                try: mt = os.path.getmtime(p)
                except OSError: mt = 0
                out.append(dict(name=n, sub=rel, path=p, mtime=mt, package=is_package(n)))

    for s in SUBDIRS:
        walk(os.path.join(party_dir, s), s, 1)
    out.sort(key=lambda f: (f['package'], not f['sub'].replace('/', '\\').endswith('1_WIP'), f['sub'], f['name'].lower()))
    return out


def _rule_of_item(name):
    n = ' ' + name.lower() + ' '
    for i, (req, _) in enumerate(RULES):
        if any(k in n for k in req): return i
    return None


def _rule_of_file(name):
    n = name.lower()
    for i, (_, words) in enumerate(RULES):
        if any(w in n for w in words): return i
    return None


def auto_match(files, items):
    """{رقم البند (n): مسار الملف} – ملف واحد لكل بند: الأحدث من اللي اسمه بيدل عليه. الباكدجات القديمة برّه."""
    by_rule = {}
    for f in files:
        if f['package']: continue
        r = _rule_of_file(f['name'])
        if r is None: continue
        if r not in by_rule or f['mtime'] > by_rule[r]['mtime']: by_rule[r] = f
    out = {}
    for it in items:
        r = _rule_of_item(it.get('name', ''))
        if r is not None and r in by_rule: out[it.get('n')] = by_rule[r]['path']
    return out
