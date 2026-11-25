# -*- coding: utf-8 -*-
"""13_PARTIES | الجهات – فولدر واحد لكل مقاول / مورد / خلاطة / مختبر، وجواه كل اللي يخصه
(طلب المستخدم 12-11):

13_PARTIES\\<CAT>\\<NN>-<ENGLISH_NAME>\\
    01_QUOTATIONS            عروض الأسعار
    02_APPROVAL\\1_WIP … 4_ARCHIVED   ملف الاعتماد
    03_CONTRACT              العقد
    04_CERTIFICATES          المستخلصات
    05_LETTERS               الخطابات
    06_DEDUCTION_MEMOS · 07_ADDITION_MEMOS · 08_FINAL_ACCOUNT   (قسم Subcontractors بيعملهم وقت الحاجة)

<CAT>: SUB مقاولين باطن · SUP موردين · PLANT خلاطات · LAB مختبرات.
الرقم واحد لكل الجهات (مابيتكررش). الاسم العربي للجهة في README.txt جوه فولدرها (سطر «AR: …») – منه البرنامج
بيعرف ورق الجهة اللي مكتوب بالعربي. الـ MAR (اعتماد المواد) بيفضل في 04_Submittals\\MAR.
"""
import os, re

ROOT_REL = '13_PARTIES'
CATS = ('SUB', 'SUP', 'PLANT', 'LAB')
CAT_NAMES = {'SUB': 'Subcontractors', 'SUP': 'Suppliers', 'PLANT': 'Batching plants', 'LAB': 'Laboratories'}
QUOTES, APPROVAL, CONTRACT, CERTS, LETTERS = '01_QUOTATIONS', '02_APPROVAL', '03_CONTRACT', '04_CERTIFICATES', '05_LETTERS'
DMEMOS, AMEMOS, FINAL = '06_DEDUCTION_MEMOS', '07_ADDITION_MEMOS', '08_FINAL_ACCOUNT'
BASE_DIRS = (QUOTES, APPROVAL, CONTRACT, CERTS, LETTERS)
STAGES = ('1_WIP', '2_SHARED', '3_PUBLISHED', '4_ARCHIVED')
# الأسامي القديمة (فولدر المقاول في 08_Cost_Payments\SUBCONTRACTORS) ← الجديدة
OLD_DIRS = {'CONTRACT': CONTRACT, 'CERTIFICATES': CERTS, 'DEDUCTION_MEMOS': DMEMOS, 'ADDITION_MEMOS': AMEMOS,
            'FINAL_ACCOUNT': FINAL, '1-Contract': CONTRACT, '2-Certificates': CERTS, '3-Deduction Memos': DMEMOS,
            '4-Addition Memos': AMEMOS, '5-Final Account': FINAL}
DIR_RX = re.compile(r'^\s*(\d+)\s*-\s*(.+?)\s*$')


def key(s):
    """للمقارنة: حروف وأرقام بس (إنجليزي صغير / عربي من غير ال والهمزات)."""
    s = str(s or '').lower()
    s = re.sub('[إأآا]', 'ا', s).replace('ة', 'ه').replace('ى', 'ي')
    return re.sub(r'[^0-9a-zء-ي]', '', s)


def en_folder(name):
    """اسم الفولدر الإنجليزي: «Sample Subcontractor» ← SAMPLE_SUBCONTRACTOR"""
    t = re.sub(r"[^A-Za-z0-9&]+", '_', str(name or '')).strip('_').upper()
    return re.sub(r'_+', '_', t)[:30].strip('_')


def root(project): return os.path.join(project, ROOT_REL)


def _aliases(d):
    out = []
    try:
        with open(os.path.join(d, 'README.txt'), encoding='utf-8-sig') as f:
            for ln in f:
                m = re.match(r'\s*(AR|ALIAS)\s*:\s*(.+)', ln, re.I)
                if m: out += [a.strip() for a in re.split(r'[|،,]', m.group(2)) if a.strip()]
    except OSError:
        pass
    return out


def all_parties(project):
    """كل الجهات: dict(cat, no, name, dir (اسم الفولدر), path, rel, aliases)."""
    out, r = [], root(project)
    for cat in CATS:
        d = os.path.join(r, cat)
        try: names = sorted(os.listdir(d))
        except OSError: continue
        for n in names:
            p = os.path.join(d, n)
            if not os.path.isdir(p) or n.startswith(('_', '.')): continue
            m = DIR_RX.match(n)
            out.append(dict(cat=cat, no=int(m.group(1)) if m else 0, name=(m.group(2) if m else n), dir=n, path=p,
                            rel=os.path.join(ROOT_REL, cat, n), aliases=_aliases(p)))
    return out


def find(project, name):
    """الجهة بالاسم الإنجليزي (أو اسم الفولدر كامل، أو الاسم العربي اللي في README)، أو None."""
    k = key(re.sub(r'^\s*\d+\s*-\s*', '', str(name or '')))
    k2 = key(name)
    if not k: return None
    for p in all_parties(project):
        if k in (key(p['name']), key(p['dir'])) or k2 == key(p['dir']) or k in [key(a) for a in p['aliases']]:
            return p
    en = re.sub(r'\([^)]*[؀-ۿ][^)]*\)|[؀-ۿ]+', ' ', str(name or ''))      # «شركة مقاول باطن (مثال) (Sample Subcontractor)»
    m = re.search(r'\(([^)]*[A-Za-z][^)]*)\)', str(name or ''))
    for cand in filter(None, (m.group(1) if m else '', en)):
        ke = key(re.sub(r'\b(co|company|est|ltd|llc|corp)\b\.?', ' ', cand, flags=re.I))
        if len(ke) >= 4:
            for p in all_parties(project):
                if ke == key(p['name']): return p
    return None


def next_no(project):
    return max([p['no'] for p in all_parties(project)] or [0]) + 1


def new_rel(project, name, cat='SUB'):
    """مسار جهة جديدة (من غير ما يتعمل): 13_PARTIES\\SUB\\16-NAME"""
    n = en_folder(name) or re.sub(r'[\\/:*?"<>|]+', '-', str(name or '').strip())[:30] or 'NEW'
    return os.path.join(ROOT_REL, cat if cat in CATS else 'SUB', f'{next_no(project):02d}-{n}')


def ensure(path, arabic=''):
    """يعمل فولدر الجهة بالفولدرات الأساسية (ولو فيه أسامي قديمة زي CONTRACT بتتغير)."""
    os.makedirs(path, exist_ok=True)
    for old, new in OLD_DIRS.items():
        po, pn = os.path.join(path, old), os.path.join(path, new)
        if os.path.isdir(po):
            try:
                if not os.path.exists(pn): os.rename(po, pn)
                else:
                    for n in os.listdir(po):
                        if not os.path.exists(os.path.join(pn, n)): os.rename(os.path.join(po, n), os.path.join(pn, n))
                    os.rmdir(po)
            except OSError:
                pass
    for d in BASE_DIRS: os.makedirs(os.path.join(path, d), exist_ok=True)
    for st in STAGES: os.makedirs(os.path.join(path, APPROVAL, st), exist_ok=True)
    rd = os.path.join(path, 'README.txt')
    if arabic and not os.path.exists(rd):
        try:
            with open(rd, 'w', encoding='utf-8-sig') as f: f.write(f'AR: {arabic}\r\n')
        except OSError:
            pass
    return path
