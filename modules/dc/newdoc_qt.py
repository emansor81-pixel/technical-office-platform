# -*- coding: utf-8 -*-
"""New Document | الخطابات الصادرة بـ Qt.
تختار الجهة (الاستشاري / الجهة المالكة / موردين ومقاولين / جهة تانية) ← الرقم اللي عليه الدور من شيت DC ←
تكتب الموضوع والنص ← «Make letter» يعمل Word على قالب الشركة (00_Registers\\00-CTR-YA.docx – الهيدر والفوتر بتوعه)
و PDF (من Word نفسه)، في 02_Correspondence\\OUT\\<الجهة>، ويكتب السطر في DC.xlsx.
القالب مابيتلمسش: بناخد أول خطاب فيه كـ«شكل» (الخطوط والمسافات) ونحط الكلام الجديد مكانه."""
import os, re, sys, json, time, shutil, zipfile, tempfile, threading, subprocess, traceback, datetime as dt
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
PLAT = os.path.dirname(os.path.dirname(HERE))
if PLAT not in sys.path: sys.path.insert(0, PLAT)

from PySide6.QtCore import Qt, QTimer, QDate
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (QWidget, QFrame, QPushButton, QStackedWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QListWidget, QListWidgetItem,
                               QComboBox, QLineEdit, QPlainTextEdit, QDateEdit, QMessageBox, QTextBrowser, QAbstractItemView,
                               QCheckBox, QScrollArea, QApplication)
import home_qt as HQ
import dc_qt as DQ
import dc_core as D
import dc_sheet as SH
from dc_io import LOG

INK, MUTED, NAVY = DQ.INK, DQ.MUTED, DQ.NAVY
TT = HQ.T
CSS = DQ.CSS.replace('QWidget#dist', 'QWidget#newdoc') + f"""
QPlainTextEdit {{ background: {TT['card2']}; color: {INK}; border: 1px solid {DQ.SOFT}; border-radius: 8px; padding: 6px;
                  font: 11pt "Segoe UI"; }}
QDateEdit {{ background: {TT['card2']}; border: 1px solid {DQ.SOFT}; border-radius: 8px; padding: 5px 9px; color: {INK}; }}
QListWidget#types {{ background: transparent; border: none; font: 11pt "Segoe UI"; }}
QListWidget#types::item {{ padding: 10px 12px; border-radius: 9px; margin-bottom: 4px; color: {INK}; }}
QListWidget#types::item:selected {{ background: {NAVY}; color: white; }}
QPushButton#tab {{ background: {TT['card2']}; color: {INK}; border: 1px solid {DQ.SOFT}; border-bottom: none;
                   border-top-left-radius: 9px; border-top-right-radius: 9px; border-bottom-left-radius: 0; border-bottom-right-radius: 0;
                   padding: 5px 10px; font: 9.5pt "Segoe UI"; text-align: left; }}
QPushButton#tab:hover {{ border-color: #3D7DB8; }}
QPushButton#tab:checked {{ background: {NAVY}; border-color: #3D7DB8; color: white; }}
QLabel#grp {{ color: {MUTED}; font: bold 8pt "Segoe UI"; padding: 0 4px 8px 6px; }}
QFrame#tabline {{ background: #3D7DB8; }}
QTextBrowser {{ background: white; color: black; border-radius: 6px; }}
"""
SETTINGS = os.path.join(HERE, 'newdoc.json')
STATUS_FILE = os.path.join(HERE, 'doc_status.json')        # الخطابات اللي اتسلّمت (مقفولة) – {'delivered': {رقم: تاريخ}}
TEMPLATE_REL = r'00_Registers\00-CTR-YA.docx'
PROJECT_NO = '000000000'
PROJECT_LINE = f'عقد تنفيذ شبكات المياه (مثال) - رقم ({PROJECT_NO})'      # رقم المشروع جنب اسمه (طلب المستخدم 21-11)
OPENING = 'إشارة إلى مشروع شبكات المياه (مثال) مقاولتنا بموجب التعميد رقم (000000000)، '
CLOSINGS = ['نأمل التكرم بالمراجعة والاعتماد.', 'نأمل التكرم بالاطلاع والإفادة.', 'نأمل التكرم بالاطلاع.', 'وذلك للعلم والإحاطة.', '']
REGARDS = 'وتفضلوا بقبول فائق الاحترام والتقدير،،،'
# kind: (الاسم، الفولدر في OUT، الجهة، العناية، شكل الرقم لو الشيت مش متاح)
TYPES = [('YA', 'Letter – the Consultant (YA)', 'YA', 'مكتب الاستشاري', 'المهندس المقيم / م. (اسم المهندس)', 'CTR-YA-WA-26-####'),
         ('CLIENT', 'Letter – CLIENT', 'CLIENT', 'الجهة المالكة', '', 'CTR-CLIENT-WA-####'),
         ('SUP', 'Letter – Suppliers / Subs', 'SUB', '', '', 'CTR-SUP C - WA-####'),
         ('OTH', 'Letter – Other', 'OTH', '', '', 'CTR-OTH- WA-###')]
ROLES = ['مدير المشروع', 'مدير المكتب الفني', 'مهندس مكتب فني', 'مدير الموقع']
NAMES = ['م. (اسم المهندس)', 'م. (اسم المهندس)', 'م. (اسم المهندس)', 'مهندس المكتب الفني']


def load_set():
    try:
        with open(SETTINGS, encoding='utf-8') as f: return json.load(f)
    except Exception:
        return {}


def save_set(s):
    try:
        with open(SETTINGS, 'w', encoding='utf-8') as f: json.dump(s, f, ensure_ascii=False, indent=1)
    except Exception:
        traceback.print_exc()


# ---------------------------------------------------------------- الـ Word من القالب (من غير مكتبات زيادة)
def _children(body):
    """عناصر الـ body اللي على أول مستوى (w:p / w:tbl / w:sectPr) كنصوص XML."""
    out, i, n = [], 0, len(body)
    while i < n:
        m = re.compile(r'<(w:p|w:tbl|w:sectPr|w:bookmarkStart|w:bookmarkEnd|w:sdt)\b').search(body, i)
        if not m: break
        tag, s = m.group(1), m.start()
        head_end = body.index('>', s)
        if body[head_end - 1] == '/':                     # <w:p/> فاضي
            out.append(body[s:head_end + 1]); i = head_end + 1; continue
        depth, j = 0, s
        pat = re.compile(rf'<{tag}\b[^>]*?(/?)>|</{tag}>')
        for mm in pat.finditer(body, s):
            if mm.group(0).startswith('</'): depth -= 1
            elif mm.group(1) != '/': depth += 1
            if depth == 0: j = mm.end(); break
        out.append(body[s:j]); i = j
    return out


def _text(p): return ''.join(re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', p))


def _rprs(p):
    """[خط أول جزء (العنوان), خط أطول جزء (الكلام)]"""
    runs = re.findall(r'<w:r\b[^>]*>.*?</w:r>', p, re.S)
    got = []
    for r in runs:
        if _text(r).strip():
            m = re.search(r'<w:rPr>.*?</w:rPr>', r, re.S); got.append((len(_text(r)), m.group(0) if m else ''))
    if not got: return ['', '']
    return [got[0][1], max(got, key=lambda x: x[0])[1]]


def _ppr(p):
    m = re.search(r'<w:pPr>.*?</w:pPr>', p, re.S)
    return m.group(0) if m else ''


def _run(rpr, t): return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>'


def _para(model, parts):
    """فقرة بنفس شكل الموديل: parts = [(نص, 'label'|'value')] – الـ label بخط أول جزء، والـ value بخط آخر جزء."""
    r = _rprs(model)
    return '<w:p>' + _ppr(model) + ''.join(_run(r[0] if k == 'label' else r[-1], t) for t, k in parts if t) + '</w:p>'


LRE, PDF_ = '‪', '‬'                     # التاريخ والأرقام من الشمال لليمين جوه السطر العربي


def _rp(size, bold=False, font='Arial'):
    b = '<w:b/><w:bCs/>' if bold else ''
    return (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>{b}'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:rtl/></w:rPr>')


def _pp(jc='right', after=0, line=276, extra='', ind=''):
    """خصائص الفقرة. الخطاب كله عربي ← الفقرة RTL (bidi)، وفي الفقرة الـ RTL «left» في Word معناها يمين الورقة
    والعكس – عشان كده jc='right' (المطلوب: يمين) بيتكتب left. (قبل كده السطور كانت بتروح شمال – ملاحظة المستخدم 21-11)"""
    j = {'right': 'left', 'left': 'right'}.get(jc, jc)
    j = f'<w:jc w:val="{j}"/>' if j else ''
    return f'<w:pPr>{extra}<w:bidi/><w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="auto"/>{ind}{j}</w:pPr>'


def build_docx(template, out, L, step=0):
    """الخطاب على ورق القالب (الهيدر والفوتر بتوعه بس) – الكلام كله بشكل ثابت مكتوب هنا.
    step: بالسالب أكبر، بالموجب أصغر (make_files بيختار أكبر مقاس يخلي الخطاب في صفحة واحدة)."""
    with zipfile.ZipFile(template) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    b0 = xml.index('<w:body>') + len('<w:body>'); b1 = xml.rindex('</w:body>')
    kids = _children(xml[b0:b1])
    sect = next((k for k in reversed(kids) if k.startswith('<w:sectPr')), '')
    if not sect: raise RuntimeError('The letter template has no page settings (00-CTR-YA.docx).')
    pg = re.search(r'<w:pgSz[^>]*w:w="(\d+)"', sect); ml = re.search(r'<w:pgMar[^>]*w:left="(\d+)"', sect)
    mr = re.search(r'<w:pgMar[^>]*w:right="(\d+)"', sect)
    tw = (int(pg.group(1)) - int(ml.group(1)) - int(mr.group(1))) if (pg and ml and mr) else 9700
    # step: 0 عادي، بالسالب أكبر (الخطاب قصير يملا الورقة)، بالموجب أصغر (ييجي في صفحة واحدة)
    hs, bs, ts = max(18, 24 - 2 * step), max(20, 28 - 2 * step), max(22, 32 - 2 * step)   # بنص البنط: 12 / 14 / 16
    line = max(228, min(360, 288 - 14 * step)); gap = max(0, min(200, 100 - 25 * step))
    MAJ = 'Sakkal Majalla'
    P = lambda runs, **k: '<w:p>' + _pp(**k) + ''.join(runs) + '</w:p>'
    R = lambda t, size, bold=False, font='Arial': f'<w:r>{_rp(size, bold, font)}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>'
    empty = lambda h=120: f'<w:p>{_pp(after=0, line=240)}<w:r>{_rp(h // 10)}<w:t></w:t></w:r></w:p>'

    def lab(label, value):                              # سطر «العنوان: القيمة» يمين
        return P([R(label, hs, True), R(value, hs, False)], after=0, line=line)

    def two(right, left):                               # «السادة / …» يمين و«المحترمين» شمال على نفس السطر (Tab لآخر السطر)
        ex = f'<w:tabs><w:tab w:val="right" w:pos="{tw}"/></w:tabs>'
        return ('<w:p>' + _pp(jc=None, after=0, line=line, extra=ex) + R(right, bs, True) + f'<w:r>{_rp(bs, True)}<w:tab/></w:r>'
                + R(left, bs, True) + '</w:p>')              # فقرة RTL: أول الكلام يمين، واللي بعد الـ Tab في آخر السطر شمال
    out_p = [empty(),
             lab('المشروع: ', PROJECT_LINE), lab('الموضوع: ', L['subject']),
             lab('التاريخ: ', LRE + L['date'] + PDF_), lab('الرقم: ', LRE + L['no'] + PDF_),
             empty(), two(f"السادة / {L['to']}", 'المحترمين')]
    if L.get('attn'): out_p.append(two(f"عناية {L['attn']}", 'المحترم'))
    out_p += [empty(), P([R('السلام عليكم ورحمة الله وبركاته', bs, True)], jc='center', after=gap, line=line)]
    out_p += [P([R(t, ts, False, MAJ)], jc='both', after=gap, line=line) for t in L['body'] if t.strip()]
    if L.get('attach'): out_p.append(P([R('مرفق طيه: ', ts, True, MAJ), R(L['attach'], ts, False, MAJ)], after=gap, line=line))
    if L.get('closing'): out_p.append(P([R(L['closing'], ts, True, MAJ)], after=gap, line=line))
    sig = f'<w:ind w:left="0" w:right="{tw // 2 + 600}"/>'           # التوقيع في النص الشمال من الصفحة
    out_p += [P([R(REGARDS, bs, True)], jc='center', after=0, line=line), empty(),
              '<w:p>' + _pp(jc='center', after=0, line=line, ind=sig) + R(L['role'], bs, True) + '</w:p>',
              '<w:p>' + _pp(jc='center', after=0, line=line, ind=sig) + R(L['name'], bs, True) + '</w:p>']
    # نص الصفحة بالنسبة لمرجع الرسمة (EMU): التصغير الأفقي للهيدر والفوتر حوالينه – يفضلوا متوسّطين
    pw = int(pg.group(1)) if pg else 11906; lm = int(ml.group(1)) if ml else 1418
    cen = {'page': pw * 635 // 2, 'column': (pw // 2 - lm) * 635, 'margin': (pw // 2 - lm) * 635,
           'leftMargin': (pw // 2 - lm) * 635, 'character': (pw // 2 - lm) * 635}
    hf = {}                                              # الهيدر والفوتر بعد التصغير والترتيب: الاسم ← XML
    with zipfile.ZipFile(template) as zin:
        for n in zin.namelist():
            if re.match(r'word/header\d*\.xml$', n):
                hf[n] = _arrange_header(_shrink(zin.read(n).decode('utf-8'), HEAD_K, cen), cen['column'], pw * 635)
            elif re.match(r'word/footer\d*\.xml$', n):
                hf[n] = _arrange_footer(zin.read(n).decode('utf-8'), cen['column'], pw * 635)
    sect = _small_margins(sect, max([_anchors_bottom(v) for n, v in hf.items() if 'header' in n] or [0]),
                          max([_anchors_bottom(v) for n, v in hf.items() if 'footer' in n] or [0]))
    new = xml[:b0] + ''.join(out_p) + sect + xml[b1:]
    tmp = out + '.part'
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == 'word/document.xml': data = new.encode('utf-8')
            elif it.filename in hf: data = hf[it.filename].encode('utf-8')
            zo.writestr(it, data)
    os.replace(tmp, out)
    return out


HEAD_EDGE = 340                                         # هامش الهيدر من حافة الورقة (twips ≈ 6 مم)
FOOT_SIDE = 170                                         # صورة الفوتر: هامش صغير من الجنبين (≈ 3 مم – أغلب الطابعات مابتطبعش للحافة)
FOOT_EDGE = 200                                         # صورة الفوتر تبعد عن آخر الورقة (twips ≈ 3.5 مم)
HEAD_K, FOOT_K = 0.82, 0.80                             # الهيدر والفوتر أصغر شوية (طلب المستخدم 19-11) – القالب نفسه مابيتغيرش


def _shrink(x, k, center=None):
    """يصغّر كل اللي في الهيدر / الفوتر بنسبة k: المقاسات والأماكن (DrawingML) وخط الكلام اللي جوه الصناديق.
    center = نص الصفحة (EMU) بالنسبة لأول عمود الكلام: التصغير الأفقي حوالين النص – عشان اللوجو والكلام يفضلوا
    متوسّطين في الورقة (قبل كده كان بيصغّر ناحية الشمال فالهيدر بيتزحزح – ملاحظة المستخدم 21-11)."""
    def pos_h(m):
        ref, inner = m.group(1), m.group(2)
        def off(mm):
            v = int(mm.group(1))
            c = center.get(ref) if isinstance(center, dict) else None
            return f'<wp:posOffset>{int(c + (v - c) * k) if c is not None else int(v * k)}</wp:posOffset>'
        return f'<wp:positionH relativeFrom="{ref}">' + re.sub(r'<wp:posOffset>(-?\d+)</wp:posOffset>', off, inner) + '</wp:positionH>'

    def pos_v(m):
        return (f'<wp:positionV relativeFrom="{m.group(1)}">'
                + re.sub(r'<wp:posOffset>(-?\d+)</wp:posOffset>', lambda mm: f'<wp:posOffset>{int(int(mm.group(1)) * k)}</wp:posOffset>', m.group(2))
                + '</wp:positionV>')

    x = re.sub(r'<wp:positionH relativeFrom="(\w+)">(.*?)</wp:positionH>', pos_h, x, flags=re.S)
    x = re.sub(r'<wp:positionV relativeFrom="(\w+)">(.*?)</wp:positionV>', pos_v, x, flags=re.S)
    for pat in (r'(<wp:extent cx=")(\d+)(")', r'(<wp:extent cx="\d+" cy=")(\d+)(")', r'(<a:ext cx=")(\d+)(")', r'(<a:ext cx="\d+" cy=")(\d+)(")',
                r'(<wp14:pctHeight>)(\d+)(</wp14:pctHeight>)'):
        x = re.sub(pat, lambda m: f'{m.group(1)}{int(int(m.group(2)) * k)}{m.group(3)}', x)
    # الخط بينزل لأقرب رقم أصغر: الكلام لازم يكفي صندوقه بعد التصغير (round كان بيقطع «Contracting Co.»)
    x = re.sub(r'(<w:sz(?:Cs)? w:val=")(\d+)(")', lambda m: f'{m.group(1)}{max(8, int(int(m.group(2)) * k))}{m.group(3)}', x)   # sz (إنجليزي) و szCs (عربي) – قبل كده الإنجليزي ماكانش بيصغر
    return x


def _set_anchor(a, cx=None, cy=None, center=None):
    """يغيّر مقاس رسمة (anchor) ومكانها الأفقي: center = نص الرسمة بالنسبة لمرجعها (EMU)."""
    m = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', a)
    if not m: return a
    ocx, ocy = int(m.group(1)), int(m.group(2))
    cx = ocx if cx is None else int(cx); cy = ocy if cy is None else int(cy)
    a = a.replace(m.group(0), f'<wp:extent cx="{cx}" cy="{cy}"', 1)
    a = re.sub(r'<a:ext cx="\d+" cy="\d+"', f'<a:ext cx="{cx}" cy="{cy}"', a)          # مقاس الشكل نفسه
    if center is not None:
        a = re.sub(r'(<wp:positionH relativeFrom="\w+">)(.*?)(</wp:positionH>)',
                   lambda mm: f'{mm.group(1)}<wp:posOffset>{int(center - cx / 2)}</wp:posOffset>{mm.group(3)}', a, count=1, flags=re.S)
    return a


def _arrange_header(x, page_c, page_w):
    """الهيدر مترتّب وبعرض الورقة (طلب المستخدم 21-11): سطور العربي متوسّطة تحت بعضها على اليمين، سطور الإنجليزي
    متوسّطة تحت بعضها على الشمال، اللوجو في النص، والخط بعرض الورقة. page_c = نص الصفحة بالنسبة لمرجع الرسمة (EMU)."""
    edge = HEAD_EDGE * 635                               # هامش من حافة الورقة
    logo = max([int(m) for a in re.findall(r'<wp:anchor\b.*?</wp:anchor>', x, re.S) if '<pic:pic' in a
                for m in re.findall(r'<wp:extent cx="(\d+)"', a)] or [900000])
    side = (page_w - logo) / 2 - edge - 120000           # عرض كل جنب (بين الحافة واللوجو)
    off = logo / 2 + 120000 + side / 2                   # بُعد نص الجنب عن نص الصفحة

    def fix(m):
        a = m.group(0)
        if '<pic:pic' in a: return _set_anchor(a, center=page_c)                         # اللوجو في النص
        if 'txbxContent' not in a: return _set_anchor(a, cx=page_w - 2 * edge, center=page_c)   # الخط بعرض الورقة
        txt = ''.join(re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', a))
        arabic = bool(re.search(r'[\u0600-\u06FF]', txt))
        a = _set_anchor(a, cx=side, center=page_c + (off if arabic else -off))
        # كل سطر جوه الصندوق في النص
        def para(pm):
            p = pm.group(0)
            if '<w:pPr>' in p:
                p = re.sub(r'<w:jc w:val="\w+"/>', '', p)
                return p.replace('</w:pPr>', '<w:jc w:val="center"/></w:pPr>', 1) if '</w:pPr>' in p else p
            return re.sub(r'(<w:p\b[^>]*>)', r'\1<w:pPr><w:jc w:val="center"/></w:pPr>', p, count=1)
        return re.sub(r'<w:p\b[^>/]*>.*?</w:p>', para, a, flags=re.S)

    return re.sub(r'<wp:anchor\b.*?</wp:anchor>', fix, x, flags=re.S)


def _arrange_footer(x, page_c, page_w):
    """صورة الفوتر ماسكة الورقة من اليمين للشمال (بنفس تناسبها)."""
    def fix(m):
        a = m.group(0)
        e = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', a)
        if not e or '<pic:pic' not in a: return a
        w = page_w - 2 * FOOT_SIDE * 635
        return _set_anchor(a, cx=w, cy=w * int(e.group(2)) / max(1, int(e.group(1))), center=page_c)
    return re.sub(r'<wp:anchor\b.*?</wp:anchor>', fix, x, flags=re.S)


def _anchors_bottom(x):
    """آخر نقطة في رسومات الهيدر (twips من أول فقرة الهيدر) – عشان الكلام يبدأ تحتها مش فوقها."""
    low = 0
    for a in re.findall(r'<wp:anchor\b.*?</wp:anchor>', x, re.S):
        v = re.search(r'<wp:positionV relativeFrom="(\w+)">.*?<wp:posOffset>(-?\d+)</wp:posOffset>', a, re.S)
        e = re.search(r'<wp:extent cx="\d+" cy="(\d+)"', a)
        if v and e and v.group(1) in ('paragraph', 'line'): low = max(low, int(v.group(2)) + int(e.group(1)))
    return low // 635


def _small_margins(sect, head_bottom=0, foot_bottom=0):
    """الكلام يبدأ تحت الهيدر الصغير على طول (مش فوقه)، والفوتر ينزل لآخر الورقة.
    head_bottom = آخر رسومات الهيدر بعد التصغير (twips من مسافة الهيدر).
    foot_bottom = ارتفاع رسومات الفوتر بعد التصغير: الفوتر يبعد عن آخر الورقة على قدّها (كانت 280 ثابتة فالصورة
    الصفرا بتتقص من تحت – ملاحظة المستخدم 21-11)، والكلام يقف فوقها."""
    def fix(m):
        t = m.group(0)
        top = int(re.search(r'w:top="(\d+)"', t).group(1)); hd = int(re.search(r'w:header="(\d+)"', t).group(1))
        new_top = int(hd + (top - hd) * HEAD_K)
        if head_bottom: new_top = max(new_top, hd + head_bottom + 60)      # تحت خط الهيدر ورقم التسجيل + مسافة صغيرة
        t = re.sub(r'w:top="\d+"', f'w:top="{new_top}"', t)
        fd = (foot_bottom + FOOT_EDGE) if foot_bottom else 280
        t = re.sub(r'w:footer="\d+"', f'w:footer="{fd}"', t)
        if foot_bottom:
            bot = int(re.search(r'w:bottom="(\d+)"', t).group(1))
            t = re.sub(r'w:bottom="\d+"', f'w:bottom="{max(bot, fd + 300)}"', t)       # الكلام فوق صورة الفوتر
        return t
    return re.sub(r'<w:pgMar\b[^>]*/>', fix, sect)


def pdf_pages(pdf):
    try:
        import pypdf
        return len(pypdf.PdfReader(pdf).pages)
    except Exception:
        return None


def _make_slow(template, docx, pdf, L, allow_two=False):
    """Word + PDF في صفحة واحدة: الخطاب القصير يكبر يملا الورقة، والطويل يصغر لحد ما ييجي في صفحة –
    ومابينزلش صفحة تانية إلا لو allow_two. يرجّع (عدد الصفحات, خطأ الـ PDF)."""
    def run(step):
        build_docx(template, docx, L, step); to_pdf(docx, pdf); return pdf_pages(pdf) or 1
    try:
        if allow_two: return run(0), None
        n = run(0)
        if n <= 1:                                       # فيه مكان ← نكبّر لحد ما يقرب يكمّل الصفحة
            best = last = 0
            for st in (-1, -2, -3):
                last = st
                if run(st) > 1: break
                best = st
            if last != best: run(best)
            return 1, None
        for st in (1, 2, 3, 4, 5):                       # طويل ← نصغّر
            n = run(st)
            if n <= 1: return 1, None
        return n, None
    except Exception as e:
        return None, str(e)


PS_FIT = r"""
$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -Encoding UTF8 $env:ND_JOB | ConvertFrom-Json
function PathOf($s) { return ($job.files | Where-Object { $_.step -eq $s }).path }
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
function Pages($s) {
  $d = $w.Documents.Open((PathOf $s), $false, $true)
  try { $d.Repaginate() | Out-Null; return [int]$d.ComputeStatistics(2) } finally { $d.Close(0) | Out-Null }
}
try {
  $pick = 0; $n = 1
  if ($job.two) { $n = Pages 0 }
  else {
    $n0 = Pages 0
    if ($n0 -le 1) { foreach ($s in -1, -2, -3) { if ((Pages $s) -gt 1) { break }; $pick = $s } }
    else { $n = $n0; $pick = 5; foreach ($s in 1, 2, 3, 4, 5) { $n = Pages $s; if ($n -le 1) { $pick = $s; break } } }
  }
  $d = $w.Documents.Open((PathOf $pick), $false, $true)
  try { $d.ExportAsFixedFormat($job.pdf, 17) | Out-Null } finally { $d.Close(0) | Out-Null }
  "FIT $pick $n"
} finally { if ($w.Documents.Count -eq 0) { $w.Quit() } }      # لو Word مفتوح فيه ملفات تانية مانقفلهوش
"""


def _make_fast(template, docx, pdf, L, allow_two=False, timeout=240):
    """نفس اختيار المقاس بتاع _make_slow، بس Word بيتفتح مرة واحدة: كل المقاسات بتتبني الأول، Word يعدّ صفحات كل
    واحد جوه نفس الفتحة ويطلّع PDF للمختار بس (طلب المستخدم 21-11: Real preview كان بيتأخر). يرجّع عدد الصفحات."""
    if os.name != 'nt': raise RuntimeError('PDF needs Word on Windows')
    d = tempfile.mkdtemp(prefix='nd_fit_')
    try:
        steps = [0] if allow_two else [0, -1, -2, -3, 1, 2, 3, 4, 5]
        files = []
        for st in steps:
            p = os.path.join(d, f's{st + 3}.docx'); build_docx(template, p, L, st); files.append(dict(step=st, path=p))
        tmp_pdf = os.path.join(d, 'out.pdf'); job = os.path.join(d, 'job.json'); ps = os.path.join(d, 'nd_fit.ps1')
        with open(job, 'w', encoding='utf-8') as fh: json.dump(dict(two=bool(allow_two), files=files, pdf=tmp_pdf), fh)
        with open(ps, 'w', encoding='utf-8-sig') as fh: fh.write(PS_FIT)
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps], env=dict(os.environ, ND_JOB=job),
                           capture_output=True, text=True, timeout=timeout, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        m = re.search(r'FIT (-?\d+) (\d+)', r.stdout or '')
        if r.returncode != 0 or not m or not os.path.isfile(tmp_pdf):
            raise RuntimeError((r.stderr or r.stdout or 'Word did not make the PDF').strip()[-400:])
        pick, n = int(m.group(1)), int(m.group(2))
        real = pdf_pages(tmp_pdf)                        # الـ PDF نفسه هو الحكم
        if real is not None: n = real
        if n > 1 and not allow_two and pick != 5: raise RuntimeError('page count mismatch')       # ← الطريقة القديمة تحسم
        src = next(x['path'] for x in files if x['step'] == pick)
        shutil.copy2(src, docx + '.part'); os.replace(docx + '.part', docx)
        shutil.copy2(tmp_pdf, pdf + '.part'); os.replace(pdf + '.part', pdf)
        return n
    finally:
        shutil.rmtree(d, ignore_errors=True)


def make_files(template, docx, pdf, L, allow_two=False):
    """Word + PDF في صفحة واحدة. الأول الطريقة السريعة (Word مرة واحدة)، ولو وقعت لأي سبب ← القديمة المجرّبة.
    يرجّع (عدد الصفحات, خطأ الـ PDF)."""
    try:
        return _make_fast(template, docx, pdf, L, allow_two), None
    except Exception:
        return _make_slow(template, docx, pdf, L, allow_two)


PS_PDF = r'''
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($env:ND_DOCX, $false, $true)
  $d.ExportAsFixedFormat($env:ND_PDF, 17)
  $d.Close(0)
} finally { if ($w.Documents.Count -eq 0) { $w.Quit() } }      # لو Word مفتوح فيه ملفات تانية مانقفلهوش
'''


def to_pdf(docx, pdf, timeout=150):
    """الـ PDF من Word نفسه (عشان الخطوط والهيدر يطلعوا زي الأصل)."""
    if os.name != 'nt': raise RuntimeError('PDF needs Word on Windows')
    d = tempfile.mkdtemp(); ps = os.path.join(d, 'nd_pdf.ps1')
    with open(ps, 'w', encoding='utf-8-sig') as f: f.write(PS_PDF)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps],
                           env=dict(os.environ, ND_DOCX=os.path.abspath(docx), ND_PDF=os.path.abspath(pdf)),
                           capture_output=True, text=True, timeout=timeout, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if r.returncode != 0 or not os.path.isfile(pdf): raise RuntimeError((r.stderr or r.stdout or 'Word did not make the PDF').strip()[-400:])
    finally:
        shutil.rmtree(d, ignore_errors=True)
    return pdf


def _bak_name(p):
    """اسم النسخة الاحتياطية: name.bak-YYYY-MM-DD، ولو موجود: b ثم c…"""
    base = f'{p}.bak-{dt.date.today():%Y-%m-%d}'
    if not os.path.exists(base): return base
    for ch in 'bcdefghijklmnopqrstuvwxyz':
        if not os.path.exists(base + ch): return base + ch
    return base + f'-{int(time.time())}'


# ---------------------------------------------------------------- الشاشة
def header(back, extra):
    t = QHBoxLayout(); t.setSpacing(14)
    b = DQ.btn('←  Document Center', back); b.setObjectName('back'); t.addWidget(b)
    tv = QVBoxLayout(); tv.setSpacing(0)
    c = QLabel('Technical Office Platform  ›  Document Center  ›  Documents'); c.setFont(HQ.font(12.5, True))
    c.setStyleSheet(f'color: {MUTED};'); tv.addWidget(c)
    tv.addWidget(QLabel(f'<span style="color:#6CB4F0; font-size:20pt; font-weight:bold;">Documents</span>'
                        f'<span style="color:{MUTED}; font-size:11pt;">   see every paper, preview it, make a new one</span>'))
    t.addLayout(tv, 1)
    for w in extra: t.addWidget(w)
    return t


class NewDocQt(QWidget):
    def __init__(self, open_section=None, run_old=None, log=None):
        super().__init__()
        self.setObjectName('newdoc'); self.setAttribute(Qt.WA_StyledBackground, True); self.setStyleSheet(CSS)
        self.open_section, self.run_old, self.log = open_section or (lambda k: None), run_old, log
        self.st = load_set(); self.kind = None; self.busy = False
        self.edit_no = None                              # رقم الخطاب اللي بيتعدّل (Edit) – None = خطاب جديد
        self.root = D.project_root()
        self._build()
        QTimer.singleShot(50, lambda: self.set_tab(0))

    def _build(self):
        v = QVBoxLayout(self); v.setContentsMargins(22, 14, 22, 16); v.setSpacing(12)
        v.addLayout(header(lambda: self.open_section('doc'),
                           [DQ.btn('Open OUT folder', self.open_out), DQ.btn('DC sheet', lambda: self._start(SH.sheet_path()))]))
        # صف الأنواع فوق (طلب المستخدم 21-11): كل نوع تابة لوحده – الخطابات وبعدها الـ submittals
        self.subm = None
        try:
            import submittal_qt
            self.subm = submittal_qt.SubmittalQt(open_section=self.open_section, log=self.log, embedded=True)
        except Exception:
            if self.log: self.log(traceback.format_exc())
        # Overview (أرقام الـ Dashboard) + صفحة القايمة (كل ورق النوع + معاينة + ＋ New) – طلب المستخدم 21-11
        self.ov = self.lp = None
        try:
            import dash_qt
            self.ov = dash_qt.DashQt(open_section=self.open_section, log=self.log, embedded=True)
        except Exception:
            if self.log: self.log(traceback.format_exc())
        try:
            import doclist_qt
            self.lp = doclist_qt.ListPage(on_new=self.open_form, on_edit=self.edit_row, on_deliver=self.deliver_row)
        except Exception:
            if self.log: self.log(traceback.format_exc())
        self.tabs, self.tab_keys, self.cur_i, self.cur_tab = [], [], 0, 0
        tr = QHBoxLayout(); tr.setSpacing(5); tr.setContentsMargins(0, 0, 0, 0)

        def grp(t):
            l = QLabel(t); l.setObjectName('grp'); l.setAlignment(Qt.AlignBottom | Qt.AlignLeft); tr.addWidget(l)

        def tab(text, key, tip=''):
            b = QPushButton(text); b.setObjectName('tab'); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor)
            b.setMinimumHeight(44)
            if tip: b.setToolTip(tip)
            i = len(self.tabs); b.clicked.connect(lambda _=False, j=i: self.set_tab(j))
            self.tabs.append(b); self.tab_keys.append(key); tr.addWidget(b)

        if self.ov: tab('▦  Overview\nall numbers', ('ov', None))
        grp('LETTERS')
        for i, (k, lab, _f, _t, _a, fmt) in enumerate(TYPES): tab(f'{lab}\n{fmt}', ('let', i))
        self.sub_codes = []
        if self.subm and self.subm.types:
            grp('SUBMITTALS')
            for t in self.subm.types:
                a, _, b2 = str(t.get('title', t.get('code', ''))).partition('  -  ')
                tab(str(t.get('code') or a.strip()), ('sub', t.get('code')), f"{a.strip()} {b2.strip()}  ·  {t.get('pmm', '')}")   # سطر واحد عشان الصف يكفي الشاشة
                self.sub_codes.append(t.get('code'))
        tr.addStretch(1)
        v.addLayout(tr)
        ln = QFrame(); ln.setObjectName('tabline'); ln.setFixedHeight(2); v.addWidget(ln)
        # شريط «جديد»: بيظهر بس وإنت بتعمل ورقة جديدة – يرجّعك للقايمة
        self.modebar = QWidget(); mb = QHBoxLayout(self.modebar); mb.setContentsMargins(0, 0, 0, 0); mb.setSpacing(10)
        mb.addWidget(DQ.btn('←  Back to list', self.back_to_list))
        self.l_mode = QLabel(''); self.l_mode.setStyleSheet(f'color: {DQ.GREEN}; font: bold 12pt "Segoe UI";'); mb.addWidget(self.l_mode)
        mb.addStretch(1); self.modebar.hide(); v.addWidget(self.modebar)
        self.stack = QStackedWidget()
        page = QWidget(); self.pg_letter = page; row = QHBoxLayout(page); row.setSpacing(14); row.setContentsMargins(0, 0, 0, 0)
        # الخانات
        fc = DQ.card(); g = QGridLayout(fc); g.setContentsMargins(16, 14, 16, 14); g.setVerticalSpacing(4); g.setHorizontalSpacing(12)
        cap = lambda t: (lambda l: (l.setStyleSheet(f'color: {MUTED}; font-size: 9pt;'), l)[1])(QLabel(t))
        self.e_no = QLineEdit(); self.e_no.setStyleSheet(f'font: bold 15pt "Segoe UI"; color: {DQ.GREEN};')
        self.e_date = QDateEdit(QDate.currentDate()); self.e_date.setCalendarPopup(True); self.e_date.setDisplayFormat('dd-MM-yyyy')
        self.l_no = cap('Number (next in DC sheet)')
        rtl = lambda w: (w.setLayoutDirection(Qt.RightToLeft), w)[1]
        self.e_to = rtl(QComboBox()); self.e_to.setEditable(True)
        self.e_attn = rtl(QComboBox()); self.e_attn.setEditable(True)
        self.e_subj = rtl(QLineEdit())
        self.e_body = QPlainTextEdit(); self.e_body.setLayoutDirection(Qt.RightToLeft)
        self.e_body.document().setDefaultTextOption(self._rtl_opt())
        self.e_att = rtl(QLineEdit())
        self.e_role = rtl(QComboBox()); self.e_role.setEditable(True)
        self.e_name = rtl(QComboBox()); self.e_name.setEditable(True)
        self.e_close = rtl(QComboBox()); self.e_close.setEditable(True); self.e_close.addItems(CLOSINGS)
        g.addWidget(self.l_no, 0, 0); g.addWidget(cap('Date'), 0, 1); g.addWidget(self.e_no, 1, 0); g.addWidget(self.e_date, 1, 1)
        g.addWidget(cap('To – السادة /'), 2, 0); g.addWidget(cap('Attention – عناية (empty = none)'), 2, 1)
        g.addWidget(self.e_to, 3, 0); g.addWidget(self.e_attn, 3, 1)
        g.addWidget(cap('Subject – الموضوع'), 4, 0, 1, 2); g.addWidget(self.e_subj, 5, 0, 1, 2)
        g.addWidget(cap('Body – النص (كل سطر فقرة، وأول سطر جاهز)'), 6, 0, 1, 2); g.addWidget(self.e_body, 7, 0, 1, 2)
        g.addWidget(cap('Attachments – مرفق طيه (empty = none)'), 8, 0, 1, 2); g.addWidget(self.e_att, 9, 0, 1, 2)
        g.addWidget(cap('Signed by – الوظيفة'), 10, 0); g.addWidget(cap('Name – الاسم'), 10, 1)
        g.addWidget(self.e_role, 11, 0); g.addWidget(self.e_name, 11, 1)
        g.addWidget(cap('Closing line'), 12, 0, 1, 2); g.addWidget(self.e_close, 13, 0, 1, 2)
        bar = QHBoxLayout(); bar.setSpacing(8)
        bar.addWidget(DQ.btn('Clear', self.clear)); bar.addStretch(1)
        self.c_two = QCheckBox('Allow 2 pages'); self.c_two.setToolTip('Without it the letter is fitted on one page (the text gets bigger / smaller).')
        bar.addWidget(self.c_two)
        self.b_prev = DQ.btn('⟳  Real preview', self.real_preview); bar.addWidget(self.b_prev)
        self.b_make = DQ.btn('Make letter (Word + PDF)', self.make, True); bar.addWidget(self.b_make)
        g.addLayout(bar, 14, 0, 1, 2)
        self.l_info = QLabel(''); self.l_info.setWordWrap(True); self.l_info.setStyleSheet(f'color: {MUTED};')
        g.addWidget(self.l_info, 15, 0, 1, 2)
        g.setRowStretch(7, 1)
        row.addWidget(fc, 3)
        # المعاينة
        pc = DQ.card(); pv = QVBoxLayout(pc); pv.setContentsMargins(12, 10, 12, 12)
        self.l_pre = QLabel('Preview – press «Real preview» (the same Word file, page by page)'); self.l_pre.setStyleSheet(f'color: {MUTED};')
        self.l_pre.setWordWrap(True); pv.addWidget(self.l_pre)
        self.sc = QScrollArea(); self.sc.setMinimumWidth(460); self.sc.setWidgetResizable(True)
        self.pages = QWidget(); self.pl = QVBoxLayout(self.pages); self.pl.setSpacing(10); self.pl.addStretch(1)
        self.sc.setWidget(self.pages); pv.addWidget(self.sc, 1)
        row.addWidget(pc, 2)
        self.stack.addWidget(page)
        if self.subm: self.stack.addWidget(self.subm)
        if self.ov: self.stack.addWidget(self.ov)
        if self.lp: self.stack.addWidget(self.lp)
        v.addWidget(self.stack, 1)
        for w in (self.e_subj, self.e_att, self.e_no): w.textChanged.connect(self.stale)
        for w in (self.e_to, self.e_attn, self.e_role, self.e_name, self.e_close): w.currentTextChanged.connect(self.stale)
        self.e_body.textChanged.connect(self.stale); self.e_date.dateChanged.connect(self.stale)
        self.e_role.addItems(self._merge(ROLES, self.st.get('roles'))); self.e_name.addItems(self._merge(NAMES, self.st.get('names')))
        self.e_role.setCurrentText(self.st.get('role', ROLES[0])); self.e_name.setCurrentText(self.st.get('name', NAMES[0]))

    @staticmethod
    def _rtl_opt():
        from PySide6.QtGui import QTextOption
        o = QTextOption(); o.setTextDirection(Qt.RightToLeft); o.setAlignment(Qt.AlignRight); return o

    @staticmethod
    def _merge(base, extra):
        out = list(base)
        for x in extra or []:
            if x and x not in out: out.append(x)
        return out

    def info(self, t, col=None):
        self.l_info.setText(t); self.l_info.setStyleSheet(f'color: {col or MUTED};')

    def _start(self, p):
        try: os.startfile(p)
        except Exception as e: QMessageBox.warning(self, 'New Document', f'Cannot open:\n{e}')

    def out_dir(self, k=None):
        t = next(t for t in TYPES if t[0] == (k or self.kind))
        return os.path.join(self.root, '02_Correspondence', 'OUT', t[2])

    def open_out(self):
        d = self.out_dir() if self.kind else os.path.join(self.root, '02_Correspondence', 'OUT')
        os.makedirs(d, exist_ok=True); self._start(d)

    # ------------------------------------------------------------ النوع والرقم
    def set_tab(self, i):
        """تابة من الصف اللي فوق: Overview ← الأرقام، أي نوع ← قايمة ورقه (ومنها «＋ New» يفتح خاناته)."""
        if not (0 <= i < len(self.tabs)): return
        for j, b in enumerate(self.tabs): b.setChecked(j == i)
        if self.edit_no:                                 # ساب التعديل بدوسة على تابة ← الخانات والرقم يرجعوا للجديد
            self.clear(); self._edit_mode(None)
        self.cur_tab = i; self.modebar.hide()
        kind, val = self.tab_keys[i]
        if kind == 'ov':
            self.stack.setCurrentWidget(self.ov); self.ov.first_show()
        elif self.lp:
            self.stack.setCurrentWidget(self.lp); self.lp.open(self.list_spec(kind, val))
        else:
            self.open_form()                             # صفحة القايمة ماتحمّلتش ← الخانات على طول زي الأول

    def _unifier_mark(self, row):
        """«اترفع على Unifier» للـ MAR / MIR (نفس العلامة اللي كانت في Materials). يرجّع النص الجديد للعمود."""
        mat = D._mat(); kind, no = row['mark']; marks = mat.load_marks(); k = mat.mkey(kind, no)
        if marks.get(k): marks.pop(k)
        else: marks[k] = dt.date.today().strftime('%d-%m-%Y')
        mat.save_marks(marks)
        self.info(f"{no}: " + (f"uploaded to Unifier ✔ {marks[k]}" if k in marks else 'Unifier mark removed'))
        return f'✔ {marks[k]}' if k in marks else '—'

    def _new_mir(self):
        """«＋ New MIR»: من توريد لسه ملوش MIR (الداتا جاهزة من الإذن) – نفس شباك الـ MIR بتاع المخزن."""
        mp = os.path.join(os.path.dirname(HERE), 'mat')
        if mp not in sys.path: sys.path.insert(0, mp)
        try:
            import mat_qt
            path = mat_qt.new_mir(self)
        except Exception as e:
            LOG.exception('New MIR failed'); return QMessageBox.critical(self, 'New MIR', f'Could not make the MIR:\n{e}')
        if path and self.lp:
            self.lp.invalidate('sub:MIR'); self.set_tab(self.cur_tab)

    def open_form(self):
        """«＋ New» من القايمة: خانات الخطاب أو متطلبات الـ submittal – في نفس المكان."""
        kind, val = self.tab_keys[self.cur_tab]
        if kind == 'sub' and val == 'MIR': return self._new_mir()        # الـ MIR بيبدأ من إذن التوريد (طلب المستخدم 21-11)
        if self.edit_no: self.clear()                    # كان بيعدّل خطاب ← الخانات تفضى للجديد
        self._edit_mode(None)
        if kind == 'let':
            self.stack.setCurrentWidget(self.pg_letter); self.cur_i = val
            self.l_mode.setText(f'New  ·  {TYPES[val][1]}'); self.pick(val)
        elif kind == 'sub' and self.subm:
            self.stack.setCurrentWidget(self.subm); self.l_mode.setText(f'New  ·  {val}')
            self.subm.pick_code(val)
        else:
            return
        self.modebar.setVisible(bool(self.lp))

    def back_to_list(self):
        if self.edit_no: self.clear()
        self._edit_mode(None)
        self.set_tab(self.cur_tab)

    # ------------------------------------------------------------ التعديل والتسليم (طلب المستخدم 21-11)
    # الورقة مفتوحة للتعديل لحد ما تتقفل: الخطاب لما يتعلّم «Delivered»، والـ submittal لما يتقرا Approved في Distribute.
    def _edit_mode(self, no):
        self.edit_no = no or None
        self.e_no.setReadOnly(bool(no))
        self.e_no.setStyleSheet(f'font: bold 15pt "Segoe UI"; color: {"#F0C860" if no else DQ.GREEN};')
        self.b_make.setText('Make letter (Word + PDF) – replace' if no else 'Make letter (Word + PDF)')
        self.l_mode.setStyleSheet(f'color: {"#F0C860" if no else DQ.GREEN}; font: bold 12pt "Segoe UI";')

    def delivered(self):
        import doclist_data as DL
        d = DL.load_status(STATUS_FILE).get('delivered')
        return d if isinstance(d, dict) else {}

    def deliver_row(self, row):
        """«Delivered» على خطاب: يتقفل. على خطاب متسلّم: يتفتح تاني (بعد سؤال)."""
        import doclist_data as DL
        no = row.get('no') or ''
        if not no: return QMessageBox.warning(self, 'Documents', 'This line has no letter number.')
        kind, val = self.tab_keys[self.cur_tab]
        if kind != 'let': return
        k = TYPES[val][0]
        st = DL.load_status(STATUS_FILE); dv = st.get('delivered') if isinstance(st.get('delivered'), dict) else {}
        key = DL.letter_key(no)
        if row.get('delivered'):
            if QMessageBox.question(self, 'Documents', f'{no} is marked delivered on {row["delivered"]}.\n\nOpen it again for editing?') != QMessageBox.Yes: return
            day = None; date = ''
        else:
            if QMessageBox.question(self, 'Documents', f'Mark {no} as delivered today?\n\nIt will be locked – no more editing.\n'
                                    'The date is written in the DC sheet.') != QMessageBox.Yes: return
            day = dt.date.today(); date = f'{day:%d-%m-%Y}'
        QApplication.setOverrideCursor(Qt.WaitCursor)      # التاريخ بيتكتب في DC.xlsx نفسه (الداتا في الإكسيل)
        try: ok, m = SH.set_delivered(k, no, day)
        except Exception as e: ok, m = False, str(e)
        finally:
            while QApplication.overrideCursor(): QApplication.restoreOverrideCursor()
        if not ok:
            return QMessageBox.warning(self, 'Documents', f'The DC sheet was not updated – nothing changed:\n\n{m}\n\n'
                                       'If DC.xlsx is open on another computer, try again in a minute.')
        if key in dv:                                    # علامة قديمة في ملف البرنامج ← خلاص مكانها الشيت
            dv.pop(key, None); st['delivered'] = dv
            try: DL.save_status(STATUS_FILE, st)
            except Exception: LOG.exception('doc_status.json')
        DL.apply_delivery(row, date); self.lp.fill()
        self.lp.invalidate(f'let:{k}')

    def _move_delivered(self, k, rows):
        """علامات «Delivered» القديمة اللي في doc_status.json ← تتكتب في DC.xlsx وتتشال من الملف (مرة واحدة لكل علامة)."""
        import doclist_data as DL
        try:
            st = DL.load_status(STATUS_FILE); dv = st.get('delivered') if isinstance(st.get('delivered'), dict) else {}
            if not dv: return
            moved = False
            for x in rows:
                no = str(x.get('no') or '').strip(); key = DL.letter_key(no)
                if not no or key not in dv: continue
                if not DL.fd(x.get('deliv')):
                    try: day = dt.datetime.strptime(str(dv[key]), '%d-%m-%Y').date()
                    except ValueError: continue
                    ok, m = SH.set_delivered(k, no, day)
                    if not ok: LOG.warning('Delivered date not moved to the DC sheet (%s): %s', no, m); continue
                    x['deliv'] = day
                dv.pop(key, None); moved = True
            if moved: st['delivered'] = dv; DL.save_status(STATUS_FILE, st)
        except Exception:
            LOG.exception('move delivered marks')

    def _letters(self, k, out):
        import doclist_data as DL
        rows = SH.read(k)
        self._move_delivered(k, rows)
        return DL.letters(rows, out, self.delivered())

    def edit_row(self, row):
        kind, val = self.tab_keys[self.cur_tab]
        if row.get('locked'): return
        if kind == 'let': return self.edit_letter(row, val)
        if kind == 'sub' and self.subm:                  # submittal: الخانات تتملى من السطر – المرفقات بتتحط من جديد
            self.stack.setCurrentWidget(self.subm); self.subm.pick_code(val)
            dc = row.get('ref') or ''
            if not dc:                                   # الأنواع اللي قايمتها من الفولدرات: الرقم من اسم الملف
                m = re.search(r'CTR-[A-Z]+\s*-?\s*WA-[\d-]+\d', str(row.get('raw') or ''))
                dc = m.group(0) if m else ''
            self.subm.prefill(no=row.get('no') or '', subject=row.get('desc') or row.get('raw') or '', party=row.get('supplier') or '', dc=dc)
            self._edit_mode(None)
            self.l_mode.setText(f"Edit  ·  {val} {row.get('no') or ''}  –  attach the papers again, then Build package")
            self.modebar.show()

    def edit_letter(self, row, val):
        docx = row.get('docx')
        if not docx or not os.path.isfile(docx):
            return QMessageBox.warning(self, 'Documents', 'No Word file made here was found for this letter, so it cannot be opened for editing.')
        try:
            import letter_read
            L = letter_read.read_letter(docx)
        except Exception as e:
            return QMessageBox.critical(self, 'Documents', f'Could not read the Word file:\n{e}')
        if not L.get('ok') and QMessageBox.question(
                self, 'Documents', 'This Word file does not look like a letter made here – some fields may be empty.\n\nOpen it anyway?') != QMessageBox.Yes:
            return
        self.stack.setCurrentWidget(self.pg_letter); self.cur_i = val
        self._edit_mode(row.get('no') or L.get('no'))
        self.pick(val)                                   # قوايم الجهة والعناية (الرقم مايتغيرش – edit_no)
        self.e_no.setText(self.edit_no); self.l_no.setText('Number (kept – editing this letter)')
        d = QDate.fromString(L.get('date') or '', 'dd-MM-yyyy')
        if d.isValid(): self.e_date.setDate(d)
        self.e_to.setCurrentText(L.get('to', '')); self.e_attn.setCurrentText(L.get('attn', ''))
        self.e_subj.setText(L.get('subject', '')); self.e_body.setPlainText('\n'.join(L.get('body') or []))
        self.e_att.setText(L.get('attach', '')); self.e_close.setCurrentText(L.get('closing', ''))
        if L.get('role'): self.e_role.setCurrentText(L['role'])
        if L.get('name'): self.e_name.setCurrentText(L['name'])
        self.l_mode.setText(f'Edit  ·  {self.edit_no}   –   read from the Word file: check the fields before «Make letter»')
        self.modebar.show()
        self.info('The old Word + PDF are kept as a backup next to the new ones.')

    # فولدرات الأنواع اللي قايمتها من الفولدرات نفسها: code: (الفولدر، فيه مجموعات؟، فيه مراحل؟)
    LIST_SRC = {'SD': (r'04_Submittals\SD', True, True), 'MS': (r'04_Submittals\MS', False, True),
                'ITP': (r'04_Submittals\ITP', False, True), 'RFI': (r'02_Correspondence\OUT\YA\RFI', False, False)}

    def list_spec(self, kind, val):
        import doclist_data as DL
        if kind == 'let':
            k, lab = TYPES[val][0], TYPES[val][1]; out = self.out_dir(k)
            return dict(key=f'let:{k}', new='New letter', cols=DL.LETTER_COLS, stretch=2, folder=out, deliver=True,
                        filters=('all', 'open', 'delivered'),
                        fetch=lambda: self._letters(k, out),
                        note=f'{lab}: from the DC sheet, newest first. The file is found by the letter number in {os.path.relpath(out, self.root)}.')
        code = val
        if code == 'MAR':
            def fetch():
                mat = D._mat(); return DL.mars(mat.load_marmir(), mat, approved=DL.approved_from_moves(D.load_moves(), 'MAR'))
            return dict(key='sub:MAR', new='New MAR', cols=DL.MAR_COLS, stretch=2, fetch=fetch, unifier=self._unifier_mark,
                        filters=('all', 'waiting', 'overdue', 'approved', 'unifier'), folder=os.path.join(self.root, '04_Submittals', 'MAR'),
                        note='From 02-LOG MAR+MIR (MAR LOG). Days = from sent to reply, or to today if there is no reply yet.')
        if code == 'MIR':
            def fetch():
                mat = D._mat(); return DL.mirs(mat.load_marmir(), mat)
            return dict(key='sub:MIR', new='New MIR', cols=DL.MIR_COLS, stretch=2, fetch=fetch, unifier=self._unifier_mark,
                        filters=('all', 'unsigned', 'signed', 'approved', 'unifier'), folder=os.path.join(self.root, '05_Quality', 'MIR'),
                        note='From 02-LOG MAR+MIR (MIR LOG). The stage comes from where the PDF is (1_WIP / 2_SHARED / 3_PUBLISHED).   '
                             '«New MIR» starts from a delivery note that has no MIR yet.')
        if code in self.LIST_SRC:
            rel, grouped, staged = self.LIST_SRC[code]; base = os.path.join(self.root, rel)
            return dict(key=f'sub:{code}', new=f'New {code}', cols=DL.FOLDER_COLS, stretch=0, folder=base,
                        fetch=lambda: DL.folder_rows(base, grouped, staged),
                        filters=('all', 'waiting', 'approved') if staged else ('all',),
                        note=f'From the folders in {rel} – there is no log with sent / reply dates for this type yet, so no day count.')
        return dict(key=f'sub:{code}', new=f'New {code}', cols=DL.FOLDER_COLS, stretch=0, fetch=lambda: [],
                    folder=os.path.join(self.root, '13_PARTIES') if code == 'SUB' else self.root,
                    note='Subcontractor papers are kept in each company folder (13_PARTIES) and followed in the SUB screen. '
                         'Press «＋ New» to prepare an approval package.' if code == 'SUB' else 'No list for this type yet.')

    def pick(self, i):
        if i < 0 or i >= len(TYPES): return
        k, lab, _f, to, attn, fmt = TYPES[i]; self.kind = k
        last = self.st.get('last', {}).get(k, {})
        self.e_to.blockSignals(True); self.e_attn.blockSignals(True)
        self.e_to.clear(); self.e_to.addItems(self._merge([to] if to else [], self.st.get('to', {}).get(k)))
        self.e_to.setCurrentText(last.get('to', to))
        self.e_attn.clear(); self.e_attn.addItems(self._merge(([attn] if attn else []) + [''], self.st.get('attn', {}).get(k)))
        self.e_attn.setCurrentText(last.get('attn', attn))
        self.e_to.blockSignals(False); self.e_attn.blockSignals(False)
        if not self.e_body.toPlainText().strip(): self.e_body.setPlainText(OPENING)
        self.st['kind'] = k; save_set(self.st)
        if self.edit_no: return                          # تعديل خطاب موجود ← رقمه زي ما هو
        self.e_no.setText(''); self.e_no.setPlaceholderText(fmt); self.l_no.setText('Number (next in DC sheet) – reading…')
        self.st['kind'] = k; save_set(self.st)
        res = {}

        def work():
            try: res['n'] = SH.next_number(k)
            except Exception as e: res['e'] = e

        th = threading.Thread(target=work, daemon=True); th.start()

        def wait():
            if th.is_alive(): return QTimer.singleShot(300, wait)
            if self.kind != k or self.edit_no: return
            if res.get('n'):
                self.e_no.setText(res['n']); self.l_no.setText('Number (next in DC sheet)')
            else:
                self.l_no.setText('Number – DC sheet not read, type it' + (f" ({res['e']})" if res.get('e') else ''))
        QTimer.singleShot(300, wait)

    # ------------------------------------------------------------ المعاينة
    def data(self):
        body = [t.strip() for t in self.e_body.toPlainText().split('\n') if t.strip()]
        return dict(no=self.e_no.text().strip(), date=self.e_date.date().toString('dd-MM-yyyy'), subject=self.e_subj.text().strip(),
                    to=self.e_to.currentText().strip(), attn=self.e_attn.currentText().strip(), body=body,
                    attach=self.e_att.text().strip(), closing=self.e_close.currentText().strip(),
                    role=self.e_role.currentText().strip(), name=self.e_name.currentText().strip())

    def stale(self, *_):
        if self.pl.count() > 1: self.l_pre.setText('Preview – changed since the last one. Press «Real preview» again.')

    def _pix_mupdf(self, pdf, w):
        """PyMuPDF: PDF pages -> QPixmap list. Returns None if PyMuPDF is missing."""
        try:
            import fitz  # PyMuPDF
        except Exception:
            try:
                import pymupdf as fitz
            except Exception:
                return None
        from PySide6.QtGui import QImage, QPixmap
        out = []
        doc = fitz.open(pdf)
        try:
            for page in doc:
                rw = page.rect.width or 595.0
                zoom = max(0.2, float(w) / rw)
                pm = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img = QImage(pm.samples, pm.width, pm.height, pm.stride,
                             QImage.Format_RGB888).copy()
                out.append(QPixmap.fromImage(img))
        finally:
            doc.close()
        return out

    def _pix_qtpdf(self, pdf, w):
        """QtPdf fallback: PDF pages -> QPixmap list. Returns None if QtPdf is missing."""
        try:
            from PySide6.QtPdf import QPdfDocument
        except Exception:
            return None
        from PySide6.QtCore import QSize
        from PySide6.QtGui import QPixmap
        out = []
        doc = QPdfDocument(self); doc.load(pdf)
        try:
            for i in range(doc.pageCount()):
                ps = doc.pagePointSize(i)
                h = int(w * ps.height() / ps.width()) if ps.width() else int(w * 1.414)
                out.append(QPixmap.fromImage(doc.render(i, QSize(w, h))))
        finally:
            doc.close()
        return out

    def show_pdf(self, pdf, note=''):
        """صفحات الـ PDF نفسه (اللي طالع من Word) جنب الخطاب"""
        while self.pl.count() > 1:
            it = self.pl.takeAt(0)
            if it.widget(): it.widget().deleteLater()
        w = max(380, self.sc.viewport().width() - 16)
        pix, engine, err = None, '', ''
        for fn, nm in ((self._pix_mupdf, 'PyMuPDF'), (self._pix_qtpdf, 'QtPdf')):
            try:
                pix = fn(pdf, w)
            except Exception as e:
                pix, err = None, f'{nm}: {e}'
            if pix:
                engine = nm; break
        if not pix:
            self.l_pre.setText(
                'Preview engine not installed – run:  pip install pymupdf'
                + (f'   ({err})' if err else '') + f'\nThe PDF is: {pdf}')
            return
        for p in pix:
            lab = QLabel(); lab.setPixmap(p); lab.setAlignment(Qt.AlignHCenter)
            self.pl.insertWidget(self.pl.count() - 1, lab)
        self.l_pre.setText((note + '   ' if note else '')
                           + f'{len(pix)} page(s) – exactly as it prints  ·  {engine}')

    def real_preview(self):
        if self.busy or not self.kind: return
        L = self.data()
        if not L['body']: return QMessageBox.warning(self, 'New Document', 'Write the letter text first.')
        tpl = os.path.join(self.root, TEMPLATE_REL)
        tmp = tempfile.mkdtemp(prefix='newdoc_')
        docx, pdf = os.path.join(tmp, 'preview.docx'), os.path.join(tmp, 'preview.pdf')
        self._run(lambda: make_files(tpl, docx, pdf, L, self.c_two.isChecked()),
                  lambda r: self._previewed(r, pdf), 'Making the preview (Word)…')

    def _previewed(self, r, pdf):
        pages, err = r if isinstance(r, tuple) else (None, str(r))
        if err or not os.path.isfile(pdf): return self.info(f'Preview not made: {err}', '#FF8A8A')
        warn = '⚠ 2 pages even after making the text smaller – shorten it, or tick «Allow 2 pages».' if (pages or 1) > 1 and not self.c_two.isChecked() else ''
        self.show_pdf(pdf, warn); self.info(warn or 'Preview ready ✔', '#FF8A8A' if warn else DQ.GREEN)

    def _run(self, fn, done, msg):
        self.busy = True; self.b_make.setEnabled(False); self.b_prev.setEnabled(False); res = {}

        def work():
            try: res['r'] = fn()
            except Exception as e: traceback.print_exc(); res['r'] = e
        th = threading.Thread(target=work, daemon=True); th.start(); t0 = time.time()

        def wait():
            if th.is_alive():
                self.info(f'{msg} {int(time.time() - t0)} s'); return QTimer.singleShot(400, wait)
            self.busy = False; self.b_make.setEnabled(True); self.b_prev.setEnabled(True)
            self.raise_(); self.activateWindow(); done(res.get('r'))
        QTimer.singleShot(300, wait)

    def clear(self):
        self.e_subj.clear(); self.e_att.clear(); self.e_body.setPlainText(OPENING); self.e_close.setCurrentIndex(0)
        self.e_date.setDate(QDate.currentDate())

    # ------------------------------------------------------------ عمل الخطاب
    def make(self):
        if self.busy or not self.kind: return
        L = self.data()
        if not L['no']: return QMessageBox.warning(self, 'New Document', 'The letter number is empty.')
        if not L['subject']: return QMessageBox.warning(self, 'New Document', 'Write the subject.')
        if not L['to']: return QMessageBox.warning(self, 'New Document', 'Write who it goes to (السادة /).')
        if not L['body']: return QMessageBox.warning(self, 'New Document', 'Write the letter text.')
        tpl = os.path.join(self.root, TEMPLATE_REL)
        if not os.path.isfile(tpl): return QMessageBox.critical(self, 'New Document', f'Template not found:\n{tpl}')
        try:
            import doclist_data as DL
            if os.path.isfile(STATUS_FILE):
                with open(STATUS_FILE, encoding='utf-8') as fh: json.load(fh)      # مش مقري ← نقف (مانبدّلش خطاب ممكن يكون متسلّم)
            when = self.delivered().get(DL.letter_key(L['no']))
            if not when and SH._tail(L['no']) is not None:                         # التسليم متسجّل في DC.xlsx
                hit = SH.find_row(self.kind, SH._tail(L['no']))
                when = DL.fd((hit or {}).get('deliv'))
        except Exception as e:
            return QMessageBox.warning(self, 'New Document', f'Could not check if this letter was delivered (DC sheet):\n{e}')
        if when:                                         # متسلّم = مقفول (حتى لو الرقم اتكتب بالإيد)
            return QMessageBox.warning(self, 'New Document', f"{L['no']} was delivered on {when} – it is locked.\n\n"
                                       'To change it: its tab → select it → «Undo delivered».')
        folder = self.out_dir(); os.makedirs(folder, exist_ok=True)
        base = re.sub(r'[\\/:*?"<>|]+', '-', L['no']).strip()
        docx, pdf = os.path.join(folder, base + '.docx'), os.path.join(folder, base + '.pdf')
        if (os.path.exists(docx) or os.path.exists(pdf)) and QMessageBox.question(
                self, 'New Document', f'{base} is already in {os.path.relpath(folder, self.root)}.\n\nReplace it?') != QMessageBox.Yes:
            return
        for p in (docx, pdf):
            if os.path.exists(p):
                try:
                    with open(p, 'ab'): pass
                except PermissionError:
                    return QMessageBox.warning(self, 'New Document', f'{os.path.basename(p)} is open – close it and try again.')
        for p in (docx, pdf):                            # القديم يتحفظ قبل ما يتبدّل: name.bak-YYYY-MM-DD ثم b، c…
            if os.path.exists(p):
                try: shutil.copy2(p, _bak_name(p))
                except Exception as e:
                    return QMessageBox.warning(self, 'New Document', f'Could not keep a backup of {os.path.basename(p)}:\n{e}')
        # يفتكر الاختيارات
        k = self.kind
        self.st.setdefault('last', {})[k] = dict(to=L['to'], attn=L['attn'])
        for key, val in (('to', L['to']), ('attn', L['attn'])):
            lst = self.st.setdefault(key, {}).setdefault(k, [])
            if val and val not in lst: lst.append(val)
        for key, val, base_l in (('roles', L['role'], ROLES), ('names', L['name'], NAMES)):
            lst = self.st.setdefault(key, [])
            if val and val not in base_l and val not in lst: lst.append(val)
        self.st['role'], self.st['name'] = L['role'], L['name']; save_set(self.st)
        d = self.e_date.date(); day = dt.date(d.year(), d.month(), d.day()); two = self.c_two.isChecked()

        def work():
            pages, perr = make_files(tpl, docx, pdf, L, two)
            if perr and not os.path.isfile(docx): raise RuntimeError(perr)
            out = dict(pages=pages, perr=perr)
            if (pages or 1) > 1 and not two: return out          # صفحتين من غير إذن ← مايتسجلش
            ent = SH.make_entry(k, L['no'], day, L['subject'], pdf if os.path.isfile(pdf) else docx, to=L['to'])
            try: out['sheet'] = SH.write([ent])
            except Exception as e: out['sheet'] = (False, str(e))
            return out
        self._run(work, lambda r: self._made(r, base, docx, pdf), 'Making the letter (Word + PDF)…')

    def _made(self, r, base, docx, pdf):
        if not isinstance(r, dict):
            self.info('The letter was not made', '#FF8A8A'); return QMessageBox.critical(self, 'New Document', str(r))
        if os.path.isfile(pdf): self.show_pdf(pdf)
        if r.get('perr'):
            return QMessageBox.warning(self, 'New Document', f'The Word file is made but not the PDF:\n{r["perr"]}\n\n{docx}')
        if (r.get('pages') or 1) > 1 and not self.c_two.isChecked():
            self.info('⚠ 2 pages – not registered', '#FF8A8A')
            return QMessageBox.warning(self, 'New Document', 'The letter is still 2 pages even after making the text smaller.\n\n'
                                       'Shorten the text, or tick «Allow 2 pages» and press Make again.\n'
                                       'It was NOT written in the DC sheet.')
        ok, _m = r.get('sheet') if isinstance(r.get('sheet'), tuple) else (False, '')
        msg = (f"{base} ready ✔  (1 page)\n\n{os.path.relpath(docx, self.root)}\n{os.path.relpath(pdf, self.root)}"
               + '\n\nDC sheet: ' + ('written ✔' if ok else 'waiting – press «Update DC sheet» in Distribute'))
        self.info(msg.replace('\n\n', '   |   ').replace('\n', '   '), DQ.GREEN)
        if QMessageBox.question(self, 'New Document', msg + '\n\nOpen the PDF to print it?') == QMessageBox.Yes:
            self._start(pdf)
        if self.lp and self.kind: self.lp.invalidate(f'let:{self.kind}')   # القايمة تتقري من جديد (فيها الخطاب ده)
        if self.edit_no and self.lp: return self.back_to_list()          # خلّص التعديل ← يرجع للقايمة
        was = self.edit_no; self._edit_mode(None)
        if was: self.clear()
        self.pick(self.cur_i)                                # الرقم اللي بعده

    def on_close(self): return True
