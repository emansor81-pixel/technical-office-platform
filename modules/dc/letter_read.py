# -*- coding: utf-8 -*-
"""يقرا خطاب اتعمل من New Document (ملف الـ Word) ويرجّع خاناته – عشان زرار «Edit» (21-11-2026).

الخطاب مبني بشكل ثابت (build_docx في newdoc_qt): المشروع / الموضوع / التاريخ / الرقم / السادة / عناية /
السلام / فقرات النص / مرفق طيه / الجملة الختامية / وتفضلوا… / الوظيفة / الاسم. من غير Qt.
"""
import re, zipfile
from xml.sax.saxutils import unescape

REGARDS_KEY = 'وتفضلوا بقبول'
SALAM_KEY = 'السلام عليكم'


def _paras(xml):
    b0 = xml.find('<w:body>'); b1 = xml.rfind('</w:body>')
    body = xml[b0 + 8:b1] if b0 >= 0 and b1 > b0 else xml
    out = []
    for m in re.finditer(r'<w:p\b[^>]*?(?:/>|>(.*?)</w:p>)', body, re.S):
        p = m.group(1) or ''
        runs = re.findall(r'<w:r\b[^>]*>(.*?)</w:r>', p, re.S)
        text = ''.join(unescape(t) for t in re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', p))
        text = re.sub(r'[‪-‮‎‏]', '', text).strip()
        bold = bool(runs) and all('<w:b/>' in r or '<w:b ' in r for r in runs if '<w:t' in r)
        if text: out.append((text, bold))
    return out


def read_letter(docx):
    """dict(no, date, subject, to, attn, body=[…], attach, closing, role, name, ok) – ok=False لو الملف مش بشكل البرنامج."""
    with zipfile.ZipFile(docx) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    L = dict(no='', date='', subject='', to='', attn='', body=[], attach='', closing='', role='', name='', ok=False)
    ps = _paras(xml)
    i_salam = next((i for i, (t, _) in enumerate(ps) if SALAM_KEY in t), -1)
    i_reg = next((i for i, (t, _) in enumerate(ps) if REGARDS_KEY in t), -1)
    head = ps[:i_salam] if i_salam >= 0 else ps
    for t, _ in head:
        if t.startswith('الموضوع:'): L['subject'] = t.split(':', 1)[1].strip()
        elif t.startswith('التاريخ:'): L['date'] = t.split(':', 1)[1].strip()
        elif t.startswith('الرقم:'): L['no'] = t.split(':', 1)[1].strip()
        elif 'السادة /' in t:                              # «المحترمين» قبلها (خطابات قديمة) أو بعدها (من 21-11)
            v = t.split('السادة /', 1)[1].strip()
            L['to'] = v[:-len('المحترمين')].strip() if v.endswith('المحترمين') else v
        elif 'عناية ' in t:
            v = t.split('عناية ', 1)[1].strip()
            L['attn'] = v[:-len('المحترم')].strip() if v.endswith('المحترم') else v
    if i_salam < 0 or i_reg < i_salam: return L                 # مش خطاب من البرنامج ← اللي اتقرا من فوق بس
    mid = ps[i_salam + 1:i_reg]
    for j, (t, bold) in enumerate(mid):
        if t.startswith('مرفق طيه:'): L['attach'] = t.split(':', 1)[1].strip()
        elif bold and j == len(mid) - 1: L['closing'] = t      # آخر سطر تقيل قبل «وتفضلوا» = الجملة الختامية
        else: L['body'].append(t)
    tail = [t for t, _ in ps[i_reg + 1:]]
    if len(tail) >= 2: L['role'], L['name'] = tail[0], tail[1]
    elif tail: L['role'] = tail[0]
    L['ok'] = bool(L['subject'] or L['body'])
    return L
