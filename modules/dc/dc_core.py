# -*- coding: utf-8 -*-
"""DC Control – صندوق الوارد: يقرا، يقترح، ينقل بعد موافقتك، ويسجّل كل نقلة عشان التراجع."""
import os, re, sys, json, time, shutil, hashlib, tempfile, datetime as dt
import dc_read as R
import dc_rules as C
import dc_claude as CL
from dc_io import LOG, save_json, load_json

HERE = os.path.dirname(os.path.abspath(__file__))
_PLAT = os.path.dirname(os.path.dirname(HERE))                       # <PLATFORM> (فيه parties.py)
if _PLAT not in sys.path: sys.path.insert(0, _PLAT)
import parties as PT  # noqa: E402
MOVES = os.path.join(HERE, 'dc_moves.json')
REVIEW = 'مراجعة'
PHOTOS = '_photos'                                   # الصورة الأصلية بتتحفظ هنا بعد ما تتحوّل PDF
CLAUDE_DIR = 'Claude'                                # الورق الصعب ← Claude يقراه في الشات ويكتب جنبه <الملف>.claude.json
NAME_ONLY = ('.zip', '.xlsx', '.xls')                   # مفيش كلام يتقري ← من اسمه (GIS، تحليل أسعار…) أو من قراية Claude
EXT = ('.pdf',) + R.IMG_EXT + NAME_ONLY                 # الـ zip (زي الـ GIS) والإكسيل بيتوزّعوا من اسمهم


def settings():
    d = HERE
    for _ in range(5):
        f = os.path.join(d, 'SETTINGS.json')
        if os.path.isfile(f):
            try:
                with open(f, encoding='utf-8-sig') as fh: return json.load(fh)
            except Exception: return {}
        d = os.path.dirname(d)
    return {}


def project_root():
    return os.environ.get('DC_PROJECT_ROOT') or settings().get('PROJECT_ROOT') or r'C:\\Projects\\SAMPLE_PROJECT'


def _pa_next():
    """رقم تحليل الأسعار الجاي: أكبر PA-NN في 08_Cost_Payments\\COST_CONTROL\\PRICE_ANALYSIS + 1."""
    d = os.path.join(project_root(), C.TYPES['PRICE'][1])
    try: ns = [int(m.group(1)) for n in os.listdir(d) for m in [re.match(r'PA-(\d+)', n)] if m]
    except OSError: ns = []
    return max(ns, default=0) + 1


C.PA_NEXT_FN = _pa_next
C.ROOT_FN = project_root                      # rules بتاخد فولدر المشروع من هنا (من غير import دائري)


def _mat():
    """قسم Materials (MAR / MIR والمخزن) – بيتحمّل وقت الحاجة بس."""
    p = os.path.join(os.path.dirname(HERE), 'mat')
    if p not in sys.path: sys.path.insert(0, p)
    import mat_core
    return mat_core


def _mir_title(no):
    try: return _mat().mir_title(no)
    except Exception: return None


C.MIR_FN = _mir_title                          # اسم الـ MIR المتوقّع من MIR LOG


def _mar_by_out(out):
    """رقم الـ M في MAR LOG لرقم الصادر ده (عمود Submittal Ref No) – None لو مش موجود."""
    try:
        for m in _mat().load_marmir()['mar']:
            if (m.get('ref') or '').replace(' ', '').upper() == out.upper(): return m['n']
    except Exception:
        return None
    return None


# ---------------------------------------------------------------- الاسم الإنجليزي لرقم الصادر
# dc_titles.json جنب البرنامج: «رقم الصادر: الاسم بالإنجليزي». الأوصاف العربي في شيت الـ DC اتترجمت فيه،
# والجديد لما تكتبه في الاسم مرة البرنامج بيحفظه هنا. الوصف اللي في الشيت بالإنجليزي أصلاً بيتاخد زي ما هو.
TITLES = os.path.join(HERE, 'dc_titles.json')                 # اللي اتترجم مع البرنامج
TITLES_USER = os.path.join(HERE, 'dc_titles_user.json')       # اللي إنت كتبته (مابيتمسحش لما البرنامج يتحدّث)
_SHEET_T = {'mt': None, 'map': {}}


def _sheet_titles():
    try:
        import dc_sheet as SH
        p = SH.sheet_path()
        mt = os.path.getmtime(p)
        if _SHEET_T['mt'] != mt:
            _SHEET_T['map'] = {str(r.get('no') or '').strip().upper(): str(r.get('desc') or '').strip() for r in SH.read('YA')}
            _SHEET_T['mt'] = mt
    except Exception:
        pass
    return _SHEET_T['map']


def out_title(out):
    from dc_io import load_json
    t = (load_json(TITLES_USER, {}) or {}).get(out) or (load_json(TITLES, {}) or {}).get(out)
    if t: return t
    d = _sheet_titles().get(out.upper(), '')
    if d and re.fullmatch(r"[A-Za-z0-9 .,&()'+_/-]+", d):        # الوصف إنجليزي في الشيت ← زي ما هو (من غير «/»)
        return ' '.join(re.sub(r'[\\/:*?"<>|]+', '-', d).split())
    return None


def save_title(out, title):
    """الاسم اللي إنت كتبته بعد رقم الصادر ← يتحفظ عشان المرة الجاية."""
    from dc_io import load_json, save_json
    title = ' '.join(str(title or '').split()).strip(' -')
    if not out or not title or out_title(out) == title: return False
    d = load_json(TITLES_USER, {}) or {}
    d[out] = title
    save_json(TITLES_USER, d)
    return True


C.OUT_TITLE_FN, C.MAR_BY_OUT_FN = out_title, _mar_by_out


def inbox():
    return os.path.join(project_root(), '00_INBOX')


def _inbox_sub(new, *old):
    """فولدرات INBOX بقت بأرقام (طلب المستخدم 12-11): 01_REVIEW · 02_CLAUDE · 03_DONE · 04_ORIGINALS · 05_DUPLICATES · 06_HANDOVER.
    لو الفولدر الجديد لسه مش موجود والقديم موجود ← نكمّل على القديم (لحد ما يتغيّر اسمه في Explorer)."""
    try:
        b = inbox()
        if not os.path.isdir(os.path.join(b, new)):
            for o in old:
                if os.path.isdir(os.path.join(b, o)): return o
    except Exception:
        pass
    return new


REVIEW = _inbox_sub('01_REVIEW', 'مراجعة')
CLAUDE_DIR = _inbox_sub('02_CLAUDE', 'Claude')
DONE_DIR = _inbox_sub('03_DONE', os.path.join(CLAUDE_DIR, '_done'))           # نتايج الشات اللي اتستخدمت
DUP_DIR = _inbox_sub('05_DUPLICATES', '_duplicates')                          # المكررات (شاشة Qt)


def _papers(d):
    try: names = sorted(os.listdir(d))
    except OSError: return []
    return [os.path.join(d, n) for n in names
            if n.lower().endswith(EXT) and not n.startswith(('~', '.')) and os.path.isfile(os.path.join(d, n))]


def list_inbox():
    """ملفات INBOX وفولدر المراجعة (من غير README والملفات المخفية).
    فولدر Claude (طلب المستخدم 12-11): الورقة مابتظهرش لحد ما Claude يكتب ملف تعريفها، وساعتها تتنقل لوحدها
    لـ 01_REVIEW (أو تتقسم لو فيها كذا ورقة) – promote_claude."""
    try: claude_dir()                                 # تعليمات فولدر Claude بتتحدث لما يتضاف نوع جديد
    except Exception: pass
    try: stuck = promote_claude()['failed']
    except Exception: LOG.exception('Claude folder promote failed'); stuck = []
    out = []
    for sub in ('', REVIEW):
        for p in _papers(os.path.join(inbox(), sub)):
            n = os.path.basename(p)
            out.append(dict(path=p, name=n, review=sub == REVIEW, part=bool(PART_RE.match(n))))
    for p, err in stuck:                              # تعريفها جاهز بس النقل وقع (الملف مفتوح؟) ← تظهر من مكانها
        out.append(dict(path=p, name=os.path.basename(p), review=False, claude=True, claude_ready=True, warn=err))
    return out


def claude_waiting():
    """عدد الورق اللي في فولدر Claude ولسه ملوش ملف تعريف."""
    return sum(1 for p in _papers(os.path.join(inbox(), CLAUDE_DIR)) if not load_chat_result(p))


def _sha1(p):
    import hashlib
    h = hashlib.sha1()
    with open(p, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''): h.update(b)
    return h.hexdigest()


def _twin(path, folder):
    """نسخة مطابقة بالظبط (نفس المحتوى) من الملف ده في الفولدر ده، أو None."""
    try: size = os.path.getsize(path)
    except OSError: return None
    if not size: return None
    h = None
    for q in _papers(folder):
        try:
            if os.path.normcase(q) == os.path.normcase(path) or os.path.getsize(q) != size: continue
            h = h or _sha1(path)
            if _sha1(q) == h: return q
        except OSError: continue
    return None


def _move_with_json(p, folder):
    """الورقة + ملف تعريفها (لو موجود) ← الفولدر. يرجّع المسار الجديد."""
    os.makedirs(folder, exist_ok=True)
    dst = free_name(folder, os.path.basename(p))
    shutil.move(p, dst)
    j = chat_result_path(p)
    if os.path.isfile(j):
        try: shutil.move(j, chat_result_path(dst))
        except OSError: LOG.exception('json move failed: %s', j)
    return dst


def promote_claude():
    """ورق فولدر Claude اللي ملف تعريفه جهز:
       ورقة واحدة ← 01_REVIEW ومعاها تعريفها · كذا ورقة ← تتقسم: الأجزاء ← 01_REVIEW والأصل ← 04_ORIGINALS
       · نسخة مطابقة لورقة موجودة في 01_REVIEW ← 05_DUPLICATES.
    يرجّع dict(review=[...], split=[(اسم، عدد)], dup=[...], failed=[(path, err)])."""
    src = os.path.join(inbox(), CLAUDE_DIR)
    rv, dd = os.path.join(inbox(), REVIEW), os.path.join(inbox(), DUP_DIR)
    res = dict(review=[], split=[], dup=[], failed=[])
    now = lambda: f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}'
    for _ in range(3):                                # لفة تانية: أجزاء قديمة بقى تعريفها في 04_ORIGINALS بعد التقسيم
        moved = False
        for p in _papers(src):
            if any(p == f for f, _ in res['failed']): continue
            chat = load_chat_result(p)
            if not chat: continue                         # لسه عند Claude
            try:
                if chat.get('bundle') and p.lower().endswith('.pdf'):
                    parts = split_bundle(p, chat['documents'], dest=rv)
                    res['split'].append((os.path.basename(p), len(parts)))
                    for q in parts:
                        tw = _twin(q, rv)
                        if tw:
                            d = _move_with_json(q, dd); res['dup'].append(d)
                            _log(dict(time=now(), act='promote', src=q, dst=d, why='same as ' + os.path.basename(tw)))
                        else: res['review'].append(q)
                else:
                    tw = _twin(p, rv)
                    d = _move_with_json(p, dd if tw else rv)
                    (res['dup'] if tw else res['review']).append(d)
                    _log(dict(time=now(), act='promote', src=p, dst=d, why='same as ' + os.path.basename(tw) if tw else 'Claude ready'))
                moved = True
            except Exception as e:
                LOG.exception('Promote failed: %s', p); res['failed'].append((p, str(e)))
        if not moved: break
    return res


# ---------------------------------------------------------------- فولدر Claude (الشات – من اشتراكك من غير API)
CLAUDE_README = """فولدر Claude
============
الورقة الصعبة، أو السكان اللي فيه كذا ورقة: دوس «Send to Claude» في شاشة الـ DC، والملف ييجي هنا.
بعدها اكتب لـ Claude في الشات: «شوف فولدر Claude».
Claude بيكتب جنب كل ملف نتيجة اسمها  <اسم الملف>.claude.json  –  ولما تدوس Distribute البرنامج بيستخدمها
(يقسّم السكان، ويقترح النوع والرقم والاسم)، وإنت اللي بتوافق زي العادي.

For Claude (in the chat)
------------------------
For every PDF/photo in this folder that has no "<file name>.claude.json" next to it: read it page by page and write
"<file name>.claude.json" (UTF-8) next to it, for example:
{"documents": [{"first_page": 1, "last_page": 13, "type": "SITE_WIR", "number": "30", "date": "2026-10-29", "company": "",
                "signed": true, "title": "Request for Base Course Compaction & Levels", "confidence": 0.95}],
 "pages": 78, "by": "Claude chat", "when": "YYYY-MM-DD HH:MM"}
A delivery note (type DN) also gets "dn", "invoice" and "items", for example:
{"first_page": 1, "last_page": 2, "type": "DN", "number": "", "date": "2026-11-02", "company": "Sample Supplier Trading Company",
 "dn": "RUH/OUT/17582", "invoice": "INV/2026/06513", "items": [{"desc": "شريط تحذير وكشف للماء 15 سم × 305 م – أزرق",
 "code": "GDWTWT15-A", "qty": 3, "unit": "رول"}], "signed": false, "title": "Delivery note + tax invoice", "confidence": 0.9}
If the paper carries the consultant's reply stamp, end "title" with the ticked decision (e.g. "- approved with comments, no
resubmission") and add "stamp_date": "YYYY-MM-DD" (the date on the stamp).
One entry per separate paper, in page order, covering all pages. A paper starts at its own form/cover/letter page; the
checklists, lab reports, sketches and survey levels after it belong to it. Types (code: meaning – what "number" is):
{types}
- OTHER: anything else
Do not move, rename or delete anything here – the DC program does that after the user confirms.
"""


def claude_dir():
    d = os.path.join(inbox(), CLAUDE_DIR)
    os.makedirs(d, exist_ok=True)
    rd = os.path.join(d, 'README.txt')
    txt = CLAUDE_README.replace('{types}', '\n'.join(f'- {c}: {h}' for c, h in CL.HINTS.items()))
    try:
        with open(rd, encoding='utf-8-sig') as fh: old = fh.read()
    except OSError: old = None
    if old != txt:                                       # أنواع جديدة (زي إذن التوريد) ← الملف يتحدث
        try:
            with open(rd, 'w', encoding='utf-8-sig') as fh: fh.write(txt)
        except OSError: pass
    return d


def chat_result_path(path): return path + '.claude.json'


def load_chat_result(path):
    """نتيجة Claude من الشات (الملف .claude.json جنب الورقة) بنفس شكل رد الـ API، أو None.
    الجزء اللي اتقسم من سكان ← بياخد المستند بتاعه من نتيجة الأصل في INBOX\\_originals."""
    c = load_json(chat_result_path(path), None)
    if not isinstance(c, dict) or not isinstance(c.get('documents'), list):
        d = _part_doc(path)
        return {'documents': [d], 'pages': d['last_page'], 'model': 'chat', 'bundle': False, 'chat': True, 'part': True} if d else None
    docs = CL.clean_docs(c['documents'], c.get('pages') if isinstance(c.get('pages'), int) else None)
    return {'documents': docs, 'pages': c.get('pages'), 'model': 'chat', 'bundle': len(docs) >= 2, 'chat': True}


ORIGINALS = _inbox_sub('04_ORIGINALS', '_originals')
PART_RE = re.compile(r'^(?P<stem>.+) - \d{2} \(p(?P<a>\d+)-(?P<b>\d+)\)(?: \(\d+\))?\.pdf$', re.I)   # «الأصل - 01 (p1-13).pdf»


def _part_doc(path):
    """قراية Claude للجزء ده (نوعه ورقمه وتاريخه) من نتيجة السكان الأصلي – بالصفحات اللي في اسم الجزء."""
    m = PART_RE.match(os.path.basename(path))
    if not m: return None
    c = load_json(chat_result_path(os.path.join(inbox(), ORIGINALS, m['stem'] + '.pdf')), None)
    if not isinstance(c, dict) or not isinstance(c.get('documents'), list): return None
    a, b = int(m['a']), int(m['b'])
    for d in CL.clean_docs(c['documents'], c.get('pages') if isinstance(c.get('pages'), int) else None):
        if d['first_page'] == a and d['last_page'] == b:
            return dict(d, first_page=1, last_page=b - a + 1)
    return None


def send_to_claude(path):
    """الورقة ← فولدر Claude (الشات يقراها). يرجّع المسار الجديد."""
    d = claude_dir()
    if os.path.dirname(os.path.abspath(path)) == os.path.abspath(d): return path
    dst = free_name(d, os.path.basename(path))
    shutil.move(path, dst)
    _log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='claude', src=path, dst=dst))
    return dst


def _chat_done(src, dst=None):
    """نتيجة الشات اتستخدمت ← تتحفظ في Claude\\_done (أو جنب الأصل في _originals)."""
    j = chat_result_path(src)
    if not os.path.isfile(j): return
    if dst: target = dst + '.claude.json'
    else:
        dd = os.path.join(inbox(), DONE_DIR); os.makedirs(dd, exist_ok=True)
        target = free_name(dd, os.path.basename(j))
    try: shutil.move(j, target); return target
    except OSError: LOG.exception('chat result move failed: %s', j)


WORK_ROOT = os.path.join(os.environ.get('TEMP') or tempfile.gettempdir(), 'dc_work')
CACHE_DIR = os.path.join(WORK_ROOT, 'cache')       # الكلام اللي اتقري من كل ملف – عشان ما يتقريش تاني
KEEP_DAYS = 7
CACHE_DAYS = 30


def clean_work(days=KEEP_DAYS, cache_days=CACHE_DAYS):
    """يمسح صور القراية الأقدم من أسبوع، والقراية المحفوظة الأقدم من شهر – بيشتغل لما التبويب يفتح."""
    if not os.path.isdir(WORK_ROOT): return 0
    n, now = 0, time.time()
    for d in os.listdir(WORK_ROOT):
        p = os.path.join(WORK_ROOT, d)
        if d == 'cache': continue
        try:
            if os.path.getmtime(p) < now - days * 86400: shutil.rmtree(p, ignore_errors=True); n += 1
        except OSError:
            pass
    if os.path.isdir(CACHE_DIR):
        for f in os.listdir(CACHE_DIR):
            p = os.path.join(CACHE_DIR, f)
            try:
                if os.path.getmtime(p) < now - cache_days * 86400: os.remove(p); n += 1
            except OSError:
                pass
    if n: LOG.info('Cleaned %d old reading file(s)', n)
    return n


def file_key(path):
    """بصمة الملف من محتواه: نفس الورقة = نفس البصمة حتى لو اسمها أو مكانها اتغير."""
    h = hashlib.sha1()
    with open(path, 'rb') as fh:
        for b in iter(lambda: fh.read(1 << 20), b''): h.update(b)
    return h.hexdigest()[:24]


def _cache_paths(key): return os.path.join(CACHE_DIR, key + '.json'), os.path.join(CACHE_DIR, key + '.pdf')


def _cache_get(key):
    j, pdf = _cache_paths(key)
    c = load_json(j, None)
    if not isinstance(c, dict) or 'text' not in c: return None
    try: os.utime(j)                                   # اتستخدم ← يتحسب من النهارده
    except OSError: pass
    return dict(text=c['text'], method=c.get('method', ''), pages=c.get('pages', 0), warn='', claude=c.get('claude'),
                scanned=c.get('scanned', c.get('method') == 'ocr'), ocr_pdf=pdf if os.path.isfile(pdf) else None, cached=True)


def _cache_put(key, r):
    if r.get('warn') or not r.get('text'): return     # قراية ناقصة (OCR مش شغال مثلاً) ما تتحفظش
    os.makedirs(CACHE_DIR, exist_ok=True)
    j, pdf = _cache_paths(key)
    if r.get('ocr_pdf') and os.path.isfile(r['ocr_pdf']) and os.path.abspath(r['ocr_pdf']) != os.path.abspath(pdf):
        shutil.copy2(r['ocr_pdf'], pdf); r['ocr_pdf'] = pdf
    save_json(j, dict(text=r['text'], method=r.get('method'), pages=r.get('pages', 0), scanned=r.get('scanned', False),
                      claude=r.get('claude'), when=f'{dt.datetime.now():%Y-%m-%d %H:%M}'), indent=None)


def drop_work(item):
    """بعد النقل: صور القراية والقراية المحفوظة للملف ده مالهاش لازمة."""
    w = item.get('work')
    if w and os.path.abspath(w).startswith(os.path.abspath(WORK_ROOT)): shutil.rmtree(w, ignore_errors=True)
    if item.get('key'):
        for p in _cache_paths(item['key']):
            try: os.remove(p)
            except OSError: pass


def _best(r, path, work, **kw):
    """قراية تانية؛ ناخدها لو اقتراحها أحسن."""
    try:
        r2 = R.read_file(path, work=work, **kw)
    except R.ReadError:
        return r
    s2 = C.classify(r2['text'], os.path.basename(path), scanned=r.get('scanned', False))   # سكان ولا لأ = من أول قراية
    if s2 and (not r['sugg'] or s2[0]['score'] > r['sugg'][0]['score']):
        r2['sugg'] = s2; r2['scanned'] = r.get('scanned', False); return r2
    return r


def _sure(r): return bool(r.get('sugg')) and r['sugg'][0]['sure']


def analyse(path):
    """قراية + اقتراح. بيرجّع dict فيه text/method/warn/sugg (قايمة اقتراحات) و work (فولدر الصور المؤقتة).
    لو الملف ده اتقري قبل كده (نفس المحتوى) بياخد الكلام المحفوظ على طول من غير قراية تانية."""
    if path.lower().endswith(NAME_ONLY):              # مفيش كلام يتقري ← من اسم الملف (GIS، تحليل أسعار …) + قراية Claude لو موجودة
        r = dict(path=path, text='', method='name', warn='', sugg=C.classify_name(path), ocr_pdf=None, work=None,
                 key=None, scanned=False, pages=0)
        chat = load_chat_result(path)
        if chat:
            r['claude'] = chat
            docs = [d for d in chat['documents'] if d.get('type') in C.TYPES]
            if docs and not (r['sugg'] and r['sugg'][0]['sure']):
                try: _merge_claude(r, docs[0], path)
                except Exception: LOG.exception('Claude merge failed: %s', path)
        return r
    os.makedirs(WORK_ROOT, exist_ok=True)
    key = file_key(path)
    r = _cache_get(key)
    if r:
        r['path'], r['work'] = path, None
        r.setdefault('scanned', r.get('method') == 'ocr')
        r['sugg'] = C.classify(r['text'], os.path.basename(path), scanned=r['scanned'])   # التصنيف بيتعاد (بالقواعد والتعلّم الجديد)
    else:
        work = tempfile.mkdtemp(prefix='w_', dir=WORK_ROOT)            # فولدر جديد لكل قراية
        try:
            r = R.read_file(path, work=work)
        except R.ReadError as e:
            LOG.warning('Cannot read %s: %s', path, e)
            return dict(path=path, text='', method='', warn=str(e), sugg=[], ocr_pdf=None, work=work, key=key)
        r['scanned'] = r.get('method') == 'ocr'                       # مفيهوش كلام مكتوب = سكان / صورة
        r['sugg'] = C.classify(r['text'], os.path.basename(path), scanned=r['scanned'])
        if not _sure(r) and r.get('pages', 0) > 2:                   # يمكن أول صفحتين غلاف – نقرا لحد 5
            r = _best(r, path, work, pages=5)
        if r.get('method') == 'text' and not _sure(r):            # الكلام اللي جوه الـ PDF ممكن يكون متلخبط – نجرب القراية كصورة
            r = _best(r, path, work, force_ocr=True, pages=3)
        if r.get('warn'): LOG.warning('Read with warning %s: %s', path, r['warn'])
        _cache_put(key, r)
        r['work'] = work
    r['key'] = key
    r['sub'] = match_sub(r['text'])                    # لو الورقة لمقاول عندنا
    if r['sub']:
        for x in r['sugg']:
            if x['code'] in C.SUB_CODES:
                x['folder'] = sub_folder(x['code'], r['sub']['folder'], r['sub']['label'])
                x['fields']['sub'] = r['sub']['label']
        for x in r['sugg']:                            # خطاب من جهة عندنا ← فولدرها\05_LETTERS
            if x['code'] == 'IN_SUB': x['folder'] = os.path.join(r['sub']['folder'], PT.LETTERS)
        if r['sugg'] and r['sugg'][0]['code'] in C.SUB_CODES and r['sugg'][0]['score'] >= 30:
            r['sugg'][0]['sure'] = True                # ورق مقاول عندنا ولقينا اسمه
    else:                                              # مش عارفين الجهة ← الفولدر فاضي وإنت تختارها
        for x in r['sugg']:
            if x['code'] in C.SUB_CODES: x['folder'] = ''
    for x in r['sugg']:                                # نسخة المستخلص الموقعة ← رقمه واسمه وفولدره من ملفات المستخلصات
        if x['code'] == 'IPC_SIGNED':
            try: _signed_target(x, r['text'])
            except Exception: LOG.exception('Signed IPC lookup failed: %s', path)
    try: _claude_help(r, path)                         # مش متأكد / سكان طويل ← نسأل Claude (لو مفتاحه متسجل)
    except Exception: LOG.exception('Claude help failed: %s', path)
    return r


# ---------------------------------------------------------------- مساعدة Claude
def _claude_help(r, path):
    """القواعد مش متأكدة ← Claude يقرا أول صفحات الورقة ويقول نوعها.
    سكان طويل من الأنواع اللي بتيجي مجمّعة (طلبات، اختبارات…) ← Claude يقرا الملف كله ويقول كل مستند من صفحة كام لكام.
    ولو الورقة اتبعتت لـ Claude في الشات (فولدر Claude) ونتيجتها موجودة ← دي اللي بتمشي (من غير API)."""
    chat = load_chat_result(path)
    if chat:
        r['claude'] = chat
        if len(chat['documents']) >= 2:
            r['parts'] = chat['documents']
            for x in r['sugg']: x['sure'] = False
            return
        docs = [d for d in chat['documents'] if d.get('type') in C.TYPES]
        if docs: _merge_claude(r, docs[0], path)
        return
    if not CL.enabled(): return
    top = r['sugg'][0] if r.get('sugg') else None
    long_scan = bool(r.get('scanned')) and r.get('pages', 0) > CL.BUNDLE_PAGES and (not top or not top['sure'] or top['code'] in CL.BUNDLE_TYPES)
    if not long_scan and _sure(r): return
    res = r.get('claude')
    if not res or (long_scan and not res.get('bundle')):
        try: res = CL.ask(path, bundle=long_scan, alt_pdf=r.get('ocr_pdf'))
        except CL.ClaudeError as e:
            r['claude_err'] = str(e); LOG.warning('Claude: %s – %s', os.path.basename(path), e); return
        r['claude'] = res
        if r.get('key'):
            try: _cache_put(r['key'], r)               # ما ندفعش تاني على نفس الملف
            except Exception: LOG.exception('cache')
    docs = [d for d in res.get('documents', []) if d.get('type') in C.TYPES]
    if len(res.get('documents', [])) >= 2:             # سكان فيه كذا مستند ← يتقسم الأول
        r['parts'] = res['documents']
        for x in r['sugg']: x['sure'] = False
        return
    if docs: _merge_claude(r, docs[0], path)


def _merge_claude(r, d, path):
    """اقتراح Claude يبقى الأول. الرقم والتاريخ من القواعد لو طلعوا، وإلا من Claude."""
    code = d['type']
    if code == 'IPC' and d.get('signed'): code = 'IPC_SIGNED'
    base = next((x for x in r['sugg'] if x['code'] == code), None)
    f = dict(base['fields']) if base else C.fields(code, r.get('text', ''), r.get('text', '')[:600], os.path.basename(path))
    if code == 'SD':                                       # نوع اللوحة والخط من عنوان Claude لو القواعد ماعرفتهمش
        num = str(d.get('number') or d.get('number_raw') or '').strip()
        for k, v in C.sd_info(f"{d.get('title') or ''} {num}").items(): f.setdefault(k, v)
        if not f.get('dwg') and re.fullmatch(r'[A-Z]{1,3}\d{0,2}(-\d{1,3})+|[A-Z]{1,3}\d{1,3}', num.upper().replace(' ', '')):
            f['dwg'] = re.sub(r'-D(\d)', r'-D-\1', 'CLIENT-WTR-SD-' + num.upper().replace(' ', ''))   # «D-1» ← CLIENT-WTR-SD-D-1
            if f['dwg'].startswith('CLIENT-WTR-SD-D-'): f.setdefault('kind', 'Detail')

    if 'WIR' in code or code.startswith('IPC'): f.pop('status', None)
    if code == 'DN':                                       # إذن التوريد: رقمه ورقم الفاتورة زي ما هم مكتوبين
        if d.get('dn'): f['dn'] = str(d['dn']).strip()
        if d.get('invoice'): f['inv'] = str(d['invoice']).strip()
    elif code == 'SUB_CONTRACT':                           # رقم العقد زي ما هو (2610/0424) والاعتمادات اللي Claude شافها
        raw = str(d.get('number_raw') or d.get('number') or '').strip()
        if raw and (not f.get('ct') or d.get('confidence', 0) >= 0.8): f['ct'] = raw
        if isinstance(d.get('approvals'), list): f['appr'] = [a for a in d['approvals'] if a in ('legal', 'sub', 'chair')]
        if isinstance(d.get('items'), list) and d['items']: f['items'] = d['items']          # بنود العقد ← تتسجل في SUB
        if isinstance(d.get('sub_info'), dict): f['sub_info'] = d['sub_info']
    elif code == 'PAYROLL':
        if d.get('payroll') in ('Workers', 'Operators'): f['kind'] = d['payroll']
        f['month'] = f.get('month') or C.month_from_name(os.path.basename(path)) or (str(d['date'])[:7] if d.get('date') else None)
        if not f['month']: f.pop('month')
        C.set_ipc(code, f, os.path.basename(path))
    elif code != 'SD' and str(d.get('number') or '').isdigit() and (not f.get('no') or d.get('confidence', 0) >= 0.8): f['no'] = int(d['number'])
    if 'WIR' not in code and not code.startswith('IPC') and code not in ('DN', 'MIR', 'UPD_CAL'):
        st = C.status_from_title(d.get('title'))           # Claude شاف الختم (المربع المتعلّم) ← أدق من القواعد
        if st: f['status'] = st
        tl = (d.get('title') or '').lower()
        if 'resubmi' in tl and 'no resubmi' not in tl: f['resubmit'] = True
        sd = d.get('stamp_date') or d.get('reply_date')    # تاريخ ختم الاستشاري (لـ MAR LOG)
        try:
            if sd: f['reply'] = dt.date.fromisoformat(str(sd)[:10]).isoformat()
        except ValueError: pass
    o = C.find_out(str(d.get('number_raw') or '')) or C.find_out(d.get('title') or '')   # رقم الصادر للاستشاري زي ما Claude شافه
    if o and code in C.OUT_NAMED and not f.get('out'):
        f['out'] = o
        if code == 'MAR':
            n = _mar_by_out(o)
            if n: f['no'] = n
    if d.get('date') and not f.get('date'):
        try: f['date'] = dt.date.fromisoformat(d['date']).isoformat()
        except ValueError: pass
    s = dict(code=code, label=C.TYPES[code][0], score=max(base['score'] if base else 0, 90), fields=f,
             folder=C.type_folder(code, f), name=C.file_name(code, f, os.path.basename(path), r.get('text', '')),
             why=['Claude' + (f" – {d.get('title')}" if d.get('title') else '')], sure=d.get('confidence', 0) >= 0.8, claude=True)
    if code == 'SITE_WIR' and not s['name']:              # الاسم من لوج الطلبات بس – مفيش ← مش متأكد ويقولك ليه
        s['note'] = C.wir_note(f); s['sure'] = False
    if C.out_note(code, f):                                # رقم صادر ملوش اسم إنجليزي ← Check
        s['note'] = C.out_note(code, f); s['sure'] = False
    if code in C.SUB_CODES:
        comp = (d.get('company') or '').strip()
        sb = match_sub(comp) if comp else r.get('sub')     # Claude قال الشركة ← مانخمّنش من الكلام (العقد فيه اسم المقاول دايماً)
        if not sb and comp and code in ('SUB_CONTRACT', 'SUB_IPC', 'QUOTE'):   # جهة جديدة ← فولدرها الإنجليزي في 13_PARTIES
            en = sub_english_folder(comp) or en_company(comp)
            r['sub'] = dict(label=comp, folder='', dir='', no=0, how='Claude')
            f['sub'], f['sub_en'] = comp, en
            s['folder'] = sub_folder(code, '', comp) or ''
            s['name'] = C.file_name(code, f, os.path.basename(path), r.get('text', ''))
            s['sure'] = s['sure'] and bool(en)
            sb = None; comp = None
        if comp is None: pass
        elif sb:
            s['folder'] = sub_folder(code, sb['folder'], sb['label']); f['sub'] = sb['label']; r['sub'] = sb
            if sb.get('name'):                             # الاسم الإنجليزي للجهة في اسم العقد/المستخلص
                f['sub_en'] = sb['name']; s['name'] = C.file_name(code, f, os.path.basename(path), r.get('text', ''))
        else:
            s['folder'] = ''; s['sure'] = False         # شركة جديدة ← اسمها الإنجليزي من الورقة جاهز، وإنت تراجعه
            if en_company(d.get('company')): f['company'] = en_company(d.get('company'))
    if code == 'IPC_SIGNED':
        if f.get('no'): s['fields']['no'] = f['no']
        try: _signed_target(s, r.get('text', ''))
        except Exception: LOG.exception('Signed IPC lookup failed: %s', path)
    r['sugg'] = [s] + [x for x in r['sugg'] if x['code'] != code]


def split_bundle(path, parts, dest=None):
    """السكان اللي فيه كذا مستند ← ملف لكل مستند في INBOX، والأصل يتحفظ في INBOX\\_originals. يرجّع الملفات الجديدة."""
    import pypdf
    rd = pypdf.PdfReader(path)
    n = len(rd.pages)
    keep = os.path.join(inbox(), ORIGINALS); os.makedirs(keep, exist_ok=True)
    orig = free_name(keep, os.path.basename(path))
    stem = os.path.splitext(os.path.basename(orig))[0]   # الأجزاء على اسم الأصل المحفوظ ← كل جزء يلاقي قراية Claude بتاعته
    out, dest = [], dest or inbox()
    os.makedirs(dest, exist_ok=True)
    for i, d in enumerate(parts, 1):
        a, b = max(1, int(d['first_page'])), min(n, int(d['last_page']))
        if a > b: continue
        w = pypdf.PdfWriter()
        for k in range(a - 1, b): w.add_page(rd.pages[k])
        dst = free_name(dest, f'{stem} - {i:02d} (p{a}-{b}).pdf')
        with open(dst, 'wb') as fh: w.write(fh)
        out.append(dst)
    shutil.move(path, orig)
    if os.path.isfile(chat_result_path(path)):
        _chat_done(path, orig)                            # نتيجة الشات تتحفظ جنب الأصل
    else:                                                 # نتيجة الـ API ← تتحفظ جنب الأصل برضه (منها الأجزاء بتاخد أرقامها)
        save_json(chat_result_path(orig), {'documents': parts, 'pages': n, 'by': 'Claude API', 'when': f'{dt.datetime.now():%Y-%m-%d %H:%M}'})
    _log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='split', src=path, dst=orig, parts=out))
    return out


# ---------------------------------------------------------------- المقاولين والموردين
SUB_ROOT = PT.ROOT_REL                                # 13_PARTIES (كل الجهات – طلب المستخدم 12-11)
OUR_IDS = {'0000000001', '000000000210003'}
SUBS_FILE = os.path.join(HERE, 'dc_subs.json')
STOP = set('شركة شركه مؤسسة مؤسسه مورد مقاول مقاولات للمقاولات المقاولات ثقب افقي أفقي خرسانه خرسانة جاهزه جاهزة مواد '
           'للتجارة التجارية المحدوده the and for company co ltd trading contracting المقاول المقاول contractorco'.split())   # اسمنا موجود في كل عقد


def _norm(t):
    t = (t or '').lower()
    for a, b in (('أ', 'ا'), ('إ', 'ا'), ('آ', 'ا'), ('ة', 'ه'), ('ى', 'ي')): t = t.replace(a, b)
    return t


def subs():
    """الجهات اللي في 13_PARTIES (مقاولين/موردين/خلاطات/مختبرات). label = «07-ASAS_AL_BENA · اساس البناء»،
    folder = فولدر الجهة نفسها (نسبةً للمشروع)."""
    out = []
    for p in PT.all_parties(project_root()):
        out.append(dict(no=p['no'], label=p['dir'] + (f" · {p['aliases'][0]}" if p['aliases'] else ''), folder=p['rel'],
                        dir=p['dir'], name=p['name'], cat=p['cat'], aliases=p['aliases']))
    return out


def _ids(text):
    """سجل تجاري (10 أرقام) ورقم ضريبي (15 رقم) – غير أرقامنا."""
    t = re.sub(r'(?<=\d)[\s-](?=\d)', '', text or '')
    ids = set(re.findall(r'(?<!\d)(?:10|40|20|70|11|50|59|46|41|45|34|51|58)\d{8}(?!\d)', t)) | set(re.findall(r'(?<!\d)3\d{13}3(?!\d)', t))
    return ids - OUR_IDS


def _load_subs():
    M = load_json(SUBS_FILE, {})
    return M if isinstance(M, dict) else {}


def remember_sub(folder, text):
    M = _load_subs(); e = M.setdefault(re.split(r'[\\/]', folder.rstrip('\\/'))[-1], {'ids': []})
    e['ids'] = sorted(set(e['ids']) | _ids(text))
    save_json(SUBS_FILE, M)


def match_sub(text):
    """أنهي مقاول دي ورقته؟ بالسجل التجاري/الرقم الضريبي اللي اتحفظ قبل كده، وبعدين باسمه."""
    L = subs()
    if not L: return None
    ids, M, nt = _ids(text), _load_subs(), _norm(text)
    def known_ids(s):                                  # المحفوظ باسم الفولدر الجديد، أو القديم (07-اساس البناء)
        out = set(M.get(s['dir'], {}).get('ids', []))
        al = {PT.key(a) for a in s.get('aliases') or []}
        for k, v in M.items():
            if PT.key(re.sub(r'^\s*\d+\s*-\s*', '', k)) in al: out |= set(v.get('ids', []))
        return out
    for s in L:
        if ids & known_ids(s): return dict(s, how='CR / VAT no.')
    best, bs = None, 0
    for s in L:
        for nm in [s.get('name', '').replace('_', ' ')] + list(s.get('aliases') or []):
            words = [w for w in re.findall(r'[\w]{3,}', _norm(nm)) if w not in STOP and not w.isdigit()]
            hit = [w for w in words if re.search(r'(?<!\w)(ال)?' + re.escape(w.removeprefix('ال')) + r'(?!\w)', nt)]
            sc = len(hit) / max(len(words), 1) if hit else 0
            if sc > bs: best, bs = dict(s, how='name: ' + ' '.join(hit)), sc
    return best if bs >= 0.5 else None


def en_company(name):
    """اسم الشركة بالإنجليزي من اللي Claude قراه («Sample Steel Co Products Co. (شركة مصنع الحديد (مثال))» ← Sample Steel Co Products)"""
    t = re.sub(r'\([^)]*[\u0600-\u06FF][^)]*\)', ' ', str(name or ''))            # اللي بين قوسين بالعربي
    t = re.sub(r'[\u0600-\u06FF]+', ' ', t)
    t = re.sub(r'\b(Co|Company|Est|Ltd|LLC|Corp)\b\.?', ' ', t, flags=re.I)
    t = ' '.join(re.sub(r"[^A-Za-z0-9 &'.-]", ' ', t).split()).strip(' .-')
    return t if len(t) >= 3 else ''


def new_sub_folder(name, cat='SUB'):
    """جهة جديدة ← 13_PARTIES\\SUB\\16-ENGLISH_NAME (الرقم اللي عليه الدور لكل الجهات)."""
    return PT.new_rel(project_root(), en_company(name) or name, cat)


def party_rel(label):
    """فولدر الجهة بالاسم (إنجليزي / عربي / «07-NAME · عربي»)، أو None."""
    p = PT.find(project_root(), str(label or '').split('·')[0].strip()) or PT.find(project_root(), label)
    return p['rel'] if p else None


SUB_PROGRAM_REG = (os.path.join('00_Registers', '10.1-SUBCONTRACTORS.xlsx'),                 # سجل قسم SUB (وقبل كده كان هنا ↓)
                   os.path.join('08_Cost_Payments', 'SUBCONTRACTORS', 'SUBCONTRACTORS.xlsx'))


def sub_english_folder(label):
    """اسم فولدر المقاول بالإنجليزي: من سجل قسم SUB، وإلا فولدر موجود في 08_Cost_Payments\\SUBCONTRACTORS
    بنفس الاسم الإنجليزي («شركة مقاول باطن (مثال) (Sample Subcontractor)» ← SAMPLE_SUBCONTRACTOR) – وإلا None"""
    p = PT.find(project_root(), str(label or '').split('·')[0].strip()) or PT.find(project_root(), label)
    if p: return p['name']                                 # الجهة في 13_PARTIES ← اسمها الإنجليزي
    try: r = _sub_reg_folder(label)
    except Exception: r = None
    if r: return r
    en = en_company(label)
    k = re.sub(r'[^a-z0-9]', '', en.lower())
    root = os.path.join(project_root(), '08_Cost_Payments', 'SUBCONTRACTORS')
    if k and os.path.isdir(root):
        for n in sorted(os.listdir(root)):
            if os.path.isdir(os.path.join(root, n)) and re.sub(r'[^a-z0-9]', '', n.lower()) == k: return n
    return None


def _sub_reg_folder(label):
    key = _norm(label).strip()
    if not key: return None
    for rel in SUB_PROGRAM_REG:
        p = os.path.join(project_root(), rel)
        if not os.path.isfile(p): continue
        try:
            import openpyxl
            wb = openpyxl.load_workbook(p, read_only=True)
            try:
                rows = list(wb['SUBS'].iter_rows(values_only=True)) if 'SUBS' in wb.sheetnames else []
            finally:
                wb.close()
        except Exception:
            return None
        if not rows: return None
        head = [str(h or '').strip() for h in rows[0]]
        if 'Name' not in head or 'Folder' not in head: return None
        i_n, i_f = head.index('Name'), head.index('Folder')
        for r in rows[1:]:
            name = str(r[i_n] or '') if i_n < len(r) else ''
            folder = str(r[i_f] or '').strip() if i_f < len(r) else ''
            if folder and key in (_norm(name).strip(), _norm(folder).strip()): return folder
        return None
    return None


SUB_PART = {'SUB': os.path.join(PT.APPROVAL, '1_WIP'), 'QUOTE': PT.QUOTES, 'SUB_CONTRACT': PT.CONTRACT,
            'SUB_IPC': PT.CERTS, 'IN_SUB': PT.LETTERS}


def sub_folder(code, sub_folder_rel, label):
    """فولدر الورقة جوه فولدر الجهة (13_PARTIES\\<CAT>\\NN-NAME): اعتماد ← 02_APPROVAL\\1_WIP، عرض سعر ← 01_QUOTATIONS،
    عقد ← 03_CONTRACT، مستخلص ← 04_CERTIFICATES، خطاب ← 05_LETTERS. الجهة مش موجودة ← فولدر جديد برقمها."""
    if code not in SUB_PART: return None
    base = sub_folder_rel if str(sub_folder_rel or '').startswith(PT.ROOT_REL) else None
    base = base or party_rel(label) or new_sub_folder(label)
    return os.path.join(base, SUB_PART[code])


def archive_old_copies(f, dst):
    """النسخة المعتمدة وصلت 3_PUBLISHED ← النسخ القديمة بنفس رقم الصادر في 1_WIP و 2_SHARED تروح 4_ARCHIVED.
    يرجّع أسامي اللي اتنقلوا."""
    out = (f or {}).get('out')
    parts = os.path.normpath(dst).split(os.sep)
    if not out or '3_PUBLISHED' not in parts: return []
    base = os.sep.join(parts[:parts.index('3_PUBLISHED')])
    arc, moved, back = os.path.join(base, '4_ARCHIVED'), [], []
    for st in ('1_WIP', '2_SHARED'):
        d = os.path.join(base, st)
        if not os.path.isdir(d): continue
        for n in os.listdir(d):
            p = os.path.join(d, n)
            if os.path.isfile(p) and out.upper() in n.upper():
                os.makedirs(arc, exist_ok=True)
                try:
                    q = free_name(arc, n); shutil.move(p, q); moved.append(n); back.append((os.path.basename(q), st))
                except OSError:
                    LOG.exception('Archive old copy failed: %s', p)
    if moved: _log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='archive', src=base, dst=arc, files=moved, back=back, out=out))
    return moved


def free_name(folder, name):
    """ما نكتبش فوق ملف موجود: لو الاسم موجود نزوّد (2)، (3)…"""
    base, ext = os.path.splitext(name)
    p, i = os.path.join(folder, name), 2
    while os.path.exists(p):
        p = os.path.join(folder, f'{base} ({i}){ext}'); i += 1
    return p


def safe_name(n):
    return re.sub(r'[\\/:*?"<>|]+', '-', n).strip().rstrip('.') or 'document.pdf'


def _log(entry):
    L = load_moves(); L.append(entry)
    save_json(MOVES, L[-2000:])


def load_moves():
    L = load_json(MOVES, [])
    return L if isinstance(L, list) else []


def to_review(path):
    """الملف اللي مش متأكد منه يروح فولدر «مراجعة» جوه INBOX."""
    d = os.path.join(inbox(), REVIEW); os.makedirs(d, exist_ok=True)
    if os.path.dirname(os.path.abspath(path)) == os.path.abspath(d): return path
    dst = free_name(d, os.path.basename(path))
    shutil.move(path, dst)
    _log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='review', src=path, dst=dst))
    return dst


def approve(item, code, folder_rel, name, guessed=None):
    """نقل الملف لمكانه بالاسم الجديد بعد موافقتك. الصورة: الـ PDF المقروء هو اللي بيتنقل والصورة تتحفظ في _photos."""
    src = item['path']
    if not os.path.isfile(src): raise FileNotFoundError(src)
    folder = os.path.join(project_root(), folder_rel)
    os.makedirs(folder, exist_ok=True)
    m = re.match(r'(13_PARTIES[\\/][^\\/]+[\\/][^\\/]+)', os.path.normpath(folder_rel))
    if m:                                                  # جهة (جديدة) ← فولدراتها الأساسية + اسمها العربي في README
        lab = str((item.get('fields') or {}).get('sub') or '').split('·')[-1].strip()
        try: PT.ensure(os.path.join(project_root(), m.group(1)), arabic=lab if re.search(r'[\u0600-\u06FF]', lab) else '')
        except OSError: LOG.exception('party folders')
    name = safe_name(name)
    if code == 'PRICE' and re.match(r'PA-\d+\s*-', name):   # الرقم يتحسب وقت النقل ← تحليلين ورا بعض ما ياخدوش نفس PA-01
        name = re.sub(r'^PA-\d+', f'PA-{_pa_next():02d}', name)
    if not os.path.splitext(name)[1]: name += '.pdf'
    dst = free_name(folder, name)
    extra = None
    ocr_pdf = item.get('ocr_pdf')
    if src.lower().endswith(R.IMG_EXT) and ocr_pdf and os.path.isfile(ocr_pdf):
        shutil.copy2(ocr_pdf, dst)
        ph = os.path.join(inbox(), PHOTOS); os.makedirs(ph, exist_ok=True)
        extra = free_name(ph, os.path.basename(src)); shutil.move(src, extra)
    else:
        if src.lower().endswith(R.IMG_EXT) and os.path.splitext(dst)[1].lower() == '.pdf':
            dst = os.path.splitext(dst)[0] + os.path.splitext(src)[1]          # مفيش OCR – ننقل الصورة زي ما هي
        try:
            shutil.move(src, dst)
        except PermissionError:
            # الملف مفتوح في برنامج تاني (PDF): ويندوز بينسخه ومايقدرش يمسح الأصل ← النسخة تتشال عشان مايبقاش فيه ملفين
            if os.path.isfile(src) and os.path.isfile(dst):
                try: os.remove(dst)
                except OSError: pass
            raise PermissionError(f'The file is open in another program (PDF viewer?) – close it and press Approve again.\n\n{src}')
    j = _chat_done(src)                                   # تعريف Claude ← 03_DONE (والورقة نفسها راحت فولدرها)
    _log(dict(time=f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}', act='move', src=src, dst=dst, photo=extra,
              type=code, guessed=guessed, fields=item.get('fields') or {}, json=j))
    try: C.learn(item.get('text', ''), guessed, code)
    except Exception: pass
    if code == 'SUB':
        try: remember_sub(folder_rel, item.get('text', ''))
        except Exception: pass
    return dst


# ---------------------------------------------------------------- MAR LOG (طلب المستخدم 12-11)
MAR_STATUS = {'approved': 'Approved', 'approved_comments': 'Approved with Comments - No Resubmission', 'rejected': 'Rejected'}


def mar_status_text(status, resubmit=False):
    if status == 'approved_comments' and resubmit: return 'Approved with Comments - Resubmit'
    return MAR_STATUS.get(status)


def _daily_backup(path):
    """نسخة احتياطية من السجل مرة في اليوم قبل أول كتابة (name.bak-YYYY-MM-DD)."""
    b = f'{path}.bak-{dt.date.today().isoformat()}'
    if not os.path.exists(b): shutil.copy2(path, b)
    return b


def mar_row(out):
    """سطر الـ MAR في MAR LOG برقم الصادر (Submittal Ref No)، أو None."""
    M = _mat()
    k = re.sub(r'\W', '', out or '').upper()
    return next((m for m in M.load_marmir()['mar'] if k and re.sub(r'\W', '', m['ref'] or '').upper() == k), None)


def mar_log_update(out, status, reply=None, resubmit=False):
    """الـ MAR رجع من الاستشاري ← سطره في MAR LOG: الحالة (L) لو فاضية، تاريخ الرد (J) لو فاضي،
    ومعادلة الأيام (K) لو فاضية. يرجّع رسالة قصيرة. الكتابة من جوه Excel (زي قسم Materials)."""
    M = _mat()
    text = mar_status_text(status, resubmit)
    if not text: return ''
    m = mar_row(out)
    if not m: return f'MAR LOG: {out} not found – add its row by hand'
    r, cells, notes = m['row'], [], []
    if not m.get('status'): cells.append(dict(addr=f'L{r}', kind='text', value=text))
    elif m['status'].strip().lower() != text.lower(): notes.append(f"status kept «{m['status']}»")
    if reply and not m.get('reply'): cells.append(dict(addr=f'J{r}', kind='date', value=reply))
    if not str(m.get('days') or '').strip(): cells.append(dict(addr=f'K{r}', kind='text', value=f'=J{r}-H{r}'))
    if not cells: return f"MAR LOG {m['no']}: already filled" + (f" ({', '.join(notes)})" if notes else '')
    _daily_backup(M.P(M.MARMIR_REL))
    M.xl_run(M.MARMIR_REL, [dict(op='set', sheet=M.S_MARLOG, cells=cells, tag='mar')])
    return f"MAR LOG {m['no']} ✔ {text}" + (f' · returned {reply}' if reply and any(c['addr'][0] == 'J' for c in cells) else '') \
        + (f" ({', '.join(notes)})" if notes else '')


def _ipc():
    """كود المستخلصات (نفس كود شاشة المستخلص) – عشان الـ DC يعلّم الاعتماد من نفس المكان."""
    base = os.path.dirname(HERE)
    for sub in ('qs', 'ipc'):
        p = os.path.join(base, sub)
        if p not in sys.path: sys.path.append(p)
    import ipc_core
    return ipc_core


def _next_ipc():
    I = _ipc()
    return I.next_ipc(out_dir=I.qs_core.project_path(I.OUT_REL))


def _ipc_dir(no):
    I = _ipc()
    return os.path.relpath(I.ipc_folder(I.qs_core.project_path(I.OUT_REL), no), project_root())


C.NEXT_IPC_FN, C.IPC_DIR_FN = _next_ipc, _ipc_dir      # المسير ← فولدر المستخلص اللي لسه ماترفعش
C.IPC_FOR_MONTH_FN = lambda mo: _ipc().ipc_for_month(mo)  # المسير ← المستخلص اللي فترته فيها شهره
C.LATEST_IPC_FN = lambda: _ipc().latest_ipc()             # الـ GIS ← آخر مستخلص


def _sub():
    """كود قسم SUB (سجل 10-LOG SUB) – عشان العقد يتسجل من نفس المكان"""
    base = os.path.dirname(HERE)
    for x in ('qs', 'sub'):
        p = os.path.join(base, x)
        if p not in sys.path: sys.path.append(p)
    import sub_core
    return sub_core


def contract_log_ready():
    """قبل ما ننقل العقد: 10-LOG SUB موجود ومش مفتوح ومفيهوش حاجة هتضيع"""
    S = _sub()
    p = S.log_writable()
    import openpyxl
    wb = openpyxl.load_workbook(p, read_only=True)
    try:
        if S.CT_SHEET in wb.sheetnames: S._ct_cols(wb[S.CT_SHEET])          # العناوين سليمة قبل ما الملف يتحرك
    finally:
        wb.close()
    return p


def after_contract(sub_en, data, dst):
    """العقد اتنقل ← سطره في 10-LOG SUB (SUB_CONTRACTS). يرجّع الناقص من الـ 3 اعتمادات."""
    return _sub().contract_record(sub_en, data['no'], data.get('date'), data['appr'], dst)


def contract_to_sub(label, ct, fields):
    """العقد اتوزع ← المقاول يتعمل/يتكمّل في قسم SUB ببنود العقد (10.1-SUBCONTRACTORS.xlsx). يرجّع (ID, بنود, جديد؟)"""
    return _sub().from_contract(label, ct['en'], ct['no'], ct.get('date'), fields.get('sub_info'), fields.get('items'))


def certify_invoice(text, ask=None):
    """فاتورة مستخلص (VN) اتوزعت ← المستخلص يتعلّم Certified في 11-LOG IPC.
    يرجّع (اتعلّم؟, رسالة). لو مبلغ الفاتورة مختلف عن المستخلص بيسأل الأول (ask)."""
    I = _ipc()
    return I.certify_from_invoice(I.parse_invoice(text), ask=ask)


def _signed_target(x, text):
    """اقتراح «مستخلص موقّع»: رقم المستخلص (من الكلام أو تاريخ نهاية الفترة) ← نفس اسم ملفه + Signed، وفي فولدره."""
    I = _ipc()
    out = I.qs_core.project_path(I.OUT_REL)
    no = I.ipc_from_text(text, out) or x['fields'].get('no')
    if not no: return
    x['fields']['no'] = no
    f = I.ipc_files(out).get(no)
    base = os.path.splitext(os.path.basename(f))[0] if f else f'IPC-{no:02d}'
    x['name'] = f'{base} - Signed.pdf'
    x['folder'] = os.path.relpath(I.ipc_folder(out, no), project_root())


def certify_signed(no, ask=None):
    """نسخة المستخلص الموقعة اتوزعت ← يسأل الأول، ولو وافقت يتعلّم Certified. يرجّع (اتعلّم؟, رسالة)."""
    return _ipc().certify_signed(no, ask=ask)


# ---------------------------------------------------------------- تقويم تحديث الشركة
def _upd():
    """مواعيد التحديث (نفس الكود اللي شاشة الحصر بتستخدمه)."""
    p = os.path.join(os.path.dirname(HERE), 'qs')
    if p not in sys.path: sys.path.append(p)
    import upd_dates
    return upd_dates


PS_PNG = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$i = [System.Drawing.Image]::FromFile($env:DC_IN)
try { $i.Save($env:DC_OUT, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $i.Dispose() }
"""


def _calendar_picture(item):
    """صورة PNG للتقويم عشان تتعرض جنب التواريخ: الـ PDF يترسم، والصورة JPG تتحول PNG. None لو مينفعش."""
    src = item['path']
    ext = os.path.splitext(src)[1].lower()
    if ext in ('.png', '.gif'): return src
    work = tempfile.mkdtemp(prefix='dc_cal_')
    try:
        if ext == '.pdf':
            pics = R.render_pages(src, 1, work)
            return pics[0] if pics else None
        if ext in R.IMG_EXT and os.name == 'nt':
            out = os.path.join(work, 'cal.png')
            import subprocess
            subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PS_PNG], capture_output=True,
                           timeout=60, env=dict(os.environ, DC_IN=os.path.abspath(src), DC_OUT=out), creationflags=R.NOWIN)
            return out if os.path.isfile(out) else None
    except Exception:
        LOG.exception('Calendar picture failed: %s', src)
    return None


def review_update_dates(parent, item):
    """تقويم تحديث الشركة: الفترات اللي اتقرت + تحديث المالك (كل ثلاثاء) جنب صورة التقويم – تراجع وتعدّل،
    ولما تدوس Save Dates تتحفظ في 00_Registers\\12-LOG UPDATE DATES.xlsx. يرجّع (رسالة, 'YYYY-MM') أو None لو لغيت."""
    U = _upd()
    src = item['path']; text = item.get('text', ''); name = os.path.basename(src)
    first = U.parse_calendar(text, name)
    reg = U.load()
    def loader(y, m):
        res = first if (y, m) == (first['year'], first['month']) else U.parse_calendar(text, name, y, m)
        warn = list(res['warn'])
        if U.month_rows(reg, U.COMPANY, y, m): warn.append(f'{y}-{m:02d} already has saved company dates – Save Dates replaces them')
        return res['company'], U.month_rows(reg, U.OWNER, y, m) or U.owner_default(y, m), warn
    pic = _calendar_picture(item)
    try:
        got = U.review(parent, first['year'], first['month'], loader, image=pic, original=src,
                       title='Company Update Calendar – check the dates')
    finally:
        if pic and os.path.basename(os.path.dirname(pic)).startswith('dc_cal_'):
            shutil.rmtree(os.path.dirname(pic), ignore_errors=True)
    if not got: return None
    y, m, company, owner = got
    U.save_month(y, m, company, owner)
    return (f'Update dates saved ({y}-{m:02d}): company {len(company)} periods, owner {len(owner)} dates', f'{y}-{m:02d}')


# ---------------------------------------------------------------- إذن التوريد والـ MIR (قسم Materials)
def page_pictures(item, n=2):
    """صور PNG لأول صفحات الورقة (عشان تتعرض في الشباك). يرجّع (القايمة، الفولدر المؤقت)."""
    src = item['path']; ext = os.path.splitext(src)[1].lower()
    work = tempfile.mkdtemp(prefix='dc_pic_')
    try:
        if ext == '.pdf': return R.render_pages(src, n, work), work
        pic = _calendar_picture(item)
        return ([pic] if pic else []), work
    except Exception:
        LOG.exception('Pictures failed: %s', src)
        return [], work


def delivery_prefill(item):
    """بيانات إذن التوريد من قراية Claude (لو موجودة) ومن القواعد."""
    s = (item.get('sugg') or [{}])[0]; f = s.get('fields') or {}
    out = dict(dn=f.get('dn', ''), inv=f.get('inv', ''), items=[])
    try: out['date'] = dt.date.fromisoformat(f['date']) if f.get('date') else None
    except ValueError: out['date'] = None
    chat = load_chat_result(item['path']) or item.get('claude_doc') or {}
    doc = next((d for d in chat.get('documents', []) if d.get('type') == 'DN'), None)
    if doc:
        out['dn'] = str(doc.get('dn') or out['dn'] or '').strip(); out['inv'] = str(doc.get('invoice') or out['inv'] or '').strip()
        out['company'] = str(doc.get('company') or '').strip()
        if doc.get('date') and not out['date']:
            try: out['date'] = dt.date.fromisoformat(doc['date'])
            except ValueError: pass
        for it in doc.get('items') or []:
            if isinstance(it, dict) and (it.get('desc') or it.get('name')):
                out['items'].append(dict(name=str(it.get('desc') or it.get('name')).strip(), code=str(it.get('code') or '').strip(),
                                         qty=it.get('qty'), unit=str(it.get('unit') or '').strip()))
    return out


def review_delivery(parent, item):
    """شباك إذن التوريد جنب صورته: الأصناف للمخزن والـ MIR. يرجّع البيانات (فيها اسم الملف) أو None لو لغيت."""
    p = os.path.join(os.path.dirname(HERE), 'mat')
    if p not in sys.path: sys.path.insert(0, p)
    import mat_app
    pics, work = page_pictures(item)
    try:
        return mat_app.delivery_dialog(parent, delivery_prefill(item), pics)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def after_delivery(data, dst):
    """إذن التوريد اتحفظ في DELIVERY ← الـ MIR (لو اتطلب) وسطور التوريدات. يرجّع (تمام، مشاكل، مسار الـ MIR)."""
    return _mat().create_delivery(data, dst)


def after_signed_mir(no, dst):
    """الـ MIR رجع متوقّع ← النسخة اللي من غير توقيع في 1_WIP تروح 4_ARCHIVED."""
    try: return _mat().archive_unsigned(no, dst)
    except Exception:
        LOG.exception('Archive unsigned MIR failed: %s', no); return []


def undo_last():
    """يرجّع آخر نقلة: الملف يرجع مكانه في INBOX بنفس اسمه القديم."""
    L = load_moves()
    for i in range(len(L) - 1, -1, -1):
        e = L[i]
        if e.get('undone') or e.get('act') not in ('move', 'review'): continue
        dst, src = e['dst'], e['src']
        if e.get('photo'):                                # صورة: نرجّع الأصل ونشيل الـ PDF المعمول
            if os.path.isfile(e['photo']): shutil.move(e['photo'], free_name(os.path.dirname(src), os.path.basename(src)))
            if os.path.isfile(dst): os.remove(dst)
        else:
            if not os.path.isfile(dst): raise FileNotFoundError(f'Moved file is no longer there:\n{dst}')
            os.makedirs(os.path.dirname(src), exist_ok=True)
            back = free_name(os.path.dirname(src), os.path.basename(src))
            shutil.move(dst, back)
            j = e.get('json')                             # تعريف Claude يرجع جنبها من 03_DONE
            if j and os.path.isfile(j) and not os.path.exists(chat_result_path(back)):
                try: shutil.move(j, chat_result_path(back))
                except OSError: LOG.exception('json restore failed: %s', j)
        e['undone'] = f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}'
        for a in L[i + 1:]:                               # النسخ القديمة اللي اتأرشفت مع النقلة دي ترجع مكانها
            if a.get('act') == 'archive' and not a.get('undone') and a.get('src') and \
                    os.path.normcase(dst).startswith(os.path.normcase(a['src'] + os.sep)):
                for n, st in a.get('back', []):
                    p = os.path.join(a['dst'], n)
                    if os.path.isfile(p): shutil.move(p, free_name(os.path.join(a['src'], st), n))
                a['undone'] = e['undone']
        save_json(MOVES, L)
        LOG.info('Undo: %s -> %s', dst, src)
        return e
    return None
