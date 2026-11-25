# -*- coding: utf-8 -*-
"""Documents – قوايم الورق لكل نوع (من غير Qt). قراية بس.

- الخطابات: من DC.xlsx (dc_sheet.read) + ملف الخطاب من 02_Correspondence\\OUT\\…
- MAR / MIR: من 02-LOG MAR+MIR (mat_core) + ملف كل واحد من فولدر مرحلته.
- الباقي (SD / MS / ITP / RFI): من الفولدرات نفسها (المرحلة = اسم الفولدر).
كل صف: dict فيه قيم الأعمدة + path (للمعاينة) + kind (ok / wait / bad / dim) + flags (للفلاتر).
"""
import os, re, json, datetime as dt

OVERDUE_DAYS = 14
STAGES = ('1_WIP', '2_SHARED', '3_PUBLISHED', '4_ARCHIVED')
STAGE_EN = {'1_WIP': 'In preparation', '2_SHARED': 'Sent', '3_PUBLISHED': 'Approved', '4_ARCHIVED': 'Archived'}
STAGE_KIND = {'1_WIP': 'dim', '2_SHARED': 'wait', '3_PUBLISHED': 'ok', '4_ARCHIVED': 'dim'}
PREVIEW_EXT = ('.pdf', '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp')

LETTER_COLS = [('no', 'NO.', 190), ('date', 'DATE', 95), ('desc', 'SUBJECT', 420), ('filename', 'FILE', 230), ('status', 'STATUS', 185)]
MAR_COLS = [('no', 'NO.', 55), ('ref', 'SUBMITTAL', 165), ('desc', 'MATERIAL', 280), ('supplier', 'SUPPLIER', 200),
            ('sent', 'SENT', 90), ('reply', 'REPLY', 90), ('days', 'DAYS', 55), ('status', 'STATUS', 110), ('unf', 'UNIFIER', 100)]
MIR_COLS = [('no', 'NO.', 150), ('date', 'DATE', 90), ('desc', 'MATERIAL', 300), ('supplier', 'SUPPLIER', 190),
            ('mat', 'MAR', 130), ('appr', 'APPROVED', 90), ('status', 'STATUS', 130), ('unf', 'UNIFIER', 100)]
FOLDER_COLS = [('name', 'NAME', 520), ('group', 'GROUP', 120), ('modified', 'MODIFIED', 100), ('status', 'STAGE', 120)]


def _date(v):
    if isinstance(v, dt.datetime): return v.date()
    return v if isinstance(v, dt.date) else None


def fd(v):
    d = _date(v)
    return f'{d:%d-%m-%Y}' if d else (str(v).strip() if v and not isinstance(v, (int, float)) else '')


def _norm(s): return re.sub(r'[\s_]+', '', str(s or '')).upper()


def _skip(n):
    n = n.lower()
    if '.bak-' in n or n.endswith('.part'): return True          # النسخ الاحتياطية مش ورق
    return n.startswith(('.', '~')) or n in ('readme.txt', 'desktop.ini', 'thumbs.db') or (n.endswith('.txt') and n.startswith('moved'))


# ---------------------------------------------------------------- التسليم والقفل (طلب المستخدم 21-11)
# الخطاب: مفيش اعتماد – بيتقفل لما المستخدم يدوس «Delivered». الـ submittal: بيتقفل لما يبقى Approved (السكان عدّى على Distribute).
def load_status(path):
    try:
        with open(path, encoding='utf-8') as f: d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_status(path, d):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def letter_key(no): return _norm(no)


def apply_delivery(row, date=''):
    """يحدّث صف الخطاب: date = تاريخ التسليم (dd-mm-yyyy) أو فاضي = لسه مفتوح."""
    row['delivered'] = date or ''
    row['locked'] = bool(date)
    row['editable'] = bool(row.get('docx')) and not date
    if date:
        row['status'] = f'🔒 Delivered {date}'; row['kind'] = 'ok'; row['flags'] = {'all', 'delivered'}
    else:
        row['status'] = ('Open – can edit' if row.get('docx') else 'Open – no Word file' if row.get('path') else 'Open – no file')
        row['kind'] = 'wait' if row.get('docx') else 'dim'; row['flags'] = {'all', 'open'}
    return row


def approved_from_moves(moves, code='MAR'):
    """أرقام الـ submittals اللي Distribute قراها Approved (من سجل الحركات dc_moves.json)."""
    out = set()
    for e in moves or []:
        if not isinstance(e, dict) or e.get('act') != 'move': continue
        if code not in (e.get('type'), e.get('guessed')): continue
        f = e.get('fields') or {}
        if not isinstance(f, dict): continue
        if str(f.get('status') or '').lower().startswith('approved'):
            try: out.add(int(f.get('no')))
            except (TypeError, ValueError): pass
    return out


# ---------------------------------------------------------------- الخطابات
def index_files(folder, depth=2):
    """كل ملفات الفولدر (لحد عمق معيّن): [(الاسم من غير مسافات كابيتال، المسار)]"""
    out = []

    def walk(d, lvl):
        try: names = os.listdir(d)
        except OSError: return
        for n in names:
            p = os.path.join(d, n)
            if os.path.isdir(p):
                if lvl < depth and not _skip(n): walk(p, lvl + 1)
            elif not _skip(n):
                out.append((_norm(n), p))

    walk(folder, 0)
    return out


def find_letter_file(no, index):
    """ملف الخطاب برقمه: الـ PDF الأول، وبعدين أي ملف."""
    key = _norm(no)
    if len(key) < 3: return ''
    tail = re.search(r'(\d{2,})$', key)
    hits = [p for n, p in index if key in n]
    if not hits and tail and re.search(r'[A-Z]', key):       # الاسم فيه آخر الرقم بس (0078 - الموضوع.pdf)
        t = tail.group(1)
        hits = [p for n, p in index if re.match(rf'^{t}(?!\d)', n)]
    for p in hits:
        if p.lower().endswith('.pdf'): return p
    return hits[0] if hits else ''


def find_letter_docx(no, index):
    key = _norm(no)
    if len(key) < 3: return ''
    for n, p in index:
        if p.lower().endswith('.docx') and n.startswith(key) and not n[len(key):len(key) + 1].isdigit(): return p
    return ''


def letters(rows, out_dir, delivered=None):
    """rows = dc_sheet.read(kind). الأحدث فوق. delivered = {رقم الخطاب (letter_key): تاريخ التسليم}."""
    delivered = delivered or {}
    idx = index_files(out_dir)
    out = []
    for x in rows:
        desc = x.get('desc')
        if desc is None or not str(desc).strip(): continue
        no = str(x.get('no') or '').strip()
        path = find_letter_file(no, idx) if no else ''
        r = dict(no=no, date=fd(x.get('date')), desc=str(desc).strip(),
                 filename=os.path.basename(path) if path else '—', path=path,
                 docx=find_letter_docx(no, idx) if no else '', _sort=x.get('row') or 0)
        when = fd(x.get('deliv')) or (delivered.get(letter_key(no), '') if no else '')     # الشيت هو الأصل؛ الملف القديم احتياطي
        out.append(apply_delivery(r, when))
    out.sort(key=lambda r: -r['_sort'])
    return out


# ---------------------------------------------------------------- MAR / MIR
def mars(mm, mat, today=None, approved=()):
    """mm = mat_core.load_marmir() ، mat = موديول mat_core (للمرحلة والملفات). approved = أرقام اتقرت Approved في Distribute."""
    today = today or dt.date.today()
    files = mat.folder_files(mat.MAR_DIR_REL)
    base = mat.P(mat.MAR_DIR_REL)
    marks = mat.load_marks()                             # «اترفع على Unifier» (كانت في Materials)
    out = []
    for m in mm.get('mar', []):
        sent, reply = _date(m.get('sent')), _date(m.get('reply'))
        hits = mat.mar_files(m, files); best = mat.best_file(hits)
        path = os.path.join(base, best[0], best[1]) if best else ''
        stage = mat.mar_stage(m, files)
        if m.get('n') in approved and stage not in ('Rejected',): stage = 'Approved'
        flags = {'all'}; days = ''
        if sent and reply: days = (reply - sent).days
        elif sent:
            days = (today - sent).days; flags.add('waiting')
            if days > OVERDUE_DAYS: flags.add('overdue')
        if stage == 'Approved': flags.add('approved'); flags.discard('waiting'); flags.discard('overdue')
        unf = marks.get(mat.mkey('MAR', m.get('no', '')))
        if stage == 'Approved' and not unf: flags.add('unifier')
        kind = ('ok' if stage == 'Approved' else 'bad' if ('overdue' in flags or stage in ('Rejected', 'Resubmit'))
                else 'wait' if 'waiting' in flags else 'dim')
        status = 'Overdue' if 'overdue' in flags else ('Waiting reply' if 'waiting' in flags else stage)
        locked = stage == 'Approved'
        if locked: status = '🔒 Approved'
        out.append(dict(no=m.get('no', ''), ref=m.get('ref', ''), desc=m.get('desc', ''), supplier=m.get('supplier', ''),
                        sent=fd(sent), reply=fd(reply), days=days, status=status, path=path, kind=kind, flags=flags,
                        locked=locked, editable=not locked, unf=f'✔ {unf}' if unf else '—', mark=('MAR', m.get('no', '')),
                        _sort=m.get('n', 0)))
    out.sort(key=lambda r: -r['_sort'])
    return out


def mirs(mm, mat):
    files = mat.folder_files(mat.MIR_DIR_REL)
    base = mat.P(mat.MIR_DIR_REL)
    marks = mat.load_marks()
    out = []
    for m in mm.get('mir', []):
        hits = mat.mir_files(m, files); best = mat.best_file(hits)
        path = os.path.join(base, best[0], best[1]) if best else ''
        stage = mat.mir_stage(m, files)
        kind = 'ok' if stage == 'Approved' else 'bad' if stage == 'Rejected' else 'wait' if stage in ('Signed', 'Waiting signature') else 'dim'
        flags = {'all'} | ({'approved'} if stage == 'Approved' else {'waiting'})
        if stage == 'Waiting signature': flags.add('unsigned')
        elif stage == 'Signed': flags.add('signed')
        unf = marks.get(mat.mkey('MIR', m.get('no', '')))
        if stage in ('Signed', 'Approved') and not unf: flags.add('unifier')
        locked = stage == 'Approved'
        out.append(dict(no=m.get('no', ''), date=fd(m.get('date')), desc=m.get('desc', ''), supplier=m.get('supplier', ''),
                        mat=m.get('mat', ''), appr=fd(m.get('appr')), status=('🔒 ' if locked else '') + stage, path=path, kind=kind, flags=flags,
                        locked=locked, editable=not locked, unf=f'✔ {unf}' if unf else '—', mark=('MIR', m.get('no', '')),
                        _sort=m.get('n', 0)))
    out.sort(key=lambda r: -r['_sort'])
    return out


# ---------------------------------------------------------------- من الفولدرات
def _first_preview(folder):
    """أول PDF / صورة جوه فولدر (لحد مستويين) – للمعاينة."""
    for n, p in sorted(index_files(folder, depth=2), key=lambda t: t[1].lower()):
        if p.lower().endswith(PREVIEW_EXT): return p
    return ''


def folder_rows(base, grouped=False, staged=True):
    """كل حاجة في فولدرات المراحل. grouped ← جوّاه مجموعات (SD: L1 / DETAILS …). staged=False ← فولدر واحد من غير مراحل."""
    out = []
    groups = [('', base)]
    if grouped:
        try:
            groups = [(g, os.path.join(base, g)) for g in sorted(os.listdir(base))
                      if os.path.isdir(os.path.join(base, g)) and not _skip(g)]
        except OSError:
            groups = []
    for gname, gdir in groups:
        for st in (STAGES if staged else ('',)):
            d = os.path.join(gdir, st) if st else gdir
            try: names = sorted(os.listdir(d))
            except OSError: continue
            for n in names:
                if _skip(n): continue
                p = os.path.join(d, n)
                isdir = os.path.isdir(p)
                if isdir and not staged: continue
                try: mt = dt.date.fromtimestamp(os.path.getmtime(p))
                except OSError: mt = None
                prev = (_first_preview(p) if isdir else (p if p.lower().endswith(PREVIEW_EXT) else ''))
                locked = st in ('3_PUBLISHED', '4_ARCHIVED')
                out.append(dict(name=('▸ ' + n) if isdir else n, group=gname, modified=fd(mt), raw=n, locked=locked, editable=not locked,
                                status=('🔒 ' if locked else '') + STAGE_EN.get(st, '—'), path=prev, open=p, kind=STAGE_KIND.get(st, 'dim'),
                                flags={'all'} | ({'waiting'} if st == '2_SHARED' else set()) | ({'approved'} if st == '3_PUBLISHED' else set()),
                                _sort=(-(STAGES.index(st) if st in STAGES else 0), mt or dt.date.min)))
    out.sort(key=lambda r: r['_sort'], reverse=True)
    return out


def match(row, cols, text, flag='all'):
    """الصف يظهر؟ (البحث في كل الأعمدة + الفلتر)"""
    if flag not in row.get('flags', {'all'}): return False
    t = (text or '').strip().lower()
    if not t: return True
    return any(t in str(row.get(k, '')).lower() for k, _, _ in cols)
