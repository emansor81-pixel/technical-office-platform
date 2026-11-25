# -*- coding: utf-8 -*-
"""Document Center Dashboard – البيانات بس (من غير Qt).

كل الأرقام من اللي موجود فعلاً: فولدرات المشروع (1_WIP / 2_SHARED / 3_PUBLISHED / 4_ARCHIVED)،
00_INBOX، وسجل 02-LOG MAR+MIR. مفيش حاجة بتتكتب – قراية بس.
"""
import os, datetime as dt

STAGES = ('1_WIP', '2_SHARED', '3_PUBLISHED', '4_ARCHIVED')
STAGE_EN = ('In preparation', 'Sent – waiting reply', 'Approved', 'Archived')
# الاسم في الشاشة: (الفولدر، فيه مجموعات جوّاه؟)
TRACKS = [('MAR', r'04_Submittals\MAR', False), ('MS', r'04_Submittals\MS', False),
          ('ITP', r'04_Submittals\ITP', False), ('SD', r'04_Submittals\SD', True),
          ('STAFF', r'04_Submittals\STAFF', False), ('MIR', r'05_Quality\MIR', False)]
OVERDUE_DAYS = 14                                    # PMM: الرد على الـ submittal خلال 14 يوم


def _skip(name):
    n = name.lower()
    return (n.startswith(('.', '~')) or n in ('readme.txt', 'desktop.ini', 'thumbs.db')
            or (n.endswith('.txt') and n.startswith('moved')))


def count_entries(folder):
    """عدد الحاجات في الفولدر (الملف = 1، والفولدر = 1 – ملف الاعتماد بيبقى فولدر كامل ساعات)."""
    try:
        return sum(1 for f in os.listdir(folder) if not _skip(f))
    except OSError:
        return 0


def stage_counts(root):
    out = []
    for name, rel, grouped in TRACKS:
        base = os.path.join(root, rel)
        tot = [0, 0, 0, 0]
        bases = [base]
        if grouped:
            try:
                bases = [os.path.join(base, g) for g in os.listdir(base)
                         if os.path.isdir(os.path.join(base, g)) and not _skip(g)]
            except OSError:
                bases = []
        for b in bases:
            for i, st in enumerate(STAGES):
                tot[i] += count_entries(os.path.join(b, st))
        out.append((name, tot))
    return out


def inbox_counts(root, review='01_REVIEW', claude='02_CLAUDE', exts=('.pdf',)):
    """(في INBOX نفسه، في المراجعة، Claude: مقري، Claude: الكل)"""
    base = os.path.join(root, '00_INBOX')

    def files(sub):
        d = os.path.join(base, sub) if sub else base
        try:
            return [f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f))
                    and f.lower().endswith(tuple(exts)) and not f.startswith(('~', '.'))]
        except OSError:
            return []

    top, rv, cl = files(''), files(review), files(claude)
    ready = sum(1 for f in cl if os.path.isfile(os.path.join(base, claude, f + '.claude.json')))
    return len(top), len(rv), ready, len(cl)


def _date(v):
    if isinstance(v, dt.datetime): return v.date()
    return v if isinstance(v, dt.date) else None


def mar_summary(mars, today=None):
    """من MAR LOG: الإجمالي، اللي اترد عليه، المستني، المتأخر (> 14 يوم)، وأقدم المستنيين."""
    today = today or dt.date.today()
    waiting = []
    replied = 0
    for m in mars:
        sent, reply = _date(m.get('sent')), _date(m.get('reply'))
        if reply: replied += 1; continue
        if not sent: continue                           # لسه ماتبعتش
        waiting.append(dict(no=m.get('no', ''), ref=m.get('ref', ''), desc=m.get('desc', ''),
                            supplier=m.get('supplier', ''), sent=sent, days=(today - sent).days,
                            status=m.get('status', '')))
    waiting.sort(key=lambda x: -x['days'])
    return dict(total=len(mars), replied=replied, waiting=len(waiting),
                overdue=sum(1 for w in waiting if w['days'] > OVERDUE_DAYS),
                not_sent=len(mars) - replied - len(waiting), oldest=waiting)


def mir_summary(mirs):
    last = max(mirs, key=lambda m: m.get('n', 0)) if mirs else None
    return dict(total=len(mirs), waiting=sum(1 for m in mirs if not _date(m.get('appr'))),
                last=(last or {}).get('no', ''), last_date=_date((last or {}).get('date')))


def collect(root, review='01_REVIEW', claude='02_CLAUDE', exts=('.pdf',), load_marmir=None):
    """كل بيانات الـ Dashboard. load_marmir = دالة بترجع {'mar': [...], 'mir': [...]} (من mat_core)."""
    d = dict(when=dt.datetime.now(), errors=[])
    d['stages'] = stage_counts(root)
    d['inbox'] = inbox_counts(root, review, claude, exts)
    d['mar'] = d['mir'] = None
    if load_marmir:
        try:
            mm = load_marmir()
            d['mar'] = mar_summary(mm.get('mar', []))
            d['mir'] = mir_summary(mm.get('mir', []))
        except Exception as e:
            d['errors'].append(f'MAR+MIR log not read: {e}')
    return d
