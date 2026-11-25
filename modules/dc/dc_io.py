# -*- coding: utf-8 -*-
"""DC Control – حفظ آمن لملفات JSON + سجل الأخطاء.

save_json: بيكتب في ملف مؤقت، وبعدين يبدّله بالأصلي مرة واحدة (os.replace) –
لو الكهربا قطعت وهو بيكتب، الملف القديم يفضل سليم.
load_json: لو الملف باظ، ياخد منه نسخة ‎.broken-…‎ ويسجّل في dc.log بدل ما يبدأ من الصفر في السكوت.
log: كل الأخطاء بتتكتب في <PLATFORM>\\logs\\dc.log (بيتقسم لما يكبر).
"""
import os, json, shutil, logging, datetime as dt
from logging.handlers import RotatingFileHandler

HERE = os.path.dirname(os.path.abspath(__file__))
PLATFORM = os.path.dirname(os.path.dirname(HERE))
LOG_DIR = os.path.join(PLATFORM, 'logs')
LOG_FILE = os.path.join(LOG_DIR, 'dc.log')


def _logger():
    lg = logging.getLogger('dc')
    if lg.handlers: return lg
    lg.setLevel(logging.INFO)
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        h = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding='utf-8')
    except Exception:
        h = logging.StreamHandler()
    h.setFormatter(logging.Formatter('%(asctime)s  %(levelname)-7s %(message)s', '%Y-%m-%d %H:%M:%S'))
    lg.addHandler(h)
    return lg


LOG = _logger()


def save_json(path, data, indent=1):
    tmp = f'{path}.tmp'
    with open(tmp, 'w', encoding='utf-8') as fh:
        json.dump(data, fh, ensure_ascii=False, indent=indent)
        fh.flush()
        try: os.fsync(fh.fileno())
        except OSError: pass
    os.replace(tmp, path)


def load_json(path, default):
    """الملف مش موجود ← القيمة الافتراضية. الملف بايظ ← نسخة منه + سطر في dc.log + القيمة الافتراضية."""
    if not os.path.isfile(path): return default
    try:
        with open(path, encoding='utf-8') as fh: return json.load(fh)
    except Exception as e:
        bad = f'{path}.broken-{dt.datetime.now():%Y%m%d-%H%M%S}'
        try: shutil.copy2(path, bad)
        except OSError: bad = '(copy failed)'
        LOG.error('Broken JSON %s (%s) – kept a copy: %s', path, e, bad)
        return default
