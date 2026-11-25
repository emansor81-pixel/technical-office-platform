# -*- coding: utf-8 -*-
"""تحديث أوتوماتيك لشيت OUR_SHEET (وشيت الاستشاري أونلاين) من غير ما البرنامج يكون مفتوح.
ويندوز بيشغّله كل ساعة (مهمة مجدولة اسمها CTR IPC Auto Update – بتتسجّل من <TOOLS>\\SETUP_IPC_AUTO.bat).
- رقم المستخلص = آخر مستخلص طالع في CLIENT_IPC + 1، والفترة من بعد آخر مستخلص لحد النهارده.
- لو مفيش طلبات جديدة اتعتمدت من آخر تشغيلة: مبيفتحش الملفات خالص.
- لو OUR_SHEET مفتوح عند حد: بيسيبه للتشغيلة الجاية.
- كل تشغيلة بتتكتب سطر في CLIENT_IPC\\HO_AUTO_LOG.txt
- نسخة احتياطية واحدة في اليوم في CLIENT_IPC\\HO_BACKUP."""
import os, sys, json, tempfile, traceback, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'qs'))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'wir'))
import qs_core  # noqa: E402
import ipc_core as C  # noqa: E402
import ipc_ho  # noqa: E402

INCLUDE_HO = True          # False = OUR_SHEET بس


def log(out_dir, msg):
    try:
        with open(os.path.join(out_dir, 'HO_AUTO_LOG.txt'), 'a', encoding='utf-8') as f:
            f.write(f'{dt.datetime.now():%Y-%m-%d %H:%M}  {msg}\n')
    except Exception:
        pass


def main():
    out_dir = qs_core.project_path(C.OUT_REL)
    template = qs_core.project_path(C.TEMPLATE_REL)
    lock = os.path.join(tempfile.gettempdir(), 'otc_ipc_auto.lock')
    try:                                                         # تشغيلة واحدة بس في نفس الوقت
        if os.path.exists(lock) and (dt.datetime.now().timestamp() - os.path.getmtime(lock)) < 3600: return
        open(lock, 'w').close()
    except Exception:
        pass
    try:
        st = ipc_ho.load_settings()
        local = ipc_ho.full_path(st['LOCAL_SHEET'])
        owner = os.path.join(os.path.dirname(local), '~$' + os.path.basename(local))
        if os.path.exists(owner):
            log(out_dir, 'اتأجل: OUR_SHEET مفتوح'); return
        files = C.ipc_files(out_dir) if os.path.isdir(out_dir) else {}
        last = max(files) if files else 0
        no = last + 1
        d2 = dt.date.today()
        d1 = d2.replace(day=1)
        if last:
            try:
                _, prev_to = C.read_todate(files[last])
                if isinstance(prev_to, dt.datetime): prev_to = prev_to.date()
                if isinstance(prev_to, dt.date): d1 = prev_to + dt.timedelta(days=1)
            except Exception:
                pass
        if d1 > d2: d1 = d2
        p = qs_core.Project(qs_core.DEFAULT_INPUT)
        p.load_requests()
        calc = C.compute(p, no, d2, out_dir, template, d1)
        n, notes, h = ipc_ho.update_ho(p, calc, no, d1, d2, backup_dir=os.path.join(out_dir, 'HO_BACKUP'),
                                       include_ho=INCLUDE_HO, tag=f'auto {d2:%Y-%m-%d}',
                                       skip_hash=st.get('AUTO_LAST_HASH', ''))
        if notes == ['مفيش جديد'] and n == 0:
            log(out_dir, f'مفيش جديد (مستخلص {no})'); return
        st = ipc_ho.load_settings(); st['AUTO_LAST_HASH'] = h; ipc_ho.save_settings(st)
        what = 'OUR_SHEET + شيت الاستشاري' if INCLUDE_HO else 'OUR_SHEET'
        log(out_dir, f'اتحدّث {what}: {n} مقطع – مستخلص {no} ({d1:%d-%m-%Y} → {d2:%d-%m-%Y})'
                     + (' | ' + ' | '.join(notes[:5]) if notes else ''))
    except Exception as e:
        log(out_dir, 'غلط: ' + ' '.join(str(e).split())[:400])
        try:
            with open(os.path.join(out_dir, 'HO_AUTO_ERROR.txt'), 'w', encoding='utf-8') as f: traceback.print_exc(file=f)
        except Exception:
            pass
    finally:
        try: os.remove(lock)
        except Exception: pass


if __name__ == '__main__':
    main()
