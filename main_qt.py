# -*- coding: utf-8 -*-
"""Project Platform – نسخة Qt (PySide6).

الشاشات اللي اتنقلت لـ Qt بتفتح هنا: الرئيسية وخريطة Document Center.
أي قسم لسه ماتنقلش (Distribute، QS، WIR، IPC، SUB، Materials) بيفتح في البرنامج القديم (main.py) في شباك لوحده.
الإعدادات (اسم المشروع ورقمه وفولدره) من نفس SETTINGS.json اللي بيقراه main.py.
لو حصل أي خطأ في شاشات Qt ← البرنامج القديم بيفتح مكانها، والخطأ في platform_error.log.
"""
import os, sys, subprocess, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
# PySide6 متسطّبة في فولدر قصير (<QTLIB>) عشان مسار بايثون الـ Store طويل وويندوز بيرفض الملفات اللي جوّاها
QTLIB = os.environ.get('QTLIB', '')
if os.path.isdir(QTLIB) and QTLIB not in sys.path: sys.path.insert(0, QTLIB)

try:
    import main as OLD                   # الإعدادات والأقسام وwrite_log (من غير ما يفتح شباك Tk)
except Exception:
    with open(os.path.join(HERE, 'platform_error.log'), 'w', encoding='utf-8') as _f: _f.write(traceback.format_exc())
    raise


def run_old(key=None):
    """يفتح البرنامج القديم (Tkinter) على قسم معيّن في شباك لوحده."""
    exe = sys.executable
    w = os.path.join(os.path.dirname(exe), 'pythonw.exe')
    if os.path.exists(w): exe = w
    code = ('import sys; sys.path.insert(0, %r); import main; a = main.Platform(); %s a.mainloop()'
            % (HERE, f'a.show({key!r});' if key else ''))
    return subprocess.Popen([exe, '-c', code], cwd=HERE)


try:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QIcon, QFont
    from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QHBoxLayout, QVBoxLayout, QStackedWidget
    import home_qt
except Exception:
    OLD.write_log(traceback.format_exc())
    run_old(); sys.exit(0)

NAVY, BG, SIDE = home_qt.T['head'], home_qt.BG, OLD.SIDE          # الألوان من home_qt (الشكل الغامق)
# الأقسام اللي اتنقلت لـ Qt (الباقي بيفتح في القديم)
QT_PAGES = {'home', 'doc'}
# أقسام Qt بتفتح في نافذة لوحدها – تقفلها ترجع للخريطة زي ما كانت.  key: (فولدرات modules، الملف، الكلاس، العنوان)
SECTIONS = {
    'dc': (['dc'], 'dc_qt', 'DistQt', 'Distribute  ·  Document Center'),
    'sub': (['sub', 'dc'], 'sub_qt', 'SubQt', 'SUB  ·  Subcontractors'),
    'wir': (['wir', 'qs', 'dc'], 'wir_qt', 'WirQt', 'WIR  ·  Inspection Requests'),
    'mar': (['mat', 'qs', 'dc'], 'mat_qt', 'MatQt', 'Store  ·  Stock / Deliveries / Issues'),
    'ipc': (['ipc', 'wir', 'qs', 'dc'], 'ipc_qt', 'IpcQt', 'IPC  ·  Payment Certificates'),
    'qs': (['qs', 'dc'], 'qs_qt', 'QsQt', 'QS  ·  Quantities'),
    'new': (['dc'], 'newdoc_qt', 'NewDocQt', 'Documents  ·  Letters  ·  Submittals'),
    'subm': (['dc'], 'submittal_qt', 'SubmittalQt', 'Submittal Builder  ·  Document Center'),
    'cycle': (['dc'], 'dash_qt', 'DashQt', 'Dashboard  ·  Document Center'),      # كارت Dashboard (مكان Submittals Cycle)
}
QT_WINDOWS = set(SECTIONS)
home_qt.HM.DOC_READY.add('new')                  # كارت New Document في Document Center بقى شغال (الخطابات)
home_qt.HM.DOC_READY.add('cycle')                # كارت Dashboard شغال (21-11)


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f'{OLD.PROJECT_NAME}  |  PO {OLD.PROJECT_NO}' if OLD.PROJECT_NO else OLD.PROJECT_NAME)
        self.setWindowIcon(app_icon())
        self.resize(1400, 800); self.setMinimumSize(1100, 640)
        self.old = {}                                    # الأقسام المفتوحة في البرنامج القديم
        self.wins = {}                                   # أقسام Qt المفتوحة في نوافذ لوحدها
        root = QWidget(); root.setAttribute(Qt.WA_StyledBackground, True); root.setObjectName('root'); root.setStyleSheet(f'#root {{ background: {BG}; }}')
        v = QVBoxLayout(root); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        v.addWidget(self._head())
        self.stack = QStackedWidget(); v.addWidget(self.stack, 1)
        self.setCentralWidget(root)
        self.current = None
        self.show_page('home')

    # ---------------- الشريط اللي فوق: ☰ (قايمة الأقسام) + اسم المشروع + إنت فين
    def _head(self):
        head = QWidget(); head.setObjectName('head'); head.setFixedHeight(48)
        head.setAttribute(Qt.WA_StyledBackground, True); head.setStyleSheet(f'#head {{ background: {NAVY}; }}')
        h = QHBoxLayout(head); h.setContentsMargins(4, 0, 16, 0); h.setSpacing(14)
        h.addSpacing(14)                                 # ☰ اتشال (طلب المستخدم) – الرجوع من الدايرة اللي فوق الخريطة
        t = QLabel(OLD.PROJECT_NAME); t.setFont(home_qt.font(19, True)); t.setStyleSheet('color: white;')
        h.addWidget(t)
        if OLD.PROJECT_NO:
            n = QLabel(f'PO {OLD.PROJECT_NO}'); n.setFont(home_qt.font(15)); n.setStyleSheet('color: #BDD7EE;')
            h.addWidget(n)
        h.addStretch(1)
        self.where = QLabel(''); self.where.setFont(home_qt.font(13)); self.where.setStyleSheet('color: #BDD7EE;')
        h.addWidget(self.where)
        return head

    def _menu(self, anchor):
        m = home_qt.nice_menu(self)
        a = m.addAction('SECTIONS'); a.setEnabled(False)
        for key, label, folder, *_ in OLD.MODULES:
            ready = key == 'home' or folder is not None
            act = m.addAction(label if ready else label + '   ·  soon', lambda k=key: self.go(k, intro=False))
            act.setEnabled(ready)
        m.exec(anchor.mapToGlobal(anchor.rect().bottomLeft()))

    # ---------------- الفتح
    def go(self, key, intro=True):
        if key.startswith('__old__:'):                   # شاشة Qt وقعت وهي مفتوحة ← نفس الشاشة في القديم
            k = key.split(':', 1)[1]
            return self.go_old_instead(k)
        if key in QT_PAGES: return self.show_page(key, intro=intro and key != self.current)
        if key in QT_WINDOWS: return self.open_window(key)
        p = self.old.get(key)
        if p and p.poll() is None:                       # مفتوح بالفعل في القديم
            self.where.setText(f'{key.upper()} is already open in its own window'); return
        try:
            self.old[key] = run_old(key)
            self.where.setText(next(m[1] for m in OLD.MODULES if m[0] == key) + '  ·  opened in its own window')
        except Exception:
            OLD.write_log(traceback.format_exc())

    def _page(self, key, intro):
        common = dict(root=OLD.PROJECT_ROOT, name=OLD.PROJECT_NAME, number=OLD.PROJECT_NO, log=OLD.write_log, intro=intro)
        if key == 'home':
            ready = {m[0] for m in OLD.MODULES if m[2]}
            return home_qt.HomeMapQt(self.go, ready, inbox_count=self._inbox_count, **common)
        return home_qt.DocMapQt(self.go, **common)

    # ---------------- الأقسام اللي بتفتح في نافذة لوحدها (Distribute)
    def open_window(self, key):
        w = self.wins.get(key)
        if w is not None:                                # مفتوحة بالفعل ← قدام
            w.showNormal(); w.raise_(); w.activateWindow(); return
        try:
            w = SectionWindow(self, key)
        except Exception:
            OLD.write_log(traceback.format_exc())
            return self.go_old_instead(key)
        self.wins[key] = w
        w.show()
        taskbar_icon(w)

    def window_closed(self, key):
        """النافذة اتقفلت ← الخريطة اللي وراها تتحدّث (العدد) من غير حركة."""
        self.wins.pop(key, None)
        cur = self.current
        if cur in QT_PAGES:
            self.current = None
            self.show_page(cur, intro=False)

    def show_page(self, key, intro=False):
        if key == self.current: return
        cur = self.stack.currentWidget()
        fn = getattr(cur, 'on_close', None)
        if fn and fn() is False: return                  # الشاشة الحالية لسه شغالة (زي Distribute وهو بيقرا)
        try:
            w = self._page(key, intro)                   # بتتبني من جديد كل مرة عشان العدد يبقى محدّث
        except Exception:
            OLD.write_log(traceback.format_exc())
            return self.go_old_instead(key)
        old = self.stack.currentWidget()
        self.stack.addWidget(w); self.stack.setCurrentWidget(w)
        if old is not None:
            self.stack.removeWidget(old); old.deleteLater()
        self.current = key
        self.where.setText(next(m[1] for m in OLD.MODULES if m[0] == key))

    def _old_once(self, key):
        """القسم في البرنامج القديم – شباك واحد بس (لو مفتوح متفتحش تاني)."""
        p = self.old.get('old:' + key)
        if p and p.poll() is None: return p
        self.old['old:' + key] = run_old(key)
        return self.old['old:' + key]

    def go_old_instead(self, key):
        """شاشة Qt وقعت ← نقول الغلط بدل ما نفتح الشكل القديم من غير ما حد يحس.

        (قبل كده كان بيفتح البرنامج القديم تلقائي – ده كان بيلخبط. طلب المستخدم 21-11.)
        """
        from PySide6.QtWidgets import QMessageBox
        log = os.path.join(HERE, 'platform_error.log')
        try:
            with open(log, encoding='utf-8') as f: last = f.read().strip().splitlines()[-1]
        except Exception:
            last = ''
        QMessageBox.critical(
            self, 'Technical Office Platform',
            f'The «{key.upper()}» screen could not open.\n\n'
            + (f'{last}\n\n' if last else '')
            + f'The full error is in:\n{log}\n\n'
              'Nothing was changed. Send this message to Claude.')
        if self.stack.count() == 0:
            if key == 'home': self.close()               # مفيش حاجة نعرضها
            else: self.show_page('home')

    def closeEvent(self, ev):
        for w in list(self.wins.values()):              # النوافذ المفتوحة تتقفل الأول (لو واحدة لسه شغالة ← استنى)
            if not w.close(): ev.ignore(); return
        fn = getattr(self.stack.currentWidget(), 'on_close', None)
        if fn and fn() is False: ev.ignore(); return
        ev.accept()

    def _inbox_count(self):
        try:
            d = os.path.join(OLD.PROJECT_ROOT, '00_INBOX')
            return sum(1 for n in os.listdir(d) if os.path.isfile(os.path.join(d, n)) and not n.lower().startswith('readme'))
        except Exception:
            return None


def app_icon():
    """أيقونة W.T.O (Water Technical Office) – ولو مش موجودة، أيقونة البرنامج القديم."""
    for n in ('wto.ico', 'app.ico'):
        p = os.path.join(HERE, 'assets', n)
        if os.path.exists(p): return QIcon(p)
    return QIcon()


AUMID = 'Contractor.WTO.Platform'


def taskbar_icon(widget):
    """يحط هوية W.T.O على النافذة نفسها (ويندوز): شريط المهام ياخد أيقونتها بدل أيقونة البايثون.
    البايثون من الـ Store ممكن يتجاهلها – ساعتها مفيش ضرر، الأيقونة بتفضل بايثون بس."""
    if os.name != 'nt': return
    try:
        import ctypes, uuid
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [('a', ctypes.c_ulong), ('b', ctypes.c_ushort), ('c', ctypes.c_ushort), ('d', ctypes.c_ubyte * 8)]

        class PKEY(ctypes.Structure):
            _fields_ = [('fmtid', GUID), ('pid', wintypes.DWORD)]

        class PROPVAR(ctypes.Structure):
            _fields_ = [('vt', ctypes.c_ushort), ('r1', ctypes.c_ushort), ('r2', ctypes.c_ushort), ('r3', ctypes.c_ushort),
                        ('p', ctypes.c_void_p), ('p2', ctypes.c_void_p)]

        def guid(t): return GUID.from_buffer_copy(uuid.UUID(t).bytes_le)
        store = ctypes.c_void_p()
        iid = guid('886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99')                     # IPropertyStore
        hr = ctypes.windll.shell32.SHGetPropertyStoreForWindow(wintypes.HWND(int(widget.winId())), ctypes.byref(iid), ctypes.byref(store))
        if hr != 0 or not store: return
        vt = ctypes.cast(store, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        SetValue = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.POINTER(PKEY), ctypes.POINTER(PROPVAR))(vt[6])
        Commit = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p)(vt[7])
        Release = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vt[2])
        fmt = guid('9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3')                      # System.AppUserModel.*
        exe = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe')
        if not os.path.exists(exe): exe = sys.executable
        ico = os.path.join(HERE, 'assets', 'wto.ico')
        keep = []
        for pid, val in ((5, AUMID), (2, f'"{exe}" "{os.path.join(HERE, "main_qt.py")}"'),
                         (4, f'{OLD.PROJECT_NO or "W.T.O"}'), (3, ico + ',0')):
            buf = ctypes.create_unicode_buffer(val); keep.append(buf)
            pv = PROPVAR(vt=31, p=ctypes.cast(buf, ctypes.c_void_p))         # VT_LPWSTR
            SetValue(store, ctypes.byref(PKEY(fmt, pid)), ctypes.byref(pv))
        Commit(store); Release(store)
    except Exception:
        OLD.write_log(traceback.format_exc())


class SectionWindow(QWidget):
    """نافذة لوحدها لقسم Qt (Distribute). ✕ أو «← Document Center» يقفلها، وترجع للخريطة زي ما كنت."""

    def __init__(self, main, key):
        super().__init__(None, Qt.Window)
        self.main, self.key = main, key
        dirs, mod, cls, title = SECTIONS[key]
        self.setWindowTitle(title)
        self.setWindowIcon(app_icon())
        self.setAttribute(Qt.WA_DeleteOnClose)
        v = QVBoxLayout(self); v.setContentsMargins(0, 0, 0, 0)
        for x in dirs:
            p = os.path.join(HERE, 'modules', x)
            if p not in sys.path: sys.path.insert(0, p)
        import importlib
        m = importlib.import_module(mod)
        self.page = getattr(m, cls)(open_section=self._section, run_old=main._old_once, log=OLD.write_log)
        v.addWidget(self.page)
        scr = main.screen().availableGeometry()
        self.resize(int(scr.width() * .9), int(scr.height() * .88))
        self.move(scr.center() - self.rect().center())

    def _section(self, k):
        if k == 'doc': self.close()                      # «← Document Center» = اقفل النافذة
        else: self.main.go(k)

    def closeEvent(self, ev):
        fn = getattr(self.page, 'on_close', None)
        if fn and fn() is False: ev.ignore(); return
        ev.accept()
        self.main.window_closed(self.key)


def excepthook(t, e, tb):
    err = ''.join(traceback.format_exception(t, e, tb))
    OLD.write_log(err); print(err)


if __name__ == '__main__':
    sys.excepthook = excepthook
    try:
        try:                                             # شريط المهام يعرض أيقونة W.T.O بدل أيقونة البايثون
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(AUMID)
        except Exception:
            pass
        app = QApplication(sys.argv)
        app.setWindowIcon(app_icon())
        app.setFont(QFont('Segoe UI', 10))
        try:                                             # البرنامج غامق دايماً (طلب المستخدم 12-11) – القوايم والرسايل كمان
            app.styleHints().setColorScheme(Qt.ColorScheme.Dark if home_qt.DARK else Qt.ColorScheme.Light)
        except Exception:
            pass
        win = Window()
        if win.stack.count() == 0: sys.exit(0)          # الرئيسية ماطلعتش ← القديم اتفتح مكانها
        win.showMaximized()
        taskbar_icon(win)
        sys.exit(app.exec())
    except SystemExit:
        raise
    except Exception:
        OLD.write_log(traceback.format_exc())
        run_old()
