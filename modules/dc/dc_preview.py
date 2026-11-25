# -*- coding: utf-8 -*-
"""Preview pane – بيعرض صفحة الـ PDF (أو الصورة) جنب الجدول.

محرك العرض: PyMuPDF (pip install pymupdf) وإن مكانش موجود يرجع لـ QtPdf.
بيتستخدم في: Distribute (dc_qt) – ويقدر يتستخدم بعدين في Submittal Builder.
"""
import os, sys

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QPushButton,
                               QHBoxLayout, QVBoxLayout, QScrollArea, QSizePolicy)

PLAT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if PLAT not in sys.path: sys.path.insert(0, PLAT)
import home_qt as HQ

BG, CARD, INK, MUTED, SOFT = HQ.BG, HQ.CARD, HQ.INK, HQ.MUTED, HQ.SOFT_LINE
TT = HQ.T
ACCENT = '#7FB3E8' if HQ.DARK else '#1F5C99'
BTN_BG = TT.get('card2', CARD)                       # خلفية الزراير
PAPER_BG = BG if HQ.DARK else '#E9EDF1'              # اللي ورا الورقة

IMG_EXT = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp'}
ZOOMS = [0.35, 0.5, 0.65, 0.8, 1.0, 1.25, 1.5, 2.0, 3.0]


def _fitz():
    """PyMuPDF لو متنصب."""
    try:
        import fitz
        return fitz
    except Exception:
        try:
            import pymupdf
            return pymupdf
        except Exception:
            return None


class PreviewPane(QWidget):
    """load(path) ← يعرض الملف. clear() ← يفضّيه."""

    def __init__(self, parent=None, width=430):
        super().__init__(parent)
        self.setObjectName('pvpane')
        self.setFixedWidth(width)
        self.path = None
        self.page = 0
        self.n = 0
        self.zoom = None          # None = Fit
        self.engine = ''
        self._doc = None
        self._img = None          # صورة (مش PDF)
        self._fz = None           # موديول PyMuPDF المستخدم
        self._rendering = False   # يمنع اللف في resize
        self._build()

    # ------------------------------------------------------------ شكل
    def _build(self):
        self.setStyleSheet(f"""
QWidget#pvpane {{ background: {CARD}; border: 1px solid {ACCENT}; border-radius: 10px; }}
QLabel#pvtitle {{ color: {ACCENT}; font: bold 10pt "Segoe UI"; letter-spacing: .5px; }}
QLabel#pvinfo  {{ color: {MUTED}; font: 9.5pt "Segoe UI"; }}
QPushButton#pvb {{ background: {BTN_BG}; color: {INK}; border: 1px solid {SOFT};
                   border-radius: 6px; padding: 3px 9px; font: 9.5pt "Segoe UI"; }}
QPushButton#pvb:hover  {{ border-color: {ACCENT}; }}
QPushButton#pvb:disabled {{ color: {MUTED}; }}
QScrollArea#pvsc {{ background: {PAPER_BG}; border: none; }}
QScrollArea#pvsc > QWidget > QWidget {{ background: {PAPER_BG}; }}
""")
        v = QVBoxLayout(self); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)

        # الهيدر: العنوان + أسهم الصفحات
        head = QWidget(); h = QHBoxLayout(head); h.setContentsMargins(12, 9, 10, 8); h.setSpacing(7)
        t = QLabel('PREVIEW'); t.setObjectName('pvtitle'); h.addWidget(t)
        h.addStretch(1)
        self.b_prev = self._btn('‹', self.prev_page, 26)
        self.l_page = QLabel('–'); self.l_page.setObjectName('pvinfo')
        self.l_page.setAlignment(Qt.AlignCenter); self.l_page.setMinimumWidth(46)
        self.b_next = self._btn('›', self.next_page, 26)
        h.addWidget(self.b_prev); h.addWidget(self.l_page); h.addWidget(self.b_next)
        v.addWidget(head)
        v.addWidget(self._line())

        # الورقة
        self.sc = QScrollArea(); self.sc.setObjectName('pvsc'); self.sc.setWidgetResizable(True)
        self.sc.setAlignment(Qt.AlignCenter)
        self.lab = QLabel('Select a paper from the list')
        self.lab.setAlignment(Qt.AlignCenter); self.lab.setWordWrap(True)
        self.lab.setStyleSheet(f'color: {MUTED}; font: 10pt "Segoe UI"; padding: 24px;')
        self.lab.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.sc.setWidget(self.lab)
        v.addWidget(self.sc, 1)

        # الفوتر: زوم + اسم المحرك
        v.addWidget(self._line())
        foot = QWidget(); f = QHBoxLayout(foot); f.setContentsMargins(10, 7, 10, 8); f.setSpacing(6)
        f.addWidget(self._btn('−', lambda: self.step_zoom(-1)))
        self.b_fit = self._btn('Fit', self.fit)
        f.addWidget(self.b_fit)
        f.addWidget(self._btn('+', lambda: self.step_zoom(+1)))
        f.addStretch(1)
        self.l_eng = QLabel(''); self.l_eng.setObjectName('pvinfo'); f.addWidget(self.l_eng)
        v.addWidget(foot)
        self._buttons(False)

    def _btn(self, text, slot, w=0):
        b = QPushButton(text); b.setObjectName('pvb'); b.clicked.connect(slot)
        b.setCursor(Qt.PointingHandCursor)
        if w: b.setFixedWidth(w)
        return b

    def _line(self):
        s = QFrame(); s.setFixedHeight(1); s.setStyleSheet(f'background: {SOFT};'); return s

    def _buttons(self, on):
        self.b_prev.setEnabled(on and self.page > 0)
        self.b_next.setEnabled(on and self.page < self.n - 1)

    # ------------------------------------------------------------ تحميل
    def clear(self, note='Select a paper from the list'):
        self._close()
        self.path = None; self.page = 0; self.n = 0; self._img = None
        self.lab.setPixmap(QPixmap())
        self.lab.setMinimumSize(0, 0); self.lab.setMaximumSize(16777215, 16777215)
        self.lab.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.lab.setText(note)
        self.l_page.setText('–'); self.l_eng.setText(''); self._buttons(False)

    def _close(self):
        if self._doc is not None:
            try: self._doc.close()
            except Exception: pass
            self._doc = None

    def load(self, path, page=0):
        """يفتح ملف (PDF أو صورة). بيرجع True لو عرض."""
        if not path or not os.path.isfile(path):
            self.clear('File not found'); return False
        if path == self.path and self._live():
            self.page = max(0, min(page, self.n - 1)); return self.render()
        self._close()
        self.path = path; self.page = 0; self._img = None
        ext = os.path.splitext(path)[1].lower()

        if ext in ('.xlsx', '.xls', '.xlsm', '.zip'):        # إكسيل / zip ← مفيش معاينة (افتحه من Open)
            self.clear(f'{"Excel" if ext != ".zip" else "Zip"} file – no preview\nuse Open to see it'); return False

        if ext in IMG_EXT:
            img = QImage(path)
            if img.isNull():
                self.clear('Could not read the picture'); return False
            self._img = img; self.n = 1; self.engine = 'image'
            return self.render()

        fz = _fitz()
        if fz is not None:
            try:
                # بنقراه في الذاكرة عشان الملف ما يفضلش مقفول (الـ Approve بينقله)
                with open(path, 'rb') as fh: data = fh.read()
                self._doc = fz.open(stream=data, filetype='pdf')
                self._fz = fz
                self.n = self._doc.page_count; self.engine = 'PyMuPDF'
                return self.render()
            except Exception as e:
                self._doc = None
                self.clear(f'Could not open the file\n{e}'); return False

        # QtPdf احتياطي
        try:
            from PySide6.QtPdf import QPdfDocument
            d = QPdfDocument(self); d.load(path)
            self._doc = d; self.n = d.pageCount(); self.engine = 'QtPdf'
            return self.render()
        except Exception:
            self.clear('Preview engine not installed –\nrun INSTALL_PREVIEW.bat\n\n'
                       f'This program runs on:\n{sys.executable}'); return False

    def _live(self):
        return self._img is not None or self._doc is not None

    # ------------------------------------------------------------ عرض
    def _target_width(self):
        w = self.sc.viewport().width() - 16
        return max(220, w)

    def render(self):
        if not self._live() or self._rendering: return False
        self._rendering = True
        try:
            if self._img is not None:
                pm = QPixmap.fromImage(self._img)
                w = self._target_width() if self.zoom is None else int(self._img.width() * self.zoom)
                pm = pm.scaledToWidth(max(60, w), Qt.SmoothTransformation)
            elif self.engine == 'PyMuPDF':
                pg = self._doc[self.page]
                rw = pg.rect.width or 595.0
                z = (self._target_width() / rw) if self.zoom is None else self.zoom
                z = max(0.08, min(z, 6.0))
                px = pg.get_pixmap(matrix=self._fz.Matrix(z, z), alpha=False)
                im = QImage(px.samples, px.width, px.height, px.stride, QImage.Format_RGB888).copy()
                pm = QPixmap.fromImage(im)
            else:                                   # QtPdf
                ps = self._doc.pagePointSize(self.page)
                w = self._target_width() if self.zoom is None else int((ps.width() or 595) * self.zoom)
                h = int(w * ps.height() / ps.width()) if ps.width() else int(w * 1.414)
                pm = QPixmap.fromImage(self._doc.render(self.page, QSize(w, h)))
        except Exception as e:
            self._rendering = False
            self.lab.setPixmap(QPixmap()); self.lab.setText(f'Could not draw the page\n{e}')
            self.l_eng.setText(''); self._buttons(False); return False

        try:
            self.lab.setText('')
            self.lab.setPixmap(pm)
            self.lab.setFixedSize(pm.size())
            self.lab.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.l_page.setText(f'{self.page + 1} / {self.n}' if self.n else '–')
            self.l_eng.setText(f'{self.engine}  ·  {self.n} page(s)' if self.n > 1 else self.engine)
            self._buttons(True)
        finally:
            self._rendering = False
        return True

    # ------------------------------------------------------------ تنقل وزوم
    def prev_page(self):
        if self.page > 0: self.page -= 1; self.render()

    def next_page(self):
        if self.page < self.n - 1: self.page += 1; self.render()

    def fit(self):
        self.zoom = None; self.render()

    def step_zoom(self, d):
        cur = self.zoom
        if cur is None:                            # من Fit ← أقرب درجة
            try:
                if self.engine == 'PyMuPDF' and self._doc is not None:
                    cur = self._target_width() / (self._doc[self.page].rect.width or 595.0)
                elif self._img is not None:
                    cur = self._target_width() / max(1, self._img.width())
                else:
                    cur = 1.0
            except Exception:
                cur = 1.0
        i = min(range(len(ZOOMS)), key=lambda k: abs(ZOOMS[k] - cur))
        i = max(0, min(i + d, len(ZOOMS) - 1))
        self.zoom = ZOOMS[i]; self.render()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self.zoom is None and self._live(): self.render()
