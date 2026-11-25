# -*- coding: utf-8 -*-
"""الشاشة الرئيسية وخريطة Document Center بـ Qt (PySide6).

نفس الخريطة اللي في home_map.py بالظبط (نفس الأماكن والألوان والكروت والسجلات والكلام) – الداتا نفسها
بتتقري من home_map.py، فأي تعديل هناك (كارت أو سجل أو شرح) بيظهر هنا لوحده. اللي اتغيّر الشكل بس:
كروت بظل وحواف مدوّرة، رسم ناعم، أسامي السجلات أزرار صغيرة، وقايمة حديثة بدل قايمة ويندوز القديمة.

- الماوس على قسم ← تفاصيله على اليمين.  على سجل ← الأقسام اللي بتستخدمه بس اللي تفضل منوّرة.
- كليك على الدايرة ← الأقسام تطلع منها (أو ترجعلها).  كليك على قسم ← يفتحه.  على سجل ← يفتح ملفه (أو قايمة).
- «الدايرة تنتقل»: كليك DOC ← الكروت ترجع للدايرة، وبعدين خريطة Document Center تطلع من دايرتها. ← Home بالعكس.
"""
import os, sys, subprocess, traceback

from PySide6.QtCore import Qt, QRectF, QPointF, QTimer, QVariantAnimation, QEasingCurve, QAbstractAnimation
from PySide6.QtGui import QPainter, QColor, QFont, QFontMetricsF, QPen, QBrush, QPainterPath, QRadialGradient, QCursor
from PySide6.QtWidgets import (QWidget, QLabel, QFrame, QVBoxLayout, QHBoxLayout, QMenu, QMessageBox,
                               QGraphicsDropShadowEffect, QSizePolicy, QScrollArea)

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path: sys.path.insert(0, HERE)
import home_map as HM            # الداتا بس (الأقسام، السجلات، الشرح، الألوان) – مفيش أي شباك Tk بيتفتح

BG, CARD, LINE, INK, MUTED = HM.BG, HM.CARD, HM.LINE, HM.INK, HM.MUTED
NAVY, NAVY2, COL = HM.NAVY, HM.NAVY2, HM.COL
# ---------------------------------------------------------------- الشكل الغامق (طلب المستخدم 12-11: «غامق بس»)
DARK = True
T = dict(card2=CARD, chip='#F6F8FB', pill='#EAF1F8', pill_hot='#D6E6F5', pill_fg=NAVY, core0='#2B4D68', core1=INK,
         back0='#3A5C78', back0_hot='#4A6F8F', ring='#FFFFFF', planned_border='#E8DCCB', wait_fg='#6B4A00', wait_bg='#FFF4D6',
         ok='#375623', bad='#C00000', title=NAVY, head='#1F4E78', menu_sel='#EAF1F8', shadow=INK)
if DARK:
    BG, CARD, LINE, INK, MUTED = '#0E151C', '#17212B', '#33424F', '#E3EAF1', '#8C9AAA'
    NAVY, NAVY2 = '#2B5F8E', '#24496C'
    COL = dict(HM.COL, dc='#5EA3E6', qs='#3CC2B4', wir='#5FC46B', ipc='#8E92F0', sub='#E58A55', mat='#D2BC4A',
               planned='#E0A24A', wlog='#F06688', arch='#9AAABB', upd='#C58BE3')
    T = dict(card2='#1C2834', chip='#1C2834', pill='#1E3448', pill_hot='#284A66', pill_fg='#9CC8F0', core0='#3D7AB3', core1='#173650',
             back0='#3D7AB3', back0_hot='#4F8CC5', ring='#2A3846', planned_border='#4A3A22', wait_fg='#F0C860', wait_bg='#3A3018',
             ok='#8FD18F', bad='#FF8A8A', title='#9CC8F0', head='#0B2238', menu_sel='#1E3448', shadow='#000000')
THEME = T                                           # نفس T (جوه _card فيه متغيّر اسمه T = أول الكارت)
FAM, SYM, MONO = 'Segoe UI', 'Segoe UI Symbol', 'Consolas'
SOFT_LINE = '#26333F' if DARK else '#E3E8EE'
# رموز ويندوز بيرسمها «إيموجي» ملوّن في Qt ← رمز قريب بيترسم بلون الكارت زي الباقي
ICON_FIX = {'✍': '✎'}

_OPENED = False      # الرئيسية اتفتحت مرة في التشغيلة دي؟ ← لما ترجعلها الكروت تفضل ظاهرة


def inbox_dirs(base):
    """فولدر المراجعة وفولدر Claude جوه INBOX: الأسامي بالأرقام (01_REVIEW / 02_CLAUDE)، أو القديمة لو لسه ماتغيّرتش."""
    def pick(new, old):
        return old if not os.path.isdir(os.path.join(base, new)) and os.path.isdir(os.path.join(base, old)) else new
    return pick('01_REVIEW', 'مراجعة'), pick('02_CLAUDE', 'Claude')


def blend(c, t, to=None):
    to = to or T['card2']
    return HM._blend(c, t, to)


def qc(c, alpha=None):
    q = QColor(c)
    if alpha is not None: q.setAlpha(alpha)
    return q


def font(px, bold=False, fam=FAM):
    f = QFont(fam)
    f.setPixelSize(max(8, int(round(px))))
    if bold: f.setWeight(QFont.Weight.Bold)
    return f


def shadow(w, blur=22, dy=3, alpha=26):
    e = QGraphicsDropShadowEffect(w)
    e.setBlurRadius(blur); e.setOffset(0, dy); e.setColor(qc(T['shadow'], min(255, alpha * (3 if DARK else 1))))
    w.setGraphicsEffect(e)


MENU_CSS = f"""
QMenu {{ background: {T['card2']}; border: 1px solid {SOFT_LINE}; border-radius: 12px; padding: 6px; }}
QMenu::item {{ padding: 8px 22px 8px 12px; border-radius: 8px; color: {INK}; font: 10pt "{FAM}"; }}
QMenu::item:selected {{ background: {T['menu_sel']}; color: {T['pill_fg']}; }}
QMenu::item:disabled {{ color: {MUTED}; font: bold 8pt "{FAM}"; padding: 6px 12px 4px 12px; }}
QMenu::separator {{ height: 1px; background: {SOFT_LINE}; margin: 4px 8px; }}
"""


def nice_menu(parent):
    """قايمة بحواف مدوّرة وظل (بدل قايمة ويندوز القديمة)."""
    m = QMenu(parent)
    m.setWindowFlags(m.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
    m.setAttribute(Qt.WA_TranslucentBackground)
    m.setStyleSheet(MENU_CSS)
    return m


# ====================================================================== الخريطة نفسها (الرسم)
class MapCanvas(QWidget):
    """الدايرة في النص، والكروت حواليها، والسجلات تحت. نفس حسابات home_map.HomeMap (بوحدات الرسم، النص = 0,0)."""

    def __init__(self, page):
        super().__init__(page)
        self.page = page
        self.setMouseTracking(True)
        self.setMinimumSize(500, 380)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.hits = []                      # (شكل، obj) – بيتملي مع كل رسمة
        self.p = 1.0 if page.expanded else 0.0   # 1 = الكروت في أماكنها، 0 = ورا الدايرة
        self.fly = None                     # الكارت اللي رايح للنص في «الدايرة تنتقل»
        self.anim = None

    # ---------------------------------------------------------------- الحسابات
    def frame(self):
        x0, y0, x1, y1 = self.page.bbox
        W, H = self.width(), self.height()
        s = max(0.5, min(W / (x1 - x0), H / (y1 - y0), 1.35))
        ox, oy = W / 2 - (x0 + x1) / 2 * s, H / 2 - (y0 + y1) / 2 * s
        return s, ox, oy

    def animate(self, opening, done=None, ms=260):
        """الكروت تطلع من الدايرة (opening) أو ترجعلها – بحركة ناعمة."""
        if self.anim: return
        a = QVariantAnimation(self)
        a.setStartValue(self.p); a.setEndValue(1.0 if opening else 0.0)
        a.setDuration(ms); a.setEasingCurve(QEasingCurve.OutCubic)
        a.valueChanged.connect(self._tick)

        def end():
            self.anim = None
            self.page.anim_done(opening)
            if done: done()
        a.finished.connect(end)
        self.anim = a
        a.start(QAbstractAnimation.DeletionPolicy.DeleteWhenStopped)

    def _tick(self, v):
        self.p = float(v)
        self.update()

    # ---------------------------------------------------------------- الرسم
    def paintEvent(self, ev):
        if getattr(self, 'failed', False): return
        qp = QPainter(self)
        try:
            self._paint(qp)
        except Exception:
            self.failed = True
            try: self.page.log and self.page.log(traceback.format_exc())
            except Exception: pass
            QTimer.singleShot(0, self.page.on_fail)
        finally:
            qp.end()

    def _paint(self, qp):
        P = self.page
        qp.setRenderHint(QPainter.Antialiasing)
        qp.setRenderHint(QPainter.TextAntialiasing)
        # الخلفية: كارت أبيض كبير بحواف مدوّرة
        qp.setPen(QPen(qc(SOFT_LINE), 1)); qp.setBrush(qc(CARD))
        qp.drawRoundedRect(QRectF(self.rect()).adjusted(1, 1, -1, -1), 18, 18)
        s, ox, oy = self.frame()
        X = lambda x: ox + x * s
        Y = lambda y: oy + y * s
        self.hits = []
        moving = self.anim is not None or self.fly is not None
        settled = self.p >= 0.999 and not moving
        lit = set(P.regs[P.hot]['uses']) if P.hot in P.regs else None

        # الخطوط: أفقي ورأسي بس، كل كارت خطه لوحده (بتظهر لما الكروت تكون في أماكنها بس)
        if settled:
            for k, d in P.sec.items():
                c = '#C9A774' if d['planned'] else COL[d['c']]
                qp.setOpacity(0.3 if (lit is not None and k not in lit) else 1.0)
                pen = QPen(qc(c), max(2.0, 3 * s)); pen.setCapStyle(Qt.RoundCap)
                qp.setPen(pen)
                for a, b, e, f in d['segs']:
                    qp.drawLine(QPointF(X(a), Y(b)), QPointF(X(e), Y(f)))
            qp.setOpacity(1.0)

        if P.BACK: self._back(qp, X, Y, s)           # الخريطة اللي جيت منها: دايرة صغيرة فوق بخط متقطّع
        # الكروت (ورا الدايرة وهي بتتحرك)، والكارت اللي رايح للنص فوقها
        if self.p > 0.001:
            for k in P.sec:
                if k != self.fly: self._card(qp, k, X, Y, s, lit)
        self._core(qp, X, Y, s)
        if self.fly and self.p > 0.001:
            self._card(qp, self.fly, X, Y, s, lit)

        if settled:
            self._legend(qp, X, Y, s)
        elif self.p <= 0.001 and not moving:
            qp.setPen(qc(MUTED)); qp.setFont(font(14 * s))
            r = QRectF(X(-300), Y(P.CORE + 14), 600 * s, 26 * s)
            qp.drawText(r, Qt.AlignCenter, 'Click to show the sections')

    def _shadow(self, qp, r, rad, strength=1.0, lift=0.0):
        """ظل ناعم تحت الكارت (طبقات شفافة)."""
        qp.setPen(Qt.NoPen)
        for i in range(7, 0, -1):
            g = i * 1.6
            qp.setBrush(qc(T['shadow'], int((5 + lift * 3) * strength * (2.5 if DARK else 1))))
            qp.drawRoundedRect(r.adjusted(-g * .6, -g * .3 + 2 + lift, g * .6, g + 2 + lift), rad + g * .6, rad + g * .6)

    def _card(self, qp, k, X, Y, s, lit):
        P, d = self.page, self.page.sec[k]
        c = COL[d['c']]
        x1, y1, x2, y2 = d['box']
        dx, dy = -d['x'] * (1 - self.p), -d['y'] * (1 - self.p)       # الحركة من/للدايرة
        x1, y1, x2, y2 = x1 + dx, y1 + dy, x2 + dx, y2 + dy
        r = QRectF(X(x1), Y(y1), (x2 - x1) * s, (y2 - y1) * s)
        hot = P.hot == k and self.anim is None
        faded = lit is not None and k not in lit
        qp.save()
        qp.setOpacity(0.32 if faded else 1.0)
        lift = 2.0 * s if hot else 0.0
        r = r.translated(0, -lift)
        self._shadow(qp, r, 14 * s, 1.0, 2.0 if hot else 0.0)
        border = qc(c) if hot else qc(THEME['planned_border'] if d['planned'] else SOFT_LINE)
        qp.setPen(QPen(border, max(2.0, 2.2 * s) if hot else 1.0))
        qp.setBrush(qc(THEME['card2']))
        qp.drawRoundedRect(r, 14 * s, 14 * s)
        L, T = r.left(), r.top()
        # الأيقونة
        tint = blend(c, .86)
        ir = QRectF(L + 16 * s, T + 16 * s, 40 * s, 40 * s)
        qp.setPen(Qt.NoPen); qp.setBrush(qc(tint)); qp.drawRoundedRect(ir, 11 * s, 11 * s)
        qp.setPen(qc(c)); qp.setFont(font(21 * s, fam=SYM))
        qp.drawText(ir, Qt.AlignCenter, ICON_FIX.get(P.ICONS.get(k, '▣'), P.ICONS.get(k, '▣')))
        # الاسم والسطر الصغير
        badge = P.badge(k)
        bw = 0.0
        if badge:
            bf = font(12.5 * s, True)
            bw = QFontMetricsF(bf).horizontalAdvance(badge[0]) + 18 * s + (14 * s if badge[3] else 0) + 8 * s
        nf = font(22 * s, True)
        # الاسم الطويل (زي Claude Folder): الشارة تنزل لسطر Shares على اليمين بدل ما الاسم يتقص – لو فيه مكان
        low = False
        if badge and QFontMetricsF(nf).horizontalAdvance(d['t']) > r.width() - 82 * s - bw:
            sf0 = font(12.5 * s)
            n = sum(1 for reg in P.regs.values() if k in reg['uses'])
            dots_end = L + 16 * s + QFontMetricsF(sf0).horizontalAdvance('Shares') + 10 * s + n * 17 * s
            low = r.right() - 14 * s - (bw - 8 * s) > dots_end + 10 * s
        qp.setFont(nf); qp.setPen(qc(c))
        avail = r.width() - 82 * s - (0 if low else bw)
        qp.drawText(QRectF(L + 68 * s, T + 14 * s, avail, 30 * s), Qt.AlignLeft | Qt.AlignVCenter,
                    QFontMetricsF(nf).elidedText(d['t'], Qt.ElideRight, avail))
        qp.setFont(font(14.5 * s)); qp.setPen(qc(MUTED))
        qp.drawText(QRectF(L + 68 * s, T + 40 * s, r.width() - 80 * s, 22 * s), Qt.AlignLeft | Qt.AlignVCenter, d['s'])
        # الشارة فوق على اليمين (زي القديم) – الاسم بيتقص قبلها لو طويل
        if badge:
            text, fg, bgc, dot = badge
            bf = font(12.5 * s, True)
            w = QFontMetricsF(bf).horizontalAdvance(text) + 18 * s + (14 * s if dot else 0)
            br = QRectF(r.right() - 14 * s - w, (T + 76 * s) if low else (T + 16 * s), w, 24 * s)
            qp.setPen(Qt.NoPen); qp.setBrush(qc(bgc)); qp.drawRoundedRect(br, 12 * s, 12 * s)
            tx = br.left() + 9 * s
            if dot:
                qp.setBrush(qc(dot)); qp.drawEllipse(QPointF(tx + 4 * s, br.center().y()), 4 * s, 4 * s); tx += 14 * s
            qp.setPen(qc(fg)); qp.setFont(bf)
            qp.drawText(QRectF(tx, br.top(), br.right() - tx, br.height()), Qt.AlignLeft | Qt.AlignVCenter, text)
        # النقط = السجلات اللي القسم بيشاركها
        sf = font(12.5 * s)
        qp.setFont(sf); qp.setPen(qc(MUTED))
        cy = T + 88 * s
        qp.drawText(QRectF(L + 16 * s, cy - 10 * s, 80 * s, 20 * s), Qt.AlignLeft | Qt.AlignVCenter, 'Shares')
        ddx = L + 16 * s + QFontMetricsF(sf).horizontalAdvance('Shares') + 10 * s
        dots = []
        for rid, reg in P.regs.items():
            if k not in reg['uses']: continue
            ctr = QPointF(ddx + 5.5 * s, cy)
            qp.setPen(QPen(qc(INK), 2.0) if P.hot == rid else Qt.NoPen)
            qp.setBrush(qc(COL[reg['col']]))
            rr = 5.5 * s + (1.5 if P.hot == rid else 0)
            qp.drawEllipse(ctr, rr, rr)
            dots.append((QRectF(ctr.x() - 7 * s, ctr.y() - 7 * s, 14 * s, 14 * s), rid))
            ddx += 17 * s
        qp.restore()
        # الضغط: النقط الأول، وبعدين الكارت كله
        if self.anim is None and self.fly is None:
            self.hits.extend(dots)
            self.hits.append((r, k))

    def _back(self, qp, X, Y, s):
        """دايرة الرجوع: اسم الخريطة اللي قبلها، متوصلة بالدايرة اللي في النص بخط متقطّع. دوسة ← ترجع."""
        P = self.page
        R = P.BACK_R * s
        c = QPointF(X(0), Y(P.BACK_Y))
        hot = P.hot_back and self.anim is None
        pen = QPen(qc(INK), max(2.0, 2.5 * s)); pen.setStyle(Qt.DashLine); pen.setCapStyle(Qt.FlatCap)
        qp.setOpacity(0.45); qp.setPen(pen)
        qp.drawLine(QPointF(c.x(), c.y() + R + 30 * s), QPointF(c.x(), Y(-P.CORE)))
        qp.setOpacity(1.0)
        qp.setPen(Qt.NoPen)
        for i in range(4, 0, -1):
            qp.setBrush(qc(T['shadow'], 9 * (3 if DARK else 1))); qp.drawEllipse(c + QPointF(0, 3), R + i * 1.8, R + i * 1.8)
        g = QRadialGradient(c - QPointF(R * .3, R * .4), R * 1.4)
        g.setColorAt(0, qc(T['back0_hot'] if hot else T['back0'])); g.setColorAt(1, qc(T['core1']))
        qp.setBrush(QBrush(g)); qp.setPen(QPen(qc(T['ring']), 3)); qp.drawEllipse(c, R, R)
        qp.setPen(qc('#FFFFFF')); qp.setFont(font(10.5 * s, True))
        qp.drawText(QRectF(c.x() - R, c.y() - R, 2 * R, 2 * R), Qt.AlignCenter, '\n'.join(P.BACK))
        bf = font(12 * s, True); w = QFontMetricsF(bf).horizontalAdvance('←  back') + 18 * s
        br = QRectF(c.x() - w / 2, c.y() + R + 6 * s, w, 22 * s)
        qp.setBrush(qc(T['pill_hot'] if hot else T['pill'])); qp.setPen(Qt.NoPen); qp.drawRoundedRect(br, 10 * s, 10 * s)
        qp.setPen(qc(T['pill_fg'])); qp.setFont(bf); qp.drawText(br, Qt.AlignCenter, '←  back')
        if self.anim is None and self.fly is None:
            self.hits.append((QRectF(c.x() - R - 4, c.y() - R - 4, 2 * R + 8, 2 * R + 40 * s), 'back'))

    def _core(self, qp, X, Y, s):
        P = self.page
        R = P.CORE * s
        c = QPointF(X(0), Y(0))
        qp.setPen(Qt.NoPen)
        for i in range(6, 0, -1):                        # ظل الدايرة
            qp.setBrush(qc(T['shadow'], 9 * (3 if DARK else 1))); qp.drawEllipse(c + QPointF(0, 4), R + i * 2.2, R + i * 2.2)
        g = QRadialGradient(c - QPointF(R * .3, R * .4), R * 1.4)
        g.setColorAt(0, qc(T['core0'])); g.setColorAt(1, qc(T['core1']))
        qp.setBrush(QBrush(g)); qp.drawEllipse(c, R, R)
        qp.setPen(qc('#FFFFFF')); qp.setFont(font(15 * s, True))
        qp.drawText(QRectF(c.x() - R, c.y() - R, 2 * R, 2 * R), Qt.AlignCenter, '\n'.join(P.CORE_TXT))
        if self.anim is None and self.fly is None:
            path = QPainterPath(); path.addEllipse(c, R, R)
            self.hits.append((path, 'core'))

    def _legend(self, qp, X, Y, s):
        P = self.page
        y = P.legend_y
        qp.setPen(qc(MUTED)); qp.setFont(font(14 * s))
        qp.drawText(QRectF(0, Y(y) - 12 * s, self.width(), 24 * s), Qt.AlignCenter,
                    'The coloured dots on a card are the registers that section shares.  Click a name to open its file.')
        nf = font(14.5 * s, True)
        fm = QFontMetricsF(nf)
        items = [(rid, r, fm.horizontalAdvance(r['name'])) for rid, r in P.regs.items()]
        gap, dot, pad = 10 * s, 10 * s, 12 * s
        widths = [pad + dot + 7 * s + w + pad for _, _, w in items]
        total = sum(widths) + gap * (len(items) - 1)
        x, cy = self.width() / 2 - total / 2, Y(y + 34)
        for (rid, r, w), bw in zip(items, widths):
            col = COL[r['col']]
            on = P.hot == rid
            rr = QRectF(x, cy - 14 * s, bw, 28 * s)
            qp.setPen(QPen(qc(col) if on else qc(SOFT_LINE), 1.5 if on else 1.0))
            qp.setBrush(qc(blend(col, .9) if on else T['chip']))
            qp.drawRoundedRect(rr, 14 * s, 14 * s)
            qp.setPen(Qt.NoPen); qp.setBrush(qc(col))
            qp.drawEllipse(QPointF(x + pad + dot / 2, cy), dot / 2, dot / 2)
            qp.setPen(qc(col if on else INK)); qp.setFont(nf)
            qp.drawText(QRectF(x + pad + dot + 7 * s, rr.top(), w + 4, rr.height()), Qt.AlignLeft | Qt.AlignVCenter,
                        r['name'])
            self.hits.append((rr, rid))
            x += bw + gap

    # ---------------------------------------------------------------- الماوس
    def obj_at(self, pos):
        pt = QPointF(pos)
        for shape, obj in self.hits:
            if shape.contains(pt): return obj
        return None

    def mouseMoveEvent(self, ev):
        self.page.on_motion(self.obj_at(ev.position()))

    def leaveEvent(self, ev):
        self.page.on_motion(None, left=True)

    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            self.page.on_click(self.obj_at(ev.position()), ev.globalPosition().toPoint())

    def mouseDoubleClickEvent(self, ev):
        pass                                             # الدبل كليك ميفتحش الحاجة مرتين


# ====================================================================== الصفحة (الخريطة + العمود اللي على اليمين)
class ClickBar(QFrame):
    def __init__(self, color, on_click=None):
        super().__init__()
        self.on_click = on_click
        self.setObjectName('bar')
        self.setStyleSheet(f'#bar {{ background: {color}; border-radius: 14px; }} QLabel {{ color: white; background: transparent; }}')
        if on_click: self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, ev):
        if self.on_click and ev.button() == Qt.LeftButton: self.on_click()


class HomeMapQt(QWidget):
    """الشاشة الرئيسية. open_section(key) بيفتح القسم (من main_qt)."""
    CORE, LANE = HM.HomeMap.CORE, HM.HomeMap.LANE
    CW, CH, MARGIN = HM.HomeMap.CW, HM.HomeMap.CH, HM.HomeMap.MARGIN
    POS = HM.HomeMap.POS
    SECTIONS, REGISTERS, INFO, ICONS = HM.SECTIONS, HM.REGISTERS, HM.INFO, HM.ICON
    CORE_TXT, TITLE, INBOX_CARD = HM.HomeMap.CORE_TXT, HM.HomeMap.TITLE, 'doc'
    DRILL = {'doc'}
    BACK, BACK_Y, BACK_R = None, 0, 34          # الخريطة اللي قبلها (Document Center بس) – دايرة صغيرة فوق

    def __init__(self, open_section, ready, root, name='', number='', inbox_count=None, log=None, intro=False):
        super().__init__()
        self.open_section, self.ready, self.root = open_section, set(ready or ()), root or ''
        self.name, self.number, self.log, self.intro = name, number, log, intro
        self.hot = self.shown = None
        self.hot_back = False
        self.next_short = ''
        try: self.inbox_n = inbox_count() if inbox_count else None
        except Exception: self.inbox_n = None
        self.expanded = self.start_open()
        self.geometry_()
        lay = QHBoxLayout(self); lay.setContentsMargins(22, 16, 22, 16); lay.setSpacing(16)
        self.cv = MapCanvas(self)
        lay.addWidget(self.cv, 1)
        self.side = QWidget(); self.side.setFixedWidth(420)
        lay.addWidget(self.side)
        self.build_side()
        self.details(None)
        self._clear = QTimer(self); self._clear.setSingleShot(True); self._clear.timeout.connect(self._maybe_clear)
        QTimer.singleShot(300, self._fill_next_safe)
        if self.intro and not self.expanded:
            QTimer.singleShot(60, self.cv, lambda: self.cv.animate(True))

    def start_open(self):
        return _OPENED and not self.intro

    def on_fail(self):
        """الرسم وقع ← البرنامج القديم بيفتح على نفس الشاشة (والخطأ في platform_error.log)."""
        self.open_section('__old__:' + ('doc' if isinstance(self, DocMapQt) else 'home'))

    def resizeEvent(self, ev):
        w = 420 if self.width() >= 1280 else 340
        if self.side.width() != w: self.side.setFixedWidth(w)
        super().resizeEvent(ev)

    # ---------------------------------------------------------------- نفس حسابات home_map
    def geometry_(self):
        self.sec = {}
        for i, (k, t, s, c) in enumerate(self.SECTIONS):
            x, y = self.POS[k]
            box = (x - self.CW / 2, y - self.CH / 2, x + self.CW / 2, y + self.CH / 2)
            self.sec[k] = dict(key=k, t=t, s=s, c=c, i=i, x=x, y=y, box=box, planned=(c == 'planned'))
        self.regs = {r['id']: dict(r) for r in self.REGISTERS}
        g, L = 6, self.LANE
        for k, d in self.sec.items():
            x, y, (x1, y1, x2, y2) = d['x'], d['y'], d['box']
            if x1 < 0 < x2:
                d['segs'] = [(0, 0, 0, (y2 + g) if y < 0 else (y1 - g))]
            elif abs(y) < self.CORE - 15:
                d['segs'] = [(0, y, (x1 - g) if x > 0 else (x2 + g), y)]
            else:
                lx = L if x > 0 else -L
                d['segs'] = [(lx, 0, lx, y), (lx, y, (x1 - g) if x > 0 else (x2 + g), y)]
        xs = [v for d in self.sec.values() for v in (d['box'][0], d['box'][2])]
        ys = [v for d in self.sec.values() for v in (d['box'][1], d['box'][3])]
        self.legend_y = max(ys) + 36
        m = self.MARGIN
        top = min(ys) - m
        if self.BACK:                                   # مكان لدايرة الرجوع فوق الكروت
            self.BACK_Y = min(ys) - 70
            top = self.BACK_Y - self.BACK_R - 12
        self.bbox = (min(xs) - m, top, max(xs) + m, self.legend_y + 34 + 14 + m)

    def badge(self, k):
        if k == self.INBOX_CARD and self.inbox_n: return (f'{self.inbox_n} waiting', T['wait_fg'], T['wait_bg'], '#E8A400')
        if k == 'qs' and self.next_short: return (self.next_short, COL['qs'], blend(COL['qs'], .88), None)
        if k not in self.ready: return ('Soon', COL['planned'], blend(COL['planned'], .88), None)
        return None

    # ---------------------------------------------------------------- العمود اللي على اليمين
    def head(self, v):
        t0 = QLabel('Technical Office Platform'); t0.setFont(font(13, True)); t0.setStyleSheet(f'color: {MUTED};')
        v.addWidget(t0)
        t1 = QLabel((self.name or 'Project').strip() or 'Project'); t1.setWordWrap(True)
        t1.setFont(font(22, True)); t1.setStyleSheet(f'color: {T["title"]};')
        v.addWidget(t1)
        ok = bool(self.root) and os.path.isdir(self.root)
        sub = (f'PO {self.number}   ·   ' if self.number else '') + \
              ('Project folder ✓' if ok else 'Project folder not found – edit PROJECT_ROOT in SETTINGS.json')
        t2 = QLabel(sub); t2.setWordWrap(True); t2.setFont(font(13))
        t2.setStyleSheet(f'color: {T["ok"] if ok else T["bad"]};')
        v.addWidget(t2)

    def inbox_text(self, n):
        return ('All good – nothing waiting' if n == 0 else f'{n} paper(s) waiting in INBOX') if n is not None else 'Ready'

    def build_side(self):
        v = QVBoxLayout(self.side); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(4)
        self.head(v)
        v.addSpacing(8)
        box = QFrame(); box.setObjectName('det')
        box.setStyleSheet(f'#det {{ background: {CARD}; border-radius: 16px; border: 1px solid {SOFT_LINE}; }}')
        shadow(box, 18, 2, 18)
        bl = QVBoxLayout(box); bl.setContentsMargins(20, 16, 20, 16)
        sc = QScrollArea(); sc.setWidgetResizable(True); sc.setFrameShape(QFrame.NoFrame)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sc.viewport().setObjectName('detv')
        sc.setStyleSheet(f'QScrollArea, #detv {{ background: {CARD}; }}')
        self.det = QLabel(); self.det.setWordWrap(True); self.det.setTextFormat(Qt.RichText)
        self.det.setAlignment(Qt.AlignLeft | Qt.AlignTop); self.det.setStyleSheet(f'background: {CARD};')
        sc.setWidget(self.det)
        bl.addWidget(sc)
        v.addWidget(box, 1)
        v.addSpacing(6)
        n = self.inbox_n
        bar = ClickBar(NAVY, (lambda: self.open_section('dc')) if self.bar_clickable(n) else None)
        hl = QHBoxLayout(bar); hl.setContentsMargins(18, 12, 18, 12)
        l1 = QLabel(self.inbox_text(n)); l1.setFont(font(15, True)); hl.addWidget(l1, 1)
        l1b = QLabel('✔' if not n else '●'); l1b.setFont(font(18, True))
        l1b.setStyleSheet(f'color: {"#7CD99B" if not n else "#FFC000"};'); hl.addWidget(l1b)
        v.addWidget(bar)
        self.b2 = ClickBar(NAVY2, self.b2_click())
        h2 = QHBoxLayout(self.b2); h2.setContentsMargins(18, 10, 18, 10)
        self.l2 = QLabel(''); self.l2.setTextFormat(Qt.PlainText); self.l2.setWordWrap(True); self.l2.setFont(font(13, True)); h2.addWidget(self.l2)
        v.addSpacing(4)
        v.addWidget(self.b2); self.b2.hide()

    def bar_clickable(self, n):
        return bool(n)

    def b2_click(self):
        return lambda: self.open_section('qs')

    def details(self, obj):
        def esc(t): return (t or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        h = []
        def title(t, c): h.append(f'<div style="font: bold 15pt \'{FAM}\'; color:{c};">{esc(t)}</div>')
        def tag(t, c=MUTED): h.append(f'<div style="font: bold 8pt \'{FAM}\'; color:{c}; letter-spacing:1px; '
                                      f'margin-top:4px; margin-bottom:8px;">{esc(t)}</div>')
        def para(t, c=INK): h.append(f'<div style="font: 10pt \'{FAM}\'; color:{c}; line-height:140%;">{esc(t)}</div>')
        def rows(head, items, c=INK):
            if not items: return
            tag(head.upper())
            h.append('<div style="margin-top:-6px;">' + ''.join(
                f'<div style="font: 10pt \'{FAM}\'; color:{c}; margin-left:6px; line-height:140%;">•&nbsp;&nbsp;{esc(t)}</div>'
                for t in items) + '</div>')
        if obj in self.sec:
            d, I = self.sec[obj], self.INFO.get(obj, {})
            planned = d['planned'] or obj not in self.ready
            title(f"{d['t']}  –  {d['s']}", COL[d['c']])
            tag('SOON  ·  NOT BUILT YET' if planned else
                ('SECTION  ·  CLICK TO OPEN ITS MAP' if obj in self.DRILL else 'SECTION  ·  CLICK TO OPEN'),
                COL['planned'] if planned else MUTED)
            para(I.get('text', ''))
            rows('Takes from', I.get('takes')); rows('Gives to', I.get('gives')); rows('Soon', I.get('soon'), COL['planned'])
        elif obj in self.regs:
            r, I = self.regs[obj], self.INFO.get(obj, {})
            title(r['name'], COL[r['col']])
            tag('SHARED REGISTER  ·  ' + ('CLICK TO CHOOSE WHAT TO OPEN' if len(r['opens']) > 1 else 'CLICK TO OPEN'))
            if I.get('path'):
                h.append(f'<div style="font: 9pt \'{MONO}\'; color:{MUTED}; margin-bottom:8px;">{esc(I["path"])}</div>')
            para(I.get('text', ''))
            rows('Used by', ['  ·  '.join(self.sec[m]['t'] for m in r['uses'] if m in self.sec)])
            rows('Opens', [o[0] for o in r['opens']])
        else:
            I = self.INFO['core']
            title(self.TITLE, INK)
            tag('HOW TO READ THE MAP')
            para(I['text'])
            hints = [('The circle on top: back to the project map' if h.startswith('← Home') else h) for h in I.get('hints') or []]
            rows('Use it', hints)
            tag('KEY')
            k1, k2 = COL['wlog'], COL['qs']
            h.append(f'<table cellspacing="0" cellpadding="3" style="font: 10pt \'{FAM}\'; color:{INK};">'
                     f'<tr><td><table border="1" cellspacing="0" cellpadding="0" style="border-color:{LINE}; border-style:solid;"><tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table></td>'
                     f'<td>&nbsp;A section – click to open it</td></tr>'
                     f'<tr><td><span style="color:{k1};">●</span><span style="color:{k2};">●</span></td>'
                     f'<td>&nbsp;The registers it shares</td></tr>'
                     f'<tr><td><span style="background:{blend(COL["planned"], .88)}; color:{COL["planned"]}; '
                     f'font-weight:bold; font-size:8pt;">&nbsp;Soon&nbsp;</span></td><td>&nbsp;Not built yet</td></tr></table>')
        self.det.setText(''.join(h))

    def _fill_next_safe(self):
        try: self.fill_next()
        except Exception:
            if self.log:
                try: self.log(traceback.format_exc())
                except Exception: pass

    def fill_next(self):
        """سطر التحديث الجاي (الشريط اللي تحت) + أقرب تاريخ (شارة كارت QS) – زي home_map بالظبط."""
        txt, short = '', ''
        try:
            p = os.path.join(HERE, 'modules', 'qs')
            if p not in sys.path: sys.path.insert(0, p)
            import upd_dates as U
            rows = U.load()
            txt = U.due_text(rows=rows)
            ds = [r['d1'] for r in (U.next_due(U.COMPANY, rows), U.next_due(U.OWNER, rows)) if r]
            if ds:
                d = min(ds)
                short = f'Next: {U.day_name(d)} {d:%d-%m}'
        except Exception:
            txt = ''
        if txt:
            self.l2.setText('Next update  ·  ' + txt); self.b2.show()
        if short and short != self.next_short:
            self.next_short = short; self.cv.update()

    # ---------------------------------------------------------------- الماوس
    def focus(self, obj):
        self.shown = obj
        self.hot = obj if (obj in self.sec or obj in self.regs) else None
        self.details(obj)
        self.cv.update()

    def on_motion(self, obj, left=False):
        if self.cv.anim or self.cv.fly: return
        hb = obj == 'back'
        if hb != self.hot_back: self.hot_back = hb; self.cv.update()
        if obj:
            self._clear.stop()
            if obj != self.shown: self.focus(obj)
            hand = obj in self.regs or obj in self.ready or obj in ('core', 'back') or obj in self.extra_clickable()
            self.cv.setCursor(Qt.PointingHandCursor if hand else Qt.ArrowCursor)
        else:
            self.cv.setCursor(Qt.ArrowCursor)
            if left: return                                  # الماوس خرج برّه الخريطة ← سيبها زي ما هي
            if self.shown is not None and not self._clear.isActive(): self._clear.start(300)

    def extra_clickable(self):
        return ()

    def _maybe_clear(self):
        pos = self.cv.mapFromGlobal(QCursor.pos())
        if self.cv.rect().contains(pos) and not self.cv.obj_at(pos): self.focus(None)

    def on_click(self, obj, gpos):
        if self.cv.anim or self.cv.fly: return
        if obj == 'core':
            self.toggle()
        elif obj == 'back':
            self.go_home()
        elif obj in self.sec:
            if obj in self.ready and obj in self.DRILL and self.expanded:
                self.fly_to(obj, lambda o=obj: self.open_section(o))
            elif obj in self.ready:
                self.open_section(obj)
        elif obj in self.regs:
            self.open_register(self.regs[obj]['name'], self.regs[obj]['opens'], gpos)
        elif obj is None:
            self.focus(None)

    # ---------------------------------------------------------------- الحركة
    def toggle(self):
        opening = not self.expanded
        self.focus(None)
        self.cv.animate(opening)

    def anim_done(self, opening):
        self.expanded = opening
        if opening and type(self) is HomeMapQt:
            global _OPENED
            _OPENED = True
        self.cv.update()

    def fly_to(self, key, done):
        """«الدايرة تنتقل»: الكروت ترجع ورا الدايرة، والكارت ده فوقها – وبعدين done() تفتح خريطته."""
        self.focus(None)
        self.cv.fly = key

        def end():
            self.cv.fly = None
            if self.isVisible(): done()
        self.cv.animate(False, end, ms=280)

    def go_home(self):
        """← Home: الكروت ترجع للدايرة، وبعدين الخريطة الكبيرة تطلع من دايرتها."""
        if self.cv.anim: return
        if self.expanded: self.cv.animate(False, lambda: self.isVisible() and self.open_section('home'))
        else: self.open_section('home')

    # ---------------------------------------------------------------- فتح ملفات السجلات
    def open_register(self, head, opens, gpos):
        if len(opens) == 1: return self.open_target(*opens[0])
        m = nice_menu(self)
        a = m.addAction(head.upper() + '  ·  OPEN'); a.setEnabled(False)
        for label, kind, val in opens:
            icon = '▤' if kind in ('file', 'dcsheet') else ('▸' if kind == 'section' else '▭')
            m.addAction(f'{icon}   {label}', lambda l=label, k=kind, v=val: self.open_target(l, k, v))
        m.exec(gpos)

    def dc_sheet(self):
        try:
            p = os.path.join(HERE, 'modules', 'dc')
            if p not in sys.path: sys.path.insert(0, p)
            import dc_sheet
            return dc_sheet.sheet_path()
        except Exception:
            return os.path.join(os.path.expanduser('~'), 'OneDrive', '01-N.W.C PROJECT', 'CO WATER', 'DC.xlsx')

    def open_target(self, label, kind, val):
        if kind == 'section': return self.open_section(val)
        if kind == 'dcsheet': path = self.dc_sheet()
        else: path = os.path.join(self.root, val.replace('\\', os.sep)) if val else self.root
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, 'Not found', f'Could not find {label}:\n\n{path or "(no project folder)"}')
            return
        try:
            path = os.path.normpath(path)
            if hasattr(os, 'startfile'): os.startfile(path)
            else: subprocess.Popen(['xdg-open', path])
        except Exception as e:
            QMessageBox.critical(self, 'Open', f'Could not open {label}:\n\n{path}\n\n{e}')


# ====================================================================== Document Center
class DocMapQt(HomeMapQt):
    """خريطة Document Center: نفس الرئيسية بكروت المركز. ← Home يرجّع للخريطة الكبيرة."""
    SECTIONS, REGISTERS, INFO, ICONS = HM.DOC_SECTIONS, HM.DOC_REGISTERS, HM.DOC_INFO, HM.DOC_ICON
    CORE_TXT, TITLE, INBOX_CARD = HM.DocMap.CORE_TXT, HM.DocMap.TITLE, 'dist'
    CW, POS = HM.DocMap.CW, HM.DocMap.POS
    DRILL = set()
    BACK = ('Project', 'Platform')               # دوسة على الدايرة اللي فوق ← الرئيسية

    def __init__(self, open_section, root, name='', number='', log=None, intro=False):
        self._claude_n = self._claude_ready = 0
        super().__init__(open_section, HM.DOC_READY, root, name, number, inbox_count=self.count, log=log, intro=intro)

    def start_open(self):
        return not self.intro

    def count(self):
        """الورق اللي مستني: INBOX + «مراجعة» + فولدر Claude (والـ Claude لوحده للكارت بتاعه)."""
        base = os.path.join(self.root, '00_INBOX')
        n = 0
        rv, cl = inbox_dirs(base)
        for sub in ('', rv, cl):
            d = os.path.join(base, sub)
            if not os.path.isdir(d): continue
            for f in os.listdir(d):
                p = os.path.join(d, f)
                if os.path.isfile(p) and f.lower().endswith(HM.DOC_EXT) and not f.startswith(('~', '.')):
                    n += 1
                    if sub == cl:
                        self._claude_n += 1
                        if os.path.isfile(p + '.claude.json'): self._claude_ready += 1
        return n

    def badge(self, k):
        if k == 'dist' and self.inbox_n: return (f'{self.inbox_n} waiting', T['wait_fg'], T['wait_bg'], '#E8A400')
        if k == 'claude' and self._claude_n:
            return (f'{self._claude_ready}/{self._claude_n} read', COL['arch'], blend(COL['arch'], .86), None)
        if k not in self.ready: return ('Soon', COL['planned'], blend(COL['planned'], .88), None)
        return None

    def head(self, v):
        t0 = QLabel('Technical Office Platform  ›  Document Center'); t0.setFont(font(13, True))
        t0.setStyleSheet(f'color: {MUTED};'); v.addWidget(t0)
        t1 = QLabel('Document Center'); t1.setFont(font(22, True)); t1.setStyleSheet(f'color: {T["title"]};'); v.addWidget(t1)
        t2 = QLabel('Papers in  ·  papers out  ·  registers'); t2.setFont(font(13)); t2.setStyleSheet(f'color: {MUTED};')
        v.addWidget(t2)

    def inbox_text(self, n):
        return ('Nothing waiting in INBOX' if n == 0 else f'{n} paper(s) waiting – Distribute') if n is not None else 'Ready'

    def bar_clickable(self, n):
        return True

    def b2_click(self):
        return None

    def fill_next(self):
        """«Waiting for you»: العقود الناقصة، والمستخلص اللي عليه الدور، وفولدر Claude."""
        lines = []
        mods = os.path.join(HERE, 'modules')
        for p in ('qs', 'sub', 'ipc'):
            q = os.path.join(mods, p)
            if q not in sys.path: sys.path.append(q)
        try:
            import sub_core
            for c in sub_core.contracts():
                if c['missing']: lines.append(f"{c['sub']} contract {c['no']} – missing {', '.join(c['missing'])}")
        except Exception:
            pass
        try:
            import ipc_core
            out = ipc_core.qs_core.project_path(ipc_core.OUT_REL)
            no = ipc_core.next_ipc(out_dir=out)
            _, res = ipc_core.upload_check(no, out)
            miss = [k for k, f in res if not f]
            lines.append(f'IPC-{no:02d} for upload – ' + (f"missing {', '.join(miss)}" if miss else 'all files there'))
        except Exception:
            pass
        if self._claude_n:
            lines.append(f'Claude folder – {self._claude_ready} of {self._claude_n} read, ready to distribute')
        if lines:
            self.l2.setText('Waiting for you\n' + '\n'.join('•  ' + t for t in lines)); self.b2.show()

    def extra_clickable(self):
        return ('dist', 'regs', 'claude')

    def anim_done(self, opening):
        self.expanded = opening
        self.cv.update()

    def on_click(self, obj, gpos):
        if self.cv.anim or self.cv.fly: return
        if obj == 'dist': return self.open_section('dc')
        if obj == 'regs':
            opens = [o for r in self.REGISTERS if r['id'] != 'inbox' for o in r['opens']]
            opens.append(('the 00_Registers folder', 'folder', '00_Registers'))
            return self.open_register('Registers', opens, gpos)
        if obj == 'claude':
            return self.open_target('the Claude folder', 'folder', os.path.join('00_INBOX', inbox_dirs(os.path.join(self.root, '00_INBOX'))[1]))
        return super().on_click(obj, gpos)
