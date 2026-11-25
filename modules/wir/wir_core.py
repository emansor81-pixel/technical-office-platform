# -*- coding: utf-8 -*-
"""Quantity of every WIR (for entering it in Unifier) – uses the same calculations as the QS module."""
import os, sys, math, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__))
QS = os.path.join(os.path.dirname(HERE), 'qs')
if QS not in sys.path: sys.path.insert(0, QS)
import qs_core as core

EPS = 1e-6
LEN_CODES = ('A-T', 'A-P', 'TOPO')
EXC_ITEMS = ['31.2.3.1', '31.2.3.2', '31.2.3.3', '31.2.3.4']
PIP_ITEMS = {300: '33.2.3.5', 400: '33.2.3.6', 600: '33.2.3.8'}


def _num(v):
    try: return float(v)
    except (TypeError, ValueError): return None


def wir_quantities(p):
    """one dict per WIR in the log (all of them, approved or not)"""
    S = p.S
    pre = str(S.get('BF_PREFIX', 'BF-')).strip()
    lay = _num(S.get('BF_LAYER')) or 0.30
    road = _num(S.get('ROAD_BASE_LAYERS'))
    if road is None: road = 0.15          # real backfill height: asphalt + one road-base layer
    cover = _num(S.get('COVER_ABOVE')) or 0.30
    rb = _num(S.get('RB_LAYER')) or 0.15
    L0 = p.seg_len
    out = []
    for q in p.reqs_all:
        code = q['type']
        r = dict(q=q, no=q['no'], unf=q.get('unf') or '', code=code, desc=q['desc'], title=q.get('title') or '', seq=q.get('seq'), a_raw=q.get('a_raw'), b_raw=q.get('b_raw'), L_raw=q.get('L_raw'), rev=q.get('rev'), dist=q['dist'], street=q['street'],
                 line=q['line'], dia=q['dia_raw'], reqdate=q['reqdate'], appdate=q['appdate'],
                 x0=None, x1=None, L=None, qty=None, unit='', item='', bands=[None] * 4, note='')
        r['status'] = 'Approved' if q['appdate'] else 'Pending'
        if q['a'] is None or q['b'] is None:
            r['note'] = 'From / To missing'; out.append(r); continue
        a, b = sorted((q['a'], q['b']))
        r['x0'], r['x1'], r['L'] = a, b, b - a
        if b - a < EPS:
            r['note'] = 'From = To'; out.append(r); continue
        q2 = dict(q, approved=True, after_cutoff=False)
        cuts = [a] + [k * L0 for k in range(int(a // L0) + 1, int(b // L0) + 1) if a + EPS < k * L0 < b - EPS] + [b]
        segs = list(zip(cuts[:-1], cuts[1:]))
        dia = int(q['dia']) if q['dia'] else None
        miss = False
        if code in LEN_CODES:
            r['qty'], r['unit'] = r['L'], 'L.M'
        elif code == str(S.get('EXC_CODE', 'EXC')).strip():
            tot, bd = 0.0, [0.0] * 4
            for x0, x1 in segs:
                e = p.calc_exc(q2, x0, x1)
                if e['total'] is None: miss = True; continue
                tot += e['total']; bd = [u + v for u, v in zip(bd, e['v'])]
            r['qty'], r['unit'], r['bands'] = tot, 'm3', bd
            r['item'] = ' / '.join(it for it, v in zip(EXC_ITEMS, bd) if v > EPS)
        elif code in ('BED-D-P', str(S.get('BED_CODE', 'COV-P')).strip()):
            below = code == 'BED-D-P'
            tot = 0.0
            for x0, x1 in segs:
                bb = p.calc_bed(q2, x0, x1)
                v = bb['below'] if below else bb['net']
                if v is None: miss = True; continue
                tot += v
            r['qty'], r['unit'], r['item'] = tot, 'm3', '31.2.4.3'
        elif code.startswith(str(S.get('PIP_CODE', 'PIP')).strip()):
            r['qty'], r['unit'], r['item'] = r['L'], 'L.M', PIP_ITEMS.get(dia, '')
        elif code.startswith(pre) and code[len(pre):].strip().isdigit():
            n = int(code[len(pre):].strip())
            od = p.od(dia)
            tot = 0.0
            for x0, x1 in segs:
                e = p.calc_exc(q2, x0, x1)
                if e['avg'] is None or not e['w'] or not od: miss = True; continue
                t = _num(S.get('BED_ROCK' if e['soil'] == core.ROCK else 'BED_OTHER')) or 0.0
                H = e['avg'] - road - (e['asph'] or 0) / 100 - t - (od + cover)
                h = min(max(H - lay * (n - 1), 0.0), lay)
                if h <= EPS: r['note'] = f'layer {n} above backfill height in part of the request'
                tot += e['L'] * e['w'] * h
            r['qty'], r['unit'], r['item'] = tot, 'm3', '31.2.4.4'
        elif code in ('RB-1', 'RB-2'):
            w = p.width(dia)
            if w: r['qty'] = r['L'] * w * rb
            else: miss = True
            r['unit'], r['item'] = 'm3', '31.2.4.5'
        else:
            r['qty'], r['unit'] = r['L'], 'L.M'
            r['note'] = 'no quantity rule for this code – length shown'
        if miss:
            r['note'] = 'missing levels / width in part of the request – quantity incomplete'
            r['miss'] = True
        out.append(r)
    return out


def load(cutoff=None):
    p = core.Project(core.DEFAULT_INPUT)
    p.load_requests()
    p.mode = 'requested'
    p.cutoff = cutoff or dt.date.today()
    return p
