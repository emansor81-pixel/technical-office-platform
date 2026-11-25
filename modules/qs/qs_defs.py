# -*- coding: utf-8 -*-
"""QS – الثوابت والدوال المشتركة (أعمدة الجداول والصيغ).

اتشالت من qs_app.py (الشاشة القديمة) عشان qs_qt.py ما يبقاش محتاجها. 21-11-2026
مفيش tkinter هنا خالص.
"""
import os, json

HERE = os.path.dirname(os.path.abspath(__file__))

UI_FILE = os.path.join(HERE, 'qs_ui.json')

INPUT_COLS = ('d0', 'd1', 'soil', 'asph')

def load_ui():
    try:
        with open(UI_FILE, encoding='utf-8') as f: return json.load(f)
    except Exception: return {}

def save_ui(d):
    try:
        with open(UI_FILE, 'w', encoding='utf-8') as f: json.dump(d, f)
    except Exception: pass

EXC_COLS = [('n', '#', 40), ('dist', 'District', 90), ('street', 'Street', 120), ('line', 'Line', 45),
            ('from', 'From', 70), ('to', 'To', 70), ('len', 'Length (m)', 75), ('d0', '✎ Depth Start', 95),
            ('d1', '✎ Depth End', 95), ('soil', '✎ Soil ▾', 70), ('asph', '✎ Asphalt (cm)', 95), ('avg', 'Avg Depth', 72), ('w', 'Width', 55), ('b1', '0–1.5', 62),
            ('b2', '1.5–2.5', 62), ('b3', '2.5–3.5', 62), ('b4', '>3.5', 55), ('tot', 'Total (m³)', 80),
            ('app', 'Approval Date', 92), ('st', 'Status', 105), ('wir', 'WIR No.', 65)]

BEDP_COLS = [('n', '#', 40), ('dist', 'District', 90), ('street', 'Street', 110), ('line', 'Line', 45), ('dia', 'Dia', 50),
             ('from', 'From', 70), ('to', 'To', 70), ('len', 'Length (m)', 75), ('soil', 'Soil', 55), ('w', 'Width', 55),
             ('t', 'Below Thk', 70), ('below', 'Below Qty', 78), ('bok', 'BED-D-P', 65), ('h', 'Around H', 68),
             ('net', 'Around Net', 80), ('cok', 'COV-P', 60), ('tot', 'Total (m³)', 80), ('done', 'Done (m³)', 80),
             ('pct', '% Done', 62), ('val', 'Done Value (SAR)', 115), ('inv', 'Invoice (m³)', 90), ('st', 'Status', 95),
             ('wir', 'WIR No.', 75)]

PIP_COLS = [('n', '#', 40), ('dist', 'District', 90), ('street', 'Street', 120), ('line', 'Line', 45),
            ('dia', 'Dia', 50), ('from', 'From', 70), ('to', 'To', 70), ('len', 'Length (m)', 75), ('test', 'Test WIR', 75),
            ('flush', 'Flushing WIR', 90), ('pct', '%', 55), ('paid', 'Paid Length (m)', 105), ('app', 'Approval Date', 92),
            ('st', 'Status', 105), ('wir', 'WIR No.', 65)]

BF_COLS = [('n', '#', 40), ('dist', 'District', 90), ('street', 'Street', 120), ('line', 'Line', 45), ('dia', 'Dia', 50),
           ('from', 'From', 70), ('to', 'To', 70), ('len', 'Length (m)', 75), ('avg', 'Avg Depth', 72), ('asph', 'Asphalt (cm)', 85),
           ('soil', 'Soil', 55), ('t', 'Bedding', 62), ('ch', 'Cover H', 62), ('Hn', 'Layers H', 72), ('H', 'Qty H (the Consultant)', 90), ('nreq', 'Layers Req.', 78),
           ('got', 'Layers Approved', 105), ('pct', '%', 55), ('w', 'Width', 55), ('qty', 'Backfill Qty (m³)', 110),
           ('paid', 'Qty by Layers (m³)', 120), ('st', 'Status', 95), ('wir', 'BF WIR No.', 90)]

RB_COLS = [('n', '#', 40), ('dist', 'District', 90), ('street', 'Street', 120), ('line', 'Line', 45), ('dia', 'Dia', 50),
           ('from', 'From', 70), ('to', 'To', 70), ('len', 'Length (m)', 75), ('w', 'Width', 60), ('th', 'Layer (m)', 70),
           ('l1', 'RB-1', 55), ('q1', 'RB-1 Qty (m³)', 100), ('l2', 'RB-2', 55), ('q2', 'RB-2 Qty (m³)', 100),
           ('qty', 'Road Base Qty (m³)', 125), ('st', 'Status', 95), ('wir', 'RB WIR No.', 90)]

ASP_COLS = [('n', '#', 40), ('dist', 'District', 90), ('street', 'Street', 120), ('line', 'Line', 45), ('dia', 'Dia', 50),
            ('from', 'From', 70), ('to', 'To', 70), ('len', 'Length (m)', 75), ('w', 'Width', 55), ('as', 'Asphalt W', 72),
            ('at', 'Thk (cm)', 62), ('w1', 'ASP-1 WIR', 75), ('q1', '32.4.1 Qty (m³)', 105), ('av', 'Lane Rest', 70),
            ('w2', 'Milling WIR', 80), ('q2', '32.4.2 Qty (m³)', 105), ('w3', 'Wearing WIR', 85), ('q3', 'Wearing (m³)', 90),
            ('mc', 'MC WIR', 60), ('rc', 'RC WIR', 60), ('st', 'Status', 95)]

STATUSES = ['All', 'Approved', 'Requested', 'Not Approved', 'After Cut-off', 'Missing Levels', 'No Width', 'Ahead of EXC', 'No Dia', 'Complete', 'Partial', 'No Layers', 'No Backfill', 'Closed', 'Below Done', 'Not Started']

SUM_COLS = [('A', 'م', 35), ('B', 'Zone', 80), ('C', 'رقم الخط', 60), ('D', 'رقم الرخصة', 75), ('E', 'من', 60),
            ('F', 'إلى', 60), ('G', 'نوع الغرفة', 70), ('H', 'طول الحفر المخصوم', 90), ('I', 'طول التمديد المخصوم', 95),
            ('J', 'طول الحفر بعد خصم الغرف', 120), ('K', 'طول التمديد بعد خصم', 105), ('L', 'عرض الحفر', 65),
            ('M', 'نوع التربة', 65), ('N', 'عمق من', 60), ('O', 'عمق إلى', 60), ('P', 'متوسط عمق الحفر', 90),
            ('Q', 'قطر الأنبوب', 70), ('R', 'عمق حتى 1.50', 75), ('S', 'عمق 1.51-2.50', 80), ('T', 'عمق 2.51-3.50', 80),
            ('U', 'عمق أكثر من 3.50', 90), ('V', 'كمية حتى 1.50', 80), ('W', 'كمية 1.51-2.50', 85), ('X', 'كمية 2.51-3.50', 85),
            ('Y', 'كمية أكثر من 3.50', 95)]

def f2(v, nd=3):
    if v is None: return '–'
    from decimal import Decimal, ROUND_HALF_UP
    d = Decimal(repr(round(v, 6))).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP)
    return f'{d:,}'

def fd(d):
    return '–' if d is None else d.strftime('%d-%m-%Y')

SUM_FREEZE = 6        # عدد أعمدة Summary الثابتة (م · Zone · رقم الخط · رقم الرخصة · من · إلى)

# ---- كانت جوه class App في الملف القديم

def tag_of(st):
    if st == 'Partial': return 'part'
    return 'ok' if st in ('Approved', 'Complete') else ('miss' if st in ('Missing Levels', 'No Width', 'Ahead of EXC', 'No Dia') else 'grey')

WK_COLS = [('item', 'Item No.', 75), ('code', 'PO Code', 75), ('desc', 'Description', 330), ('unit', 'Unit', 45),
           ('rate', 'Rate (SAR)', 80), ('cq', 'Total Qty (to date)', 120),
           ('rq', 'Reported to Company (to date)', 190),          # من تقارير الإنتاج الأسبوعية (▲ = أكتر من المحسوب)
           ('tq', 'To Report to CTR', 130),                               # المحسوب − المرفوع (السالب = صفر) – 15-11
           ('cv', 'Total Value to date (SAR)', 170),
           ('pq', 'Previous Qty', 100), ('pv', 'Previous Value (SAR)', 140), ('dq', 'Qty This Period', 110),
           ('dv', 'Value This Period (SAR)', 160)]
