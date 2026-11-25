# -*- coding: utf-8 -*-
"""يبني تمبلت المستخلص الثابت (11.1-TEMPLATE IPC.xlsx) بنفس تنسيق مستخلص المقاول (CAVAR + VO)
من جدول كميات أمر الشراء. التنسيق بيتاخد صف بصف من ملف المرجع REF_FORMAT.xlsx اللي جنب الملف ده.
"""
import os, re, copy, shutil, zipfile, tempfile
import openpyxl
from openpyxl.worksheet.pagebreak import Break, RowBreak

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(HERE, 'REF_FORMAT.xlsx')
COVER = 'CAVAR'
IPC = 'VO'
HELP = 'طريقة الشغل'
QSUM = 'QS SUMMARY'
# أسماء الموقّعين: بتتكتب مرة واحدة في الغلاف (CAVAR) عمود L، وكل خانات التوقيع بتاخد منها
SIGNERS = [('L3', 'المقاول – مدير المشروع'), ('L4', 'المقاول – حاسب الكميات'), ('L5', 'المقاول – مهندس المكتب الفني'),
           ('L6', 'الاستشاري – مهندس المشروع'), ('L7', 'الاستشاري – مهندس مقيم المشروع'), ('L8', 'الاستشاري – حاسب الكميات'),
           ('L9', 'الاستشاري – مهندس الموقع'), ('L10', 'المالك – مدير المشروع'), ('L11', 'المالك – مدير البرنامج')]


def NAME(cell):
    return f'="الاسم : "&{COVER}!${cell[0]}${cell[1:]}'
COLS = 'ABCDEFGHIJKLM'
CODE_COL = 'N'          # كود البند (برّه منطقة الطباعة) – البرنامج بيطابق بيه
VAT = 0.15

DIVISIONS = [(26, 'الأعمال الكهربائية'), (31, 'الأعمال الترابية'), (32, 'أعمال ترميم الأسطح'),
             (33, 'أعمال المرافق'), (40, 'أعمال نظام التحكم')]
PAGE_BUDGET = 540      # ارتفاع محتوى صفحة البنود (نقطة) زي المرجع

# صفوف النماذج في شيت VO بتاع المرجع
P = dict(info=8, sumhdr=13, sumrow=14, sumtot=21, sg1=22, sg2=23, sg3=24, gap1=25, gap2=26, gap3=27,
         dh1=28, dh2=29, drow=30, dtot=33, dpct=34, dsg1=35, dsg2=36, dsg3=37, dgap1=38, dgap2=39, dgap3=40,
         ih1=41, ih2=42, idiv=43, ihead=45, item=46, ptot=55, psg1=56, psg2=57, psg3=58, pgap1=59, pgap2=60, pgap3=61,
         mq=101)


# ------------------------------------------------------------------ جدول الكميات
def read_boq(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    divs = [{'no': n, 'name': nm, 'rows': []} for n, nm in DIVISIONS]
    k = 0
    for r in range(4, ws.max_row + 1):
        item, code, en, ar, unit, unit_ar, qty, rate = (ws[f'{c}{r}'].value for c in 'BCDEFGHI')
        txt = ' '.join(str(v) for v in (item, en, ar) if v is not None)
        if isinstance(item, str) and item.strip() in ('ما قبله', 'Item'): continue
        if 'ينقل إلى صفحة الخلاصة' in txt or 'ينقل الى صفحة الخلاصة' in txt:
            k = min(k + 1, len(divs) - 1); continue
        if 'ينقل إلى الصفحة التالية' in txt or 'ينقل الى الصفحة التالية' in txt: continue
        if en and str(en).strip().upper().startswith('DIVISION'):
            m = re.search(r'DIVISION\s+(\d+)', str(en))
            if m:
                for i, d in enumerate(divs):
                    if d['no'] == int(m.group(1)): k = i
            continue
        if item is None and en is None and ar is None: continue
        item = str(item).replace('\xa0', ' ').strip() if item is not None else ''
        desc = re.sub(r'\s+', ' ', str(ar or en or '').replace('\xa0', ' ')).strip()
        if isinstance(qty, (int, float)) and isinstance(rate, (int, float)) and code:
            u = str(unit or '').strip(); ua = str(unit_ar or '').strip()
            divs[k]['rows'].append({'item': item, 'code': int(code), 'desc': desc, 'unit': u,
                                    'unit_txt': f'{u}/{ua}' if ua else u, 'qty': qty, 'rate': rate})
        elif desc:
            divs[k]['rows'].append({'item': item, 'head': True, 'desc': desc})
    return divs


# ------------------------------------------------------------------ نسخ تنسيق صف من المرجع
class Cloner:
    def __init__(self, src, dst):
        self.s, self.d = src, dst
        self.m1 = {}   # صف → دمج صف واحد
        for m in src.merged_cells.ranges:
            self.m1.setdefault((m.min_row, m.max_row), []).append(m)

    def row(self, sr, dr, values=None, height=None):
        for c in COLS:
            a, b = self.s[f'{c}{sr}'], self.d[f'{c}{dr}']
            if a.has_style:
                b.font = copy.copy(a.font); b.fill = copy.copy(a.fill); b.border = copy.copy(a.border)
                b.alignment = copy.copy(a.alignment); b.number_format = a.number_format
        for m in self.m1.get((sr, sr), []):
            if m.min_col <= 13:
                self.d.merge_cells(start_row=dr, start_column=m.min_col, end_row=dr, end_column=min(m.max_col, 13))
        self.d.row_dimensions[dr].height = height if height is not None else self.s.row_dimensions[sr].height
        for k, v in (values or {}).items():
            self.d[f'{k}{dr}'] = v
        return dr

    def block(self, sr0, sr1, dr0):
        for i in range(sr1 - sr0 + 1): self.row(sr0 + i, dr0 + i)
        for m in self.s.merged_cells.ranges:
            if m.min_row >= sr0 and m.max_row <= sr1 and m.min_row != m.max_row and m.min_col <= 13:
                self.d.merge_cells(start_row=m.min_row - sr0 + dr0, start_column=m.min_col,
                                   end_row=m.max_row - sr0 + dr0, end_column=min(m.max_col, 13))
        return dr0 + (sr1 - sr0)


def head_height(desc):
    n = len(desc or '')
    return 40.9 if n <= 70 else 67.5 if n <= 150 else 115.15 if n <= 300 else 147.6 if n <= 420 else 190


# ------------------------------------------------------------------ بناء التمبلت
def build_template(boq_path, out, ref=REF):
    divs = read_boq(boq_path)
    src_wb = openpyxl.load_workbook(ref)
    wb = openpyxl.load_workbook(ref)
    for n in list(wb.sheetnames):
        if n not in (COVER, IPC): del wb[n]
    build_cover(wb[COVER])
    S, ws = src_wb[IPC], wb[IPC]
    # نفضّي VO من صف 8 لتحت ونشيل أعمدة المساعدة
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= 8 or m.min_col >= 14: ws.unmerge_cells(str(m))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            if c.row >= 8 or c.column >= 14:
                c.value = None; c.style = 'Normal'; c.hyperlink = None
    for r in range(8, ws.max_row + 1): ws.row_dimensions[r].height = None
    ws.row_breaks = RowBreak()
    for a in ('D2', 'I2', 'M2', 'M1'): ws[a].value = None
    ws['A2'] = 'الجهة المالكة'; ws['A3'] = 'وحدة أعمال المدينة'; ws['A4'] = 'الإدارة العامة للمشاريع بالقطاع الأوسط'
    ws['E2'] = f'={COVER}!B9'; ws['K2'] = f'={COVER}!B11'
    ws['B6'] = 'المستخلص الجاري رقم'; ws['D6'] = f'={COVER}!F6'; ws['E6'] = 'الفترة من '
    ws['F6'] = f'=IF({COVER}!E7="","",{COVER}!E7)'; ws['H6'] = 'الى'; ws['I6'] = f'=IF({COVER}!E8="","",{COVER}!E8)'
    C = Cloner(S, ws)
    info = {}
    r = 8
    # ---------------- صفحة 1: المبالغ
    for lab, val in (('    الاستشارى :', '=E2'), ('    المقــاول   :', '=K2'), ('    المشــروع :', f'={COVER}!B5'),
                     ('    تاريخ بدء المشروع :', None), ('    التاريخ المحدد لإنهاء المشروع :', None)):
        C.row(P['info'] if lab.strip() != 'المشــروع :' else 10, r, {'A': lab, 'E': val}); r += 1
    info['start_row'], info['end_row'] = r - 2, r - 1
    C.row(P['sumhdr'], r, {'A': 'الوصف', 'E': 'المبلغ', 'G': ' القيمة بالاحــــرف'}); r += 1
    sum_lines = ['اجمالي قيمة المشروع', 'قيمة الأعمال المنفذة من العقد حتى تاريخه', 'قيمة الأعمال المنفذة سابقا',
                 'قيمة الاعمال خلال الفترة', 'ضريبة القيمة المضافة 15 %',
                 'إجمالي المبلغ المستحق للمقاول في هذه الدفعة بعد اضافة الضريبة 15 %']
    sr = {}
    for i, lab in enumerate(sum_lines):
        C.row(P['sumtot'] if i == len(sum_lines) - 1 else P['sumrow'], r, {'A': lab}); sr[i] = r; r += 1
    info['sum_rows'] = sr
    r = sign_page(C, r, P['sg1'], P['sg2'], P['sg3'], (P['gap1'], P['gap2'], P['gap3']))
    ws.row_breaks.append(Break(id=r - 2))
    # ---------------- صفحة 2: الأقسام
    dh = r
    C.block(P['dh1'], P['dh2'], r)
    ws[f'A{r}'] = 'البند'; ws[f'B{r}'] = 'الوصف'; ws[f'F{r}'] = 'القيمة بالعقد '; ws[f'I{r}'] = 'المنفذ بالريال السعودى'
    ws[f'F{r + 1}'] = '( ريال سعودى )'; ws[f'I{r + 1}'] = 'المنفذ السابق'; ws[f'K{r + 1}'] = 'خلال الفترة'; ws[f'M{r + 1}'] = 'الإجمالى'
    r += 2
    div_rows = {}
    for no, name in DIVISIONS:
        C.row(P['drow'], r, {'A': f'رقم ({no})', 'B': f'إجمالي {name}'}); div_rows[no] = r; r += 1
    dt_ = r
    C.row(P['dtot'], r, {'B': 'إجمالى قيمة الأعمال'})
    d0, d1 = div_rows[DIVISIONS[0][0]], div_rows[DIVISIONS[-1][0]]
    ws[f'F{r}'] = f'=SUM(F{d0}:H{d1})'; ws[f'I{r}'] = f'=SUM(I{d0}:J{d1})'; ws[f'K{r}'] = f'=M{r}-I{r}'
    ws[f'M{r}'] = f'=ROUND(SUM(M{d0}:M{d1}),2)'
    r += 1
    C.row(P['dpct'], r, {'B': 'النسبة المئـــــــوية %', 'F': f'=F{dt_}/F{dt_}', 'I': f'=ROUND(I{dt_}/F{dt_},4)',
                         'K': f'=M{r}-I{r}', 'M': f'=ROUND(M{dt_}/F{dt_},4)'})
    r += 1
    r = sign_page(C, r, P['dsg1'], P['dsg2'], P['dsg3'], (P['dgap1'], P['dgap2'], P['dgap3']), head=dh)
    ws.row_breaks.append(Break(id=r - 2))
    info['div_total'] = dt_
    # ربط صفحة المبالغ
    ws[f'E{sr[0]}'] = f'=F{dt_}'; ws[f'E{sr[1]}'] = f'=M{dt_}'; ws[f'E{sr[2]}'] = f'=I{dt_}'
    ws[f'E{sr[3]}'] = f'=K{dt_}'; ws[f'E{sr[4]}'] = f'=ROUND(E{sr[3]}*{VAT},2)'; ws[f'E{sr[5]}'] = f'=ROUND(E{sr[3]}+E{sr[4]},2)'
    # ---------------- صفحات البنود
    page = 3
    item_rows = []
    for d in divs:
        rows = d['rows']; ranges = []
        i = 0; carry = None; first_page = True
        while i < len(rows) or first_page:
            C.block(P['ih1'], P['ih2'], r)
            ws[f'A{r}'] = 'رقم البند'; ws[f'B{r}'] = 'وصف البند'; ws[f'F{r}'] = 'الوحدة'; ws[f'G{r}'] = 'الكمية التقديرية'
            ws[f'H{r}'] = 'الكمية بالجدول المعدل '; ws[f'I{r}'] = 'الكميات المنفذة'; ws[f'L{r}'] = 'سعر الوحدة'; ws[f'M{r}'] = 'القيمة الإجمالية'
            ws[f'I{r + 1}'] = 'السابقة'; ws[f'J{r + 1}'] = 'خلال الفترة'; ws[f'K{r + 1}'] = 'حتى تاريخه'
            r += 2; top = r; used = 0
            if carry:
                C.row(P['mq'], r, {'B': 'ماقبله', 'M': f'=M{carry}'}); r += 1; used += 29
            if first_page:
                C.row(P['idiv'], r, {'A': str(d['no']), 'B': f"{d['no']} {d['name']}"}); r += 1; used += 41
                first_page = False
            a0 = r
            while i < len(rows):
                it = rows[i]
                h = head_height(it['desc']) if it.get('head') else max(40.9, head_height(it['desc']) if len(it['desc']) > 70 else 40.9)
                if used + h > PAGE_BUDGET and r > a0: break
                if it.get('head'):
                    C.row(P['ihead'], r, {'A': it['item'], 'B': it['desc']}, height=h)
                else:
                    C.row(P['item'], r, {'A': it['item'], 'B': it['desc'], 'F': it['unit_txt'], 'G': it['qty'], 'H': f'=G{r}',
                                         'J': f'=K{r}-I{r}', 'L': it['rate'], 'M': f'=ROUND(L{r}*K{r},2)'}, height=h)
                    ws[f'{CODE_COL}{r}'] = it['code']
                    item_rows.append((r, it['code'], d['no']))
                used += h; r += 1; i += 1
            ranges.append((a0, r - 1))
            last = i >= len(rows)
            txt = (f"إجمالي {d['name']} ينقل لصفحة رقم ( 2 )" if last else 'المجموع ( يرحل للصفحة التالية)')
            C.row(P['ptot'], r, {'A': txt, 'M': f'=SUM(M{top}:M{r - 1})'})
            ptot = r; r += 1
            r = item_sign(C, r)
            ws.row_breaks.append(Break(id=r - 2)); page += 1
            carry = ptot
        # ربط القسم بصفحة الأقسام
        dr = div_rows[d['no']]
        ws[f'F{dr}'] = '=' + '+'.join(f'SUMPRODUCT(H{a}:H{b},L{a}:L{b})' for a, b in ranges)
        ws[f'I{dr}'] = '=' + '+'.join(f'SUMPRODUCT(I{a}:I{b},L{a}:L{b})' for a, b in ranges)
        ws[f'K{dr}'] = f'=M{dr}-I{dr}'
        ws[f'M{dr}'] = f'=M{ptot}'
    ws.print_area = f'A8:M{r - 1}'
    ws.column_dimensions[CODE_COL].hidden = True
    info['item_rows'] = item_rows
    # ---------------- 90% (برّه الطباعة)
    ws['O2'] = 'نسبة الانجاز من العقد'; ws['O3'] = f'=M{dt_ + 1}'; ws['O3'].number_format = '0.00%'
    ws['O4'] = f'=IF(O3>0.9,"تنبيه: المستخلصات الجارية عدّت 90% من قيمة العقد","")'
    build_qs_summary(wb, divs)
    build_help(wb)
    # تنضيف بواقي ملف المرجع: روابط خارجية، أسماء بايظة (#REF!)، لينكات
    wb._external_links = []
    for n in list(wb.defined_names.keys()):
        try:
            if '#REF!' in str(wb.defined_names[n].attr_text): del wb.defined_names[n]
        except Exception: pass
    for sh in wb.worksheets:
        sh._hyperlinks = []
    wb.active = 0
    tmp = out + '.tmp.xlsx'
    wb.save(tmp)
    add_cover_logos(tmp, out, ref)
    os.remove(tmp)
    return out, divs, info


OK_TXT = 'الأعمال منفذة طبقاً لشروط ومواصفات العقد'


def sign_page(C, r, s1, s2, s3, gaps, head=None):
    """توقيعات صفحة المبالغ وصفحة الأقسام (زي نموذج المستخلص): المقاول | الاستشاري (اتنين) | المالك (اتنين)"""
    C.row(s1, r, {'A': f'={COVER}!B11', 'C': '=$E$2', 'H': '="المالك : "&$A$2&" / جهاز الإشراف على المشروع"'})
    C.row(s2, r + 1, {'A': OK_TXT, 'C': OK_TXT, 'H': OK_TXT})
    C.d[f'H{r + 1}'].font = copy.copy(C.d[f'C{r + 1}'].font); C.d[f'H{r + 1}'].alignment = copy.copy(C.d[f'C{r + 1}'].alignment)
    C.row(s2, r + 2, {'A': 'الوظيفة : مدير المشروع', 'C': 'الوظيفة : مهندس المشروع', 'E': 'الوظيفة : مهندس مقيم المشروع',
                      'J': 'الوظيفة : مدير المشروع', 'L': 'الوظيفة : مدير البرنامج'})
    C.row(s3, r + 3, {'A': NAME('L3'), 'C': NAME('L6'), 'E': NAME('L7'), 'J': NAME('L10'), 'L': NAME('L11')})
    C.row(s3, r + 4, {'A': 'التوقيع :', 'C': 'التوقيع :', 'E': 'التوقيع :', 'J': 'التوقيع :', 'L': 'التوقيع :'})
    for i, g in enumerate(gaps): C.row(g, r + 5 + i, height=45 if i == 0 else None)
    return r + 5 + len(gaps)


def item_sign(C, r):
    """توقيعات صفحات البنود (زي نموذج المستخلص): المقاول (مدير المشروع + حاسب كميات) | الاستشاري (مهندس المشروع + حاسب الكميات)"""
    C.row(P['psg1'], r, {'A': '=$K$2', 'F': '=$E$2'})
    C.row(P['psg2'], r + 1, {'A': OK_TXT, 'F': OK_TXT})
    C.row(P['psg2'], r + 2, {'A': 'الوظيفة : مدير المشروع', 'C': 'الوظيفة : حاسب كميات', 'F': 'الوظيفة : مهندس المشروع',
                             'J': 'الوظيفة : حاسب الكميات'})
    C.row(P['psg3'], r + 3, {'A': NAME('L3'), 'C': NAME('L4'), 'F': NAME('L6'), 'J': NAME('L8')})
    C.row(P['psg3'], r + 4, {'A': 'التوقيع :', 'C': 'التوقيع :', 'F': 'التوقيع :', 'J': 'التوقيع :'})
    for i, g in enumerate((P['pgap1'], P['pgap2'], P['pgap3'])): C.row(g, r + 5 + i, height=45 if i == 0 else None)
    return r + 8


def build_qs_summary(wb, divs):
    """شيت «ملخص الحصر» (QS SUMMARY): كل بنود أمر الشراء بالسابق وخلال الفترة والإجمالي والقيم وأرقام الطلبات."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    ws = wb.create_sheet(QSUM)
    ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False; ws.sheet_view.zoomScale = 70
    NAVY, LIGHT, DIV, TOTF = '1F4E78', 'DDEBF7', 'BDD7EE', 'FFF2CC'
    thin = Side(style='thin', color='8EA9C1'); med = Side(style='medium', color=NAVY)
    B = Border(left=thin, right=thin, top=thin, bottom=thin)
    F = lambda sz=11, b=False, c='000000': Font(name='Arial', size=sz, bold=b, color=c)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center', wrap_text=True)
    widths = dict(A=12, B=11, C=58, D=9, E=13, F=12, G=12, H=12, I=12, J=16, K=16, L=17, M=10, N=24)
    for c, w in widths.items(): ws.column_dimensions[c].width = w
    # ---------------- الرأس
    ws.row_dimensions[1].height = 72
    ws.merge_cells('C1:K1')
    ws['C1'] = 'ملخص الحصر  |  Quantity Summary'; ws['C1'].font = F(22, True, NAVY); ws['C1'].alignment = CEN
    ws.merge_cells('C2:K2')
    ws['C2'] = f"={COVER}!B5"; ws['C2'].font = F(13, True); ws['C2'].alignment = CEN
    ws.row_dimensions[2].height = 24
    ws.merge_cells('A3:D3'); ws.merge_cells('E3:I3'); ws.merge_cells('J3:N3')
    ws['A3'] = 'المالك : الجهة المالكة – الإدارة العامة للمشاريع بالقطاع الأوسط'
    ws['E3'] = f'={COVER}!B9'; ws['J3'] = f'={COVER}!B11'
    for a in ('A3', 'E3', 'J3'): ws[a].font = F(11, True, NAVY); ws[a].alignment = CEN
    ws.row_dimensions[3].height = 22
    ws.merge_cells('A4:N4')
    ws['A4'] = (f'="المستخلص الجاري رقم ( "&{COVER}!F6&" )     الفترة من  "&TEXT({COVER}!E7,"dd-mm-yyyy")&"  إلى  "'
                f'&TEXT({COVER}!E8,"dd-mm-yyyy")')
    ws['A4'].font = F(14, True, 'FFFFFF'); ws['A4'].fill = PatternFill('solid', fgColor=NAVY); ws['A4'].alignment = CEN
    ws.row_dimensions[4].height = 28
    for c in 'ABCDEFGHIJKLMN':
        ws[f'{c}1'].border = Border(top=med); ws[f'{c}4'].border = Border(bottom=med)
    ws.row_dimensions[5].height = 6
    # ---------------- عناوين الجدول (صفين)
    h1 = {'A': 'رقم البند\nItem No', 'B': 'الكود\nCode', 'C': 'وصف البند\nDescription', 'D': 'الوحدة\nUnit',
          'E': 'الكمية التعاقدية\nBOQ Qty', 'F': 'الكميات المنفذة  |  Executed Qty', 'I': 'سعر الوحدة\nUnit Rate',
          'J': 'القيم المنفذة (ريال)  |  Executed Value (SAR)', 'M': 'نسبة التنفيذ\nProgress', 'N': 'أرقام الطلبات\nWIR No.'}
    h2 = {'F': 'السابقة\nPrevious', 'G': 'خلال الفترة\nThis Period', 'H': 'الإجمالية\nTotal',
          'J': 'السابقة\nPrevious', 'K': 'خلال الفترة\nThis Period', 'L': 'الإجمالية\nTotal'}
    for c, v in h1.items(): ws[f'{c}6'] = v
    for c, v in h2.items(): ws[f'{c}7'] = v
    for c in 'ABCDEIMN': ws.merge_cells(f'{c}6:{c}7')
    ws.merge_cells('F6:H6'); ws.merge_cells('J6:L6')
    for r in (6, 7):
        ws.row_dimensions[r].height = 34
        for c in 'ABCDEFGHIJKLMN':
            x = ws[f'{c}{r}']; x.font = F(10, True, 'FFFFFF'); x.fill = PatternFill('solid', fgColor=NAVY); x.alignment = CEN; x.border = B
    # ---------------- البنود
    r = 8; tot_rows = []; items = {}
    for d in divs:
        ws.merge_cells(f'A{r}:N{r}')
        ws[f'A{r}'] = f"القسم {d['no']}  –  {d['name']}"
        for c in 'ABCDEFGHIJKLMN':
            x = ws[f'{c}{r}']; x.fill = PatternFill('solid', fgColor=DIV); x.font = F(12, True, NAVY); x.alignment = RGT; x.border = B
        ws.row_dimensions[r].height = 22
        r += 1; a0 = r
        for it in d['rows']:
            if it.get('head'): continue
            ws[f'A{r}'] = it['item']; ws[f'B{r}'] = it['code']; ws[f'C{r}'] = it['desc']; ws[f'D{r}'] = it['unit_txt'].split('/')[-1] or it['unit']
            ws[f'E{r}'] = it['qty']; ws[f'I{r}'] = it['rate']
            ws[f'G{r}'] = f'=H{r}-F{r}'
            ws[f'J{r}'] = f'=ROUND(F{r}*I{r},2)'; ws[f'K{r}'] = f'=L{r}-J{r}'; ws[f'L{r}'] = f'=ROUND(H{r}*I{r},2)'
            ws[f'M{r}'] = f'=IF(E{r}=0,"",H{r}/E{r})'
            n = len(it['desc'])
            ws.row_dimensions[r].height = 20 if n <= 60 else 32 if n <= 130 else 46 if n <= 220 else 62
            for c in 'ABCDEFGHIJKLMN':
                x = ws[f'{c}{r}']; x.border = B; x.font = F(10); x.alignment = RGT if c in 'CN' else CEN
            for c in 'EFGH': ws[f'{c}{r}'].number_format = '#,##0.00;-#,##0.00;"-"'
            for c in 'IJKL': ws[f'{c}{r}'].number_format = '#,##0.00;-#,##0.00;"-"'
            ws[f'M{r}'].number_format = '0.0%;-0.0%;"-"'
            if (r - a0) % 2: 
                for c in 'ABCDEFGHIJKLMN': ws[f'{c}{r}'].fill = PatternFill('solid', fgColor='F5F9FC')
            items[it['code']] = r
            r += 1
        ws.merge_cells(f'A{r}:I{r}')
        ws[f'A{r}'] = f"إجمالي {d['name']}"
        for c in 'JKL': ws[f'{c}{r}'] = f'=SUM({c}{a0}:{c}{r - 1})'
        for c in 'ABCDEFGHIJKLMN':
            x = ws[f'{c}{r}']; x.fill = PatternFill('solid', fgColor=LIGHT); x.font = F(11, True); x.border = B
            x.alignment = RGT if c == 'A' else CEN
            if c in 'JKL': x.number_format = '#,##0.00'
        tot_rows.append(r); r += 1
    # ---------------- الإجماليات
    r += 1
    lines = [('إجمالي قيمة الأعمال  |  Total Value', lambda c: '=' + '+'.join(f'{c}{t}' for t in tot_rows)),
             ('ضريبة القيمة المضافة 15%  |  VAT 15%', lambda c: f'=ROUND({c}{r}*{VAT},2)'),
             ('الإجمالي شامل الضريبة  |  Total incl. VAT', lambda c: f'={c}{r}+{c}{r + 1}')]
    for i, (lab, fn) in enumerate(lines):
        rr = r + i
        ws.merge_cells(f'A{rr}:I{rr}'); ws[f'A{rr}'] = lab
        for c in 'JKL': ws[f'{c}{rr}'] = fn(c)
        for c in 'ABCDEFGHIJKLMN':
            x = ws[f'{c}{rr}']; x.border = B; x.font = F(12, True, 'FFFFFF' if i == 2 else '000000')
            x.fill = PatternFill('solid', fgColor=NAVY if i == 2 else TOTF); x.alignment = RGT if c == 'A' else CEN
            if c in 'JKL': x.number_format = '#,##0.00'
        ws.row_dimensions[rr].height = 24
    grand = r
    r += 4
    # ---------------- التوقيعات (زي ملخص حصر المقاول): المقاول | الاستشاري
    ws.merge_cells(f'A{r}:E{r}'); ws.merge_cells(f'F{r}:N{r}')
    ws[f'A{r}'] = f'={COVER}!B11'; ws[f'F{r}'] = f'={COVER}!B9'
    for a in (f'A{r}', f'F{r}'): ws[a].font = F(12, True, NAVY); ws[a].alignment = CEN
    blocks = [('A', 'B', 'الوظيفة : حاسب الكميات', 'L4'), ('C', 'E', 'الوظيفة : مدير المشروع', 'L3'),
              ('F', 'I', 'الوظيفة : حاسب الكميات', 'L8'), ('J', 'N', 'الوظيفة : مهندس الموقع', 'L9')]
    for c0, c1, job, nm in blocks:
        for k, v in enumerate((job, NAME(nm), 'التوقيع :'), 1):
            ws.merge_cells(f'{c0}{r + k}:{c1}{r + k}'); ws[f'{c0}{r + k}'] = v
            ws[f'{c0}{r + k}'].font = F(11, k == 1); ws[f'{c0}{r + k}'].alignment = CEN
    for k in range(1, 4): ws.row_dimensions[r + k].height = 24
    ws.row_dimensions[r + 3].height = 40
    last = r + 4
    ws.freeze_panes = 'D8'
    ws.print_title_rows = '1:7'
    ws.print_area = f'A1:N{last}'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.3; ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = 'صفحة &P من &N'
    return items, grand


def build_cover(ws):
    for m in list(ws.merged_cells.ranges):
        if m.min_col >= 10: ws.unmerge_cells(str(m))
    for row in ws.iter_rows():
        for c in row:
            if c.column >= 10: c.value = None; c.style = 'Normal'
    ws['B1'] = 'المملكة العربية السعودية'
    ws['B2'] = 'الجهة المالكة وحـــــدة أعمـــــــال الريـــــــاض'
    ws['B3'] = 'الإدارة العامة للمشاريع بالقطاع الأوسط'
    ws['B5'] = 'مشروع شبكات المياه (مثال) ( أمر شراء 000000000 )'
    ws['D6'] = 'المستخلص الجارى رقم'; ws['F6'] = 1
    ws['D7'] = 'الفترة من'; ws['E7'] = None; ws['G7'] = 'م'
    ws['D8'] = 'إلي'; ws['E8'] = None; ws['G8'] = 'م'
    for a in ('E7', 'E8'): ws[a].number_format = 'dd-mm-yyyy'
    ws['B9'] = 'الاستشاري : مكتب الاستشاري'
    ws['B11'] = 'المقاول : شركة المقاول'
    # جدول أسماء الموقّعين (برّه منطقة الطباعة)
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    j = ws.column_dimensions['J']; j.min = 10; j.max = 10
    ws.column_dimensions['K'].width = 34; ws.column_dimensions['L'].width = 34
    thin = Side(style='thin', color='808080'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws['K2'] = 'أسماء الموقّعين – اكتبها هنا مرة واحدة'
    ws['K2'].font = Font(name='Arial', size=14, bold=True, color='1F4E78')
    for cell, lab in SIGNERS:
        r = int(cell[1:])
        a, b = ws[f'K{r}'], ws[f'L{r}']
        a.value = lab; a.font = Font(name='Arial', size=12, bold=True); a.border = B
        a.alignment = Alignment(horizontal='right', vertical='center')
        b.font = Font(name='Arial', size=12); b.border = B; b.fill = PatternFill('solid', fgColor='E2EFDA')
        b.alignment = Alignment(horizontal='right', vertical='center')


RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
DRAW_T = RNS + '/drawing'


def _sheet_part(z, name):
    wbx = z.read('xl/workbook.xml').decode('utf8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    m = re.search(rf'<sheet [^>]*name="{name}"[^>]*r:id="(rId\d+)"|<sheet [^>]*r:id="(rId\d+)"[^>]*name="{name}"', wbx)
    if not m: return None
    rid = m.group(1) or m.group(2)
    m = re.search(rf'<Relationship [^>]*Id="{rid}"[^>]*Target="([^"]+)"|<Relationship [^>]*Target="([^"]+)"[^>]*Id="{rid}"', rels)
    t = (m.group(1) or m.group(2)).lstrip('/')
    return t if t.startswith('xl/') else 'xl/' + t


def _get_drawing(z, name):
    """(drawing xml, rels xml, [media]) لشيت معيّن، أو None"""
    sp = _sheet_part(z, name)
    if not sp: return None
    rp = sp.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if rp not in z.namelist(): return None
    m = re.search(r'Target="\.\./drawings/([^"]+)"', z.read(rp).decode('utf8'))
    if not m: return None
    xml = z.read('xl/drawings/' + m.group(1)).decode('utf8')
    rels = z.read('xl/drawings/_rels/' + m.group(1) + '.rels').decode('utf8')
    return xml, rels, re.findall(r'Target="\.\./media/([^"]+)"', rels)


LOGO_SPECS = {
    # (رقم الصورة: 0 الجهة المالكة، 1 الاستشاري، 2 المقاول, العمود, الصف, العرض المتاح بالبكسل, الارتفاع المتاح EMU, ارتفاع اللوجو EMU)
    IPC: [(0, 3, 1, 211, 1600200, 1350000), (1, 8, 1, 384, 1600200, 1350000), (2, 12, 1, 211, 1600200, 1350000)],
    QSUM: [(0, 0, 0, 171, 914400, 780000), (1, 11, 0, 124, 914400, 700000), (2, 13, 0, 173, 914400, 780000)],
}


def _logo_drawing(cover_xml, spec):
    pics = re.findall(r'<xdr:pic>.*?</xdr:pic>', cover_xml, re.S)
    if len(pics) < 3: return None
    EMU = 9525
    out = []
    for i, (k, col, row, wpx, havail, h) in enumerate(spec):
        pic = pics[k]
        cx, cy = map(int, re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', pic).groups())
        w = int(h * cx / cy); hh = h
        maxw = wpx * EMU - 60000
        if w > maxw: hh = int(h * maxw / w); w = maxw
        xoff = max(0, (wpx * EMU - w) // 2); yoff = max(0, (havail - hh) // 2)
        pic = re.sub(r'<a:off x="\d+" y="\d+"/>', '<a:off x="0" y="0"/>', pic)
        pic = re.sub(r'<a:ext cx="\d+" cy="\d+"/>', f'<a:ext cx="{w}" cy="{hh}"/>', pic, count=1)
        pic = re.sub(r'<xdr:cNvPr id="\d+"', f'<xdr:cNvPr id="{10 + i}"', pic)
        out.append(f'<xdr:oneCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{xoff}</xdr:colOff><xdr:row>{row}</xdr:row>'
                   f'<xdr:rowOff>{yoff}</xdr:rowOff></xdr:from><xdr:ext cx="{w}" cy="{hh}"/>{pic}<xdr:clientData/></xdr:oneCellAnchor>')
    head = cover_xml[:cover_xml.index('>', cover_xml.index('<xdr:wsDr')) + 1]
    return head + ''.join(out) + '</xdr:wsDr>'


def add_cover_logos(src_xlsx, out, ref, specs=None):
    """openpyxl بيشيل الصور، فبنرجّع اللوجوهات: الغلاف (CAVAR) + رأس شيت المستخلص (VO).
    specs = {اسم الشيت: مواصفات اللوجوهات} لشيتات تانية (زي الحصر الكلي) – الصور بتتاخد من غلاف التمبلت."""
    zr = zipfile.ZipFile(ref)
    cover = _get_drawing(zr, COVER)
    if not cover:
        zr.close(); shutil.copyfile(src_xlsx, out); return
    parts = {COVER: ('logo1.xml',) + cover}
    specs = specs or {IPC: LOGO_SPECS[IPC], QSUM: LOGO_SPECS[QSUM]}
    for i, (sh, spec) in enumerate(specs.items(), 2):
        x = _logo_drawing(cover[0], spec)
        if x: parts[sh] = (f'logo{i}.xml', x, cover[1], cover[2])
    media = {}
    for nm, (fn, xml, rels, med) in parts.items():
        for m in med:
            media[m if m.startswith('logo_') else 'logo_' + m] = zr.read('xl/media/' + m)
    zi = zipfile.ZipFile(src_xlsx)
    names = zi.namelist()
    targets = {}
    for nm in parts:
        sp = _sheet_part(zi, nm)
        if sp: targets[sp] = nm
    zo_tmp = out + '.z'
    zo = zipfile.ZipFile(zo_tmp, 'w', zipfile.ZIP_DEFLATED)
    done_rels = set()
    for n in names:
        if n.startswith('xl/drawings/logo') or n.startswith('xl/media/logo_'): continue
        data = zi.read(n)
        if n in targets:
            s = data.decode('utf8')
            s = re.sub(r'<drawing r:id="[^"]+"\s*/>', '', s)
            if 'xmlns:r=' not in s[:600]:
                s = s.replace('<worksheet ', f'<worksheet xmlns:r="{RNS}" ', 1)
            tag = '<drawing r:id="rIdLogo"/>'
            if '<legacyDrawing' in s: s = s.replace('<legacyDrawing', tag + '<legacyDrawing', 1)
            elif '<tableParts' in s: s = s.replace('<tableParts', tag + '<tableParts', 1)
            elif '<extLst' in s: s = s.replace('<extLst', tag + '<extLst', 1)
            else: s = s.replace('</worksheet>', tag + '</worksheet>')
            data = s.encode('utf8')
        elif n.startswith('xl/worksheets/_rels/'):
            sp = n.replace('worksheets/_rels/', 'worksheets/')[:-5]
            if sp in targets:
                s = data.decode('utf8')
                s = re.sub(r'<Relationship [^>]*Type="' + re.escape(DRAW_T) + r'"[^>]*/>', '', s)
                s = s.replace('</Relationships>', f'<Relationship Id="rIdLogo" Type="{DRAW_T}" Target="../drawings/{parts[targets[sp]][0]}"/></Relationships>')
                data = s.encode('utf8'); done_rels.add(sp)
        elif n == '[Content_Types].xml':
            s = data.decode('utf8')
            have_ext = {e.lower() for e in re.findall(r'<Default [^>]*Extension="([^"]+)"', s)}
            for m in media:
                ext = m.rsplit('.', 1)[-1]
                if ext.lower() in have_ext: continue
                ct = 'image/png' if ext.lower() == 'png' else 'image/jpeg'
                s = s.replace('</Types>', f'<Default Extension="{ext}" ContentType="{ct}"/></Types>'); have_ext.add(ext.lower())
            for nm, (fn, *_r) in parts.items():
                if _sheet_part(zi, nm) is None: continue          # شيت مش موجود (زي QS SUMMARY) – من غير لوجو
                pn = f'/xl/drawings/{fn}'
                if pn not in s:
                    s = s.replace('</Types>', f'<Override PartName="{pn}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
            data = s.encode('utf8')
        zo.writestr(n, data)
    for sp, nm in targets.items():
        if sp not in done_rels:
            rn = sp.replace('worksheets/', 'worksheets/_rels/') + '.rels'
            zo.writestr(rn, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                            f'<Relationship Id="rIdLogo" Type="{DRAW_T}" Target="../drawings/{parts[nm][0]}"/></Relationships>')
    for nm, (fn, xml, rels, med) in parts.items():
        if _sheet_part(zi, nm) is None: continue
        zo.writestr('xl/drawings/' + fn, xml)
        for m in med:
            nmm = m if m.startswith('logo_') else 'logo_' + m
            rels = rels.replace(f'../media/{m}"', f'../media/{nmm}"')
        zo.writestr('xl/drawings/_rels/' + fn + '.rels', rels)
    for m, b in media.items(): zo.writestr('xl/media/' + m, b)
    zo.close(); zi.close(); zr.close()
    os.replace(zo_tmp, out)


HELP_LINES = [
    ('طريقة عمل المستخلص', 'title'),
    ('الملف ده هو التمبلت الثابت للمستخلص، بنفس تنسيق مستخلصات المقاول. مكانه: 00_Registers\\11.1-TEMPLATE IPC.xlsx. متعدّلش فيه، خد منه نسخة.', ''),
    ('', ''),
    ('أولاً: بالبرنامج (الأسهل)', 'h'),
    ('1. افتح البرنامج ← تبويب «IPC». اكتب رقم المستخلص والفترة، ودوس «احسب» تشوف المستخلص والخلاصة والحصر الكلي.', ''),
    ('2. دوس «طلّع المستخلص»: البرنامج بياخد نسخة من التمبلت ويملاها (حتى تاريخه من الطلبات المعتمدة، والسابق من المستخلص اللي قبله، والمبالغ بالحروف).', ''),
    ('3. النسخة بتتحفظ في 08_Cost_Payments\\CLIENT_IPC باسم IPC-رقم المستخلص، وبتتسجل سطر في 00_Registers\\11-LOG IPC.xlsx.', ''),
    ('4. ورا شيت VO بتلاقي نفس شاشات البرنامج (الملخص، الحفر، الفرشة، المواسير، الردم، قاعدة الطريق، الطلبات) عشان أي رقم يتراجع.', ''),
    ('', ''),
    ('ثانياً: من غير البرنامج (بإيدك)', 'h'),
    ('1. خد نسخة من التمبلت وسمّيها IPC-رقم المستخلص وحطها في 08_Cost_Payments\\CLIENT_IPC.', ''),
    ('2. شيت CAVAR: اكتب رقم المستخلص (F6) والفترة من (E7) وإلى (E8).', ''),
    ('3. شيت VO – صفحات البنود: عمود «السابقة» = «حتى تاريخه» من المستخلص اللي قبله. عمود «حتى تاريخه» = الكمية المعتمدة لحد نهاية الفترة.', ''),
    ('4. «خلال الفترة» والقيم وصفحة الأقسام وصفحة المبالغ بتتحسب لوحدها. المبالغ بالحروف بتتكتب بإيدك (البرنامج بيكتبها لوحده).', ''),
    ('5. سجّل المستخلص في 11-LOG IPC.xlsx.', ''),
    ('', ''),
    ('ملاحظات', 'h'),
    ('• الأسعار والكميات من جدول كميات أمر الشراء 000000000 (0,000,000 ريال بدون الضريبة).', ''),
    ('• «الكمية بالجدول المعدل» = الكمية التقديرية لحد ما يطلع جدول معدّل؛ لو اتعدّل اكتب الكمية الجديدة مكان المعادلة.', ''),
    ('• المستخلص الختامي ميقلّش عن 10% من قيمة العقد (أمر الشراء)، وفي خانة O4 في شيت VO تنبيه لو الجاري عدّى 90%.', ''),
    ('• مفيش دفعة مقدمة ولا خصومات غير الضريبة 15%. الضمان النهائي 5% خطاب بنك.', ''),
]


def build_help(wb):
    from openpyxl.styles import Font, Alignment
    ws = wb.create_sheet(HELP)
    ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 3; ws.column_dimensions['B'].width = 140
    for i, (t, k) in enumerate(HELP_LINES, 2):
        c = ws[f'B{i}']; c.value = t; c.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
        c.font = Font(name='Arial', size=16 if k == 'title' else 12 if k == 'h' else 11, bold=bool(k), color='1F4E78' if k else '000000')


if __name__ == '__main__':
    import sys
    o, d, info = build_template(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else REF)
    for x in d:
        items = [r for r in x['rows'] if not r.get('head')]
        print(x['no'], x['name'], len(items), round(sum(r['qty'] * r['rate'] for r in items), 2))
    print({k: v for k, v in info.items() if k != 'item_rows'}, len(info['item_rows']))
