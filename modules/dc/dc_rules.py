# -*- coding: utf-8 -*-
"""DC Control – يعرف نوع الورقة من الكلام اللي اتقري، ويطلّع رقمها وتاريخها وحالتها والفولدر والاسم.

كل نوع ليه علامات بوزن؛ أعلى مجموع يكسب. العلامات القوية (أرقام الأكواد) وزنها كبير،
والكلمات العامة وزنها صغير. اللي بتصحّحه بيتحفظ في dc_learn.json ويتحسب في المرة الجاية.
"""
import os, re, json, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
LEARN = os.path.join(HERE, 'dc_learn.json')
SP = r'[\s\-–_]*'                                         # مسافات وشرط OCR بين أجزاء الكود

# ---------------------------------------------------------------- الأنواع
# code: (الاسم بالعربي، الفولدر نسبةً لفولدر المشروع – {st} = حالة الاعتماد)
TYPES = {
    'UNF_WIR': ('طلب استلام يونيفاير', r'05_Quality\WIR\5_UNIFIER_APPROVED'),
    'SITE_WIR': ('طلب استلام موقع', r'05_Quality\WIR\{st}'),
    'MIR': ('طلب استلام مواد MIR', r'05_Quality\MIR\{st}'),
    'MAR': ('طلب اعتماد مواد MAR', r'04_Submittals\MAR\{st}'),
    'MS': ('طريقة تنفيذ MS', r'04_Submittals\MS\{st}'),
    'ITP': ('خطة فحص ITP', r'04_Submittals\ITP\{st}'),
    'SD': ('رسومات SD', r'04_Submittals\SD\{sdg}\{st}'),          # {sdg} = الخط (L1…) أو DETAILS / CHAMBER_MECH
    'SUB': ('اعتماد مقاول باطن', r'13_PARTIES'),                     # فولدر الجهة\02_APPROVAL\1_WIP (بيتحدد مع الجهة)
    'QUOTE': ('عرض سعر', r'13_PARTIES'),
    'PROD': ('تقرير الإنتاج الأسبوعي (سيستم الشركة)', r'07_Planning_Progress\REPORTS_CTR\WEEKLY'),   # عمود «Reported» في QS                               # فولدر الجهة\01_QUOTATIONS
    'NCR': ('عدم مطابقة NCR', r'05_Quality\NCR'),
    'TEST': ('تقرير اختبار', r'05_Quality\TESTS\{st}'),
    'CAL': ('شهادة معايرة', r'05_Quality\CALIBRATION'),
    'RFI': ('طلب معلومات RFI', r'02_Correspondence\OUT\YA\RFI'),
    'OUT_YA': ('خطاب صادر للاستشاري', r'02_Correspondence\OUT\YA'),
    'OUT_CLIENT': ('خطاب صادر للمياه', r'02_Correspondence\OUT\CLIENT'),
    'OUT_SUB': ('خطاب صادر لمورد/مقاول', r'02_Correspondence\OUT\SUB'),
    'OUT_OTH': ('خطاب صادر لجهة أخرى', r'02_Correspondence\OUT\OTH'),
    'IN_YA': ('خطاب وارد من الاستشاري', r'02_Correspondence\IN\YA'),
    'IN_CLIENT': ('وارد من المياه', r'02_Correspondence\IN\CLIENT'),
    'IN_SUB': ('وارد من مورد/مقاول', r'02_Correspondence\IN\SUB'),
    'IN_OTH': ('وارد من جهة أخرى', r'02_Correspondence\IN\OTH'),
    'PERMIT': ('تصريح / رخصة حفر', r'12_Permits_Utilities'),
    'IPC': ('مستخلص', r'08_Cost_Payments\CLIENT_IPC'),
    'IPC_INV': ('فاتورة مستخلص (VN)', r'08_Cost_Payments\CLIENT_IPC\{ipc}'),     # {ipc} = 1-IPC، 2-IPC…
    'IPC_SIGNED': ('مستخلص موقّع (سكان)', r'08_Cost_Payments\CLIENT_IPC\{ipc}'),
    'SUB_IPC': ('مستخلص مقاول باطن', r'13_PARTIES'),                 # فولدر الجهة\04_CERTIFICATES
    'UPD_CAL': ('تقويم تحديث الشركة', r'08_Cost_Payments\QS\UPDATE_CALENDAR'),          # مواعيد التحديث الشهرية
    'DN': ('إذن توريد', r'09_Procurement_Materials\DELIVERY'),                         # الاسم زي عمود DL في سجل المخزن
    'SUB_CONTRACT': ('عقد مقاول باطن', r'13_PARTIES'),               # فولدر الجهة\03_CONTRACT + الـ 3 اعتمادات
    'PAYROLL': ('مسير رواتب (مع المستخلص)', r'08_Cost_Payments\CLIENT_IPC\{ipc}'),   # فولدر المستخلص اللي فترته فيها شهر المسير
    'GIS': ('ملف GIS (أز بلت) مع المستخلص', r'08_Cost_Payments\CLIENT_IPC\{ipc}'),    # zip – فولدر آخر مستخلص
    'PRICE': ('تحليل أسعار', r'08_Cost_Payments\COST_CONTROL\PRICE_ANALYSIS'),          # PA-01 - … (طلب المستخدم 23-11)
}
SUB_CODES = ('SUB', 'SUB_IPC', 'SUB_CONTRACT', 'QUOTE')             # الأنواع اللي بتروح فولدر الجهة نفسها (13_PARTIES)

# أسامي الأنواع في شاشة البرنامج (الواجهة إنجليزي)
EN = {'UNF_WIR': 'Unifier WIR', 'SITE_WIR': 'Site WIR', 'MIR': 'MIR', 'MAR': 'MAR', 'MS': 'Method Statement',
      'ITP': 'ITP', 'SD': 'Shop Drawing', 'SUB': 'Subcontractor', 'NCR': 'NCR', 'TEST': 'Test Report',
      'CAL': 'Calibration', 'RFI': 'RFI', 'OUT_YA': 'Letter OUT – the Consultant', 'OUT_CLIENT': 'Letter OUT – CLIENT',
      'OUT_SUB': 'Letter OUT – Supplier/Sub', 'OUT_OTH': 'Letter OUT – Other', 'IN_YA': 'Letter IN – the Consultant',
      'IN_CLIENT': 'IN – CLIENT', 'IN_SUB': 'IN – Supplier/Sub', 'IN_OTH': 'IN – Other', 'PERMIT': 'Permit',
      'IPC': 'IPC', 'IPC_INV': 'IPC Invoice (VN)', 'IPC_SIGNED': 'Signed IPC (scan)', 'SUB_IPC': 'Subcontractor IPC',
      'UPD_CAL': 'Company Update Calendar', 'DN': 'Delivery Note', 'SUB_CONTRACT': 'Subcontract', 'PAYROLL': 'Payroll', 'GIS': 'GIS (as-built)',
      'QUOTE': 'Quotation', 'PROD': 'Production Report', 'PRICE': 'Price Analysis'}

CONTRACT_HEAD = r'عقد\s*(مقاولة|مقاول)\s*(من\s*)?(ال)?باطن|عقد\s*تأجير\s*(عمالة|معدات)|عقد\s*توريد\s*وتركيب|Sub-?contract\s*Agreement'
PAYROLL_HEAD = r'مسير\s*(ال)?رواتب|كشف\s*(ال)?رواتب|مسير\s*(ال)?(عمال|العاملين|المشغلين|عمالة)|Pay\s*-?roll|Salary\s*Sheet|Wages\s*Sheet'

# (نوع، نمط، وزن)
RULES = [
    ('UNF_WIR', rf'WIR{SP}AP{SP}3500', 60), ('UNF_WIR', rf'Civil{SP}\d{{5}}', 50),
    ('UNF_WIR', r'Works Inspection Request', 8), ('UNF_WIR', r'Client Authority|الجهة المالكة', 6),
    ('SITE_WIR', rf'PMF{SP}010{SP}QUA{SP}WIR', 60), ('SITE_WIR', rf'PMF{SP}010{SP}QUA{SP}013', 30), ('SITE_WIR', r'طلب استلام', 10),
    ('SITE_WIR', r'Works Inspection Request', 8), ('SITE_WIR', r'Request for (Inspection|Pipe Laying|Bedding|Cover|Backfill|Asphalt|Road Base|Sub ?grade)', 15),
    ('MIR', rf'MIR{SP}\d', 40), ('MIR', r'Material Inspection Request|طلب (استلام|فحص) مواد', 40),
    ('MAR', rf'MAR{SP}\d|M\d{{1,2}}{SP}Material Approval', 30), ('MAR', r'Material Approval Request|اعتماد مواد|طلب اعتماد مادة', 35),
    ('MS', r'METHOD STATEMENT\s+(of|for)', 40), ('MS', rf'QUA{SP}004', 20), ('MS', r'طريقة (العمل|التنفيذ)', 15),
    ('ITP', r'Inspection (and|&) Test(ing)? Plan', 40), ('ITP', rf'ITPS?{SP}AP3500', 10),
    ('SD', r'Shop Drawing Submittal', 30), ('SD', rf'CLIENT{SP}W(?:TR|RT){SP}S[HD]', 40),
    ('PROD', r'نتاج[\s\S]{0,3000}المنجزة', 45),                         # تقرير الإنتاج من سيستم الشركة (الكمية المنجزة)
    ('PRICE', r'تحليل\s*(ال)?[اأإ]سعار|Price\s*Analysis|Rate\s*Analysis|Unit\s*Rate\s*Breakdown', 50),
    ('QUOTE', r'Quotation|عرض\s*(ال)?سعر|عرض\s*(ال)?أسعار|Price\s*Offer|Financial\s*(and\s*Technical\s*)?Offer|Technical\s*and\s*Financial\s*Offer', 40),
    ('SUB', r'Subcontractor Approval|اعتماد مقاول (من )?(ال)?باطن|كمقاول باطن', 35),
    ('SUB', r'سجل تجاري|السجل التجاري|Commercial Registration|شهادة تسجيل.{0,20}(ضريبة|القيمة المضافة)|VAT Registration|العنوان الوطني|عنوان وطني|National Address|Company Profile|بروفايل|Pre-?Qualification', 30),
    ('NCR', r'Non[\s-]?Conformance|عدم مطابقة|NCR', 40),
    ('TEST', r'Test Report|Field Density|Proctor|Compaction|دمك|تقرير (اختبار|فحص)|معمل|Laboratory', 25),
    ('CAL', r'Calibration Certificate|شهادة معايرة|Calibration', 40),
    ('RFI', r'Request for Information|طلب معلومات|\bRFI\b', 45),
    ('OUT_YA', rf'CTR{SP}YA{SP}WA{SP}\d', 45), ('OUT_CLIENT', rf'CTR{SP}CLIENT{SP}WA{SP}\d', 45),
    ('OUT_SUB', rf'CTR{SP}(SUP|SUB){SP}WA{SP}\d', 45), ('OUT_OTH', rf'CTR{SP}OTH{SP}WA{SP}\d', 45),
    ('IN_YA', r'CONSULTANT\s*OFFICE|الاستشاري للاستشارات|consultant\.example', 30), ('IN_YA', rf'cons{SP}26{SP}\d+', 45),
    ('IN_CLIENT', r'محضر|المحضر التنسيقي|خطاب.*الجهة المالكة', 20),
    ('PERMIT', r'رخصة حفر|تصريح حفر|Excavation Permit|رخصة رقم', 35),
    ('IPC', r'مستخلص\s*(ال)?(جار[يى]|رقم|ختامي)|Interim Payment Certificate|\bIPC\b', 30),
    ('IPC_INV', r'فاتورة\s*ضريبية', 30),          # فاتورة المستخلص من نظام المياه (بتتأكد في classify)
    ('SUB_IPC', r'(مستخلص|Payment Certificate|Invoice|فاتورة).{0,400}(مقاول (من )?الباطن|المقاول الرئيسي|Sub-?contract)|(مقاول (من )?الباطن|المقاول الرئيسي|Sub-?contract).{0,400}(مستخلص|Payment Certificate)', 45),
    # تقويم تحديث الشركة: «الفترة الأولى/الثانية …» و«يتم الرفع على النظام» (الكلام في الـ PDF ساعات بيطلع بالمقلوب)
    ('UPD_CAL', r'يتم\s*الرفع\s*على\s*النظام|النظام\s*على\s*الرفع\s*يتم', 35),
    ('UPD_CAL', r'الفترة\s*ال[أا]ولى|ال[أا]ولى\s*الفترة', 20), ('UPD_CAL', r'الفترة\s*الثانية|الثانية\s*الفترة', 20),
    ('UPD_CAL', r'SUNDAY\s+MONDAY\s+TUESDAY|إجمالي\s*الأيام', 10),
    # إذن توريد المورد (ومعاه فاتورته ساعات): الأصناف دي بتتسجل في المخزن ويتعمل لها MIR
    ('SUB_CONTRACT', CONTRACT_HEAD, 50),
    ('SUB_CONTRACT', r'الطرف\s*الأول.{0,400}الطرف\s*الثاني', 20),
    ('PAYROLL', PAYROLL_HEAD, 45),
    ('DN', r'Delivery\s*Note|إذن\s*(توريد|تسليم|استلام)|سند\s*(تسليم|استلام)|\bDN\s*No|Delivery\s*Order', 45),
    ('DN', r'TAX\s*INVOICE|فاتورة\s*ضريبية', 10),
]
COMPANY = r'المقاول|Contractor Co'   # اسم شركتك (بيتغير من الإعدادات)
STATUS = [('rejected', r'Rejected|مرفوض|Code\s*C\b|Not Approved|غير معتمد'),
          ('approved_comments', r'Approved (with|as) (comments|noted)|معتمد (مع|ب)ملاحظات|Code\s*B\b'),
          ('approved', r'^\s*APPROVED\s*$|Code\s*A\b|(?<![\w])معتمد(?![\w])(?! من)')]


# ---------------------------------------------------------------- استخراج
def _num(s):
    return int(re.sub(r'\D', '', s) or 0)


def find_date(t):
    for p, order in ((r'(20\d\d)\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d\s?\d?)', 'ymd'),
                     (r'(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(20\d\d)', 'dmy')):
        for m in re.finditer(p, t):
            a, b, c = (_num(x) for x in m.groups())
            y, mo, d = (a, b, c) if order == 'ymd' else (c, b, a)
            try: return dt.date(y, mo, d)
            except ValueError: continue
    return None


MONTHS_EN = ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
MONTHS_AR = (('يناير', 1), ('فبراير', 2), ('مارس', 3), ('ابريل', 4), ('أبريل', 4), ('مايو', 5), ('يونيو', 6), ('يوليو', 7),
             ('اغسطس', 8), ('أغسطس', 8), ('سبتمبر', 9), ('اكتوبر', 10), ('أكتوبر', 10), ('نوفمبر', 11), ('ديسمبر', 12))


def cal_month(t, src_name=''):
    """'YYYY-MM' من أول سطور التقويم (November 2026 / نوفمبر)، وبعدين من اسم الملف (NOV-2026)."""
    head = '\n'.join([l for l in (t or '').splitlines() if l.strip()][:3])
    m = next((i for i, n in enumerate(MONTHS_EN, 1) if re.search(n, head, re.I)), None) or \
        next((k for w, k in MONTHS_AR if w in head), None)
    y = re.search(r'(20\d\d)', head)
    base = os.path.basename(src_name or '')
    if not m:
        m = next((i for i, n in enumerate(MONTHS_EN, 1) if re.search(rf'(?<![A-Za-z]){n[:3]}(?:{n[3:]})?(?![A-Za-z])', base, re.I)), None)
    y = int(y.group(1)) if y else (int(re.search(r'(20\d\d)', base).group(1)) if re.search(r'(20\d\d)', base) else None)
    return f'{y}-{m:02d}' if (y and m) else None


# ---------------------------------------------------------------- الـ Shop Drawing (طلب المستخدم 12-11)
# تابع للخط (فولدر الخط L1…): Layout / Plan & Profile · Connection · Chamber STR
# عام: Detail ← SD\DETAILS · Chamber MECH ← SD\CHAMBER_MECH  (القديم اتنقل SD\OLD بإيد المستخدم 12-11)
SD_LINE_KINDS = ('Layout', 'Plan & Profile', 'Connection', 'Chamber STR')
SD_GENERAL = {'Detail': 'DETAILS', 'Chamber MECH': 'CHAMBER_MECH'}
SD_KINDS = SD_LINE_KINDS + tuple(SD_GENERAL)
SD_DWG = re.compile(rf'CLIENT{SP}W(?:TR|RT){SP}S[DH](?:{SP}[A-Z]{{1,3}}\d{{0,2}})*(?:{SP}\d{{1,3}})+', re.I)   # WRT = غلطة كتابة شايعة


def sd_info(text, src_name=''):
    """نوع اللوحة والخط وكود اللوحة من الكلام واسم الملف."""
    t = ' '.join(((text or '')[:4000] + ' ' + os.path.basename(src_name or '')).split())
    t = re.sub(r'(?<![A-Za-z])([DL])!', r'\g<1>1', t)               # «SD-D!» = D-1 (الـ ! هي 1 من غير Shift)
    out = {}
    m = SD_DWG.search(t)
    if m:
        d = re.sub(r'\s*[-–_]\s*|\s+', '-', m.group(0).strip()).upper().replace('-WRT-', '-WTR-')
        out['dwg'] = re.sub(r'-D(\d)', r'-D-\1', d)                     # تفاصيل: D-1 (مش D1)
    m = re.search(rf'S[DH]{SP}(L{SP}\d{{1,2}})(?!\d)', t, re.I) or re.search(r'(?<![A-Z0-9])(L\s?-?\s?\d{1,2})(?![\d.])', t)
    if m: out['line'] = re.sub(r'\W|_', '', m.group(1)).upper()
    ch = re.search(rf'chamber|غرف|S[DH]{SP}CH(?![A-Z])', t, re.I)
    if ch:
        if re.search(r'reinforc|rebar|concrete|slab|structur|انشائ|إنشائ|تسليح|خرسان', t, re.I): out['kind'] = 'Chamber STR'
        elif re.search(r'valve|fitting|flange|mechanic|pipe work|ميكانيك|محبس|محابس', t, re.I): out['kind'] = 'Chamber MECH'
        else: out['kind'] = 'Chamber'                              # غرفة بس مش عارف ميكانيكا ولا إنشائي ← Check
    elif re.search(r'connection|connect(ed|ing)? to|to ex(isting)?\b|tie[\s-]?in|ربط', t, re.I): out['kind'] = 'Connection'
    elif re.search(rf'\bdetails?\b|thrust\s*block|typical|standard|S[DH]{SP}D{SP}\d', t, re.I): out['kind'] = 'Detail'
    elif re.search(r'plan\s*(&|and)?\s*profile|longitudinal|profile', t, re.I): out['kind'] = 'Plan & Profile'
    elif re.search(r'layout|alignment|general plan|key plan', t, re.I) or out.get('line'): out['kind'] = 'Layout'
    return out


def sd_group(f):
    """فولدر اللوحة جوه SD: الخط أو DETAILS / CHAMBER_MECH، أو '' لو مش معروف."""
    k = f.get('kind')
    if k in SD_GENERAL: return SD_GENERAL[k]
    if k in SD_LINE_KINDS and f.get('line'): return f['line']
    return ''


def sd_note(f):
    k = f.get('kind')
    if not k: return 'Shop drawing: choose the kind (Line / Detail / Connection / Chamber) – type it at the start of the name'
    if k == 'Chamber': return 'Chamber: mechanical (CHAMBER_MECH) or structural (per line)? – fix the folder and name'
    if k in SD_LINE_KINDS and not f.get('line'): return f'{k}: which line (L1, L2…)? – type it in the folder'
    return None


def status_from_title(title):
    """حالة الاعتماد من عنوان Claude: «… approved with comments» / «… approved» / «… rejected»."""
    t = (title or '').lower()
    if re.search(r'not approved|rejected|code\s*c\b|revise (and|&) resubmit', t): return 'rejected'
    if re.search(r'approved (with|as|per|subject to) (the )?(comments|noted)|approved as noted|code\s*b\b', t): return 'approved_comments'
    if re.search(r'\bapproved\b|code\s*a\b', t): return 'approved'
    return None


def fields(code, t, head, src_name=''):
    f = {}
    if code == 'SD':
        f.update(sd_info(t, src_name))
    if code == 'PROD':                                        # فترة التقرير (من / إلى) – الشهري بيروح MONTHLY (15-11)
        ds = sorted(set(re.findall(r'20\d{2}-\d{2}-\d{2}', t)))
        if ds: f['d0'], f['d1'] = ds[0], ds[-1]
    if code == 'UNF_WIR':
        m = re.search(rf'Civil{SP}(\d{{5}})', t) or re.search(rf'AP3500{SP}Water{SP}.*?(\d{{5}})', t)
        if m: f['no'] = int(m.group(1))
        m = re.search(r'Request for [^\n]*?-\s*(\d{2,3})\s*(?:\n|$)', t)
        if m: f['site_no'] = int(m.group(1))
    elif code == 'SITE_WIR':
        m = re.search(rf'PMF{SP}010{SP}QUA{SP}WIR\W{{0,3}}(\d{{1,3}})', t)
        if m: f['no'] = int(m.group(1))
    elif code.startswith('OUT_'):
        m = re.search(rf'CTR{SP}(YA|CLIENT|SUP|SUB|OTH){SP}WA{SP}(26{SP})?([\d\s]{{2,9}})', t)
        if m: f['no'] = _num(m.group(3)); f['to'] = m.group(1); f['yr'] = bool(m.group(2))
    elif code == 'IN_YA':
        m = re.search(rf'cons{SP}26{SP}(\d+)', t, re.I)
        if m: f['no'] = int(m.group(1))
    elif code == 'SUB_CONTRACT':                              # رقم العقد زي ما هو مكتوب (2610/0424)
        m = re.search(r'عقد\s*(?:رقم|No\.?)\s*[:：(]?\s*(\d[\d/\-]{2,15}\d)', t) or \
            re.search(r'Contract\s*No\.?\s*[:：]?\s*([A-Z0-9][A-Z0-9/\-]{2,15})', t, re.I)
        if m: f['ct'] = m.group(1)
    elif code in SUB_CODES:
        pass
    elif code == 'PAYROLL':                                   # عمال ولا مشغلين، وشهر المسير (المستخلص بيتحدد من الشهر)
        f['kind'] = payroll_kind(t, head)
        mo = cal_month(head)
        if mo: f['month'] = mo
    elif code == 'MIR':
        m = re.search(rf'MIR{SP}(\d{{1,4}})', t)
        if m: f['no'] = int(m.group(1))
    elif code == 'MAR':
        m = re.search(rf'\bM{SP}?(\d{{1,3}})\b', t)
        if m: f['no'] = int(m.group(1))
    elif code == 'RFI':
        m = re.search(rf'RFI{SP}(\d{{1,3}})', t)
        if m: f['no'] = int(m.group(1))
    elif code == 'IPC_SIGNED':                                # رقم المستخلص من الغلاف (والـ DC بيتأكد منه بتاريخ الفترة)
        m = re.search(r'مستخلص\s*(?:ال)?(?:جار[يى]\s*)?رقم\s*[:(\[]?\s*(\d{1,3})(?!\d)', t) or re.search(r'IPC\s*(?:No\.?|#|-)\s*0?(\d{1,3})(?!\d)', t)
        if m: f['no'] = int(m.group(1))
    elif code == 'UPD_CAL':                                   # شهر التقويم من العنوان (مش من «إجمالي الأيام لشهر …»)
        mo = cal_month(t)
        if mo: f['month'] = mo
    elif code == 'DN':                                        # رقم إذن التوريد ورقم الفاتورة زي ما هم
        m = re.search(r'Delivery\s*Note\s*#?\s*[:：]?\s*([A-Z0-9]{2,6}(?:\s*/\s*[A-Z0-9]{2,8}){1,3})', t, re.I) or \
            re.search(r'\bDN\s*No\.?\s*[:：]?\s*([A-Z0-9][A-Z0-9/\-]{3,})', t, re.I)
        if m: f['dn'] = re.sub(r'\s+', '', m.group(1))
        m = re.search(r'\b(INV\s*/\s*\d{4}\s*/\s*\d{3,})', t, re.I)
        if m: f['inv'] = re.sub(r'\s+', '', m.group(1)).upper()
    elif code == 'IPC_INV':                                   # رقم المستخلص + نهاية فترته (لاسم الملف)
        m = re.search(r'رقم\s*المستخلص\s*[:：]?\s*(\d{1,3})(?!\d)', t) or re.search(r'مستخلص\s*رقم\s*[:：]?\s*(\d{1,3})(?!\d)', t)
        if m: f['no'] = int(m.group(1))
        m = re.search(r'فترة\s*تنفيذ\s*المستخلص\s*[:：]?\s*(20\d\d-\d\d-\d\d)', t)
        if m: f['to'] = m.group(1)
    if code in OUT_NAMED:                                     # رقم الصادر للاستشاري اللي على الورقة
        o = find_out(t)
        if o:
            f['out'] = o
            if code == 'OUT_YA': f['no'] = int(o[-4:])       # السجل والاسم نفس الرقم
            if code == 'MAR' and MAR_BY_OUT_FN:              # رقم الـ M من MAR LOG بالرقم ده (أدق من «M» في الكلام)
                try: n = MAR_BY_OUT_FN(o)
                except Exception: n = None
                if n: f['no'] = n
    d = find_date(head) or find_date(t)
    if d: f['date'] = d.isoformat()
    hits = [st for st, p in STATUS if re.search(p, t, re.M if st == 'approved' else re.I | re.M)]
    if len(hits) == 1: f['status'] = hits[0]                  # الختم مطبوع فيه كل الاختيارات ← أكتر من واحد = مش عارف (Claude بيحدد)
    m = re.search(r'Request for [A-Za-z ]+?(?=\s*(?:-|FROM|TO|\n|$))', t)
    if m and 'WIR' in code: f['work'] = ' '.join(m.group(0).split())
    return f


def file_name(code, f, src_name, text=''):
    """الأسامي إنجليزي وبنفس شكل الملفات اللي عندنا في كل فولدر."""
    ext = os.path.splitext(src_name)[1].lower() or '.pdf'
    ext = '.pdf' if ext in ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp') else ext
    no, date = f.get('no'), f.get('date', '')
    t = ' '.join((text or '').split())
    if code == 'SD' and (f.get('kind') or f.get('dwg')):                         # Plan & Profile - CLIENT-WTR-SD-L1-002-2.pdf
        k = {'Chamber': 'Chamber'}.get(f.get('kind'), f.get('kind') or 'Shop Drawing')
        ref = f.get('dwg') or f.get('out') or ''
        title = (f.get('title') or (out_title(f['out']) if f.get('out') and not f.get('dwg') else '') or '').strip()[:100].strip()
        return safe(' - '.join(x for x in (k, ref, title or ('' if ref else date)) if x) or 'Shop Drawing') + ext
    if code in OUT_NAMED and f.get('out'):                                       # CTR-YA-WA-26-0035 - Method Statement ….pdf
        title = (f.get('title') or out_title(f['out']) or '').strip()[:120].strip()
        pre = f'M{no} - ' if code == 'MAR' and no else ''
        return safe(f"{pre}{f['out']}{' - ' + title if title else ''}") + ext
    if code == 'UNF_WIR' and no: return f'{no:05d}{ext}'                       # 00019.pdf
    if code == 'SITE_WIR':                                                       # 021 -  Dis  AL-SHAFA - … (من لوج الطلبات بس)
        title = wir_title(no) if no else None
        return safe(title) + ext if title else ''                                # مفيش رقم / مش في اللوج ← فاضي (wir_note بتقول ليه)
    if code == 'OUT_OTH' and no: return f'{no:03d}{ext}'                         # 002.pdf
    if code.startswith('OUT_') and no: return f"CTR-{f.get('to', 'YA')}-WA-{'26-' if f.get('yr', True) else ''}{no:04d}{ext}"
    if code == 'IN_YA' and no: return f'cons-26-{no}{ext}'
    if code == 'MIR' and no:                                                     # CLIENT-CTR-MIR-003 - الوصف (من MIR LOG)
        title = MIR_FN(no) if MIR_FN else None
        return safe(title) + ext if title else f'CLIENT-CTR-MIR-{no:03d}{ext}'
    if code == 'DN':                                                             # الاسم الأخير بيتحدد في شباك إذن التوريد
        return safe(f"DN {f['dn']}") + ext if f.get('dn') else f"Delivery Note{' - ' + date if date else ''}{ext}"
    if code == 'MAR' and no: return f'M{no}-Material Approval Request{ext}'
    if code == 'RFI' and no: return f'RFI-{no}{ext}'
    if code == 'MS':                                                             # P1-MF-010-QUA-004_03 Method Statement - … - REV 0
        m = re.search(r'(P1\s*-\s*MF\s*-\s*010\s*-\s*QUA\s*-\s*\d+_\d+)\s*(Method Statement\s*-\s*.+?REV\s*\d+)', t, re.I)
        if m: return safe(re.sub(r'\s*-\s*', '-', m.group(1)) + ' ' + m.group(2)) + ext
    if code == 'SD':
        m = re.search(r'CLIENT\s*-\s*WTR\s*-\s*SH\s*-\s*[A-Z]+\s*-\s*\d+', t)
        if m: return re.sub(r'\s+', '', m.group(0)) + ext
    if code == 'SUB':                                                            # ورق المورد بنوعه
        for p, n in ((r'VAT Registration|شهادة تسجيل.{0,20}القيمة المضافة', 'VAT Certificate'),
                     (r'سجل تجاري|السجل التجاري|Commercial Registration', 'Commercial Registration'),
                     (r'العنوان الوطني|عنوان وطني|National Address', 'National Address'),
                     (r'Company Profile|بروفايل|Profile', 'Company Profile'), (r'Pre-?Qualification|اعتماد', 'Previous Approval')):
            if re.search(p, t[:700], re.I) or re.search(p, os.path.basename(src_name or ''), re.I): return n + ext
        return 'Company Documents' + ext                                         # التاريخ اللي جوه البروفايل مش تاريخ الورقة
    if code == 'IPC_INV' and no:                                                 # IPC-01 (31-10-2026) - VN.pdf
        try: to = dt.date.fromisoformat(f.get('to', '')).strftime(' (%d-%m-%Y)')
        except ValueError: to = ''
        return f'IPC-{no:02d}{to} - VN{ext}'
    if code == 'IPC_SIGNED' and no: return f'IPC-{no:02d} - Signed{ext}'           # الـ DC بيكمّل التاريخ من ملف المستخلص
    if code == 'UPD_CAL':                                                        # 2026-11 Company Update Calendar.pdf
        mo = f.get('month') or cal_month(text, src_name)
        return f"{mo + ' ' if mo else ''}Company Update Calendar{ext}"
    if code == 'SUB_CONTRACT':                                                   # CT-2610-0424 - <Sub> - <المرحلة> (التاريخ)
        return f"CT-{ct_no(f.get('ct')) or 'XXXX'} - {f.get('sub_en') or 'Subcontract'}{' - ' + f['stage'] if f.get('stage') else ''}" \
               f"{' (' + date + ')' if date else ''}{ext}"
    if code == 'GIS':                                                            # IPC-01 - GIS.zip
        return f"{'IPC-%02d - ' % f['ipc'] if f.get('ipc') else ''}GIS{os.path.splitext(src_name)[1].lower() or '.zip'}"
    if code == 'PAYROLL':                                                        # IPC-02 - Payroll - Workers 10-2026.pdf
        mo = f.get('month') or (date[:7] if date else '')
        mo = f'{mo[5:7]}-{mo[:4]}' if len(mo) == 7 else ''
        m = re.search(r'(?<!\d)(0?[1-9]|1[0-2])[-_. ](20\d\d)(?!\d)', os.path.basename(src_name or ''))
        if m and not f.get('month'): mo = f'{int(m.group(1)):02d}-{m.group(2)}'        # «العاملين 10-2026.pdf»
        pre = f"IPC-{f['ipc']:02d} - " if f.get('ipc') else ''
        return f"{pre}Payroll - {f.get('kind') or 'Workers'}{' ' + mo if mo else ''}{ext}"
    if code == 'PRICE':                                                          # PA-01 - تحليل اسعار.xlsx
        stem = re.sub(r'^PA-\d+\s*-\s*', '', os.path.splitext(os.path.basename(src_name or ''))[0]).strip() or 'Price Analysis'
        n = f.get('no')
        if not n and PA_NEXT_FN:
            try: n = PA_NEXT_FN()
            except Exception: n = None
        return safe(f"PA-{int(n or 1):02d} - {stem}") + ext
    if code == 'SUB_IPC': return f"Subcontractor IPC{' - ' + date if date else ''}{ext}"
    if code == 'PROD':                                                           # 2026-11-07 - Production Report.pdf
        ds = sorted(set(re.findall(r'20\d{2}-\d{2}-\d{2}', t)))
        end = ds[-1] if ds else (date or '')
        if prod_monthly(f):                                                      # 2026-10 - Monthly Production Report.pdf
            return safe(f"{f['d1'][:7]} - Monthly Production {'by Executor' if 'منفذ' in t else 'Report'}") + ext
        return safe(f"{end} - Production {'by Executor' if 'منفذ' in t else 'Report'}".strip(' -')) + ext
    if code == 'QUOTE':                                                          # Quotation - RQ-MET-26-240 - 2026-11-08.pdf
        m = re.search(r'\b(?:RQ|QT|QUO)-[A-Z]{1,6}-\d{2}-\d{2,6}\b', t[:3000]) or \
            re.search(r'(?:Quotation|Offer|Ref(?:erence)?)\s*(?:No\.?|#)\s*[:：]?\s*([A-Z0-9/\-]{0,12}\d[A-Z0-9/\-]{0,12})(?![A-Za-z0-9])', t[:3000], re.I)
        ref = re.sub(r'[/ ]', '-', (m.group(m.lastindex or 0)) if m else '').strip('-')
        if not ref:
            m = re.search(r'\b(RQ|QT)-[A-Z]+-\d{2}-\d+', os.path.basename(src_name or ''), re.I)
            ref = m.group(0).upper() if m else ''
        if not ref and no: ref = f'No. {no}'                                     # رقم العرض من قراية Claude
        return safe('Quotation' + (' - ' + ref if ref else '') + (' - ' + date if date else '')) + ext
    if code == 'IN_CLIENT': return f"CLIENT{' - ' + date if date else ''}{ext}"
    return f"{EN[code]}{' - ' + date if date else ''}{ext}"


_WIR = {'mt': None, 'map': {}, 'err': ''}
PA_NEXT_FN = None               # dc_core: رقم تحليل الأسعار الجاي (PA-01، PA-02…) من فولدر PRICE_ANALYSIS
ROOT_FN = None                  # dc_core بيحط هنا دالة فولدر المشروع (من غير ما rules تستورد core)
MIR_FN = None                   # dc_core: اسم الـ MIR من MIR LOG (قسم Materials)
NEXT_IPC_FN = None              # dc_core: رقم المستخلص اللي لسه ماترفعش (من 11-LOG IPC)
IPC_FOR_MONTH_FN = None         # dc_core: المستخلص اللي فترته فيها الشهر ده (YYYY-MM)
LATEST_IPC_FN = None            # dc_core: آخر مستخلص اتعمل
OUT_TITLE_FN = None             # dc_core: الاسم الإنجليزي لرقم الصادر (dc_titles.json أو شيت الـ DC)
MAR_BY_OUT_FN = None            # dc_core: رقم الـ M من MAR LOG برقم الصادر
# الورق اللي بيتقدّم للاستشاري برقم صادر: اسمه = رقم الصادر + الوصف بالإنجليزي (والـ MAR: M14 - رقم الصادر - المادة)
OUT_NAMED = ('MS', 'ITP', 'SD', 'MAR', 'SUB', 'TEST', 'CAL', 'RFI', 'NCR', 'OUT_YA')
OUT_RX = re.compile(r'CTR\W{0,3}YA\W{0,3}WA\W{0,3}26\W{0,3}(\d{4})(?!\d)', re.I)


def find_out(t):
    """رقم الصادر للاستشاري اللي على الورقة (CTR-YA-WA-26-0035) – من غير «رقم التقديم السابق»."""
    t = t or ''
    ms = list(OUT_RX.finditer(t))
    for i, m in enumerate(ms):
        # أقرب عنوان للرقم على نفس السطر (قبله أو بعده – السطر العربي بيتقري بالعكس): «السابق / Previous» ← مش رقمنا
        a = max(t.rfind('\n', 0, m.start()) + 1, m.start() - 60, ms[i - 1].end() if i else 0)
        e = t.find('\n', m.end()); e = min(e if e >= 0 else len(t), m.end() + 60, ms[i + 1].start() if i + 1 < len(ms) else len(t))
        best = None                                           # (المسافة، سابق؟)
        for seg, off, before in ((t[a:m.start()], a, True), (t[m.end():e], m.end(), False)):
            for x in re.finditer(r'(?P<p>Previous(?:\s*Submittal\s*(?:Ref\.?\s*)?No\.?)?|رقم\s*التقديم\s*السابق|السابق)|(?P<o>رقم\s*التقديم(?!\s*السابق)|Submittal\s*(?:Ref\.?\s*)?No|Serial\s*No)',
                                 seg, re.I):
                d = (m.start() - (off + x.end())) if before else x.start() + 5      # العنوان اللي قبل الرقم أولى
                if best is None or d < best[0]: best = (d, bool(x.group('p')))
        if best and best[1]: continue
        return f'CTR-YA-WA-26-{m.group(1)}'
    return None


def out_title(out):
    if not out or not OUT_TITLE_FN: return None
    try: return OUT_TITLE_FN(out)
    except Exception: return None


def out_note(code, f):
    """ورقة برقم صادر ملوش اسم إنجليزي ← Check، وتكتب الاسم بعد الرقم مرة والبرنامج يفتكره."""
    if code == 'SD':
        n = sd_note(f)
        if n: return n
        if f.get('dwg'): return None                          # الاسم من كود اللوحة – مش محتاج اسم الصادر
    if code in OUT_NAMED and f.get('out') and not (f.get('title') or out_title(f['out'])):
        return f"{f['out']} has no English title yet – type it after the number in the name (it will be remembered)"
    return None
IPC_DIR_FN = None               # dc_core: فولدر المستخلص ده نسبةً لفولدر المشروع (CLIENT_IPC\2-IPC)


def month_from_name(src_name):
    """«العاملين 10-2026.pdf» ← 2026-10"""
    m = re.search(r'(?<!\d)(0?[1-9]|1[0-2])[-_. ](20\d\d)(?!\d)', os.path.basename(src_name or ''))
    return f'{m.group(2)}-{int(m.group(1)):02d}' if m else None


def set_ipc(code, f, src_name=''):
    """المسير ← المستخلص اللي فترته فيها شهره (وإلا اللي لسه ماترفعش). الـ GIS ← آخر مستخلص."""
    try:
        if code == 'PAYROLL':
            mo = f.get('month') or month_from_name(src_name) or (f.get('date') or '')[:7]
            if mo and not f.get('month'): f['month'] = mo
            n = None
            try: n = IPC_FOR_MONTH_FN(mo) if (mo and IPC_FOR_MONTH_FN) else None
            except Exception: n = None
            n = n or (NEXT_IPC_FN() if NEXT_IPC_FN else None)
            if n: f['ipc'] = int(n)
        elif code == 'GIS' and LATEST_IPC_FN:
            n = LATEST_IPC_FN()
            if n: f['ipc'] = int(n)
    except Exception:
        pass


def classify_name(src_name):
    """ملف من غير كلام (zip …): من اسمه بس. GIS ← فولدر المستخلص."""
    n = os.path.basename(src_name or '')
    if re.search(r'تحليل\s*(ال)?[اأإ]سعار|Price\s*Analysis|Rate\s*Analysis', n, re.I):      # تحليل أسعار (إكسيل) ← COST_CONTROL
        f = {}
        return [dict(code='PRICE', label=TYPES['PRICE'][0], score=60, fields=f, folder=type_folder('PRICE', f),
                     name=file_name('PRICE', f, n), why=['file name'], sure=True)]
    if not re.search(r'(?<![A-Za-z])GIS(?![A-Za-z])|as[\s_-]*built|أز\s*بلت', n, re.I): return []
    f = {}
    set_ipc('GIS', f, n)
    return [dict(code='GIS', label=TYPES['GIS'][0], score=60, fields=f, folder=type_folder('GIS', f),
                 name=file_name('GIS', f, n), why=['file name'], sure=bool(f.get('ipc')))]


def ct_no(no):
    """رقم العقد في اسم الملف: 2610/0424 ← 2610-0424"""
    return re.sub(r'[\\/:*?"<>|\s]+', '-', str(no or '').strip()).strip('-')


def payroll_kind(t, head=''):
    """مسير العمال ولا المشغلين (السواقين ومشغلين المعدات)"""
    op = r'مشغل|المشغلين|سائق|السائقين|Operators?|Drivers?'
    wk = r'(?<![ء-ي])(عمال|العمال|العاملين|العمالة)(?![ء-ي])|Workers|Labou?r'      # «أعمال» مش عمال
    if re.search(op, head or '', re.I): return 'Operators'
    for x in (head, t):
        a, b = re.search(op, x or '', re.I), re.search(wk, x or '', re.I)
        if a and (not b or a.start() < b.start()): return 'Operators'
        if b: return 'Workers'
    return 'Workers'
WIR_LOG = '01-LOG WIR.xlsm'


def wir_title(no):
    """اسم طلب الموقع زي ما هو في لوج الطلبات (عمود S NAM). لو اللوج ماتقراش السبب بيتحفظ في _WIR['err']."""
    try:
        if ROOT_FN is None: return None
        src = os.path.join(ROOT_FN(), '00_Registers', WIR_LOG)
        mt = os.path.getmtime(src)
        if _WIR['mt'] != mt:
            import openpyxl, shutil, tempfile
            fd, tmp = tempfile.mkstemp(prefix='dc_wir_', suffix='.xlsm'); os.close(fd); shutil.copy2(src, tmp)
            try:
                wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True); ws = wb['SW']
                hdr = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=2, max_row=2, max_col=40, values_only=True))]
                fc = hdr.index('S NAM') if 'S NAM' in hdr else 15
                m = {}
                for r in ws.iter_rows(min_row=4, max_col=fc + 1, values_only=True):
                    try: m[int(r[1])] = str(r[fc]).strip()
                    except (TypeError, ValueError): pass
                wb.close(); _WIR.update(mt=mt, map=m)
            finally:
                try: os.remove(tmp)
                except OSError: pass
        _WIR['err'] = ''
        return _WIR['map'].get(int(no))
    except Exception as e:
        _WIR['err'] = f'{type(e).__name__}: {e}'[:200]
        return None


def wir_note(f):
    """ليه اسم طلب الموقع فاضي – بيظهر في شاشة الـ DC."""
    no = f.get('no')
    if not no: return 'WIR No. not found – type it in «No.» and the name comes from the WIR log'
    if _WIR.get('err'): return f'Could not read {WIR_LOG}: {_WIR["err"]}'
    return f'WIR {no} is not in {WIR_LOG} (column S NAM) – add it there, or type the name'


def safe(n):
    return re.sub(r'[\\/:*?"<>|]+', '-', n).strip()


def prod_monthly(f):
    """تقرير الإنتاج فترته شهر (25 يوم أو أكتر)؟"""
    try: return (dt.date.fromisoformat(f['d1']) - dt.date.fromisoformat(f['d0'])).days >= 24
    except Exception: return False


def type_folder(code, f):
    """فولدر النوع: {st} = حالة الاعتماد، {ipc} = فولدر المستخلص (1-IPC…)، {sub} بيتحدد مع المقاول."""
    if code in ('PAYROLL', 'GIS'):                            # فولدر المستخلص بتاعه (CLIENT_IPC\IPC-1 …)
        n = f.get('ipc')
        if not n: return TYPES[code][1].replace('\\{ipc}', '')
        if IPC_DIR_FN:
            try: return IPC_DIR_FN(n)
            except Exception: pass
        return TYPES[code][1].replace('{ipc}', f'{n}-IPC')
    if code == 'PROD' and prod_monthly(f):                    # تقرير شهري ← REPORTS_CTR\MONTHLY\2026-10
        return r'07_Planning_Progress\REPORTS_CTR\MONTHLY' + '\\' + f['d1'][:7]
    if code == 'SD':                                          # جوه SD: الخط أو DETAILS / CHAMBER_MECH – مش معروف ← فاضي وإنت تختار
        g = sd_group(f)
        return TYPES[code][1].replace('{sdg}', g).replace('{st}', status_folder(f, code)) if g else ''
    return (TYPES[code][1].replace('{st}', status_folder(f, code)).replace('\\{sub}', '')
            .replace('{ipc}', f"{f['no']}-IPC" if code in ('IPC_INV', 'IPC_SIGNED') and f.get('no') else '').rstrip('\\'))


def status_folder(f, code=''):
    if code == 'SITE_WIR': return '3_PUBLISHED'               # نسخ الموقع المعتمدة
    if code == 'MIR': return '2_SHARED'                       # الـ MIR رجع متوقّع ← مستني الرفع على اليونيفاير والاعتماد
    if code == 'SUB': return '1_WIP'
    if f.get('status') in ('approved', 'approved_comments'): return '3_PUBLISHED'
    if code in OUT_NAMED and not f.get('out') and not f.get('no') and not f.get('status'):
        return '1_WIP'                                        # لسه مالوش رقم صادر = نسخة شغل (ماتقدّمتش)
    return '2_SHARED'


# ---------------------------------------------------------------- التعلّم
def _tokens(t):
    return set(w for w in re.findall(r'[A-Za-z]{4,}|[ء-ي]{4,}', t[:1500]))


def load_learn():
    from dc_io import load_json
    L = load_json(LEARN, None)
    return L if isinstance(L, dict) and 'examples' in L else {'examples': [], 'stats': {}}


def learn(text, guessed, actual, path=LEARN):
    """إنت أكدت أو صحّحت: نحفظ كلمات الورقة ونوعها الصح، ونعدّ هو كان صح ولا غلط."""
    L = load_learn()
    L['examples'] = (L['examples'] + [{'type': actual, 'tokens': sorted(_tokens(text))[:120]}])[-400:]
    s = L['stats'].setdefault(actual, {'ok': 0, 'fixed': 0})
    s['ok' if guessed == actual else 'fixed'] += 1
    from dc_io import save_json
    save_json(path, L)


def _learned_scores(text):
    L = load_learn(); tk = _tokens(text); sc = {}
    if not tk: return sc
    for ex in L.get('examples', []):
        e = set(ex['tokens'])
        if not e: continue
        j = len(tk & e) / len(tk | e)
        if j > 0.25: sc[ex['type']] = max(sc.get(ex['type'], 0), int(j * 60))
    return sc


# ---------------------------------------------------------------- التصنيف
def classify(text, src_name='', scanned=False):
    """يرجّع قايمة اقتراحات مترتبة: dict(code, label, score, sure, fields, folder, name, why)."""
    t = text or ''
    head = t[:600]
    sc, why = {}, {}
    for code, p, w in RULES:
        if re.search(p, t, re.I):
            sc[code] = sc.get(code, 0) + w; why.setdefault(code, []).append(p.split('|')[0][:30])
    # الخطاب: لو جواب المقاول في الراس وفيه رقم صادر = صادر. لو راس الاستشاري ومفيش رقم صادر = وارد
    ours = (re.search(COMPANY, head, re.I) or '0000000001' in head) and not re.search(r'CONSULTANT|الاستشاري', t[:300], re.I)
    if not ours:
        for c in [c for c in sc if c.startswith('OUT_')]: sc[c] -= 25          # رقمنا مذكور كمرجع بس
    if ours and 'IN_YA' in sc: sc['IN_YA'] -= 20
    # كلمة Method Statement جوه فورم الطلب مش معناها إن الورقة MS
    if any(c in sc and sc[c] >= 50 for c in ('UNF_WIR', 'SITE_WIR')):
        for c in ('MS', 'ITP', 'IN_CLIENT', 'IN_YA', 'TEST'): sc.pop(c, None)
    if sc.get('UNF_WIR', 0) >= 50: sc.pop('SITE_WIR', None)
    if 'SUB_IPC' in sc: sc.pop('IPC', None)
    # العقد والمسير: بالوزن الكامل لو العنوان في أول الورقة بس – خطاب أو محضر أو مستخلص بيذكرهم مش هو العقد/المسير
    letter = any(sc.get(c, 0) >= 20 for c in sc if c.startswith(('OUT_', 'IN_')))
    if 'SUB_CONTRACT' in sc:
        if 'SUB_IPC' in sc or re.search(r'مستخلص', t) or any(sc.get(c, 0) >= 45 for c in sc if c.startswith('OUT_')):
            sc.pop('SUB_CONTRACT')
        elif not re.search(CONTRACT_HEAD, head, re.I): sc['SUB_CONTRACT'] -= 35
    if 'PAYROLL' in sc:
        if re.search(r'محضر|الموضوع|Subject\s*:|Minutes', head, re.I) or (letter and not re.search(PAYROLL_HEAD, t[:100], re.I)):
            sc.pop('PAYROLL')                                 # خطاب أو محضر بيطلب المسير – مش المسير نفسه
        elif not re.search(PAYROLL_HEAD, head, re.I): sc['PAYROLL'] = 15
    # إذن توريد (ومعاه فاتورة المورد): مش مستخلص ولا فاتورة مستخلص – والعميل اللي فيه المقاول
    if sc.get('DN', 0) >= 45 and not re.search(r'مستخلص', t):
        for c in ('IPC_INV', 'SUB_IPC', 'IPC', 'IN_SUB', 'IN_OTH'): sc.pop(c, None)
        if re.search(COMPANY, t, re.I) or '000000000210003' in t: sc['DN'] += 20
    # فاتورة ضريبية فيها «مستخلص» والجهة المالكة ورقم أمر الشراء = فاتورة مستخلصنا (اعتماده)
    if 'IPC_INV' in sc and 'مستخلص' in t and re.search(r'الجهة المالكة', t) and '000000000' in t:
        sc['IPC_INV'] += 35; sc.pop('IPC', None); sc.pop('SUB_IPC', None)
    if ours and 'SUB' in sc and not re.search(r'كمقاول باطن|اعتماد مقاول', t): sc['SUB'] -= 20      # خطابنا مش ورق المورد
    # المستخلص لو جاي سكان (أو اسمه Signed/موقع) = النسخة الموقعة – بس لو المستخلص هو أعلى حاجة (مش خطاب عن المستخلص)
    if 'IPC' in sc and sc['IPC'] >= max(sc.values()) and (scanned or re.search(r'signed|موق[ّ]?ع', src_name or '', re.I)):
        sc['IPC_SIGNED'] = sc.pop('IPC') + 30; why['IPC_SIGNED'] = why.pop('IPC', []) + ['scan']
    for c, v in _learned_scores(t).items():
        sc[c] = sc.get(c, 0) + v; why.setdefault(c, []).append('learned')
    out = []
    for code, s in sorted(sc.items(), key=lambda kv: -kv[1]):
        if s <= 0: continue
        f = fields(code, t, head, src_name)
        if code == 'PAYROLL': set_ipc(code, f, src_name)
        if 'WIR' in code or code in ('IPC_INV', 'IPC_SIGNED', 'UPD_CAL', 'MIR', 'DN'): f.pop('status', None)   # الفورم فيه كلمة Approved/Rejected كاختيارات
        folder = type_folder(code, f)
        o = dict(code=code, label=TYPES[code][0], score=s, fields=f, folder=folder,
                 name=file_name(code, f, src_name, t), why=why.get(code, []))
        if code == 'SITE_WIR' and not o['name']: o['note'] = wir_note(f)    # مفيش اسم من اللوج ← مش متأكد
        if out_note(code, f): o['note'] = out_note(code, f)                  # رقم صادر ملوش اسم إنجليزي ← Check
        out.append(o)
    for i, o in enumerate(out):
        nxt = out[i + 1]['score'] if i + 1 < len(out) else 0
        o['sure'] = i == 0 and o['score'] >= 45 and o['score'] - nxt >= 20 and not o.get('note')
    return out
