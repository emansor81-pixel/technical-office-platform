# -*- coding: utf-8 -*-
"""المستخلص: ياخد نسخة من التمبلت الثابت ويملاها من الحصر، ويحط وراها نفس شيتات البرنامج، ويسجّل في 11-LOG IPC."""
import os, re, sys, copy, glob, shutil, tempfile, datetime as dt
import openpyxl
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'qs'))
import qs_core  # noqa: E402
from ipc_template import COVER, IPC, HELP, QSUM, CODE_COL, DIVISIONS, VAT, add_cover_logos  # noqa: E402

TEMPLATE_REL = r'00_Registers\11.1-TEMPLATE IPC.xlsx'
OUT_REL = r'08_Cost_Payments\CLIENT_IPC'
LOG_REL = r'00_Registers\11-LOG IPC.xlsx'
WIR_PDF_REL = r'05_Quality\WIR\5_UNIFIER_APPROVED'   # الطلبات المعتمدة من اليونيفاير – بتتلزق ورا الحصر الكلي في الـ PDF
WIR_PAPER_REL = r'05_Quality\WIR\3_PUBLISHED'
WIR_MERGE_REL = r'05_Quality\WIR\6_MERG_UN_SIT_WIR'   # كل طلب (يونيفاير + ورقي) في ملف، ومتجمعين zip لكل نوع مع كل مستخلص         # الطلبات الورقية المعتمدة – كل واحد بيتلزق ورا نسخة اليونيفاير بتاعته
QS_SHEET = 'QS-TOTAL'

PS_PDF = r'''$ErrorActionPreference = 'Stop'
$src = $env:IPC_XLSX
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int p);'
  $xpid = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xpid)
  if ($env:IPC_PIDFILE) { Set-Content -Path $env:IPC_PIDFILE -Value $xpid }
} catch {}
try {
  $wb = $xl.Workbooks.Open($src, 0, $true)
  $wb.Application.Calculate()
  foreach ($job in ($env:IPC_JOBS -split ';')) {
    $keep = ($job -split '\|')[0] -split ','; $out = ($job -split '\|')[1]
    foreach ($k in $keep) { $wb.Worksheets.Item([int]$k).Visible = -1 }
    for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
      if ($keep -notcontains [string]$i) { $wb.Worksheets.Item($i).Visible = 0 }
    }
    $wb.ExportAsFixedFormat(0, $out, 0, $true, $false)
  }
  $wb.Close($false)
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
'''


def export_pdf(xlsx, sheets, timeout=240, pdf=None):
    """PDF واحد من شيتات معيّنة. pdf = اسم الملف (الافتراضي نفس اسم الإكسيل)."""
    pdf = pdf or os.path.splitext(xlsx)[0] + '.pdf'
    return export_pdfs(xlsx, [(sheets, pdf)], timeout)[0]


def export_pdfs(xlsx, jobs, timeout=240):
    """كذا PDF من نفس ملف الإكسيل في فتحة واحدة لإكسيل: jobs = [(شيتات, مسار الـ PDF), ...].
    بيطبع في فولدر مؤقت الأول وبعدين بينقل الملفات (عشان لو PDF مفتوح في Foxit أو المعاينة ما يعلّقش)،
    ولو إكسيل خد أكتر من timeout ثانية بيتقفل لوحده."""
    import subprocess
    if os.name != 'nt': raise RuntimeError('PDF export works on Windows only (needs Excel)')
    names = openpyxl.load_workbook(xlsx, read_only=True).sheetnames
    for _, pdf in jobs:
        if os.path.exists(pdf):
            try:
                with open(pdf, 'ab'): pass
            except PermissionError:
                raise RuntimeError(f'The PDF is open (Foxit or Explorer preview) – close it and try again:\n{os.path.basename(pdf)}')
    tmpd = tempfile.mkdtemp()
    pidf = os.path.join(tmpd, 'xl.pid')
    spec, tmps = [], []
    for i, (sheets, pdf) in enumerate(jobs):
        keep = [str(names.index(n) + 1) for n in sheets if n in names]
        t = os.path.join(tmpd, f'out{i}.pdf'); tmps.append(t)
        spec.append(','.join(keep) + '|' + t)
    ps = os.path.join(tmpd, 'ipc_pdf.ps1')
    with open(ps, 'w', encoding='utf-8-sig') as f: f.write(PS_PDF)
    env = dict(os.environ, IPC_XLSX=os.path.abspath(xlsx), IPC_JOBS=';'.join(spec), IPC_PIDFILE=pidf)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps], env=env,
                           capture_output=True, text=True, timeout=timeout, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except subprocess.TimeoutExpired:
        try:
            xp = int(open(pidf).read().strip())
            subprocess.run(['taskkill', '/F', '/PID', str(xp)], capture_output=True, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        except Exception:
            pass
        shutil.rmtree(tmpd, ignore_errors=True)
        raise RuntimeError(f'Excel took more than {timeout // 60} minutes to export the PDF and was stopped. The Excel file is fine – try again or print it from Excel.')
    if r.returncode != 0 or not all(os.path.exists(t) for t in tmps):
        shutil.rmtree(tmpd, ignore_errors=True)
        raise RuntimeError((r.stderr or r.stdout or 'Excel could not export the PDF').strip()[-400:])
    try:
        for t, (_, pdf) in zip(tmps, jobs): shutil.move(t, pdf)
    except PermissionError as e:
        raise RuntimeError(f'The PDF is open – close it and try again:\n{e.filename}')
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)
    return [pdf for _, pdf in jobs]

# ---------------------------------------------------------------- تفقيط
_ONES = ['', 'واحد', 'اثنان', 'ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة', 'عشرة',
         'أحد عشر', 'اثنا عشر', 'ثلاثة عشر', 'أربعة عشر', 'خمسة عشر', 'ستة عشر', 'سبعة عشر', 'ثمانية عشر', 'تسعة عشر']
_TENS = ['', '', 'عشرون', 'ثلاثون', 'أربعون', 'خمسون', 'ستون', 'سبعون', 'ثمانون', 'تسعون']
_HUND = ['', 'مائة', 'مائتان', 'ثلاثمائة', 'أربعمائة', 'خمسمائة', 'ستمائة', 'سبعمائة', 'ثمانمائة', 'تسعمائة']


def _below_1000(n):
    h, r = divmod(n, 100)
    parts = []
    if h: parts.append(_HUND[h])
    if r:
        if r < 20: parts.append(_ONES[r])
        else:
            t, o = divmod(r, 10)
            parts.append((_ONES[o] + ' و' + _TENS[t]) if o else _TENS[t])
    return ' و'.join(parts)


def _scale(n, one, two, few, many):
    """n = عدد الآلاف/الملايين (1..999). 11–99 في آخره → منصوب (ألفاً)، غير كده مجرور (ثلاثمائة ألف)"""
    if n == 1: return one
    if n == 2: return two
    if 3 <= n <= 10: return f'{_below_1000(n)} {few}'
    last = n % 100
    if 3 <= last <= 10: return f'{_below_1000(n)} {few}'
    return f'{_below_1000(n)} {many}' + ('اً' if last >= 11 else '')


def _noun(n, sing, dual, plural, acc):
    last = n % 100
    if n == 1: return f'{sing}'
    if n == 2: return dual
    if 3 <= last <= 10: return plural
    if last >= 11: return acc
    return sing


def words_int(n):
    if n == 0: return 'صفر'
    parts = []
    b, n = divmod(n, 10 ** 9)
    m, n = divmod(n, 10 ** 6)
    k, n = divmod(n, 1000)
    if b: parts.append(_scale(b, 'مليار', 'ملياران', 'مليارات', 'مليار'))
    if m: parts.append(_scale(m, 'مليون', 'مليونان', 'ملايين', 'مليون'))
    if k: parts.append(_scale(k, 'ألف', 'ألفان', 'آلاف', 'ألف'))
    if n: parts.append(_below_1000(n))
    return ' و'.join(parts)


def tafqit(amount):
    """1234.5 → فقط ألف ومائتان وأربعة وثلاثون ريالاً سعودياً وخمسون هللة لا غير"""
    amount = round(float(amount or 0), 2)
    sar = int(amount); hal = int(round((amount - sar) * 100))
    s = f'فقط {words_int(sar)} ' + _noun(sar, 'ريال سعودي', 'ريالان سعوديان', 'ريالات سعودية', 'ريالاً سعودياً')
    if hal: s += f' و{words_int(hal)} ' + _noun(hal, 'هللة', 'هللتان', 'هللات', 'هللة')
    return s + ' لا غير'


# التفقيط كمعادلة إكسيل (زي الاستشاري – Excel 365/2021): بيتغير لوحده لو الرقم اتغير. E14 = خانة المبلغ
TAFQIT_F = '=_xlfn.LET(_xlpm.xx,E14,_xlpm.vv,ROUND(ABS(_xlpm.xx),2),_xlpm.nn,INT(_xlpm.vv),_xlpm.hl,ROUND((_xlpm.vv-_xlpm.nn)*100,0),_xlpm.bb,INT(_xlpm.nn/10^9),_xlpm.mi,MOD(INT(_xlpm.nn/10^6),1000),_xlpm.th,MOD(INT(_xlpm.nn/1000),1000),_xlpm.oo,MOD(_xlpm.nn,1000),_xlpm.mm,MOD(_xlpm.nn,100),_xlpm.parts,_xlfn.TEXTJOIN(" و",TRUE,IF(_xlpm.bb=0,"",IF(_xlpm.bb=1,"مليار",IF(_xlpm.bb=2,"ملياران",_xlfn.LET(_xlpm.hh,INT(_xlpm.bb/100),_xlpm.rr,MOD(_xlpm.bb,100),_xlpm.uu,MOD(_xlpm.rr,10),_xlpm.hw,CHOOSE(_xlpm.hh+1,"","مائة","مائتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"),_xlpm.uw,CHOOSE(_xlpm.uu+1,"","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة"),_xlpm.tw,CHOOSE(INT(_xlpm.rr/10)+1,"","عشرة","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"),_xlpm.rw,IF(_xlpm.rr=0,"",IF(_xlpm.rr<10,_xlpm.uw,IF(_xlpm.rr=10,"عشرة",IF(_xlpm.rr<20,CHOOSE(_xlpm.uu,"أحد","اثنا","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة")&" عشر",IF(_xlpm.uu=0,_xlpm.tw,_xlpm.uw&" و"&_xlpm.tw))))),IF(AND(_xlpm.hw<>"",_xlpm.rw<>""),_xlpm.hw&" و"&_xlpm.rw,_xlpm.hw&_xlpm.rw))&" "&IF(AND(MOD(_xlpm.bb,100)>=3,MOD(_xlpm.bb,100)<=10),"مليارات",IF(MOD(_xlpm.bb,100)>=11,"ملياراً","مليار"))))),IF(_xlpm.mi=0,"",IF(_xlpm.mi=1,"مليون",IF(_xlpm.mi=2,"مليونان",_xlfn.LET(_xlpm.hh,INT(_xlpm.mi/100),_xlpm.rr,MOD(_xlpm.mi,100),_xlpm.uu,MOD(_xlpm.rr,10),_xlpm.hw,CHOOSE(_xlpm.hh+1,"","مائة","مائتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"),_xlpm.uw,CHOOSE(_xlpm.uu+1,"","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة"),_xlpm.tw,CHOOSE(INT(_xlpm.rr/10)+1,"","عشرة","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"),_xlpm.rw,IF(_xlpm.rr=0,"",IF(_xlpm.rr<10,_xlpm.uw,IF(_xlpm.rr=10,"عشرة",IF(_xlpm.rr<20,CHOOSE(_xlpm.uu,"أحد","اثنا","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة")&" عشر",IF(_xlpm.uu=0,_xlpm.tw,_xlpm.uw&" و"&_xlpm.tw))))),IF(AND(_xlpm.hw<>"",_xlpm.rw<>""),_xlpm.hw&" و"&_xlpm.rw,_xlpm.hw&_xlpm.rw))&" "&IF(AND(MOD(_xlpm.mi,100)>=3,MOD(_xlpm.mi,100)<=10),"ملايين",IF(MOD(_xlpm.mi,100)>=11,"مليوناً","مليون"))))),IF(_xlpm.th=0,"",IF(_xlpm.th=1,"ألف",IF(_xlpm.th=2,"ألفان",_xlfn.LET(_xlpm.hh,INT(_xlpm.th/100),_xlpm.rr,MOD(_xlpm.th,100),_xlpm.uu,MOD(_xlpm.rr,10),_xlpm.hw,CHOOSE(_xlpm.hh+1,"","مائة","مائتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"),_xlpm.uw,CHOOSE(_xlpm.uu+1,"","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة"),_xlpm.tw,CHOOSE(INT(_xlpm.rr/10)+1,"","عشرة","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"),_xlpm.rw,IF(_xlpm.rr=0,"",IF(_xlpm.rr<10,_xlpm.uw,IF(_xlpm.rr=10,"عشرة",IF(_xlpm.rr<20,CHOOSE(_xlpm.uu,"أحد","اثنا","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة")&" عشر",IF(_xlpm.uu=0,_xlpm.tw,_xlpm.uw&" و"&_xlpm.tw))))),IF(AND(_xlpm.hw<>"",_xlpm.rw<>""),_xlpm.hw&" و"&_xlpm.rw,_xlpm.hw&_xlpm.rw))&" "&IF(AND(MOD(_xlpm.th,100)>=3,MOD(_xlpm.th,100)<=10),"آلاف",IF(MOD(_xlpm.th,100)>=11,"ألفاً","ألف"))))),IF(_xlpm.oo=0,"",_xlfn.LET(_xlpm.hh,INT(_xlpm.oo/100),_xlpm.rr,MOD(_xlpm.oo,100),_xlpm.uu,MOD(_xlpm.rr,10),_xlpm.hw,CHOOSE(_xlpm.hh+1,"","مائة","مائتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"),_xlpm.uw,CHOOSE(_xlpm.uu+1,"","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة"),_xlpm.tw,CHOOSE(INT(_xlpm.rr/10)+1,"","عشرة","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"),_xlpm.rw,IF(_xlpm.rr=0,"",IF(_xlpm.rr<10,_xlpm.uw,IF(_xlpm.rr=10,"عشرة",IF(_xlpm.rr<20,CHOOSE(_xlpm.uu,"أحد","اثنا","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة")&" عشر",IF(_xlpm.uu=0,_xlpm.tw,_xlpm.uw&" و"&_xlpm.tw))))),IF(AND(_xlpm.hw<>"",_xlpm.rw<>""),_xlpm.hw&" و"&_xlpm.rw,_xlpm.hw&_xlpm.rw)))),_xlpm.rl,IF(_xlpm.nn=0,IF(_xlpm.hl>0,"","صفر ريال سعودي"),IF(_xlpm.nn=1,"ريال سعودي واحد",IF(_xlpm.nn=2,"ريالان سعوديان",_xlpm.parts&" "&IF(AND(_xlpm.mm>=3,_xlpm.mm<=10),"ريالات سعودية",IF(_xlpm.mm>=11,"ريالاً سعودياً","ريال سعودي"))))),_xlpm.ha,IF(_xlpm.hl=0,"",IF(_xlpm.nn=0,""," و")&IF(_xlpm.hl=1,"هللة واحدة",IF(_xlpm.hl=2,"هللتان",_xlfn.LET(_xlpm.hh,INT(_xlpm.hl/100),_xlpm.rr,MOD(_xlpm.hl,100),_xlpm.uu,MOD(_xlpm.rr,10),_xlpm.hw,CHOOSE(_xlpm.hh+1,"","مائة","مائتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"),_xlpm.uw,CHOOSE(_xlpm.uu+1,"","واحدة","اثنتان","ثلاث","أربع","خمس","ست","سبع","ثماني","تسع"),_xlpm.tw,CHOOSE(INT(_xlpm.rr/10)+1,"","عشرة","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"),_xlpm.rw,IF(_xlpm.rr=0,"",IF(_xlpm.rr<10,_xlpm.uw,IF(_xlpm.rr=10,"عشر",IF(_xlpm.rr<20,CHOOSE(_xlpm.uu,"إحدى","اثنتا","ثلاث","أربع","خمس","ست","سبع","ثماني","تسع")&" عشرة",IF(_xlpm.uu=0,_xlpm.tw,IF(_xlpm.uu=1,"إحدى",_xlpm.uw)&" و"&_xlpm.tw))))),IF(AND(_xlpm.hw<>"",_xlpm.rw<>""),_xlpm.hw&" و"&_xlpm.rw,_xlpm.hw&_xlpm.rw))&" "&IF(AND(_xlpm.hl>=3,_xlpm.hl<=10),"هللات","هللة")))),"فقط "&_xlpm.rl&_xlpm.ha&" لا غير")'


def tafqit_formula(cell):
    return TAFQIT_F.replace('E14', cell)


def plan_box(ws, no):
    """جدول صغير برا الطباعة جنب الملخص (زي الاستشاري): نسبة المستخلصات | نسبة الإنجاز (من الخطة) | الفرق."""
    from openpyxl.styles import Font, Alignment, Border, Side
    pr = next((r for r in range(1, ws.max_row + 1) if 'النسبة المئ' in str(ws.cell(r, 2).value or '')), None)
    if not pr: return
    try:
        import ipc_ho
        plan = ipc_ho.load_settings().get('PLAN', {}).get(str(int(no)))
    except Exception:
        plan = None
    md = Side(style='medium'); B = Border(left=md, right=md, top=md, bottom=md)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for k, t in zip('RST', ('نسبة المستخلصات', 'نسبة الإنجاز', 'الفرق')):
        x = ws[f'{k}13']; x.value = t; x.font = Font(name='Calibri', size=12, bold=True); x.alignment = CEN; x.border = B
    ws['R14'] = f'=M{pr}'
    ws['S14'] = plan / 100 if plan is not None else None
    ws['T14'] = '=IF(S14="","",(S14-R14)/S14)'
    for k in 'RST':
        x = ws[f'{k}14']; x.number_format = '0.00%'; x.font = Font(name='Calibri', size=11); x.alignment = CEN; x.border = B
    ws.column_dimensions['S'].width = max(ws.column_dimensions['S'].width or 0, 20.86)


# ---------------------------------------------------------------- نسخ شيت من ملف لملف
def copy_sheet(src, dst_wb, title=None):
    ws = dst_wb.create_sheet(title or src.title)
    for row in src.iter_rows():
        for c in row:
            if c.__class__.__name__ == 'MergedCell': continue
            n = ws.cell(c.row, c.column, c.value)
            if c.has_style:
                n.font = copy.copy(c.font); n.fill = copy.copy(c.fill); n.border = copy.copy(c.border)
                n.alignment = copy.copy(c.alignment); n.number_format = c.number_format
    for m in src.merged_cells.ranges: ws.merge_cells(str(m))
    for k, d in src.column_dimensions.items():
        ws.column_dimensions[k].width = d.width; ws.column_dimensions[k].hidden = d.hidden
    for k, d in src.row_dimensions.items():
        if d.height: ws.row_dimensions[k].height = d.height
    ws.freeze_panes = src.freeze_panes
    ws.sheet_view.rightToLeft = src.sheet_view.rightToLeft
    try:
        ws.page_setup.orientation = src.page_setup.orientation; ws.page_setup.paperSize = src.page_setup.paperSize
        ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        if src.print_title_rows: ws.print_title_rows = src.print_title_rows
        if src.print_area:
            pa = src.print_area if isinstance(src.print_area, str) else src.print_area[0]
            ws.print_area = pa.split('!')[-1].replace("'", '')
    except Exception:
        pass
    return ws


# ---------------------------------------------------------------- المستخلصات السابقة
def ipc_files(out_dir):
    """{رقم المستخلص: ملف الإكسيل} – في CLIENT_IPC نفسه أو في فولدر جواه (1-IPC، 2-IPC…)."""
    res = {}
    found = glob.glob(os.path.join(out_dir, 'IPC-*.xlsx')) + glob.glob(os.path.join(out_dir, '*', 'IPC-*.xlsx'))
    for p in found:
        m = re.match(r'IPC-(\d+) \(\d{2}-\d{2}-\d{4}\)\.xlsx$', os.path.basename(p))
        if not m: continue                                   # QS-TOTAL ونسخ تانية
        k = int(m.group(1))
        if k not in res or os.path.getmtime(p) > os.path.getmtime(res[k]): res[k] = p
    return res


def ipc_folder(out_dir, no):
    """فين ملفات المستخلص: لو موجود ← نفس فولدره، ولو جديد ← CLIENT_IPC\\{رقم}-IPC (زي 1-IPC)."""
    f = ipc_files(out_dir).get(no)
    return os.path.dirname(f) if f else os.path.join(out_dir, f'{no}-IPC')


# ---------------------------------------------------------------- حالة المستخلص في 11-LOG IPC (اعتماد)
CERTIFIED = ('Certified', 'Paid')


def ipc_status(log=None):
    """{رقم المستخلص: {'status', 'cert', 'invoice'}} من 11-LOG IPC (من نسخة، فبيشتغل حتى لو اللوج مفتوح)."""
    log = log or qs_core.project_path(LOG_REL)
    out = {}
    if not os.path.exists(log): return out
    fd, tmp = tempfile.mkstemp(prefix='ipc_log_', suffix='.xlsx'); os.close(fd)
    try:
        shutil.copy2(log, tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        for r in wb['IPC'].iter_rows(min_row=12, max_col=21, values_only=True):
            r = list(r) + [None] * 21
            try: no = int(r[1])
            except (TypeError, ValueError): continue
            out[no] = dict(status=str(r[14] or '').strip(), cert=r[12], invoice=str(r[20] or '').strip(),
                           todate=r[5], prev=r[6])
        wb.close()
    finally:
        try: os.remove(tmp)
        except OSError: pass
    return out


def mark_certified(no, date=None, invoice=None, log=None, fill_only=False):
    """المستخلص اتعتمد: Status = Certified (Paid بيفضل Paid)، وتاريخ الاعتماد لو اتكتب، ورقم الفاتورة في عمود Invoice No.
    fill_only = يكمّل التاريخ ورقم الفاتورة لو فاضيين بس (مايغيّرش اللي مكتوب)."""
    from copy import copy
    log = log or qs_core.project_path(LOG_REL)
    wb = openpyxl.load_workbook(log)
    ws = wb['IPC']
    r = next((i for i in range(12, ws.max_row + 1) if ws[f'B{i}'].value in (no, str(no))), None)
    if r is None: raise ValueError(f'IPC No. {no} is not in 11-LOG IPC')
    h = ws['U11']
    if not h.value:
        h.value = 'Invoice No.'; src = ws['Q11']
        h.font, h.fill, h.border, h.alignment = copy(src.font), copy(src.fill), copy(src.border), copy(src.alignment)
        ws.column_dimensions['U'].width = 18
    if str(ws[f'O{r}'].value or '').strip() not in CERTIFIED: ws[f'O{r}'] = 'Certified'
    if date and not (fill_only and ws[f'M{r}'].value):
        ws[f'M{r}'] = date; ws[f'M{r}'].number_format = 'dd-mm-yyyy'
    if invoice and not (fill_only and ws[f'U{r}'].value):
        ws[f'U{r}'] = str(invoice).strip(); ws[f'U{r}']._style = copy(ws[f'Q{r}']._style)
    wb.save(log)


# ---------------------------------------------------------------- المستخلص الجاي ← فولدر جاهز للرفع على سيستم الجهة المالكة
DONE_UPLOAD = ('Uploaded', 'Under Review', 'Certified', 'Paid')       # اترفع خلاص
STATUS_LIST = 'Submitted,Uploaded,Under Review,Certified,Paid'
# (الاسم في الشاشة، شكل اسم الملف جوه فولدر المستخلص) – {p} = IPC-02 …
UPLOAD_LIST = [('IPC Excel', r'^{p} \(\d\d-\d\d-\d{{4}}\)\.xlsx$'),
               ('QS-TOTAL PDF', r'^{p}\b.*QS-TOTAL\.pdf$'),
               ('QS-SUMMARY PDF', r'^{p}\b.*QS-SUMMARY\.pdf$'),
               ('WIR PDF', r'^{p}\b.*- WIR\.pdf$'),
               ('WIR zips', r'^{p}\b(?!.*GIS).*\.zip$'),
               ('GIS zip', r'^{p}\b.*GIS.*\.zip$'),
               ('Payroll – Workers', r'^{p} - Payroll - Workers\b.*\.pdf$'),
               ('Payroll – Operators', r'^{p} - Payroll - Operators\b.*\.pdf$')]


def next_ipc(log=None, out_dir=None):
    """المستخلص اللي لسه ماترفعش: أول واحد في 11-LOG IPC حالته فاضية أو Submitted – وإلا الرقم اللي بعد آخر واحد."""
    log = log or qs_core.project_path(LOG_REL)
    if not os.path.isfile(log): raise FileNotFoundError(f'Not found: {log}')     # من غير اللوج كل المستخلصات تبان مارفعتش
    st = ipc_status(log)
    files = ipc_files(out_dir) if out_dir and os.path.isdir(out_dir) else {}
    for no in sorted(set(st) | set(files)):                # مستخلص اتعمل ولسه مش في اللوج = لسه ماترفعش
        if st.get(no, {}).get('status') not in DONE_UPLOAD: return no
    return max(list(st) + list(files) + [0]) + 1


def ipc_periods(log=None):
    """{رقم المستخلص: (من، لحد)} من 11-LOG IPC (عمود C و D) – من نسخة، فبيشتغل حتى لو اللوج مفتوح."""
    log = log or qs_core.project_path(LOG_REL)
    out = {}
    if not os.path.exists(log): return out
    fd, tmp = tempfile.mkstemp(prefix='ipc_log_', suffix='.xlsx'); os.close(fd)
    try:
        shutil.copy2(log, tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        for r in wb['IPC'].iter_rows(min_row=12, max_col=4, values_only=True):
            r = list(r) + [None] * 4
            try: no = int(r[1])
            except (TypeError, ValueError): continue
            a, b = (x.date() if isinstance(x, dt.datetime) else x for x in (r[2], r[3]))
            if isinstance(a, dt.date) and isinstance(b, dt.date): out[no] = (a, b)
        wb.close()
    finally:
        try: os.remove(tmp)
        except OSError: pass
    return out


def ipc_for_month(month, log=None):
    """المستخلص اللي فترته فيها الشهر ده («2026-10») – مسير أكتوبر ← IPC-01 (11-08 لحد 31-10).
    الشهر بعد آخر مستخلص ← المستخلص الجاي. None لو مش عارفين."""
    try:
        y, m = (int(x) for x in str(month)[:7].split('-'))
        m0 = dt.date(y, m, 1); m1 = (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))
    except (ValueError, TypeError):
        return None
    P = ipc_periods(log)
    if not P: return None
    hit = [n for n, (a, b) in P.items() if a <= m1 and b >= m0]
    if hit:
        full = [n for n in hit if P[n][0] <= m1 <= P[n][1]]      # اللي فيه آخر الشهر
        return max(full or hit)
    if m0 > max(b for _, b in P.values()): return max(P) + 1
    return None


def latest_ipc(out_dir=None):
    """آخر مستخلص اتعمل (ملفه) – أو آخر واحد في اللوج."""
    out_dir = out_dir or qs_core.project_path(OUT_REL)
    files = ipc_files(out_dir) if os.path.isdir(out_dir) else {}
    if files: return max(files)
    st = ipc_status()
    return max(st) if st else None


def upload_check(no, out_dir=None):
    """قايمة الرفع للمستخلص ده: [(الاسم، [الملفات اللي لقيناها])] – من فولدره (CLIENT_IPC\\2-IPC …)."""
    out_dir = out_dir or qs_core.project_path(OUT_REL)
    folder = ipc_folder(out_dir, no)
    names = sorted(os.listdir(folder)) if os.path.isdir(folder) else []
    pre = f'IPC-{no:02d}'
    res = []
    for label, pat in UPLOAD_LIST:
        rx = re.compile(pat.format(p=re.escape(pre)), re.I)
        res.append((label, [n for n in names if rx.search(n)]))
    return folder, res


def mark_uploaded(no, date=None, log=None):
    """المستخلص اترفع على سيستم الجهة المالكة ← 11-LOG IPC: Status = Uploaded وتاريخ الرفع في Submitted Date.
    (Under Review / Certified / Paid بيفضلوا زي ما هم.) قايمة الحالات في إكسل بتزيد «Uploaded»."""
    log = log or qs_core.project_path(LOG_REL)
    wb = openpyxl.load_workbook(log)
    ws = wb['IPC']
    r = next((i for i in range(12, ws.max_row + 1) if ws[f'B{i}'].value in (no, str(no))), None)
    if r is None: raise ValueError(f'IPC No. {no} is not in 11-LOG IPC – make the IPC first')
    cur = str(ws[f'O{r}'].value or '').strip()
    if cur in DONE_UPLOAD and cur != 'Uploaded': return False
    ws[f'O{r}'] = 'Uploaded'
    ws[f'E{r}'] = date or dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    ws[f'E{r}'].number_format = 'dd-mm-yyyy'
    for v in ws.data_validations.dataValidation:
        if v.formula1 and 'Submitted' in v.formula1 and 'Uploaded' not in v.formula1:
            v.formula1 = f'"{STATUS_LIST}"'
    wb.save(log)
    return True


# ---------------------------------------------------------------- فاتورة المستخلص (VN) ← اعتماد أوتوماتيك
# لما المستخلص يتعتمد، نظام الجهة المالكة بيطلّع فاتورة ضريبية فيها «مستخلص رقم N». البرنامج يقراها ويعلّم المستخلص Certified.
PO_NO = '000000000'
MONEY = r'(\d{1,3}(?:,\d{3})+\.\d{2}|\d+\.\d{2})'
INV_NAME = re.compile(r'(\bVN\b|invoice|فاتورة)', re.I)


def parse_invoice(text):
    """{'ipc', 'invoice', 'date', 'period_to', 'net', 'vat', 'total', 'ours'} من نص الفاتورة – أو {} لو مش فاتورة مستخلص."""
    import unicodedata
    t = unicodedata.normalize('NFKC', text or '').translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
    if not re.search(r'فاتورة\s*ضريبية', t) or 'مستخلص' not in t: return {}
    out = {'ours': PO_NO in t}
    m = re.search(r'رقم\s*المستخلص\s*[:：]?\s*(\d{1,3})(?!\d)', t) or re.search(r'مستخلص\s*رقم\s*[:：]?\s*(\d{1,3})(?!\d)', t)
    if m: out['ipc'] = int(m.group(1))
    m = re.search(r'(?<!\d)(\d{14})(?!\d)', t)                     # رقم الفاتورة (14 رقم)
    if m: out['invoice'] = m.group(1)
    def day(s):
        try: return dt.date.fromisoformat(s)
        except ValueError: return None
    m = re.search(r'فترة\s*تنفيذ\s*المستخلص\s*[:：]?\s*(20\d\d-\d\d-\d\d)', t)
    if m: out['period_to'] = day(m.group(1))
    m = re.search(r'تاريخ\s*بداية\s*العقد\s*[:：]?\s*(20\d\d-\d\d-\d\d)', t)
    other = {m.group(1) if m else None, out['period_to'].isoformat() if out.get('period_to') else None}
    d = next((x for x in re.findall(r'20\d\d-\d\d-\d\d', t) if x not in other), None)   # تاريخ الفاتورة
    if d: out['date'] = day(d)
    def money(label):
        m = re.search(label + r'[^\d\n]{0,40}?' + MONEY, t)
        return float(m.group(1).replace(',', '')) if m else None
    out['net'], out['vat'], out['total'] = money(r'الصافي'), money(r'ضريبة\s*القيمة\s*المضافة'), money(r'صافي\s*قيمة\s*الفاتورة')
    return out


def read_invoice(path):
    """نص الفاتورة (PDF مكتوب) ← parse_invoice."""
    try:
        import pypdf
        rd = pypdf.PdfReader(path)
        return parse_invoice('\n'.join((p.extract_text() or '') for p in rd.pages[:3]))
    except Exception:
        return {}


def certify_from_invoice(inv, ask=None, log=None):
    """فاتورة مستخلص ← المستخلص يتعلّم Certified (تاريخ الفاتورة + رقمها) لو مبلغها زي اللي قدمناه؛
    ولو مختلف يسأل الأول (ask بترجّع True/False). يرجّع (اتعلّم؟, رسالة)."""
    no = inv.get('ipc') if inv else None
    if not no: return False, 'This is not an IPC invoice (IPC number not found).'
    if not inv.get('ours'): return False, f'The invoice is not for PO {PO_NO}.'
    st = ipc_status(log).get(no)
    if st is None: return False, f'IPC No. {no} is not in 11-LOG IPC.'
    if st['status'] in CERTIFIED:
        if st.get('invoice'): return False, f'IPC No. {no} is already {st["status"]}.'
        d = inv.get('date')                         # اتعتمد من النسخة الموقعة ← نكمّل رقم الفاتورة وتاريخها
        mark_certified(no, dt.datetime.combine(d, dt.time()) if d else None, inv.get('invoice'), log, fill_only=True)
        return True, f"IPC No. {no}: invoice {inv.get('invoice', '')} added ✔"
    try: ours = float(st.get('todate') or 0) - float(st.get('prev') or 0)
    except (TypeError, ValueError): ours = None
    net = inv.get('net')
    if net is None or ours is None or abs(net - ours) > 0.5:
        sar = lambda v: '?' if v is None else f'{v:,.2f}'
        q = (f"Invoice {inv.get('invoice', '?')} – IPC No. {no}\n\n"
             f"Invoice amount (excl. VAT): {sar(net)}\n"
             f"Our IPC amount: {sar(ours)}\n\n"
             f"The amounts are different. Mark IPC No. {no} as Certified anyway?")
        if not (ask and ask(q)): return False, f'IPC No. {no}: the invoice amount is different – not marked Certified.'
    d = inv.get('date')
    mark_certified(no, dt.datetime.combine(d, dt.time()) if d else None, inv.get('invoice'), log)
    return True, f"IPC No. {no} is Certified ✔ (invoice {inv.get('invoice', '')})"


# ---------------------------------------------------------------- نسخة المستخلص الموقعة (سكان)
SIGNED_NAME = re.compile(r'(signed|موق[ّ]?عة?)', re.I)


def ipc_from_text(text, out_dir=None):
    """رقم المستخلص من كلام نسخة متصوّرة: «المستخلص الجاري رقم N» – ولو مش واضح، بتاريخ نهاية الفترة
    (اللي في اسم ملف المستخلص: IPC-01 (31-10-2026)). يرجّع الرقم أو None."""
    import unicodedata
    t = unicodedata.normalize('NFKC', text or '').translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789'))
    ends = {}
    for k, pth in (ipc_files(out_dir) if out_dir and os.path.isdir(out_dir) else {}).items():
        m = re.search(r'\((\d{2})-(\d{2})-(\d{4})\)', os.path.basename(pth))
        if m: ends[k] = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    days = set()
    for a, b, c in re.findall(r'(\d{1,4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,4})', t):
        y, mo, d = (int(a), int(b), int(c)) if len(a) == 4 else (int(c), int(b), int(a))
        try: days.add(dt.date(y, mo, d))
        except ValueError: pass
    by_date = [k for k, e in ends.items() if e in days]
    m = re.search(r'مستخلص\s*(?:ال)?(?:جار[يى]\s*)?رقم\s*[:(\[]?\s*(\d{1,3})(?!\d)', t) or re.search(r'IPC\s*(?:No\.?|#|-)\s*0?(\d{1,3})(?!\d)', t)
    no = int(m.group(1)) if m else None
    if no and (no in ends or not ends): return no
    if len(by_date) == 1: return by_date[0]
    return no


def certify_signed(no, ask=None, log=None, source=''):
    """النسخة الموقعة: البرنامج مش بيشوف التوقيع نفسه ← يسأل الأول، ولو وافقت يتعلّم Certified (من غير تاريخ ولا فاتورة)."""
    st = ipc_status(log).get(no)
    if st is None: return False, f'IPC No. {no} is not in 11-LOG IPC.'
    if st['status'] in CERTIFIED: return False, f'IPC No. {no} is already {st["status"]}.'
    q = (f"Signed copy of IPC No. {no}" + (f":\n{source}" if source else '') +
         f"\n\nThe program can't see the signatures and stamps.\nMark IPC No. {no} as Certified?")
    if not (ask and ask(q)): return False, f'IPC No. {no}: not marked Certified.'
    mark_certified(no, None, None, log)
    return True, f'IPC No. {no} is Certified ✔ (signed copy)'


def scan_invoices(out_dir, ask=None, log=None, skip=()):
    """في فولدرات المستخلصات: فاتورة VN ← Certified (أو تكمّل رقم الفاتورة لو اتعتمد من النسخة الموقعة)،
    ونسخة موقعة (Signed / موقع في الاسم) ← يسأل الأول. يرجّع [(رقم المستخلص, اسم الملف, اتعلّم؟, رسالة)]."""
    res, st = [], ipc_status(log)
    for no, xlsx in sorted(ipc_files(out_dir).items()):
        s = st.get(no, {})
        done = s.get('status') in CERTIFIED
        if done and s.get('invoice'): continue                          # معتمد وفاتورته متسجلة – مفيش جديد
        pdfs = [p for p in sorted(glob.glob(os.path.join(os.path.dirname(xlsx), '*.pdf'))) if (no, os.path.basename(p)) not in skip]
        hit = False
        for pdf in pdfs:                                                 # 1) الفاتورة
            b = os.path.basename(pdf)
            if not INV_NAME.search(b): continue
            inv = read_invoice(pdf)
            if not inv or inv.get('ipc') != no or not inv.get('ours'): continue
            ok, msg = certify_from_invoice(inv, ask=ask, log=log)
            res.append((no, b, ok, msg)); hit = True
            break
        if hit or done: continue
        for pdf in pdfs:                                                 # 2) النسخة الموقعة
            b = os.path.basename(pdf)
            if not SIGNED_NAME.search(b) or INV_NAME.search(b): continue
            ok, msg = certify_signed(no, ask=ask, log=log, source=b)
            res.append((no, b, ok, msg))
            break
    return res


def read_todate(path):
    """{كود البند: الكمية حتى تاريخه} من مستخلص قديم + تاريخ نهاية فترته"""
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb[IPC]
    out = {}
    for r in range(8, ws.max_row + 1):
        code, q = ws[f'{CODE_COL}{r}'].value, ws[f'K{r}'].value
        if isinstance(code, (int, float)) and isinstance(q, (int, float)): out[int(code)] = q
    return out, wb[COVER]['E8'].value


def template_layout(template):
    """بيقرا شكل VO: البنود (بالقسم)، صفوف المبالغ، وصفوف الأقسام."""
    wb = openpyxl.load_workbook(template, read_only=True)
    ws = wb[IPC]
    divnos = {str(n) for n, _ in DIVISIONS}
    names = dict(DIVISIONS)
    items, sum_rows, cur = [], [], None
    for r, row in enumerate(ws.iter_rows(min_row=1, max_col=14, values_only=True), 1):
        A, B = row[0], row[1]
        if r >= 13 and r <= 30 and isinstance(A, str) and len(sum_rows) < 6 and (
                A.startswith('اجمالي قيمة') or A.startswith('قيمة') or A.startswith('ضريبة') or A.startswith('إجمالي المبلغ')):
            sum_rows.append(r)
        if isinstance(A, str) and A.strip() in divnos and isinstance(B, str) and B.startswith(A.strip()):
            cur = int(A.strip())
        code = row[13]
        if isinstance(code, (int, float)):
            items.append(dict(r=r, div=cur, divname=names.get(cur, ''), item=A, desc=B, unit=row[5], qty=row[6] or 0,
                              rate=row[11] or 0, code=int(code)))
    wb.close()
    return items, sum_rows


def quantities(project, no, date_to, out_dir):
    """(الكميات حتى تاريخه {كود: كمية}, السابق {كود: كمية}, ملف المستخلص السابق, تاريخه, items, notes)"""
    old_mode, old_cut = project.mode, project.cutoff
    project.mode = 'approved'; project.cutoff = date_to
    try:
        items, notes = project.build(date_to)
        todate = {}
        for it, code, desc, unit, q, rate, v in project.boq_rows(items):
            todate[int(code)] = todate.get(int(code), 0) + (q or 0)
        todate = {k: round(v, 3) for k, v in todate.items()}
    finally:
        project.mode, project.cutoff = old_mode, old_cut
    prev, prev_path, prev_to = {}, None, None
    files = ipc_files(out_dir)
    older = [k for k in files if k < no]
    if older:
        prev_path = files[max(older)]
        prev, prev_to = read_todate(prev_path)
    return todate, prev, prev_path, prev_to, items, notes


def compute(project, no, date_to, out_dir=None, template=None, date_from=None):
    """كل اللي المستخلص محتاجه: البنود + الأقسام + المبالغ (هو نفسه اللي بيتعرض في البرنامج وبينزل في الإكسيل)."""
    template = template or qs_core.project_path(TEMPLATE_REL)
    out_dir = out_dir or qs_core.project_path(OUT_REL)
    todate, prev, prev_path, prev_to, items, notes = quantities(project, no, date_to, out_dir)
    lay, sum_rows = template_layout(template)
    rows, less = [], []
    for it in lay:
        # المستخلص بيشتغل بدقة «زي المعروض» (خانتين عشريتين) زي ملفات المقاول
        p = round(prev.get(it['code'], 0), 2); t = round(todate.get(it['code'], p), 2)
        if it['code'] in todate and t + 0.005 < p: less.append(it['item'])
        rows.append(dict(it, prev=p, todate=t, period=round(t - p, 2), vprev=round(p * it['rate'], 2),
                         vtodate=round(t * it['rate'], 2), vcontract=it['qty'] * it['rate']))
    for x in rows: x['vperiod'] = round(x['vtodate'] - x['vprev'], 2)
    divs = []
    for n, name in DIVISIONS:
        rr = [x for x in rows if x['div'] == n]
        divs.append(dict(no=n, name=name, contract=sum(x['vcontract'] for x in rr), prev=sum(x['vprev'] for x in rr),
                         todate=sum(x['vtodate'] for x in rr), period=sum(x['vperiod'] for x in rr)))
    contract = sum(d['contract'] for d in divs); tod = round(sum(d['todate'] for d in divs), 2)
    prv = round(sum(d['prev'] for d in divs), 2); per = round(tod - prv, 2); vat = round(per * VAT, 2)
    amounts = [('اجمالي قيمة المشروع', contract), ('قيمة الأعمال المنفذة من العقد حتى تاريخه', tod),
               ('قيمة الأعمال المنفذة سابقا', prv), ('قيمة الاعمال خلال الفترة', per),
               ('ضريبة القيمة المضافة 15 %', vat), ('إجمالي المبلغ المستحق للمقاول في هذه الدفعة بعد اضافة الضريبة 15 %', round(per + vat, 2))]
    wirs = wir_by_code(project, items, date_from if prev_path else None, date_to)   # أول مستخلص = كل اللي قبله
    for x in rows: x['wirs'] = ', '.join(wirs.get(x['code'], []))
    return dict(rows=rows, divs=divs, amounts=amounts, wirs=wirs, sum_rows=sum_rows, contract=contract, todate=tod, prev=prv, period=per,
                vat=vat, total=round(per + vat, 2), pct=(tod / contract if contract else 0), prev_file=prev_path, prev_to=prev_to,
                less_than_prev=less, items=items, notes=notes)


def wir_by_code(project, items, d_from, d_to):
    """{كود البند: [أرقام الطلبات المعتمدة في الفترة]} – الحفر بالفئة، الفرشة، التمديد (+الاختبار)، الردم، قاعدة الطريق"""
    def inper(d):
        if not d: return False
        if hasattr(d, 'date') and not isinstance(d, dt.date): d = d.date()
        if isinstance(d, dt.datetime): d = d.date()
        return (d_from is None or d >= d_from) and d <= d_to
    appd = {qs_core.wno(q['no']): q['appdate'] for q in project.reqs_all}
    out = {}
    def add(code, no):
        no = qs_core.wno(no)
        if code and no and inper(appd.get(no)):
            lst = out.setdefault(int(code), [])
            if no not in lst: lst.append(no)
    bands = {int(b[0]): b[2] for b in project.boq}
    for r in items.get('exc', []):
        if r.get('status') != 'Approved': continue
        for i, v in enumerate(r.get('v') or []):
            if v: add(bands.get(i + 1), r['q']['no'])
    bed = qs_core.PO_ITEMS['bed'][1]
    for r in items.get('bed', []):
        if r.get('status') == 'Approved': add(bed, r['q']['no'])
    for r in items.get('pip', []):
        if r.get('status') != 'Approved': continue
        it = qs_core.PO_ITEMS.get(int(r['dia'])) if r.get('dia') else None
        if it:
            add(it[1], r['q']['no'])
            for x in (r.get('test'), r.get('flush')):
                for n in str(x or '').split(','): add(it[1], n.strip())
    for key, lst in (('bf', items.get('bf', [])), ('rb', items.get('rb', []))):
        code = qs_core.PO_ITEMS[key][1]
        for r in lst:
            for n in str(r.get('wirs') or '').split(','): add(code, n.strip())
    # الأسفلت: ASP-1 ← 32.4.1 ومعاه البرايم كوت MC | الكشط ASP-2-ML ← 32.4.2 ومعاه RC وطلب الطبقة السطحية
    # الطلبات اللي معاه بتنزل في نفس مستخلص طلبه الأساسي (لو معتمدة)، ولو اتعتمدت بعده تنزل في مستخلص اعتمادها
    def upto(no):
        no = qs_core.wno(no); d = appd.get(no)
        if hasattr(d, 'date') and not isinstance(d, dt.date): d = d.date()
        if isinstance(d, dt.datetime): d = d.date()
        return bool(no) and d is not None and d <= d_to
    split = lambda v: [x.strip() for x in str(v or '').split(',') if x.strip()]
    for r in items.get('asp', []):
        for key, main, extra in (('asp1', r.get('w1'), (r.get('wmc'),)), ('asp2', r.get('w2'), (r.get('wrc'), r.get('w3')))):
            mains = split(main)
            if not mains: continue
            code = qs_core.PO_ITEMS[key][1]
            now = any(inper(appd.get(qs_core.wno(n))) for n in mains)
            for n in mains: add(code, n)
            for n in (x for v in extra for x in split(v)):
                if not now: add(code, n)
                elif upto(n):
                    lst = out.setdefault(int(code), []); n = qs_core.wno(n)
                    if n not in lst: lst.append(n)
    def k(x):
        try: return (0, float(x))
        except ValueError: return (1, x)
    return {c: sorted(v, key=k) for c, v in out.items()}


def preview(project, no, date_to, out_dir=None, template=None):
    c = compute(project, no, date_to, out_dir, template)
    return [x for x in c['rows'] if x['prev'] or x['todate']], c['prev_file'], c['prev_to']


# ---------------------------------------------------------------- عمل المستخلص
def make_ipc(project, no, date_from, date_to, out_dir=None, template=None, log=None, wir_rows=None, calc=None):
    """project = qs_core.Project جاهز. يرجّع (مسار الملف, ملخص)."""
    template = template or qs_core.project_path(TEMPLATE_REL)
    out_dir = out_dir or qs_core.project_path(OUT_REL)
    log = log if log is not None else qs_core.project_path(LOG_REL)
    if not os.path.exists(template): raise FileNotFoundError(f'Template not found:\n{template}')
    if ipc_status(log).get(no, {}).get('status') in CERTIFIED:
        raise ValueError(f'IPC No. {no} is Certified – an approved IPC cannot be made again.')
    os.makedirs(out_dir, exist_ok=True)
    c = calc or compute(project, no, date_to, out_dir, template, date_from)
    dest = ipc_folder(out_dir, no)                  # ملفات المستخلص في فولدره (1-IPC، 2-IPC…)
    os.makedirs(dest, exist_ok=True)
    # شيتات البرنامج (نفس تصدير الحصر، على نفس التاريخ)
    tmpd = tempfile.mkdtemp()
    qs_out = os.path.join(tmpd, 'qs.xlsx')
    old_mode, old_cut = project.mode, project.cutoff
    project.mode = 'approved'; project.cutoff = date_to
    try:
        project.export(c['items'], list(c['notes']), qs_out)
    finally:
        project.mode, project.cutoff = old_mode, old_cut
    wb = openpyxl.load_workbook(template)
    cv = wb[COVER]
    cv['F6'] = no; cv['E7'] = date_from; cv['E8'] = date_to
    for cell, name in SIGNER_NAMES.items():          # أسماء الموقّعين في الغلاف ← بتطلع في كل التوقيعات
        if not cv[cell].value: cv[cell] = name
    ws = wb[IPC]
    for x in c['rows']:
        ws[f"I{x['r']}"] = x['prev'] if x['prev'] else 0
        ws[f"K{x['r']}"] = x['todate'] if x['todate'] else 0
    for r, (lab, amt) in zip(c['sum_rows'], c['amounts']):
        ws[f'G{r}'] = tafqit_formula(f'E{r}') if abs(amt or 0) > 0.004 else None      # صفر ← فاضي زي الاستشاري
    rows_map, pages, _ = compact_vo(ws, c, VO_MODE, no, date_to)      # ورق أقل: البنود ورا بعض + مرحّل محسوب
    tidy_signatures(ws)                               # التوقيعات: الوظيفة والاسم بس (من غير كلمات «الوظيفة/الاسم/التوقيع») ومن غير سطر «الأعمال منفذة…»
    number_pages(wb, no, date_to, pages)
    vo_extras(ws, c, rows_map)
    plan_box(ws, no)                                  # نسبة المستخلصات | نسبة الإنجاز | الفرق (برا الطباعة)                        # قيم السابق/خلال الفترة/حتى تاريخه برا الطباعة + تلوين «حتى تاريخه» مظبوط
    # ملخص الحصر (QS SUMMARY) اتشال من المستخلص (2026-11-02): «خلاصة حصر المستخلص» في OUR_SHEET بتحسبه أوتوماتيك
    if QSUM in wb.sheetnames: del wb[QSUM]
    # ملفين منفصلين: المستخلص (CAVAR + VO) + الحصر الكلي لوحده (باتجاهه الأصلي، مش بيقلب الشاشة)
    if HELP in wb.sheetnames: del wb[HELP]
    # من غير صفحة فاضية في الآخر: منطقة الطباعة لحد آخر صف فيه كلام، ومن غير فاصل صفحات بعده
    last = max((c.row for row in ws.iter_rows(max_col=13) for c in row if c.value not in (None, '')), default=ws.max_row)
    try:
        a = str(ws.print_area).split('!')[-1].replace('$', '').split(':')[0]
        ws.print_area = f"{a}:M{last}"
    except Exception:
        pass
    ws.row_breaks.brk = [b for b in ws.row_breaks.brk if b.id < last]
    umap = unifier_map(project)                     # أرقام الطلبات في الملخص والحصر = أرقام اليونيفاير
    wb.active = 0
    base = f'IPC-{no:02d} ({date_to:%d-%m-%Y})'
    out = os.path.join(dest, base + '.xlsx')
    qs_file = os.path.join(dest, base + ' - QS-TOTAL.xlsx')
    tmp = os.path.join(tmpd, 'ipc.xlsx')
    wb.save(tmp)
    try:
        add_cover_logos(tmp, out, template)
    except Exception:
        shutil.copyfile(tmp, out)
    src = openpyxl.load_workbook(qs_out)
    for n in list(src.sheetnames):
        if n != 'QS-TOTAL': del src[n]
    src.active = 0
    qws = src['QS-TOTAL']
    site = {}                                       # أرقام طلبات الموقع (قبل اليونيفاير) – عشان رقم المستخلص INV NO
    for r in range(5, qws.max_row + 1):
        for col in range(52, 52 + len(qs_core.WIR_COLS)):
            v = qws.cell(r, col).value
            if v not in (None, '') and not (isinstance(v, str) and v.startswith('=')):
                site[(r + QS_HEAD_ROWS, col)] = v
                qws.cell(r, col).value = to_unifier(v, umap)
    qspec = dress_qs_total(qws, no, date_from, date_to)
    try: qs_inv_columns(qws, project, site, inv_lookup(project, no, date_to, out_dir))   # UN NO + INV NO + متابعة المستخلصات
    except Exception: pass
    qs_consultant_colors(qws)
    sspec = build_summary(src, c, no, date_from, date_to, umap, project)      # ملخص الحصر: شيت تاني في ملف الحصر الكلي
    qtmp = os.path.join(tmpd, 'qs.xlsx')
    src.save(qtmp)
    try:
        add_cover_logos(qtmp, qs_file, template, specs={'QS-TOTAL': qspec, SUMS: sspec})
    except Exception:
        shutil.copyfile(qtmp, qs_file)
    shutil.rmtree(tmpd, ignore_errors=True)
    pdf, pdf_err, qs_pdf, sum_pdf = None, None, None, None
    try:                                           # 4 ملفات PDF: المستخلص (الغلاف + VO)، الملخص، الحصر الكلي، والطلبات (WIR)
        pdf = export_pdf(out, [COVER, IPC])
        sum_pdf, qs_pdf = export_pdfs(qs_file, [([SUMS], os.path.join(dest, base + ' - QS-SUMMARY.pdf')),
                                                (['QS-TOTAL'], os.path.join(dest, base + ' - QS-TOTAL.pdf'))])
    except Exception as e:
        pdf_err = str(e)
    wir_pack = None
    if pdf or qs_pdf:                                   # طلبات الاستلام في ملف لوحده: … - WIR.pdf
        try: wir_pack = make_wir_pdf(os.path.join(dest, base + ' - WIR.pdf'), c, unf=umap)
        except Exception as e: wir_pack = dict(error=str(e))
    try: wir_zips = make_wir_zips(project, c, base, dest, unf=umap)    # zip لكل نوع في 6_MERG_UN_SIT_WIR + نسخة جنب المستخلص
    except Exception as e: wir_zips = dict(error=str(e))
    summary = {'pdf': pdf, 'pdf_error': pdf_err, 'qs_file': qs_file, 'qs_pdf': qs_pdf, 'sum_pdf': sum_pdf, 'wir_pack': wir_pack, 'wir_zips': wir_zips,
               'no': no, 'from': date_from, 'to': date_to, 'contract': c['contract'], 'todate': c['todate'],
               'prev': c['prev'], 'period': c['period'], 'vat': c['vat'], 'total': c['total'],
               'prev_file': c['prev_file'], 'less_than_prev': c['less_than_prev'], 'file': out}
    if log and os.path.exists(log):
        try: log_ipc(log, summary)
        except PermissionError: summary['log_error'] = '11-LOG IPC is open – close it and log again'
    return out, summary


def add_wir_sheet(wb, rows, cut):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    ws = wb.create_sheet('برنامج - الطلبات')
    ws.sheet_view.rightToLeft = True
    thin = Side(style='thin', color='808080'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr = ['م', 'رقم الطلب', 'الحي', 'الشارع', 'الخط', 'الكود', 'من', 'إلى', 'الطول', 'القطر', 'تاريخ الطلب', 'تاريخ التسليم',
           'رقم اليونيفاير', 'الكمية', 'الوحدة', 'البند', 'محسوب في المستخلص']
    for i, h in enumerate(hdr, 1):
        c = ws.cell(1, i, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78'); c.border = B
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for j, r in enumerate(rows, 2):
        ok = bool(r.get('appdate')) and r['appdate'] <= cut
        vals = [j - 1, r.get('no'), r.get('dist'), r.get('street'), r.get('line'), r.get('code'), r.get('a_raw'), r.get('b_raw'),
                r.get('L_raw'), r.get('dia'), r.get('reqdate'), r.get('appdate'), r.get('unf'), r.get('qty'), r.get('unit'),
                r.get('item'), 'نعم' if ok else 'لا']
        for i, v in enumerate(vals, 1):
            c = ws.cell(j, i, v); c.border = B; c.alignment = Alignment(horizontal='center')
            if isinstance(v, (dt.date, dt.datetime)): c.number_format = 'dd-mm-yyyy'
            if isinstance(v, float): c.number_format = '#,##0.000'
            if not ok: c.fill = PatternFill('solid', fgColor='EDEDED')
    for i, w in enumerate([5, 9, 12, 18, 6, 10, 8, 8, 8, 6, 11, 11, 10, 11, 6, 14, 10], 1):
        ws.column_dimensions[CL(i)].width = w
    ws.freeze_panes = 'A2'


def log_ipc(path, s):
    wb = openpyxl.load_workbook(path)
    ws = wb['IPC']
    r = None
    for i in range(12, ws.max_row + 1):
        v = ws[f'B{i}'].value
        if v in (s['no'], str(s['no'])) or v is None:
            r = i; break
    if r is None: r = ws.max_row + 1
    ws[f'B{r}'] = s['no']; ws[f'C{r}'] = s['from']; ws[f'D{r}'] = s['to']
    ws[f'F{r}'] = round(s['todate'], 2); ws[f'G{r}'] = round(s['prev'], 2); ws[f'K{r}'] = s['vat']
    if not ws[f'O{r}'].value: ws[f'O{r}'] = 'Submitted'
    ws[f'Q{r}'] = os.path.basename(s['file'])
    for c in 'CD': ws[f'{c}{r}'].number_format = 'dd-mm-yyyy'
    wb.save(path)


def log_plan(path, no, pct, plan_value, diff):
    """تحديث مهندس التخطيط في 11-LOG IPC: R نسبة التخطيط، S قيمته، T الفرق"""
    if not path or not os.path.exists(path): return
    from copy import copy
    wb = openpyxl.load_workbook(path)
    ws = wb['IPC']
    for col, h in (('R', 'Planned Progress %'), ('S', 'Planned Value (SAR)'), ('T', 'IPC vs Plan %')):
        c = ws[f'{col}11']
        if not c.value:
            c.value = h; src = ws['Q11']
            c.font, c.fill, c.border, c.alignment = copy(src.font), copy(src.fill), copy(src.border), copy(src.alignment)
            ws.column_dimensions[col].width = 16
    r = None
    for i in range(12, ws.max_row + 1):
        if ws[f'B{i}'].value in (no, str(no)): r = i; break
    if r is None: return
    ws[f'R{r}'] = pct / 100; ws[f'R{r}'].number_format = '0.00%'
    ws[f'S{r}'] = round(plan_value, 2); ws[f'S{r}'].number_format = '#,##0.00'
    ws[f'T{r}'] = round(diff, 4); ws[f'T{r}'].number_format = '+0.00%;-0.00%;0.00%'
    wb.save(path)


# ---------------------------------------------------------------- طلبات اليونيفاير ورا الحصر الكلي
def wir_pdfs(folder=None):
    """{رقم الطلب: ملف PDF} – الاسم لازم يبدأ برقم الطلب (021 - ...). الملفات اللي بتبدأ بـ X- بتتجاهل."""
    folder = folder or qs_core.project_path(WIR_PDF_REL)
    out = {}
    if not os.path.isdir(folder): return out
    for fn in os.listdir(folder):
        if not fn.lower().endswith('.pdf'): continue
        m = re.match(r'\s*0*(\d+)\s*(?:[-_ .]|$)', fn)
        if not m: continue
        n = int(m.group(1)); p = os.path.join(folder, fn)
        if n not in out or os.path.getmtime(p) > os.path.getmtime(out[n]): out[n] = p
    return out


def _pypdf():
    try:
        import pypdf; return pypdf
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, '-m', 'pip', 'install', '--user', '-q', 'pypdf'], capture_output=True, timeout=300,
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        import importlib, site
        try: sys.path.append(site.getusersitepackages())
        except Exception: pass
        importlib.invalidate_caches()
        import pypdf; return pypdf


def unifier_map(project):
    """{رقم طلبنا: رقم اليونيفاير} من عمود «رقم اليونيفاير» في لوج الطلبات (ملفات 5_UNIFIER_APPROVED متسمّية برقم اليونيفاير)"""
    out = {}
    for q in getattr(project, 'reqs_all', []) or []:
        try: out[int(float(str(q.get('no')).strip()))] = int(float(str(q.get('unf')).strip()))
        except (TypeError, ValueError): pass
    return out


def wir_type_names(project):
    """{كود النوع: اسمه بالإنجليزي} من شيت DAT في لوج الطلبات (عمود B = الإنجليزي، C = الكود)."""
    out = {}
    try:
        src = str(project.S['WIR_LOG_PATH'])
        src = src if os.path.isabs(src) else qs_core.project_path(src)
        tmp = os.path.join(tempfile.gettempdir(), 'ipc_dat_copy' + os.path.splitext(src)[1])
        shutil.copy2(src, tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        for r in wb['DAT'].iter_rows(max_col=3, values_only=True):
            if r[1] and r[2]: out[re.sub(r'\s+', '', str(r[2])).upper()] = str(r[1]).strip()
        wb.close()
    except Exception:
        pass
    return out


def _sta(x):
    """130 ← 0+130 ، 1024.43 ← 1+024.43"""
    try: x = float(x)
    except (TypeError, ValueError): return ''
    km, m = divmod(x, 1000)
    return f'{int(km)}+{m:06.2f}'.replace('.00', '') if abs(m - round(m)) > 1e-9 else f'{int(km)}+{int(round(m)):03d}'


def wir_file_name(q, n, umap, types):
    """اسم ملف الطلب المدموج: «UN-00016 - Request for … - 0+130 to 0+168» (لو ملوش يونيفاير: «SITE-015 - …»)."""
    u = umap.get(n) if umap else None
    head = f'UN-{u:05d}' if u else f'SITE-{n:03d}'
    typ = types.get(re.sub(r'\s+', '', str((q or {}).get('type') or '')).upper()) or str((q or {}).get('type') or '').strip()
    a, b = (q or {}).get('a'), (q or {}).get('b')
    st = f'{_sta(min(a, b))} to {_sta(max(a, b))}' if isinstance(a, (int, float)) and isinstance(b, (int, float)) else ''
    name = ' - '.join(x for x in (head, typ, st) if x)
    return re.sub(r'[\\/:*?"<>|]+', '-', name)[:150]


def to_unifier(text, umap):
    """«3, 7, 15 (T: 19)» ← أرقام طلبات الموقع تتحوّل لأرقام اليونيفاير. الطلب اللي ملوش رقم يونيفاير بيفضل برقمه وجنبه * """
    if text in (None, '') or not umap: return text
    def rep(m):
        n = int(m.group(0))
        return str(umap[n]) if n in umap else f'{n}*'
    return re.sub(r'\d+', rep, str(text))


def make_wir_pdf(out_pdf, calc, folder=None, paper_folder=None, unf=None, project=None):
    """ملف PDF لوحده (… - WIR.pdf) فيه طلبات الاستلام بتاعة المستخلص ده بس: لكل طلب نسخة اليونيفاير وبعدها الورقة المعتمدة.
    الترتيب برقم الطلب، وكل ملف ليه علامة (Bookmark)."""
    order = set()
    for x in calc['rows']:
        for w in str(x.get('wirs') or '').split(','):
            w = w.strip()
            if w.isdigit(): order.add(int(w))
    order = sorted(order)
    if not order: return dict(added=[], missing=[], missing_paper=[])
    uni = wir_pdfs(folder)
    paper = wir_pdfs(paper_folder or qs_core.project_path(WIR_PAPER_REL))
    if unf is not None:                                  # ملف اليونيفاير برقم اليونيفاير، مش برقم طلبنا
        uni = {n: uni[unf[n]] for n in order if n in unf and unf[n] in uni}
    missing = [n for n in order if n not in uni]
    missing_paper = [n for n in order if n not in paper]
    if not any(n in uni or n in paper for n in order): return dict(added=[], missing=missing, missing_paper=missing_paper)
    if os.path.exists(out_pdf):
        try:
            with open(out_pdf, 'ab'): pass
        except PermissionError:
            return dict(error=f'{os.path.basename(out_pdf)} is open – close it and make the IPC again')
    pypdf = _pypdf()
    w = pypdf.PdfWriter()
    added, bad = [], []
    for n in order:
        got = False
        head = f'UN-{unf[n]:05d}' if unf and unf.get(n) else f'SITE-{n:03d}'
        for src, tag in ((uni, 'Unifier'), (paper, 'Site')):
            if n not in src: continue
            try: w.append(src[n], outline_item=f'{head} – {tag}'); got = True
            except Exception: bad.append(f'{n} ({tag})')
        if got: added.append(n)
    if not added: return dict(added=[], missing=missing, missing_paper=missing_paper, bad=bad)
    tmp = out_pdf + '.tmp'
    with open(tmp, 'wb') as f: w.write(f)
    os.replace(tmp, out_pdf)
    return dict(file=out_pdf, added=added, missing=missing, missing_paper=missing_paper, bad=bad)


# فولدر كل نوع طلب في 6_MERG_UN_SIT_WIR (بكود الطلب في اللوج – عمود NAM). المسار والرفع المساحي مبيتجمعوش.
MERGE_FOLDERS = [('PIP-T', 'Test Report'), ('PIP', 'Pipe Installation'), ('EXC', 'Excavation'), ('BED', 'Sand Bedding'),
                 ('COV', 'Sand Bedding'), ('BF', 'Backfill of Subgrade and Subbase'), ('RB-1', 'Backfill of Subgrade and Subbase'),
                 ('RB-2', 'Backfill of Base Course'), ('ASP', 'Asphalt'), ('MC', 'Asphalt'), ('RC', 'Asphalt'), ('CH', 'ALL CHAMBER')]


def merge_folder(code):
    c = re.sub(r'\s+', '', str(code or '')).upper()
    for pre, f in MERGE_FOLDERS:
        if c.startswith(pre): return f
    return None


def make_wir_zips(project, calc, base, out_dir, unf=None, folder=None, paper_folder=None, merge_root=None):
    """لطلبات المستخلص ده بس: كل طلب (اليونيفاير وبعده الورقي) في PDF واحد باسم الطلب، ومتجمعين zip لكل نوع
    جوه فولدره في 6_MERG_UN_SIT_WIR، ونسخة من كل zip في فولدر «IPC-xx (…) - WIR» جنب المستخلص."""
    import zipfile
    order = set()
    for x in calc['rows']:
        for w in str(x.get('wirs') or '').split(','):
            w = w.strip()
            if w.isdigit(): order.add(int(w))
    order = sorted(order)
    if not order: return dict(zips={})
    info = {}
    for q in getattr(project, 'reqs_all', []) or []:
        try: info[int(float(str(q['no']).strip()))] = q
        except (ValueError, TypeError): pass
    uni = wir_pdfs(folder)
    paper = wir_pdfs(paper_folder or qs_core.project_path(WIR_PAPER_REL))
    if unf is not None: uni = {n: uni[unf[n]] for n in order if n in unf and unf[n] in uni}
    root = merge_root or qs_core.project_path(WIR_MERGE_REL)
    groups, skipped, empty = {}, [], []
    for n in order:
        f = merge_folder(info.get(n, {}).get('type'))
        if not f: skipped.append(n); continue
        if n not in uni and n not in paper: empty.append(n); continue
        groups.setdefault(f, []).append(n)
    pypdf = _pypdf()
    types = wir_type_names(project)
    tmpd = tempfile.mkdtemp()
    copy_dir = os.path.join(out_dir, base + ' - WIR')
    zips, bad = {}, []
    try:
        for f, nos in groups.items():
            zp = os.path.join(tmpd, f'{base} - {f}.zip')
            with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
                used = set()
                for n in nos:
                    w = pypdf.PdfWriter()
                    for src in (uni, paper):
                        if n in src:
                            try: w.append(src[n])
                            except Exception: bad.append(f'{n} ({os.path.basename(src[n])})')
                    if not len(w.pages): continue
                    name = wir_file_name(info.get(n), n, unf, types)       # UN-00016 - نوع الطلب بالإنجليزي - المحطات
                    if name in used: name += f' ({n})'
                    used.add(name)
                    t = os.path.join(tmpd, 'one.pdf')
                    with open(t, 'wb') as fh: w.write(fh)
                    z.write(t, name + '.pdf')
            for d in (os.path.join(root, f), copy_dir):
                os.makedirs(d, exist_ok=True)
                shutil.copyfile(zp, os.path.join(d, os.path.basename(zp)))
            zips[f] = nos
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)
    return dict(zips=zips, copy_dir=copy_dir if zips else None, skipped=skipped, empty=empty, bad=bad,
                missing=[n for n in order if n not in uni and n not in skipped],
                missing_paper=[n for n in order if n not in paper and n not in skipped])


# ---------------------------------------------------------------- شكل الحصر الكلي (A3 + رأس + توقيعات)
QS_SIGNERS = [   # (الجهة, [(الوظيفة, الاسم), ...]) – زي التوقيعات اللي في شيت الاستشاري
    ('المقاول : شركة المقاول', [('حاسب الكميات', 'م/ (اسم المهندس)'), ('مدير المشروع', 'م/ (اسم المهندس)')]),
    ('الاستشاري : مكتب الاستشاري', [('حاسب الكميات', 'م/ (اسم المهندس)'), ('مهندس الموقع', 'م/ (اسم المهندس)')]),
]
QS_HEAD_ROWS = 6


def _shift_down(ws, n):
    """يزوّد n صف فوق الشيت ويعدّل المعادلات والدمج وارتفاع الصفوف (openpyxl مبيعملش ده لوحده)."""
    from openpyxl.formula.translate import Translator
    merges = [str(m) for m in ws.merged_cells.ranges]
    for m in merges: ws.unmerge_cells(m)
    heights = {r: d.height for r, d in ws.row_dimensions.items() if d.height}
    cells = [(c.row, c.column, c.value) for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('=')]
    ws.insert_rows(1, n)
    for r, col, f in cells:
        c = ws.cell(r + n, col)
        c.value = Translator(f, origin=f'{CL(col)}{r}').translate_formula(f'{CL(col)}{r + n}')
    for r in list(ws.row_dimensions): ws.row_dimensions[r].height = None
    for r, h in heights.items(): ws.row_dimensions[r + n].height = h
    from openpyxl.utils.cell import range_boundaries
    for m in merges:
        c1, r1, c2, r2 = range_boundaries(m)
        ws.merge_cells(start_row=r1 + n, start_column=c1, end_row=r2 + n, end_column=c2)
    if ws.freeze_panes:
        from openpyxl.utils.cell import coordinate_from_string
        col, row = coordinate_from_string(ws.freeze_panes); ws.freeze_panes = f'{col}{row + n}'


QS_WIDTHS = {'A': 3.29, 'B': 11.29, 'C': 9.0, 'D': 12.86, 'E': 7.57, 'F': 10.0, 'G': 8.0, 'H': 11.0, 'I': 13.0, 'J': 12.57, 'K': 13.0,
             'L': 9.29, 'M': 8.71, 'N': 10.29, 'O': 13.0, 'P': 10.86, 'Q': 9.14, 'R': 8.43, 'S': 10.0, 'T': 9.71, 'U': 7.71, 'V': 11.14,
             'W': 10.0, 'X': 13.0, 'Y': 9.0, 'Z': 10.29, 'AA': 9.71, 'AB': 10.29, 'AC': 13.43, 'AD': 11.57, 'AE': 12.29, 'AF': 11.14,
             'AG': 13.29, 'AH': 9.29, 'AI': 13.29, 'AJ': 9.29, 'AK': 13.29, 'AL': 9.0, 'AM': 16.0, 'AN': 8.43, 'AO': 9.0, 'AP': 8.14,
             'AQ': 9.0, 'AR': 8.14, 'AS': 6.86, 'AT': 9.0, 'AU': 8.14, 'AV': 6.86, 'AW': 9.0, 'AX': 8.43, 'AY': 11.57}


def dress_qs_total(ws, no, date_from, date_to):
    """رأس زي المستخلص (لوجوهات + بيانات المشروع + رقم المستخلص وفترته) فوق الحصر، وتوقيعات تحته، وطباعة A3.
    بيرجّع مواصفات اللوجوهات لـ add_cover_logos."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.properties import PageSetupProperties
    NAVY, LIGHT = '1F4E78', 'DDEBF7'
    F = 'Arial'
    ncol = 51        # الطباعة والـ PDF: A..AY زي شيت الاستشاري. أعمدة أرقام الطلبات (AZ..BD) بتفضل في الإكسيل بس، برا منطقة الطباعة
    for col, w in QS_WIDTHS.items():                          # عرض الأعمدة زي ملف الاستشاري
        ws.column_dimensions[col].width = w
    for r in range(1, 5):
        for col in 'GHI': ws[f'{col}{r}'].alignment = ws[f'{col}{r}'].alignment.copy(wrap_text=True)
    last = max((c.row for row in ws.iter_rows() for c in row if c.value not in (None, '')), default=4)
    H = QS_HEAD_ROWS
    _shift_down(ws, H)
    last += H
    # الأعمدة: نقسم العرض تلات أجزاء للرأس
    def wpx(i):
        w = ws.column_dimensions[CL(i)].width
        return int((w if w else 8.43) * 7 + 5)
    widths = [wpx(i) for i in range(1, ncol + 1)]
    tot = sum(widths); acc, cuts = 0, []
    for i, w in enumerate(widths, 1):
        acc += w
        if len(cuts) < 2 and acc >= tot * (len(cuts) + 1) / 3: cuts.append(i)
    boxes = [(1, cuts[0]), (cuts[0] + 1, cuts[1]), (cuts[1] + 1, ncol)]      # شمال: المقاول | نص: الاستشاري | يمين: المالك
    thin = Side(style='thin', color='9DB3C8'); med = Side(style='medium', color=NAVY)
    for r, h in zip(range(1, H + 1), (24, 24, 24, 30, 26, 8)): ws.row_dimensions[r].height = h
    texts = [
        ['المقاول', 'شركة المقاول', 'Contractor Co.'],
        ['الاستشاري', 'مكتب الاستشاري', 'Consultant Engineering Office'],
        ['الجهة المالكة', 'وحدة أعمال المدينة', 'الإدارة العامة للمشاريع بالقطاع الأوسط'],
    ]
    for (c1, c2), lines in zip(boxes, texts):
        for k, t in enumerate(lines):
            r = k + 1
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
            cell = ws.cell(r, c1, t)
            cell.font = Font(name=F, size=13 if k == 0 else 11, bold=k < 2, color=NAVY if k == 0 else '1F2D3D')
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for r in range(1, 4):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).border = Border(top=med if r == 1 else None, bottom=med if r == 3 else None,
                                              left=med if c == c1 else None, right=med if c == c2 else None)
    # العنوان
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=ncol)
    c = ws.cell(4, 1, 'الحصر الكلي   —   مشروع شبكات المياه (مثال) ( أمر شراء 000000000 )')
    c.font = Font(name=F, size=16, bold=True, color=NAVY); c.alignment = Alignment(horizontal='center', vertical='center')
    # شريط رقم المستخلص
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=ncol)
    f = lambda d: d.strftime('%d-%m-%Y') if hasattr(d, 'strftime') else str(d or '')
    c = ws.cell(5, 1, f'المستخلص الجاري رقم ( {no} )      الفترة من  {f(date_from)}  إلى  {f(date_to)}')
    c.font = Font(name=F, size=13, bold=True, color='FFFFFF'); c.alignment = Alignment(horizontal='center', vertical='center')
    for col in range(1, ncol + 1): ws.cell(5, col).fill = PatternFill('solid', fgColor=NAVY)
    # التوقيعات تحت الإجمالي: المقاول شمال، الاستشاري يمين، كل واحد اتنين جنب بعض
    r0 = last + 3
    halves = [(ncol // 2 + 2, ncol), (1, ncol // 2 - 1)]       # المقاول يمين، الاستشاري شمال
    for (party, people), (h1, h2) in zip(QS_SIGNERS, halves):
        ws.merge_cells(start_row=r0, start_column=h1, end_row=r0, end_column=h2)
        c = ws.cell(r0, h1, party); c.font = Font(name=F, size=13, bold=True, color=NAVY)
        c.alignment = Alignment(horizontal='center', vertical='center')
        for col in range(h1, h2 + 1):
            ws.cell(r0, col).fill = PatternFill('solid', fgColor=LIGHT); ws.cell(r0, col).border = Border(bottom=med)
        span = (h2 - h1 + 1) // len(people)
        for j, (job, name) in enumerate(people):
            a = h1 + j * span; b = h2 if j == len(people) - 1 else a + span - 1
            for k, t in enumerate((job, name, '', '')):          # الوظيفة والاسم بس – والسطر الفاضي للإمضا
                r = r0 + 1 + k
                ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
                cell = ws.cell(r, a, t if t else None)
                cell.font = Font(name=F, size=12, bold=k == 0, color='1F2D3D')
                cell.alignment = Alignment(horizontal='right', vertical='center', indent=3)    # الوظيفة/الاسم/التوقيع تحت بعض على اليمين
            for col in range(a + 2, b - 1):
                ws.cell(r0 + 4, col).border = Border(bottom=Side(style='dotted', color='7F8C9A'))
    for r, h in zip(range(r0, r0 + 5), (26, 22, 22, 30, 14)): ws.row_dimensions[r].height = h
    end = r0 + 4
    # الطباعة: A3 بالعرض، كل الأعمدة في عرض صفحة، الرأس بيتكرر
    ws.page_setup.paperSize = 8; ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.3; ws.page_margins.top = 0.35; ws.page_margins.bottom = 0.45
    ws.page_margins.header = 0.15; ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True
    ws.print_area = f'A1:{CL(ncol)}{end}'
    ws.print_title_rows = f'1:{H + 4}'
    ws.oddFooter.center.text = 'صفحة &P من &N'; ws.oddFooter.center.size = 9
    ws.oddFooter.right.text = '&D'; ws.oddFooter.right.size = 9
    # اللوجوهات: (صورة: 0 الجهة المالكة، 1 الاستشاري، 2 المقاول, العمود, الصف, العرض المتاح px, الارتفاع المتاح EMU, ارتفاع اللوجو EMU)
    hav = int(72 * 12700)                                   # 3 صفوف × 24 نقطة
    lw = sum(widths[:3]); rw = sum(widths[-3:])
    mid = boxes[1][0] - 1
    return [(2, 0, 0, lw, hav, int(hav * .85)),
            (1, mid, 0, sum(widths[mid:mid + 3]), hav, int(hav * .8)),
            (0, ncol - 3, 0, rw, hav, int(hav * .85))]


# ---------------------------------------------------------------- ملخص الحصر (زي «خلاصة حصر المستخلص» عند الاستشاري)
SUMS = 'QS-SUMMARY'


def _unit_ar(u):
    """الوحدة بالعربي بس (عدد/Nr ← عدد)."""
    for p in str(u or '').replace('\n', '/').split('/'):
        if any('؀' <= ch <= 'ۿ' for ch in p): return p.strip()
    return u


def qs_links(wb, project):
    """كود البند ← معادلة الكمية الإجمالية من الحصر الكلي (زي الاستشاري): ROUNDDOWN لإجمالي العمود – نفس تقريب المستخلص."""
    if 'QS-TOTAL' not in wb.sheetnames or project is None: return {}
    q = wb['QS-TOTAL']
    t = next((r for r in range(QS_HEAD_ROWS + 5, q.max_row + 1) if q.cell(r, 1).value == 'الإجمالي'), None)
    if not t: return {}
    P = qs_core.PO_ITEMS
    m = {}
    for b in getattr(project, 'boq', []) or []:
        band, code = b[0], b[2]
        col = {1: 'V', 2: 'W', 3: 'X', 4: 'Y'}.get(int(band))
        if col: m[int(code)] = [col]
    m[P['bed'][1]] = ['AA', 'AE']
    for d, col in ((300, 'AF'), (400, 'AH'), (600, 'AJ')): m[P[d][1]] = [col]
    m[P['bf'][1]] = ['AN']
    m[P['rb'][1]] = ['AP', 'AR']
    m[P['asp1'][1]] = ['AU']                           # سفلتة عادية
    m[P['asp2'][1]] = ['AX']                           # سفلتة بالفرادة والكشاطة (الطبقة السطحية AY مش داخلة زي الاستشاري)
    return {k: "=ROUNDDOWN(ROUND(" + '+'.join(f"'QS-TOTAL'!{c}{t}" for c in v) + ",4),0)" for k, v in m.items()}


def build_summary(wb, c, no, date_from, date_to, umap, project=None):
    """شيت «ملخص الحصر» (QS-SUMMARY) في ملف الحصر الكلي، من الشمال لليمين زي شيت الاستشاري: كل بنود أمر الشراء بالسابق وخلال الفترة والإجمالي والقيم وأرقام طلبات اليونيفاير.
    عليه فلتر: الطباعة والـ PDF بالبنود اللي ليها كمية خلال الفترة بس. بيرجّع مواصفات اللوجوهات."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.properties import PageSetupProperties
    from openpyxl.worksheet.filters import CustomFilters, CustomFilter, FilterColumn
    if SUMS in wb.sheetnames: del wb[SUMS]
    ws = wb.create_sheet(SUMS)
    ws.sheet_view.rightToLeft = False; ws.sheet_view.showGridLines = False; ws.sheet_view.zoomScale = 80
    NAVY, LIGHT, TOTF = '1F4E78', 'DDEBF7', 'FFF2CC'
    thin = Side(style='thin', color='8EA9C1'); med = Side(style='medium', color=NAVY)
    B = Border(left=thin, right=thin, top=thin, bottom=thin)
    F = lambda sz=11, b=False, col='1F2D3D': Font(name='Arial', size=sz, bold=b, color=col)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center', wrap_text=True)
    COLS = 'ABCDEFGHIJKLMN'
    W = dict(A=12, B=11, C=52, D=9, E=12, F=11, G=11, H=12, I=11, J=15, K=15, L=15, M=30, N=11)
    links = qs_links(wb, project)                    # الإجمالية مربوطة بالحصر الكلي
    for k, w in W.items(): ws.column_dimensions[k].width = w
    # ---- رأس المستخلص (زي VO بس متعكس): الجهة المالكة شمال | الاستشاري | المقاول يمين + شريط رقم المستخلص
    boxes = [(1, 3, 2, 3), (4, 9, 4, 8), (10, 14, 10, 13)]          # (من, لحد, نص من, نص لحد)
    texts = [['الجهة المالكة', 'وحدة أعمال المدينة', 'الإدارة العامة للمشاريع بالقطاع الأوسط'],
             ['', 'الاستشاري : مكتب الاستشاري', ''],
             ['', 'المقاول : شركة المقاول', '']]
    for (a, b, t1, t2), lines in zip(boxes, texts):
        for k, t in enumerate(lines):
            ws.merge_cells(start_row=k + 1, start_column=t1, end_row=k + 1, end_column=t2)
            x = ws.cell(k + 1, t1, t or None); x.font = F(11, True); x.alignment = CEN
        for r in range(1, 4):
            for cc in range(a, b + 1):
                ws.cell(r, cc).border = Border(top=med if r == 1 else None, bottom=med if r == 3 else None,
                                               left=med if cc == a else None, right=med if cc == b else None)
    for r, h in zip(range(1, 7), (24, 24, 24, 8, 26, 30)): ws.row_dimensions[r].height = h
    fmt = lambda d: d.strftime('%d-%m-%Y') if hasattr(d, 'strftime') else str(d or '')
    ws.merge_cells('B5:M5'); x = ws['B5']
    x.value = f'المستخلص الجاري رقم   {no}          الفترة من   {fmt(date_from)}          إلى   {fmt(date_to)}'
    x.font = F(12, True); x.alignment = CEN
    for cc in 'BCDEFGHIJKLM':
        ws[f'{cc}5'].fill = PatternFill('solid', fgColor='B4C6E7')
        ws[f'{cc}5'].border = Border(top=med, bottom=med, left=med if cc == 'B' else None, right=med if cc == 'M' else None)
    ws.merge_cells('A6:N6'); x = ws['A6']
    x.value = 'ملخص الحصر   —   مشروع شبكات المياه (مثال) ( أمر شراء 000000000 )'
    x.font = F(15, True, 'FF0000'); x.alignment = CEN
    for k in COLS: ws[f'{k}6'].fill = PatternFill('solid', fgColor='F2F2F2')         # زي عنوان شيت الاستشاري
    # ---- العناوين
    h1 = {'A': 'رقم البند', 'B': 'الكود', 'C': 'وصف البند', 'D': 'الوحدة', 'E': 'الكمية التعاقدية', 'F': 'الكميات المنفذة',
          'I': 'سعر الوحدة', 'J': 'المنفذ بالريال السعودي', 'M': f'أرقام طلبات اليونيفاير للمستخلص {no:02d}', 'N': 'مراجعة'}
    h2 = {'F': 'السابقة', 'G': 'خلال الفترة', 'H': 'الإجمالية', 'J': 'القيمة الإجمالية', 'K': 'خلال الفترة', 'L': 'السابق'}
    for k, v in h1.items(): ws[f'{k}7'] = v
    for k, v in h2.items(): ws[f'{k}8'] = v
    for k in 'ABCDEIMN': ws.merge_cells(f'{k}7:{k}8')
    ws.merge_cells('F7:H7'); ws.merge_cells('J7:L7')
    for r in (7, 8):
        ws.row_dimensions[r].height = 28
        for k in COLS:
            hf = 'FFC000' if k == 'G' else '92D050' if k == 'H' else 'D9E1F2'         # ألوان عناوين الاستشاري
            x = ws[f'{k}{r}']; x.font = F(11, True, '000000'); x.fill = PatternFill('solid', fgColor=hf); x.alignment = CEN; x.border = B
    # ---- البنود (كلها) – اللي ملهاش كمية خلال الفترة متخبية بالفلتر
    NUM = '#,##0.00;-#,##0.00;"-"'
    r = a0 = 9
    for i, it in enumerate(c['rows']):
        w = to_unifier(it['wirs'], umap) if it.get('wirs') else ''
        vals = dict(A=it['item'], B=it['code'], C=it['desc'], D=_unit_ar(it['unit']), E=it['qty'], F=it['prev'],
                    G=f'=H{r}-F{r}', H=links.get(int(it['code'] or 0), it['todate']), I=it['rate'], J=f'=ROUND(H{r}*I{r},2)', K=f'=J{r}-L{r}',
                    L=f'=ROUND(F{r}*I{r},2)', M=w or None,
                    N=f'=IF(G{r}<0,"NOT OK",IF(E{r}=0,"OK",IF(H{r}<=E{r},"OK","NOT OK")))')
        live = abs(it['period']) > 0.004
        for k, v in vals.items():
            x = ws[f'{k}{r}']; x.value = v; x.border = B; x.font = F(10, live and k in 'CG')
            x.alignment = RGT if k == 'C' else CEN
            if k in 'EFGHIJKL': x.number_format = NUM
            if k in 'GK': x.fill = PatternFill('solid', fgColor='FFFF00')            # خلال الفترة أصفر زي الاستشاري
        n = len(str(it['desc'] or ''))
        ws.row_dimensions[r].height = 20 if n <= 55 else 32 if n <= 120 else 46 if n <= 200 else 62 if n <= 300 else 90
        if not live: ws.row_dimensions[r].hidden = True
        r += 1
    last = r - 1
    ws.auto_filter.ref = f'A8:N{last}'
    ws.auto_filter.filterColumn.append(FilterColumn(colId=6, customFilters=CustomFilters(
        customFilter=[CustomFilter(operator='notEqual', val='0')])))
    # ---- الإجماليات (المجموع بيعدّ كل البنود حتى المتخبية)
    lines = [('القيمة الإجمالية', lambda k: f'=SUM({k}{a0}:{k}{last})'),
             ('ضريبة القيمة المضافة 15 %', lambda k: f'=ROUND({k}{r}*{VAT},2)'),
             ('إجمالي قيمة الأعمال المنفذة بعد إضافة الضريبة', lambda k: f'={k}{r}+{k}{r + 1}')]
    for i, (lab, fn) in enumerate(lines):
        rr = r + i
        ws.merge_cells(f'A{rr}:I{rr}'); ws[f'A{rr}'] = lab
        for k in 'JKL': ws[f'{k}{rr}'] = fn(k)
        ws.merge_cells(f'M{rr}:N{rr}')
        for k in COLS:
            x = ws[f'{k}{rr}']; x.border = B; x.font = F(12, True, '000000')
            x.alignment = RGT if k == 'A' else CEN
            if k in 'JKL': x.number_format = '#,##0.00'
        ws.row_dimensions[rr].height = 24
    # ---- التوقيعات: المقاول يمين، الاستشاري شمال (زي الحصر الكلي)
    r0 = r + 5
    for (party, people), (h1_, h2_, spans) in zip(QS_SIGNERS, [(8, 14, [(8, 10), (11, 14)]), (1, 6, [(1, 3), (4, 6)])]):
        ws.merge_cells(start_row=r0, start_column=h1_, end_row=r0, end_column=h2_)
        x = ws.cell(r0, h1_, party); x.font = F(13, True, NAVY); x.alignment = CEN
        for cc in range(h1_, h2_ + 1):
            ws.cell(r0, cc).fill = PatternFill('solid', fgColor='D0CECE'); ws.cell(r0, cc).border = Border(bottom=med)
        for (a, b), (job, name) in zip(spans, people):
            for k, t in enumerate((job, name, '')):              # الوظيفة والاسم بس – والسطر الفاضي للإمضا
                ws.merge_cells(start_row=r0 + 1 + k, start_column=a, end_row=r0 + 1 + k, end_column=b)
                x = ws.cell(r0 + 1 + k, a, t); x.font = F(12, k == 0)
                x.alignment = Alignment(horizontal='right', vertical='center', indent=3)          # تحت بعض على اليمين
    for rr, h in zip(range(r0, r0 + 4), (26, 22, 22, 34)): ws.row_dimensions[rr].height = h
    end = r0 + 4
    # ---- الطباعة: A4 بالعرض زي المستخلص، كل الأعمدة في عرض صفحة، الرأس والعناوين بتتكرر
    ws.freeze_panes = 'D9'
    ws.page_setup.paperSize = 9; ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.3; ws.page_margins.top = 0.35; ws.page_margins.bottom = 0.45
    ws.page_margins.header = 0.15; ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True
    ws.print_area = f'A1:N{end}'; ws.print_title_rows = '1:8'
    ws.oddFooter.center.text = 'صفحة &P من &N'; ws.oddFooter.center.size = 9
    # ---- اللوجوهات: (0 الجهة المالكة، 1 الاستشاري، 2 المقاول, العمود, الصف, العرض المتاح px, الارتفاع المتاح EMU, ارتفاع اللوجو EMU)
    px = lambda k: int(W[k] * 7 + 5)
    hav = int(72 * 12700)
    return [(0, 0, 0, px('A'), hav, int(hav * .85)), (1, 8, 0, px('I'), hav, int(hav * .75)), (2, 13, 0, px('N'), hav, int(hav * .85))]



# ---------------------------------------------------------------- زي شيت الاستشاري: ألوانه، رقم المستخلص INV NO، متابعة المستخلصات
def _fill(c):
    from openpyxl.styles import PatternFill
    return PatternFill('solid', fgColor=c)


def vo_extras(ws, c, rows_map=None):
    """شيت VO: (1) تلات خانات برا الطباعة جنب كل بند: قيمة السابق، خلال الفترة، حتى تاريخه (بالريال).
    (2) التلوين الشرطي القديم كان ماسك صفوف العناوين («حتى تاريخه» بيطلع أحمر) – بيتعمل من جديد على صفوف البنود بس."""
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.formatting.rule import CellIsRule
    thin = Side(style='thin', color='000000'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cols = ('P', 'Q', 'R')
    for k in cols: ws.column_dimensions[k].width = 18
    heads = [cell.row for cell in ws['I'] if cell.value == 'السابقة']
    for hr in heads:
        ws.merge_cells(f'P{hr - 1}:R{hr - 1}')
        ws[f'P{hr - 1}'] = 'القيمة المنفذة (ريال)'
        for k, t in zip(cols, ('السابقة', 'خلال الفترة', 'حتى تاريخه')): ws[f'{k}{hr}'] = t
        for rr in (hr - 1, hr):
            for k in cols:
                x = ws[f'{k}{rr}']; x.fill = _fill('D9E1F2'); x.font = Font(name='Arial', size=10, bold=True); x.alignment = CEN; x.border = B
    rows = sorted(rows_map.values()) if rows_map is not None else [x['r'] for x in c['rows']]
    for r in rows:
        for k, src in zip(cols, ('I', 'J', 'K')):
            x = ws[f'{k}{r}']; x.value = f'=ROUND({src}{r}*L{r},2)'; x.number_format = '#,##0.00;-#,##0.00;"-"'
            x.border = B; x.alignment = CEN; x.font = Font(name='Arial', size=10)
    ws.conditional_formatting = ConditionalFormattingList()
    if rows:
        rng = lambda col: ' '.join(f'{col}{r}' for r in rows)
        ws.conditional_formatting.add(rng('J'), CellIsRule(operator='greaterThan', formula=['0'], fill=_fill('DDEBF7')))
        ws.conditional_formatting.add(rng('Q'), CellIsRule(operator='greaterThan', formula=['0'], fill=_fill('DDEBF7')))
        first = rows[0]
        ws.conditional_formatting.add(rng('K'), CellIsRule(operator='greaterThan', formula=[f'$G{first}'], fill=_fill('FF0000')))


def inv_lookup(project, no, date_to, out_dir):
    """دالة: رقم طلب الموقع ← رقم المستخلص اللي نزل فيه (بتاريخ الاعتماد وفترات المستخلصات)."""
    ends = []
    for k, pth in sorted((ipc_files(out_dir) if out_dir and os.path.isdir(out_dir) else {}).items()):
        if k >= no: continue
        try:
            d = read_todate(pth)[1]
            if isinstance(d, dt.datetime): d = d.date()
            if isinstance(d, dt.date): ends.append((k, d))
        except Exception:
            pass
    d2 = date_to.date() if isinstance(date_to, dt.datetime) else date_to
    ends.append((no, d2))
    appd = {}
    for q in getattr(project, 'reqs_all', []) or []:
        try: n = int(float(str(q['no']).strip()))
        except (ValueError, TypeError): continue
        a = q.get('appdate')
        if isinstance(a, dt.datetime): a = a.date()
        if isinstance(a, dt.date): appd[n] = a
    def inv(n):
        a = appd.get(n)
        if not a: return None
        for k, e in ends:
            if a <= e: return k
        return None
    return inv


WORK_SHORT = {'الحفر': 'حفر', 'الفرشة': 'فرشة', 'الفرشة أسفل الماسورة': 'فرشة', 'الفرشة حول الماسورة': 'فرشة حول', 'التمديد': 'تمديد', 'اختبار الأنابيب': 'اختبار', 'الردم': 'ردم', 'قاعدة الطريق': 'قاعدة طريق',
              'الأسفلت': 'أسفلت', 'الكشط': 'كشط', 'الطبقة السطحية': 'طبقة سطحية'}


def qs_inv_columns(ws, project, site, inv):
    """أعمدة أرقام الطلبات (AZ..BD) بتبقى أزواج زي الاستشاري: «رقم الطلب UN NO» + «رقم المستخلص INV NO»،
    وبعدهم جدول «متابعة المستخلصات». كله برا منطقة الطباعة (الإكسيل بس)."""
    from openpyxl.styles import Font, Alignment, Border, Side
    thin = Side(style='thin', color='000000'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    H = QS_HEAD_ROWS; h1, h3, h4 = H + 1, H + 3, H + 4            # صفوف العناوين 7..10
    total = next((r for r in range(h4 + 1, ws.max_row + 1) if ws.cell(r, 1).value == 'الإجمالي'), None)
    if not total: return
    data = range(h4 + 1, total)
    names = [str(ws.cell(h3, c).value or '').strip() for c in range(52, 52 + len(qs_core.WIR_COLS))]
    vals = {(r, c): ws.cell(r, c).value for r in data for c in range(52, 52 + len(qs_core.WIR_COLS))}
    for m in [str(m) for m in ws.merged_cells.ranges if m.min_col >= 52]: ws.unmerge_cells(m)
    for r in range(h1, total + 1):
        for c in range(52, 85): ws.cell(r, c).value = None
    # الاختبار مكتوب جوه خانة التمديد «18 (T: 19)» – بيطلع في عمود لوحده «اختبار الأنابيب» زي شيت الاستشاري
    split = lambda v: (lambda m: (m.group(1).strip(), m.group(2).strip()) if m else (v, None))(re.match(r'^(.*?)\s*\(T:\s*(.*?)\)\s*$', str(v))) if v not in (None, '') else (v, None)
    groups = []                                   # (الاسم, {صف: (رقم اليونيفاير, رقم الموقع)})
    for i, nm in enumerate(names):
        main, test = {}, {}
        for r in data:
            u, ut = split(vals[(r, 52 + i)]); sv, st = split(site.get((r, 52 + i)))
            main[r] = (u, sv)
            if ut not in (None, ''): test[r] = (ut, st)
        groups.append((nm, main))
        if nm == 'التمديد': groups.append(('اختبار الأنابيب', test))
    nums = lambda v: [int(x) for x in re.findall(r'\d+', str(v or ''))]
    hdr = lambda cell, fill='D9E1F2', blue=False: (setattr(cell, 'fill', _fill(fill)), setattr(cell, 'border', B), setattr(cell, 'alignment', CEN),
                                                    setattr(cell, 'font', Font(name='Arial', size=10, bold=True, color='0000FF' if blue else '000000')))
    # ---- الأزواج: UN NO | INV NO
    last_c = 52 + 2 * len(groups) - 1
    ws.merge_cells(start_row=h1, start_column=52, end_row=h1 + 1, end_column=last_c)
    ws.cell(h1, 52, 'طلبات الاستلام WIR UNIFIER')
    invs = {}
    for i, (nm, gv) in enumerate(groups):
        cu = 52 + 2 * i; ci = cu + 1
        ws.merge_cells(start_row=h3, start_column=cu, end_row=h3, end_column=ci)
        ws.cell(h3, cu, nm); ws.cell(h4, cu, 'رقم الطلب UN NO'); ws.cell(h4, ci, 'رقم المستخلص INV NO')
        for r in data:
            v, sv = gv.get(r, (None, None))
            if isinstance(v, str) and v.strip().isdigit(): v = int(v.strip())
            ws.cell(r, cu, v)
            iv = sorted({k for k in (inv(n) for n in nums(sv)) if k})
            if iv:
                ws.cell(r, ci, iv[0] if len(iv) == 1 else ', '.join(map(str, iv)))
                for k in iv: invs.setdefault(r, {}).setdefault(k, []).append(WORK_SHORT.get(nm, nm))
            for c in (cu, ci):
                x = ws.cell(r, c); x.border = B; x.alignment = CEN; x.font = Font(name='Arial', size=10)
            if v not in (None, ''):
                ws.cell(r, cu).fill = _fill('A9D08E'); ws.cell(r, ci).fill = _fill('A9D08E')     # الرقمين أخضر زي الاستشاري
        # رقم الطلب المتكرر في صفوف ورا بعض = خانة واحدة مدموجة (زي الاستشاري)
        rs = list(data); i0 = 0
        while i0 < len(rs):
            v = ws.cell(rs[i0], cu).value; j = i0
            while j + 1 < len(rs) and v not in (None, '') and ws.cell(rs[j + 1], cu).value == v: j += 1
            if j > i0:
                for rr in rs[i0 + 1:j + 1]: ws.cell(rr, cu).value = None
                ws.merge_cells(start_row=rs[i0], start_column=cu, end_row=rs[j], end_column=cu)
            i0 = j + 1
    for c in range(52, last_c + 1):                     # صف الإجمالي ممدود لحد أعمدة الطلبات
        x = ws.cell(total, c); x.fill = _fill('FFF2CC'); x.border = B; x.alignment = CEN; x.font = Font(name='Arial', size=10, bold=True)
    for r in range(h1, h4 + 1):
        for c in range(52, last_c + 1): hdr(ws.cell(r, c))
    for c in range(52, last_c + 1): ws.column_dimensions[CL(c)].width = 13
    # ---- متابعة المستخلصات
    last = max([k for d in invs.values() for k in d] or [1])
    n_ipc = max(6, last)
    c0 = last_c + 2
    ws.merge_cells(start_row=h1, start_column=c0, end_row=h1, end_column=c0 + n_ipc)
    ws.cell(h1, c0, 'متابعة المستخلصات')
    ws.merge_cells(start_row=h1 + 1, start_column=c0, end_row=h4, end_column=c0)
    ws.cell(h1 + 1, c0, 'رقم المستخلص (آخر مستخلص)')
    ws.merge_cells(start_row=h1 + 1, start_column=c0 + 1, end_row=h3 - 1 if h3 - 1 > h1 + 1 else h1 + 1, end_column=c0 + n_ipc)
    ws.cell(h1 + 1, c0 + 1, 'المنفذ بالمستخلص رقم')
    for k in range(1, n_ipc + 1):
        ws.merge_cells(start_row=h3, start_column=c0 + k, end_row=h4, end_column=c0 + k)
        ws.cell(h3, c0 + k, k)
    for r in range(h1, h4 + 1):
        for c in range(c0, c0 + n_ipc + 1): hdr(ws.cell(r, c), 'C4D79B', blue=(r >= h3 and c > c0))
    for r in data:
        d = invs.get(r, {})
        ws.cell(r, c0, max(d) if d else None)
        for k in range(1, n_ipc + 1):
            ws.cell(r, c0 + k, '+'.join(d[k]) if k in d else None)
        for c in range(c0, c0 + n_ipc + 1):
            x = ws.cell(r, c); x.fill = _fill('C4D79B'); x.border = B; x.alignment = CEN; x.font = Font(name='Arial', size=10, bold=c == c0)
    ws.column_dimensions[CL(c0)].width = 14
    for k in range(1, n_ipc + 1): ws.column_dimensions[CL(c0 + k)].width = 22
    ws.auto_filter.ref = f'{CL(c0)}{h4}:{CL(c0 + n_ipc)}{total - 1}'


QS_TINT = {   # ألوان شيت الاستشاري («حصر مجمع») لكل عمود في صفوف البيانات
    'A': 'E2EFDA', 'M': 'E2EFDA', 'V': 'E2EFDA', 'W': 'E2EFDA', 'X': 'E2EFDA', 'Y': 'E2EFDA', 'AA': 'E2EFDA', 'AC': 'E2EFDA',
    'AF': 'E2EFDA', 'AH': 'E2EFDA', 'AJ': 'E2EFDA', 'AP': 'E2EFDA', 'AR': 'E2EFDA', 'AU': 'E2EFDA', 'AX': 'E2EFDA', 'AY': 'E2EFDA'}
QS_BLUE = ('M', 'AG', 'AI', 'AK', 'AM')      # خانات يدوي – خط أزرق
QS_PCT = ('AG', 'AI', 'AK', 'AM')            # عناوين النسب – أخضر فاقع


def qs_consultant_colors(ws):
    """ألوان الحصر الكلي = نسخة من ألوان شيت الاستشاري: العنوان رمادي فاتح وكلامه أحمر، العناوين لبني،
    النسب أخضر فاقع، الكميات المحسوبة أخضر فاتح، الإجمالي أصفر فاتح، التوقيعات رمادي."""
    from openpyxl.styles import Font
    H = QS_HEAD_ROWS; ncol = 51
    for r in (4, 5):
        for c in range(1, ncol + 1):
            x = ws.cell(r, c); x.fill = _fill('F2F2F2')
            if x.value is not None: x.font = Font(name=x.font.name or 'Arial', size=x.font.size or 13, bold=True, color='FF0000')
    for r in range(H + 1, H + 5):
        for c in range(1, ncol + 1):
            x = ws.cell(r, c); x.fill = _fill('66FF66' if r == H + 4 and CL(c) in QS_PCT else 'D9E1F2')
    total = next((r for r in range(H + 5, ws.max_row + 1) if ws.cell(r, 1).value == 'الإجمالي'), None)
    if not total: return
    for r in range(H + 5, total):
        for c in range(1, ncol + 1):
            L = CL(c); x = ws.cell(r, c)
            if L in QS_TINT: x.fill = _fill(QS_TINT[L])
            elif x.fill is not None and x.fill.fill_type: x.fill = _fill('FFFFFF')
            if L in QS_BLUE: x.font = x.font.copy(color='0000FF')
    for c in range(1, ncol + 1):
        if CL(c) not in ('A', 'M'): ws.cell(total, c).fill = _fill('FFF2CC')
    for r in range(total + 1, ws.max_row + 1):
        for c in range(1, ncol + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and (v.startswith('المقاول :') or v.startswith('الاستشاري :')):
                for cc in range(1, ncol + 1):
                    y = ws.cell(r, cc)
                    if y.fill is not None and y.fill.fill_type: y.fill = _fill('D0CECE')
                break


# ---------------------------------------------------------------- VO مضغوط (ورق أقل) + ترقيم الصفحات
SIGNER_NAMES = {'L3': 'م/ (اسم المهندس)', 'L4': 'م/ (اسم المهندس)', 'L6': 'م/ (اسم المهندس)', 'L7': 'م/ (اسم المهندس)',
                'L8': 'م/ (اسم المهندس)', 'L10': 'م/ (اسم المهندس)', 'L11': 'م/ (اسم المهندس)'}
VO_MODE = 'all_sign'   # 'all_sign' = كل البنود والتوقيعات في كل صفحة | 'all' = كل البنود والتوقيعات في الآخر | 'live' = البنود اللي ليها كميات بس
VO_START = 44          # أول صف بعد صفحتي المبالغ والخلاصة


def compact_vo(ws, c, mode=None, no=None, date_to=None):
    """بيعيد رص صفحات البنود في VO: البنود ورا بعض، والعنوان والمرحّل (يرحل/ماقبله) بيتحطوا لوحدهم عند كل فاصل صفحة
    محسوب بارتفاع الصفوف. بيرجّع {كود البند: الصف الجديد} وعدد صفحات المستخلص (للترقيم)."""
    from copy import copy
    from openpyxl.worksheet.pagebreak import Break, RowBreak
    mode = mode or VO_MODE
    end = ws.max_row
    t = lambda r, col='A': (str(ws[f'{col}{r}'].value).strip() if ws[f'{col}{r}'].value is not None else '')
    def snap(r, two=False):
        cells = [(cc, ws.cell(r, cc).value, copy(ws.cell(r, cc)._style)) for cc in range(1, 15)]
        merges = [(m.min_col, m.max_col, m.max_row - m.min_row) for m in ws.merged_cells.ranges
                  if m.min_row == r and (m.max_row == r or two)]
        return dict(cells=cells, merges=merges, h=ws.row_dimensions[r].height or 15)
    hr = next(r for r in range(VO_START, end + 1) if t(r) == 'رقم البند')
    head = [snap(hr - 1), snap(hr, True), snap(hr + 1)]
    carry_r = next(r for r in range(VO_START, end + 1) if t(r).startswith('المجموع'))
    maq_r = next(r for r in range(VO_START, end + 1) if t(r, 'B') == 'ماقبله')
    sig0 = next(r for r in range(VO_START, end + 1) if t(r) == '=$K$2')
    sig1 = next(r for r in range(sig0, end + 1) if t(r) == 'التوقيع :')
    carry, maq = snap(carry_r), snap(maq_r)
    sig = [snap(r) for r in range(sig0, sig1 + 2) if not t(r).startswith('الأعمال منفذة')]
    divnos = {str(n) for n, _ in DIVISIONS}
    SIGS = ('=$K$2', 'الأعمال منفذة', 'الوظيفة', '="الاسم', 'التوقيع')
    content, r = [], VO_START
    while r <= end:
        a = t(r)
        if a == 'رقم البند': r += 2; continue
        if t(r, 'B') == 'ماقبله' or a.startswith('المجموع') or a.startswith(SIGS) or all(ws.cell(r, cc).value in (None, '') for cc in range(1, 15)):
            r += 1; continue
        code = ws[f'N{r}'].value
        kind = 'item' if isinstance(code, (int, float)) else 'total' if a.startswith('إجمالي') else 'div' if a in divnos else 'head'
        live = kind == 'item' and any(isinstance(ws[f'{k}{r}'].value, (int, float)) and abs(ws[f'{k}{r}'].value) > 1e-9 for k in 'IK')
        content.append(dict(snap(r), kind=kind, num=a, code=int(code) if kind == 'item' else None, live=live))
        r += 1
    if mode == 'live':
        keep_items = [x['num'] for x in content if x['kind'] == 'item' and x['live']]
        out, i = [], 0
        while i < len(content):
            x = content[i]
            if x['kind'] == 'item':
                if x['live']: out.append(x)
            elif x['kind'] == 'head':
                j = i                                         # قسم الرأس: لحد الرأس اللي بعد بند
                while j < len(content) and content[j]['kind'] == 'head': j += 1
                k = j
                while k < len(content) and content[k]['kind'] == 'item': k += 1
                sec = [y for y in content[j:k] if y['live']]
                if sec or any(n.startswith(x['num'] + '.') for n in keep_items): out.append(x)
            else:
                out.append(x)
            i += 1
        content = out
    # ---- امسح القديم
    for m in [m for m in list(ws.merged_cells.ranges) if m.min_row >= VO_START]: ws.unmerge_cells(str(m))
    blank = copy(ws.cell(1, 30)._style)
    for rr in range(VO_START, end + 1):
        for cc in range(1, 19):
            x = ws.cell(rr, cc); x.value = None; x._style = copy(blank)
        ws.row_dimensions[rr].height = None
    # ---- اكتب (الدمج في الآخر عشان صفوف العنوان المدموجة على صفين)
    pending = []
    def put(s, row, values=None):
        for cc, v, st in s['cells']:
            x = ws.cell(row, cc); x._style = copy(st)
            x.value = (values or {}).get(cc, v)
        for c1, c2, dr in s['merges']:
            if c2 > c1 or dr: pending.append((row, c1, row + dr, c2))
        ws.row_dimensions[row].height = s['h']
    def fix(v, old, new):                                 # معادلات الصف (H/J/M) بتشاور على صفها
        return re.sub(rf'([A-Z]){old}\b', lambda m: f'{m.group(1)}{new}', v) if isinstance(v, str) and v.startswith('=') else v
    titles = sum((ws.row_dimensions[k].height or 15) for k in range(2, 8))
    ws.page_margins.bottom = 0.35; ws.page_margins.footer = 0.12
    scale = (ws.page_setup.scale or 40) / 100
    cap = ((595.28 - (ws.page_margins.top + ws.page_margins.bottom) * 72) / scale - titles) * 0.97
    sig_h = sum(s['h'] for s in sig)
    breaks = [b.id for b in ws.row_breaks.brk if b.id < VO_START]
    rows_map, div_tot, div_span = {}, {}, {}
    state = dict(r=VO_START, used=0.0, seg=None, div=None, first=True)
    def new_page():
        if not state['first']:
            breaks.append(state['r'] - 1)
        state['first'] = False; state['used'] = 0.0
        for s in head:
            put(s, state['r']); state['used'] += s['h']; state['r'] += 1
        if state['seg'] is not None:                       # ماقبله = المرحّل من الصفحة اللي فاتت
            put(maq, state['r'], {13: f"=M{state['carry']}"}); state['seg'] = state['r']
            state['used'] += maq['h']; state['r'] += 1
    def close_page(with_sig):
        if state['seg'] is not None:
            put(carry, state['r'], {13: f"=SUM(M{state['seg']}:M{state['r'] - 1})"}); state['carry'] = state['r']; state['r'] += 1
        if with_sig:
            for s in sig: put(s, state['r']); state['r'] += 1
    new_page()
    for i, x in enumerate(content):
        reserve = (0 if x['kind'] == 'total' else carry['h']) + (sig_h if mode == 'all_sign' else 0)
        need = x['h'] + (content[i + 1]['h'] if x['kind'] in ('div', 'head') and i + 1 < len(content) else 0)   # الرأس ميتسابش لوحده آخر الصفحة
        if state['used'] + need + reserve > cap and state['used'] > sum(s['h'] for s in head) + 1:
            close_page(mode == 'all_sign'); new_page()
        row = state['r']
        vals = {cc: fix(v, None, None) for cc, v, _ in []}
        if x['kind'] == 'item':
            old = None
            for cc, v, _ in x['cells']:
                if isinstance(v, str) and v.startswith('='):
                    m = re.search(r'[A-Z](\d+)', v)
                    if m: old = m.group(1); break
            vals = {cc: (re.sub(rf'([A-Z]){old}(?!\d)', lambda m: f'{m.group(1)}{row}', v) if old and isinstance(v, str) and v.startswith('=') else v)
                    for cc, v, _ in x['cells']}
            rows_map[x['code']] = row
        elif x['kind'] == 'div':
            state['div'] = int(x['num']); state['seg'] = row; div_span[state['div']] = [row, row]
        elif x['kind'] == 'total':
            seg = state['seg'] if state['seg'] is not None else row
            vals = {13: f'=SUM(M{seg}:M{row - 1})'}
            if state['div'] is not None:
                div_tot[state['div']] = row; div_span[state['div']][1] = row
            state['seg'] = None
        put(x, row, vals)
        state['used'] += x['h']; state['r'] += 1
    # التوقيعات في الآخر
    if mode == 'all_sign':
        close_page(True)
    else:
        if state['used'] + sig_h > cap: breaks.append(state['r'] - 1)
        for s in sig: put(s, state['r']); state['r'] += 1
    last = state['r'] - 1
    for r1, c1, r2, c2 in pending: ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    # ---- صفحة الخلاصة (صفوف الأقسام) تشاور على الأماكن الجديدة
    divs = {d['no']: d for d in c['divs']}
    for rr in range(8, VO_START):
        m = re.match(r'رقم \((\d+)\)', t(rr))
        if not m: continue
        n = int(m.group(1))
        if n in div_tot:
            a, b = div_span[n]
            ws[f'F{rr}'] = round(divs[n]['contract'], 2) if n in divs else ws[f'F{rr}'].value
            ws[f'I{rr}'] = f'=SUMPRODUCT(I{a}:I{b},L{a}:L{b})'
            ws[f'M{rr}'] = f'=M{div_tot[n]}'
    # ---- الطباعة والترقيم
    ws.row_breaks = RowBreak()
    for b in sorted(set(breaks)): ws.row_breaks.append(Break(id=b))
    a0 = str(ws.print_area).split('!')[-1].replace('$', '').split(':')[0] if ws.print_area else 'A8'
    ws.print_area = f'{a0}:M{last}'
    pages = 1 + len(ws.row_breaks.brk) + 1             # الغلاف + صفحات VO
    return rows_map, pages, last


def number_pages(wb, no, date_to, pages):
    """رقم الصفحة تحت على اليمين في الغلاف وVO: «مستخلص رقم 01 – حتى 31-10-2026 – صفحة 3 من 9»."""
    d = date_to.strftime('%d-%m-%Y') if hasattr(date_to, 'strftime') else str(date_to)
    for i, name in enumerate((COVER, IPC)):
        if name not in wb.sheetnames: continue
        ws = wb[name]
        if i == 1:
            ws.page_setup.firstPageNumber = 2; ws.page_setup.useFirstPageNumber = True
        else:
            ws.page_margins.bottom = max(ws.page_margins.bottom or 0, 0.35); ws.page_margins.footer = 0.12
        f = ws.oddFooter.left if ws.sheet_view.rightToLeft else ws.oddFooter.right    # الشيت يمين لشمال: إكسيل بيعكس أجزاء الفوتر
        f.text = f'مستخلص رقم {no:02d} – حتى {d} – صفحة &P من {pages}'
        f.size = 10



def tidy_signatures(ws):
    """التوقيعات في VO: بتتكتب الوظيفة والاسم بس – من غير كلمات «الوظيفة :» و«الاسم :» و«التوقيع :» (سطر التوقيع بيفضل فاضي للإمضا)،
    وسطر «الأعمال منفذة طبقاً لشروط ومواصفات العقد» بيتخفى."""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=13):
        for x in row:
            v = x.value
            if not isinstance(v, str): continue
            w = v.strip()
            if w.startswith('الوظيفة'):
                x.value = re.sub(r'^الوظيفة\s*:\s*', '', w)
            elif w.startswith('=') and '"الاسم : "&' in w.replace('"الاسم :"&', '"الاسم : "&'):
                x.value = w.replace('"الاسم : "&', '').replace('"الاسم :"&', '')
            elif w.startswith('التوقيع'):
                x.value = None
            elif w.startswith('الأعمال منفذة'):
                ws.row_dimensions[x.row].hidden = True