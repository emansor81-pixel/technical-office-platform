# -*- coding: utf-8 -*-
"""Materials | MAR و MIR والمخزن.
- بيقرا 00_Registers\\02-LOG MAR+MIR.xlsm (MAR LOG و MIR LOG) و 00_Registers\\04-LOG STOCK.xlsx
  (المخزون، التوريدات، الصرف، القوائم) زي ما هم.
- مرحلة كل MAR و MIR من مكان ملفه (1_WIP / 2_SHARED / 3_PUBLISHED / 4_ARCHIVED) ومن حالته في السجل.
- علامة «اترفع يونيفاير» محفوظة جوه البرنامج (mat_status.json) زي تبويب الطلبات.
- الكتابة في السجلات بتتعمل من خلال إكسيل نفسه، فالمعادلات والماكرو والقوايم بتفضل زي ما هي:
  إذن التوريد ← سطر MIR + PDF الـ MIR (شيت MIR وإذن التوريد وراه) + سطور التوريدات (والمادة والشركة الجديدة في القوايم)،
  والصرف للموقع ← سطر في الصرف. لو السجل مفتوح عندك في إكسيل البرنامج بيكتب فيه ويحفظه.
"""
import os, re, sys, json, shutil, tempfile, subprocess, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
_qs = os.path.join(os.path.dirname(HERE), 'qs')
if _qs not in sys.path: sys.path.insert(0, _qs)
import qs_core as Q  # noqa: E402

MARMIR_REL = r'00_Registers\02-LOG MAR+MIR.xlsm'
STOCK_REL = r'00_Registers\04-LOG STOCK.xlsx'
MIR_DIR_REL = r'05_Quality\MIR'
MAR_DIR_REL = r'04_Submittals\MAR'
DELIVERY_REL = r'09_Procurement_Materials\DELIVERY'
STAGES = ('1_WIP', '2_SHARED', '3_PUBLISHED', '4_ARCHIVED')
MARKS = os.path.join(HERE, 'mat_status.json')         # علامة «اترفع يونيفاير»
PENDING = os.path.join(HERE, 'mat_pending.json')      # سطور مخزن ماتكتبتش (إكسيل كان مشغول) – بتتكتب بعدين
NOWIN = 0x08000000 if os.name == 'nt' else 0
MIR_PREFIX = 'CLIENT-CTR-MIR-'
UNITS = ('م.ط', 'عدد', 'رول', 'علبة', 'طقم', 'م²', 'م³', 'طن', 'كجم', 'لتر')
IN_STATUS = ('وصل', 'لم يصل')
BAD = r'[\\/:*?"<>|]+'
NAME_MAX = 90                                        # أقصى طول لاسم ملف الـ MIR (من غير .pdf)

S_STOCK, S_DELIV, S_ISSUE, S_LISTS = 'المخزون', 'التوريدات', 'الصرف', 'القوائم'
S_MARLOG, S_MIRLOG, S_MIRFORM = 'MAR LOG', 'MIR LOG', 'MIR'


def P(rel): return Q.project_path(rel)
def safe(n): return re.sub(r'\s+', ' ', re.sub(BAD, '-', str(n or ''))).strip(' .-')


def free_name(folder, name):
    base, ext = os.path.splitext(name)
    p, i = os.path.join(folder, name), 2
    while os.path.exists(p):
        p = os.path.join(folder, f'{base} ({i}){ext}'); i += 1
    return p


def _s(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime('%d-%m-%Y')
    return str(v).strip()


def _d(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    return None


def _n(v):
    if v in (None, ''): return None
    try: return float(v)
    except (TypeError, ValueError): return None


def num_text(v):
    if v is None: return ''
    return f'{v:,.0f}' if abs(v - round(v)) < 1e-9 else f'{v:,.2f}'


# ---------------------------------------------------------------- القراية
_CACHE = {}


def _load(rel, reader):
    """السجل بيتقري من نسخة مؤقتة (لو مفتوح في إكسيل) – ومحفوظ لحد ما الملف يتغير."""
    import openpyxl, warnings
    src = P(rel)
    mt = os.path.getmtime(src)
    hit = _CACHE.get(rel)
    if hit and hit[0] == mt: return hit[1]
    fd, tmp = tempfile.mkstemp(prefix='mat_', suffix=os.path.splitext(src)[1]); os.close(fd)
    try:
        shutil.copy2(src, tmp)
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
        try: data = reader(wb)
        finally: wb.close()
    finally:
        try: os.remove(tmp)
        except OSError: pass
    _CACHE[rel] = (mt, data)
    return data


def _read_marmir(wb):
    mar, mir = [], []
    for i, r in enumerate(wb[S_MARLOG].iter_rows(min_row=9, max_col=15, values_only=True), 9):
        no = _s(r[0])
        if not re.fullmatch(r'M\d+', no) or not (_s(r[1]) or _s(r[3])): continue
        mar.append(dict(row=i, no=no, n=int(no[1:]), ref=_s(r[1]), rev=_s(r[2]), desc=_s(r[3]), supplier=_s(r[4]).strip('\t '),
                        disc=_s(r[5]), spec=_s(r[6]), sent=_d(r[7]), to_client=_d(r[8]), reply=_d(r[9]), days=_s(r[10]),
                        status=_s(r[11]), remarks=_s(r[12]), factory=_s(r[13]), prev=_s(r[14])))
    for i, r in enumerate(wb[S_MIRLOG].iter_rows(min_row=9, max_col=16, values_only=True), 9):
        no = _s(r[0]); m = re.search(r'MIR\D*(\d+)', no)
        if not m: continue
        mir.append(dict(row=i, no=no, n=int(m.group(1)), rev=_s(r[1]), date=_d(r[2]), desc=_s(r[3]), origin=_s(r[4]),
                        model=_s(r[5]), serial=_s(r[6]), qty=_s(r[7]), unit=_s(r[8]), mat=_s(r[9]), appr=_d(r[10]),
                        supplier=_s(r[11]), boq=_s(r[12]), disc=_s(r[13]), status=_s(r[14]), pdf=_s(r[15])))
    return dict(mar=mar, mir=mir)


def _read_stock(wb):
    items, dels, iss, comp, types = [], [], [], [], []
    for i, r in enumerate(wb[S_STOCK].iter_rows(min_row=3, max_col=12, values_only=True), 3):
        if not _s(r[1]): continue
        items.append(dict(row=i, name=_s(r[1]), code=_s(r[2]), unit=_s(r[3]), type=_s(r[4]), need=_n(r[5]), got=_n(r[6]),
                          wait=_n(r[7]), out=_n(r[8]), bal=_n(r[9]), left=_n(r[10]), status=_s(r[11])))
    for i, r in enumerate(wb[S_DELIV].iter_rows(min_row=3, max_col=16, values_only=True), 3):
        if not _s(r[6]): continue
        dels.append(dict(row=i, tr=_s(r[1]), date=_d(r[2]), dn=_s(r[3]), company=_s(r[4]), type=_s(r[5]), name=_s(r[6]),
                         code=_s(r[7]), unit=_s(r[8]), qty=_n(r[9]), status=_s(r[10]), mir=_s(r[11]), note=_s(r[12]), dl=_s(r[15])))
    for i, r in enumerate(wb[S_ISSUE].iter_rows(min_row=3, max_col=9, values_only=True), 3):
        if not _s(r[2]): continue
        iss.append(dict(row=i, date=_d(r[1]), name=_s(r[2]), code=_s(r[3]), unit=_s(r[4]), qty=_n(r[5]),
                        route=_s(r[6]), who=_s(r[7]), note=_s(r[8])))
    for r in wb[S_LISTS].iter_rows(min_row=2, max_col=2, values_only=True):
        if _s(r[0]): comp.append(_s(r[0]))
        if _s(r[1]): types.append(_s(r[1]))
    return dict(items=items, deliveries=dels, issues=iss, companies=comp, types=types)


def load_marmir(): return _load(MARMIR_REL, _read_marmir)
def load_stock(): return _load(STOCK_REL, _read_stock)


def mir_title(no):
    """اسم ملف الـ MIR زي ما هو في MIR LOG (عمود «اسم ملف الـ PDF») – None لو الرقم مش في السجل."""
    try:
        for m in load_marmir()['mir']:
            if m['n'] == int(no): return safe(m['pdf'] or f"{m['no']} - {m['desc']}")
    except Exception:
        return None
    return None


# ---------------------------------------------------------------- المراحل
def folder_files(rel):
    root, out = P(rel), {}
    for st in STAGES:
        d = os.path.join(root, st)
        try:
            out[st] = sorted(n for n in os.listdir(d) if os.path.isfile(os.path.join(d, n))
                             and not n.startswith(('~', '.')) and n.lower() != 'readme.txt')
        except OSError:
            out[st] = []
    return out


def _starts(name, key):
    return name.upper().startswith(key.upper()) and not name[len(key):len(key) + 1].isdigit()


def mir_files(m, files):
    return [(st, n) for st in STAGES for n in files.get(st, []) if _starts(n, m['no'])]


def mar_files(m, files):
    pat = re.compile(rf'^M\s*{m["n"]}(?!\d)', re.I)
    hits = [(st, n) for st in STAGES for n in files.get(st, []) if pat.match(n)]
    if m.get('ref'):              # ملفين بنفس رقم M ← اللي فيه رقم تقديم تاني مش بتاعه
        hits = [h for h in hits if m['ref'] in h[1] or not re.search(r'CTR-[A-Z]+-WA-\d', h[1])]
    return hits


MIR_STAGES = ('Not prepared', 'Waiting signature', 'Signed', 'Approved', 'Rejected')
MAR_STAGES = ('Preparing', 'Submitted', 'Approved', 'Resubmit', 'Rejected')


def mir_stage(m, files):
    st = (m.get('status') or '').lower(); where = {s for s, _ in mir_files(m, files)}
    if 'reject' in st: return 'Rejected'
    if 'approved' in st or '3_PUBLISHED' in where: return 'Approved'
    if '2_SHARED' in where: return 'Signed'
    if '1_WIP' in where: return 'Waiting signature'
    return 'Not prepared'


def mar_stage(m, files):
    st = (m.get('status') or '').lower(); where = {s for s, _ in mar_files(m, files)}
    if 'reject' in st: return 'Rejected'
    if 'resubmit' in st and 'no resubmission' not in st: return 'Resubmit'
    if 'approved' in st or '3_PUBLISHED' in where: return 'Approved'
    if m.get('sent') or '2_SHARED' in where: return 'Submitted'
    return 'Preparing'


def best_file(files_hits):
    """أحدث مرحلة فيها ملف (المعتمد قبل المتوقّع قبل اللي بيتجهز)."""
    for st in ('3_PUBLISHED', '2_SHARED', '1_WIP', '4_ARCHIVED'):
        for s, n in files_hits:
            if s == st: return s, n
    return None


# ---------------------------------------------------------------- علامة اليونيفاير
def mkey(kind, no): return f'{kind}|{no}'


def load_marks():
    try:
        with open(MARKS, encoding='utf-8') as f: return json.load(f)
    except Exception:
        return {}


def save_marks(m):
    tmp = MARKS + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(m, f, ensure_ascii=False, indent=1)
    os.replace(tmp, MARKS)


# ---------------------------------------------------------------- مساعدات التوريد
def next_mir(mirs):
    return f'{MIR_PREFIX}{max([m["n"] for m in mirs] or [0]) + 1:03d}'


def dn_file_name(company, dn, deliveries):
    """زي عمود DL في التوريدات: رقم توريد الشركة دي (01، 02…) - الشركة - رقم الإذن."""
    c = (company or '').strip().lower()
    seen = []
    for d in deliveries:
        if d['company'].strip().lower() == c and d['dn'] and d['dn'] not in seen: seen.append(d['dn'])
    k = seen.index(dn) + 1 if dn in seen else len(seen) + 1
    return safe(f'{k:02d}-{company.strip()}-{dn}')


_AR = str.maketrans('أإآةىئؤ', 'اااهييو')
STOP = {'شركه', 'شركة', 'للتجاره', 'التجاريه', 'مؤسسه', 'للمقاولات', 'المحدوده', 'مصنع', 'company', 'trading', 'the', 'and',
        'ltd', 'est', 'factory', 'co', 'for', 'with', 'من', 'الي', 'على', 'في', 'مم', 'قطر'}


def _toks(s):
    s = (s or '').lower().translate(_AR)
    return {w for w in re.findall(r'[a-zء-ي]{3,}', s)} - STOP


def match_mar(company, texts, mars):
    """أقرب اعتماد مواد للتوريد ده: نفس المورد + كلام الوصف. يرجّع (الأفضل أو None، القايمة مترتبة)."""
    ct = _toks(company)
    tt = set()
    for t in texts or (): tt |= _toks(t)
    def score(m):
        return 3 * len(ct & _toks(m['supplier'] + ' ' + m['factory'])) + 2 * len(tt & _toks(m['desc'])) + \
            (1 if 'approved' in m['status'].lower() else 0)
    ranked = sorted(mars, key=lambda m: (-score(m), m['n']))
    best = ranked[0] if ranked and score(ranked[0]) >= 3 else None
    return best, ranked


def mar_from_history(company, names, deliveries, mirs, mars):
    """الـ MAR اللي اتعمل عليه MIR قبل كده لنفس الصنف من نفس المورد (التوريدات ← رقم الـ MIR ← عمود MAT في MIR LOG)."""
    c = (company or '').strip().lower(); names = set(names or ())
    by_no = {m['no'].strip().upper(): m for m in mirs}
    count = {}
    for d in deliveries:
        if not d.get('mir') or d['name'] not in names or d['company'].strip().lower() != c: continue
        mir = by_no.get(d['mir'].strip().upper())
        ref = (mir or {}).get('mat', '').strip()
        if ref: count[ref] = count.get(ref, 0) + 1
    for ref in sorted(count, key=lambda r: -count[r]):
        hit = next((m for m in mars if m['ref'].strip().upper() == ref.upper() or m['no'].upper() == ref.upper()), None)
        if hit: return hit
    return None


def mar_label(m): return f"{m['no']} – {m['ref']} – {m['desc'][:60]}".strip(' –')


# ---------------------------------------------------------------- الكتابة من خلال إكسيل
PS_XL = r"""
$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -Encoding UTF8 $env:MAT_JOB | ConvertFrom-Json
$res = @{ rows = @(); name = ''; pdf = ''; error = ''; saved = $false }
$xl = $null; $wb = $null; $mine = $false; $openedHere = $false
function Norm($s) { return ([string]$s -replace '\s', '').ToUpper() }
function Sheet($name) {
  foreach ($s in $wb.Worksheets) { if ((Norm ($s.Name)) -eq (Norm ($name))) { return $s } }
  throw ('Sheet not found: ' + $name)
}
function LastRow($ws, $col) { return $ws.Range($col + '1048576').End(-4162).Row }
function FreeRow($ws, $first, $cols) {
  $last = [int]$first
  foreach ($c in $cols) { $r = LastRow $ws $c; if ($r -gt $last) { $last = $r } }
  for ($r = [int]$first; $r -le $last + 1; $r++) {
    $busy = $false
    foreach ($c in $cols) { if ([string]$ws.Range($c + $r).Formula -ne '') { $busy = $true; break } }
    if (-not $busy) { return $r }
  }
  return $last + 1
}
function Put($ws, $addr, $c) {
  $cell = $ws.Range($addr)
  if ($c.kind -eq 'date') {
    $cell.Value2 = [DateTime]::ParseExact([string]$c.value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).ToOADate()
    if ([string]$cell.NumberFormat -eq 'General') { $cell.NumberFormat = 'dd/mm/yyyy' }
  }
  elseif ($c.kind -eq 'num') { $cell.Value2 = [double]$c.value }
  else { $cell.Value2 = [string]$c.value }
}
try {
  try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $xl = $null }
  $fname = [IO.Path]::GetFileName($job.path)
  if ($xl) { foreach ($w in $xl.Workbooks) { if ($w.FullName -eq $job.path -or $w.Name -eq $fname) { $wb = $w } } }
  if (-not $wb) {
    $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $mine = $true
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($job.path, 0); $openedHere = $true
    if ($wb.ReadOnly) { throw 'READONLY' }
  }
  # 1) everything is found first, so nothing is half written
  $plan = @()
  foreach ($op in $job.ops) { $plan += @{ op = $op; ws = (Sheet $op.sheet) } }
  if ($job.export) { $null = Sheet $job.export.sheet }
  # 2) write
  foreach ($p in $plan) {
    $op = $p.op; $ws = $p.ws
    if ($op.op -eq 'ensure') {
      $found = 0; $last = LastRow $ws $op.col
      for ($r = [int]$op.first; $r -le $last; $r++) { if ((Norm ($ws.Range($op.col + $r).Text)) -eq (Norm ($op.value))) { $found = $r; break } }
      if ($found -eq 0) {
        $r = FreeRow $ws $op.first @($op.col)
        $ws.Range($op.col + $r).Value2 = [string]$op.value
        foreach ($c in $op.cells) { if ([string]$ws.Range($c.col + $r).Formula -eq '') { Put $ws ($c.col + $r) $c } }
        $res.rows += ('ADD ' + $op.tag + ' ' + $r)
      } else { $res.rows += ('HAS ' + $op.tag + ' ' + $found) }
    }
    elseif ($op.op -eq 'append') {
      $r = FreeRow $ws $op.first $op.check
      foreach ($c in $op.cells) {
        if ([string]$ws.Range($c.col + $r).Formula -ne '' -and -not $c.force) { continue }
        Put $ws ($c.col + $r) $c
      }
      $res.rows += ('ROW ' + $op.tag + ' ' + $r)
    }
    elseif ($op.op -eq 'set') {
      foreach ($c in $op.cells) { Put $ws $c.addr $c }
      $res.rows += ('SET ' + $op.tag)
    }
  }
  $xl.Calculate()
  $wb.Save(); $res.saved = $true
  # 3) PDF of a form sheet (MIR) after choosing its number
  if ($job.export) {
    $e = $job.export; $ws = Sheet $e.sheet
    foreach ($c in $e.set) { $ws.Range($c.cell).Value2 = [string]$c.value }
    $xl.Calculate()
    $name = [string]$ws.Range($e.name_cell).Value2
    if (-not $name -or $name -like '#*') { $name = [string]$e.fallback }
    $res.name = $name
    $ws.ExportAsFixedFormat(0, $e.pdf, 0, $true, $false)
    $res.pdf = $e.pdf
    $wb.Save()
  }
} catch {
  $msg = [string]$_.Exception.Message
  if ($msg -like '*READONLY*') { $msg = 'READONLY' }
  $res.error = $msg
} finally {
  if ($openedHere -and $wb) { try { $wb.Close($false) } catch {} }
  if ($mine -and $xl) { try { $xl.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}
$res | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $env:MAT_OUT
"""


class XlError(RuntimeError):
    pass


def xl_run(rel, ops, export=None, timeout=300):
    """يكتب في سجل إكسيل من خلال إكسيل نفسه. يرجّع dict (rows, name, pdf, saved). أي مشكلة ← XlError برسالة واضحة."""
    path = os.path.abspath(P(rel))
    base = os.path.basename(path)
    if os.name != 'nt': raise XlError('Writing to the Excel registers needs Excel on Windows')
    if not os.path.isfile(path): raise XlError(f'Not found: {path}')
    tmpd = tempfile.mkdtemp(prefix='mat_xl_')
    job, out, ps = (os.path.join(tmpd, n) for n in ('job.json', 'out.json', 'mat_xl.ps1'))
    with open(job, 'w', encoding='utf-8') as f: json.dump(dict(path=path, ops=ops, export=export), f, ensure_ascii=False)
    with open(ps, 'w', encoding='utf-8-sig') as f: f.write(PS_XL)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps],
                           env=dict(os.environ, MAT_JOB=job, MAT_OUT=out), capture_output=True, timeout=timeout, creationflags=NOWIN)
        try:
            with open(out, encoding='utf-8-sig') as f: res = json.load(f)
        except Exception:
            err = (r.stderr or b'').decode('utf-8', 'replace').strip() or (r.stdout or b'').decode('utf-8', 'replace').strip()
            raise XlError(f'Excel did not answer for {base}: {err[-300:] or "no reply"}')
        if res.get('error'):
            e = res['error']
            if e == 'READONLY': raise XlError(f'{base} is open somewhere else (another Excel window or computer) – close it there and try again.')
            raise XlError(f'{base}: {e[-300:]}')
        rows = res.get('rows') or []
        res['rows'] = [rows] if isinstance(rows, str) else rows
        _CACHE.pop(rel, None)
        return res
    except subprocess.TimeoutExpired:
        raise XlError(f'Excel took too long with {base} – check that Excel is not waiting for an answer (a message), then try again.')
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


def merge_pdfs(parts, out):
    import pypdf
    w = pypdf.PdfWriter()
    for p in parts:
        for pg in pypdf.PdfReader(p).pages: w.add_page(pg)
    tmp = out + '.tmp'
    with open(tmp, 'wb') as fh: w.write(fh)
    os.replace(tmp, out)
    return out


# ---------------------------------------------------------------- اللي بيتكتب
def _c(col, value, kind='text'): return dict(col=col, kind=kind, value=value)


def delivery_ops(d):
    """d: date, dn, inv, company, type, items[{name, code, qty, unit, status, note}], mir_no."""
    ops = [dict(op='ensure', sheet=S_LISTS, col='A', first=2, value=d['company'], cells=[], tag='company')]
    if d.get('type'): ops.append(dict(op='ensure', sheet=S_LISTS, col='B', first=2, value=d['type'], cells=[], tag='type'))
    for it in d['items']:
        cells = [c for c in (_c('D', it.get('unit') or ''), _c('E', d.get('type') or '')) if c['value']]
        ops.append(dict(op='ensure', sheet=S_STOCK, col='B', first=3, value=it['name'], cells=cells, tag='item'))
    for it in d['items']:
        cells = [_c('C', d['date'].isoformat(), 'date'), _c('D', d['dn']), _c('E', d['company']), _c('G', it['name']),
                 _c('J', it['qty'], 'num'), _c('K', it.get('status') or IN_STATUS[0])]
        if d.get('type'): cells.append(_c('F', d['type']))
        if d.get('mir_no'): cells.append(_c('L', d['mir_no']))
        inv = d.get('inv') or ''
        note = ' – '.join(x for x in (it.get('code'), it.get('note'), inv if inv.upper().startswith('INV') else (f'INV {inv}' if inv else '')) if x)
        if note: cells.append(_c('M', note))
        ops.append(dict(op='append', sheet=S_DELIV, first=3, check=['C', 'D', 'G', 'J'], cells=cells, tag='delivery'))
    return ops


def mir_ops(m):
    q = _n(m.get('qty'))
    cells = [_c('A', m['no']), _c('B', _n(m.get('rev')) or 1, 'num'), _c('C', m['date'].isoformat(), 'date'), _c('D', m['desc']),
             _c('E', m.get('origin', '')), _c('F', m.get('model', '')), _c('G', m.get('serial', '')),
             _c('H', q, 'num') if q is not None else _c('H', m.get('qty', '')), _c('I', m.get('unit', '')), _c('J', m.get('mat', ''))]
    return [dict(op='append', sheet=S_MIRLOG, first=9, check=['A'], cells=[c for c in cells if str(c['value']).strip()], tag='mir')]


def issue_ops(x):
    cells = [_c('B', x['date'].isoformat(), 'date'), _c('C', x['name']), _c('F', x['qty'], 'num')]
    for col, k in (('G', 'route'), ('H', 'who'), ('I', 'note')):
        if x.get(k): cells.append(_c(col, x[k]))
    return [dict(op='append', sheet=S_ISSUE, first=3, check=['B', 'C', 'F'], cells=cells, tag='issue')]


def make_mir(m, attach=None):
    """سطر الـ MIR في MIR LOG + PDF من شيت MIR، وإذن التوريد وراه ← 05_Quality\\MIR\\1_WIP. يرجّع مسار الـ PDF."""
    tmpd = tempfile.mkdtemp(prefix='mat_mir_')
    form = os.path.join(tmpd, 'mir.pdf')
    try:
        res = xl_run(MARMIR_REL, mir_ops(m), export=dict(sheet=S_MIRFORM, set=[dict(cell='D6', value=m['no'])], name_cell='K2',
                                                          fallback=f"{m['no']} - {m['desc']}", pdf=form))
        if not os.path.isfile(form): raise XlError('Excel saved the MIR row but did not make the PDF – use «Save MIR as PDF» in the sheet')
        folder = os.path.join(P(MIR_DIR_REL), '1_WIP'); os.makedirs(folder, exist_ok=True)
        base = safe(res.get('name') or f"{m['no']} - {m['desc']}")
        if len(base) > NAME_MAX: base = base[:NAME_MAX].rstrip(' .-–+')       # اسم طويل ← المسار يعدّي حد ويندوز (260)
        out = free_name(folder, base + '.pdf')
        merge_pdfs([form] + [a for a in (attach or []) if a and os.path.isfile(a)], out)
        return out
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


def create_delivery(d, dn_pdf=None):
    """إذن التوريد اتحفظ: (1) الـ MIR لو مطلوب، (2) سطور التوريدات. يرجّع (رسايل تمام، مشاكل، مسار الـ MIR)."""
    ok, bad, mir_path = [], [], None
    if d.get('mir'):
        try:
            mir_path = make_mir(d['mir'], [dn_pdf])
            d['mir_no'] = d['mir']['no']
            ok.append(f"{d['mir']['no']} ready for signature (05_Quality\\MIR\\1_WIP)")
        except Exception as e:
            d['mir_no'] = ''
            bad.append(f'MIR not made: {e}')
    try:
        xl_run(STOCK_REL, delivery_ops(d))
        ok.append(f"{len(d['items'])} item(s) added to 04-LOG STOCK")
    except Exception as e:
        add_pending('delivery', d)
        bad.append(f'Stock not written yet: {e} – it waits in Materials («Write waiting rows»)')
    return ok, bad, mir_path


def mir_for_delivery(rows, m, dn_pdf=None):
    """MIR لتوريد متسجل قبل كده: سطر + PDF، ورقمه يتكتب قدام أصناف التوريد في التوريدات."""
    path = make_mir(m, [dn_pdf])
    xl_run(STOCK_REL, [dict(op='set', sheet=S_DELIV, cells=[dict(addr=f'L{r}', kind='text', value=m['no']) for r in rows], tag='mir')])
    return path


def write_issue(x):
    return xl_run(STOCK_REL, issue_ops(x))


def archive_unsigned(no, keep_path):
    """الـ MIR رجع متوقّع ← النسخة اللي من غير توقيع في 1_WIP تروح 4_ARCHIVED. يرجّع الملفات اللي اتنقلت."""
    key = f'{MIR_PREFIX}{int(no):03d}'
    wip, arch = (os.path.join(P(MIR_DIR_REL), s) for s in ('1_WIP', '4_ARCHIVED'))
    moved = []
    if not os.path.isdir(wip): return moved
    for n in os.listdir(wip):
        p = os.path.join(wip, n)
        if os.path.isfile(p) and _starts(n, key) and os.path.abspath(p) != os.path.abspath(keep_path or ''):
            os.makedirs(arch, exist_ok=True)
            dst = free_name(arch, n); shutil.move(p, dst); moved.append(dst)
    return moved


def find_dn_file(dl):
    """ملف إذن التوريد في DELIVERY باسم عمود DL."""
    if not dl: return None
    d = P(DELIVERY_REL)
    try:
        for n in os.listdir(d):
            if os.path.splitext(n)[0].lower() == safe(dl).lower() and n.lower().endswith('.pdf'): return os.path.join(d, n)
    except OSError:
        pass
    return None


# ---------------------------------------------------------------- السطور اللي مستنية
def _jsonable(d):
    return json.loads(json.dumps(d, default=lambda o: o.isoformat() if isinstance(o, (dt.date, dt.datetime)) else str(o)))


def load_pending():
    try:
        with open(PENDING, encoding='utf-8') as f:
            L = json.load(f)
        return L if isinstance(L, list) else []
    except Exception:
        return []


def _save_pending(L):
    tmp = PENDING + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f: json.dump(L, f, ensure_ascii=False, indent=1)
    os.replace(tmp, PENDING)


def add_pending(kind, data):
    _save_pending(load_pending() + [dict(kind=kind, data=_jsonable(data), at=f'{dt.datetime.now():%Y-%m-%d %H:%M}')])


def flush_pending():
    """يكتب السطور اللي اتأجلت. يرجّع (اتكتب كام، رسالة آخر مشكلة)."""
    L, left, done, err = load_pending(), [], 0, ''
    for p in L:
        d = dict(p['data'])
        try:
            d['date'] = dt.date.fromisoformat(str(d['date'])[:10])
            xl_run(STOCK_REL, delivery_ops(d) if p['kind'] == 'delivery' else issue_ops(d))
            done += 1
        except Exception as e:
            left.append(p); err = str(e)
    _save_pending(left)
    return done, err
