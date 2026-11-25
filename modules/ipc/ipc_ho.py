# -*- coding: utf-8 -*-
"""ربط شيت الاستشاري بالبرنامج.
البرنامج بيكتب في «حصر مجمع» الخانات اليدوي بس (القطاع، المحطات، العمق، القطر، السُمك، النسب، أرقام الطلبات)،
ومعادلات الاستشاري هي اللي بتحسب. كل كمية مربوطة برقم الطلب بتاعها: مفيش رقم طلب ← الكمية صفر.
نفس الخطوات بتتعمل على شيتنا (OUR_SHEET) الأول وبعدين على شيت الاستشاري أونلاين (consultant_sheet)،
والصفوف بتزيد لوحدها فوق صف الإجمالي بمعادلات صف 11. الإكسيل نفسه هو اللي بيفتح ويحسب ويحفظ."""
import os, sys, json, shutil, tempfile, subprocess, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'qs'))
import qs_core  # noqa: E402

SETTINGS = os.path.join(HERE, 'ipc_settings.json')
DEFAULT_HO = 'https://d.docs.live.net/93505DDF9DAC9A38/consultant_sheet.xlsx'          # شيت الاستشاري أونلاين
DEFAULT_LOCAL = r'08_Cost_Payments\CLIENT_IPC\CLIENT_IPC\OUR_SHEET.xlsx'        # شيتنا (جوه فولدر المشروع)
FIRST_ROW = 11
# تحديث شيت الاستشاري أونلاين متوقف (2026-11-02): الزرار بيحدّث OUR_SHEET بس. True = يرجع يحدّث الاتنين
HO_ENABLED = True
QS_SHEET, SUM_SHEET = 'حصر مجمع', 'خلاصة حصر'

# الخانات اليدوي في كل صف
INPUT_COLS = ['B', 'C', 'E', 'F', 'H', 'I', 'M', 'N', 'O', 'Q', 'Z', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AT', 'AW']
KEEP_COLS = ['D', 'G']          # رقم الرخصة ونوع الغرفة: لو اتكتبوا يدوي بيفضلوا مع نفس المقطع
# أعمدة رقم الطلب (UN) – رقم المستخلص (INV) هو العمود اللي بعده
UN_COLS = ['BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ', 'BS', 'BU', 'BW', 'BY', 'CA', 'CC', 'CE', 'CG',
           'CI', 'CK', 'CM', 'CO', 'CQ', 'CS']          # آخر 6 = الردم طبقة 7..12 (البرنامج بيزوّدهم بعد الطبقة السطحية)
# طبقات الردم BF-1..BF-12 ← أعمدتها في شيت الاستشاري
BF_COLS = ['BM', 'BO', 'BQ', 'BS', 'BU', 'BW', 'CI', 'CK', 'CM', 'CO', 'CQ', 'CS']
NEW_LAYERS = list(range(7, 13))
# الكمية ← رقم الطلب اللي بيشغّلها (الفرشة تحت وحوالين الماسورة بند واحد بيتصرف مع طلب «الفرشة حول الماسورة»)
GATES = {'V': 'BC', 'W': 'BC', 'X': 'BC', 'Y': 'BC', 'AA': 'BE', 'AE': 'BK',     # الفرشة أسفل ← BE (BED-D-P)، حول ← BK (COV-P)
          'AF': 'BG', 'AH': 'BG',
         'AN': 'BM', 'AP': 'BY', 'AR': 'CA', 'AU': 'CC', 'AX': 'CE', 'AY': 'CG'}
# أسماء الأحياء في اللوج (إنجليزي) → زي ما مكتوبة في شيت الاستشاري (عشان ملخص الأحياء يشتغل)
ZONE_AR = {'ZONE A': 'المنطقة أ', 'ZONE B': 'المنطقة ب', 'ZONE C': 'المنطقة ج', 'ZONE D': 'المنطقة د'}   # أسماء الأحياء في اللوج → زي ما مكتوبة في شيت الاستشاري


def zone_ar(v):
    s = str(v or '').upper().replace('AL-', '').replace('AL ', '').replace('-', '').replace(' ', '')
    for k, ar in ZONE_AR.items():
        if k in s: return ar
    return v


def load_settings():
    try:
        with open(SETTINGS, encoding='utf-8') as f: d = json.load(f)
    except Exception: d = {}
    d.setdefault('HO_SHEET', DEFAULT_HO)
    d.setdefault('LOCAL_SHEET', DEFAULT_LOCAL)
    return d


def save_settings(d):
    with open(SETTINGS, 'w', encoding='utf-8') as f: json.dump(d, f, ensure_ascii=False, indent=1)


def is_url(p): return str(p or '').lower().startswith(('http://', 'https://'))


def full_path(p):
    if is_url(p) or os.path.isabs(p) or (len(p) > 1 and p[1] == ':'): return p
    return qs_core.project_path(p)


def _k(v):
    if isinstance(v, (int, float)): return ('%.3f' % v).rstrip('0').rstrip('.')
    return '' if v is None else str(v)


def _r(v, n=6):
    return round(v, n) if isinstance(v, float) else v


def _to_unf(text, umap):
    import re
    if text in (None, '') or not umap: return text
    return re.sub(r'\d+', lambda m: str(umap[int(m.group(0))]) if int(m.group(0)) in umap else f'{m.group(0)}*', str(text))


def input_rows(project, items):
    """صف لكل مقطع حفر معتمد: الخانات اليدوي + أرقام الطلبات. نفس ترتيب summary_rows."""
    rows = project.summary_rows(items)
    exc = sorted([r for r in items['exc'] if r['q']['approved']],
                 key=lambda r: (qs_core.kstr(r['q']['dist']), r['q']['key'][0], r['x0']))
    S = project.S
    rb_t = qs_core.num(S.get('RB_LAYER')) or 0.15
    umap = {}
    for q in getattr(project, 'reqs_all', []) or []:
        try: umap[int(float(str(q.get('no')).strip()))] = int(float(str(q.get('unf')).strip()))
        except (TypeError, ValueError): pass
    pre = str(S.get('BF_PREFIX', 'BF-')).strip(); rpre = str(S.get('RB_PREFIX', 'RB-')).strip()
    bf_ivs = {k: project.ivs_code(f'{pre}{k}') for k in range(1, len(BF_COLS) + 1)}
    rb_ivs = {k: project.ivs_code(f'{rpre}{k}') for k in (1, 2)}
    # الأسفلت: CC = ASP-1 (الأسفلت) | CE = ASP-2-ML (الكشط) | CG = طلب الطبقة السطحية
    asp_ivs = {c: (project.ivs_code(project.code(k)) if project.code(k) else {})
               for c, k in (('CC', 'ASP1_CODE'), ('CE', 'ASP2_CODE'), ('CG', 'ASP_WEAR_CODE'))}
    wear_t = qs_core.num(S.get('WEAR_TH')) or 0.04
    out, notes = [], []
    for R, e in zip(rows, exc):
        J = R.get('J') or 0
        v = {c: None for c in INPUT_COLS}
        un = {c: '' for c in UN_COLS}
        v.update(B=zone_ar(R.get('B')), C=R.get('C'), E=_r(R.get('E')), F=_r(R.get('F')), M=R.get('M'),
                 N=_r(R.get('N')), O=_r(R.get('O')), Q=R.get('Q'))
        v['Z'] = qs_core.num(S.get('BED_ROCK' if e['soil'] == qs_core.ROCK else 'BED_OTHER'))
        v['AO'] = v['AQ'] = rb_t
        v['AT'] = _r((e.get('asph') or 0) / 100)
        v['AW'] = _r(wear_t)
        F = R.get('_f') or {}
        if J:
            for cq, cp in (('AF', 'AG'), ('AH', 'AI'), ('AJ', 'AK')):
                if R.get(cq): v[cp] = _r(R[cq] * R[cp] / J)          # النسبة على طول المقطع كله
            if R.get('AM') is not None and F.get('bf'): v['AM'] = _r(R['AM'] * F['bf']['ol'] / J)
            b = F.get('bed')
            if b and 0 < b['ol_b'] < J - 1e-6:
                notes.append(f"{R.get('C')} {R.get('E')}-{R.get('F')}: bedding below pipe approved on {b['ol_b']:g} m only of {J:g} m")
            if b and 0 < b['ol_a'] < J - 1e-6:
                notes.append(f"{R.get('C')} {R.get('E')}-{R.get('F')}: bedding around pipe approved on {b['ol_a']:g} m only of {J:g} m")
            rb = F.get('rb')
            if rb and (0 < rb['ol1'] < J - 1e-6 or 0 < rb['ol2'] < J - 1e-6):
                notes.append(f"{R.get('C')} {R.get('E')}-{R.get('F')}: road base approved on part of the section only")
            a = F.get('asp')
            if a and any(0 < a[k] < J - 1e-6 for k in ('ol1', 'ol2', 'ol3')):
                notes.append(f"{R.get('C')} {R.get('E')}-{R.get('F')}: asphalt approved on part of the section only")
        if str(R.get('Q') or '').strip().startswith('600'):
            notes.append(f"{R.get('C')} {R.get('E')}-{R.get('F')}: Ø600 line – the consultant's 600 pipe column has no formula")
        un['BC'] = qs_core.wno(R.get('AZ'))
        pp = str(R.get('BC') or '')
        if ' (T: ' in pp:
            pp, tt = pp.split(' (T: ', 1); un['BI'] = tt.rstrip(')')
        un['BG'] = pp
        un['BE'] = R.get('BA') or ''            # Sand Bedding Under Pipe ← BED-D-P
        un['BK'] = R.get('BB') or ''            # Bedding And Pipe Surround ← COV-P
        key, x0, x1 = e['q']['key'], e['x0'], e['x1']
        def hit(ivs):
            return qs_core.wjoin(no for a, b, no in ivs.get(key, []) if qs_core.overlap(a, b, x0, x1) > qs_core.EPS)
        for k, c in enumerate(BF_COLS, 1): un[c] = hit(bf_ivs[k])
        un['BY'] = hit(rb_ivs[1]); un['CA'] = hit(rb_ivs[2])
        for c, ivs in asp_ivs.items(): un[c] = hit(ivs)
        un = {c: _to_unf(x, umap) for c, x in un.items()}          # «رقم الطلب UN NO» = رقم اليونيفاير
        out.append(dict(key=f"{R.get('C') or ''}|{_k(R.get('E'))}|{_k(R.get('F'))}", v=v, un=un))
    return out, notes


PS = r'''$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -Encoding UTF8 $env:HO_JOB | ConvertFrom-Json
function Find-Sheet($wb, $part) { foreach ($s in $wb.Worksheets) { if ($s.Name.Trim() -like "*$part*") { return $s } }; throw "Sheet not found: $part" }
function ColN($ws, $L) { return $ws.Range("$($L)1").Column }
function Key($ws, $r) {
  $inv = [Globalization.CultureInfo]::InvariantCulture
  $e = $ws.Cells.Item($r, 5).Value2; $f = $ws.Cells.Item($r, 6).Value2
  if ($e -is [double]) { $e = [string]::Format($inv, '{0:0.###}', $e) }; if ($f -is [double]) { $f = [string]::Format($inv, '{0:0.###}', $f) }
  return "$([string]$ws.Cells.Item($r, 3).Value2)|$e|$f"
}
function Put($cell, $v) {
  if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { $cell.ClearContents() | Out-Null; return }
  if ($v -is [decimal] -or $v -is [int64] -or $v -is [int32] -or $v -is [single] -or $v -is [double]) {
    $cell.Formula = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:R}', [double]$v)
  } else {
    $cell.Formula = [string]$v
  }
}
function Do-Book($xl, $t) {
  $wb = $xl.Workbooks.Open($t.path, 0, $false)
  try {
    if ($wb.ReadOnly) { throw "The file opened read-only – close it in Excel and try again: $($t.path)" }
    if ($t.backup) { $wb.SaveCopyAs($t.backup) }
    $ws = Find-Sheet $wb $job.qs_sheet
    $first = [int]$job.first_row; $tot = 0
    for ($r = $first; $r -le $first + 5000; $r++) { if ([string]$ws.Cells.Item($r, 3).Value2 -match '^\s*Total\s*$') { $tot = $r; break } }
    if ($tot -eq 0) { throw 'Total row not found in the QS sheet' }
    # 00) أعمدة الردم طبقة 7..12 بعد «الطبقة السطحية» (مرة واحدة بس)
    $wc = 0
    for ($c = 50; $c -le 120; $c++) { if ([string]$ws.Cells.Item(6, $c).Value2 -match 'Wearing') { $wc = $c; break } }
    if ($wc -eq 0) { throw 'Wearing Course column not found in the header row' }
    $wa = $ws.Cells.Item(6, $wc).Address($false, $false)
    if ($wa -ne 'CG6') { throw "Unexpected WIR column layout (Wearing Course at $wa)" }
    $nc = $wc + 2; $L = $job.new_layers.Count
    if (-not ([string]$ws.Cells.Item(6, $nc).Value2 -match 'طبقة 7')) {
      $ws.Range($ws.Cells.Item(1, $nc), $ws.Cells.Item(1, $nc + 2 * $L - 1)).EntireColumn.Insert() | Out-Null
      $src = $ws.Range($ws.Cells.Item(6, $wc), $ws.Cells.Item($tot, $wc + 1))
      for ($i = 0; $i -lt $L; $i++) {
        $c0 = $nc + 2 * $i
        $src.Copy() | Out-Null
        $ws.Cells.Item(6, $c0).PasteSpecial(-4122) | Out-Null
        $ws.Columns($c0).ColumnWidth = $ws.Columns($wc).ColumnWidth; $ws.Columns($c0 + 1).ColumnWidth = $ws.Columns($wc + 1).ColumnWidth
        $lay = $job.new_layers[$i]
        $ws.Cells.Item(6, $c0).Value2 = "الردم طبقة $lay Backfill $lay"
        $ws.Cells.Item(9, $c0).Value2 = 'رقم الطلب UN NO'; $ws.Cells.Item(9, $c0 + 1).Value2 = 'رقم المستخلص INV NO'
      }
      $t5 = [string]$ws.Range("BC5").Value2
      $ws.Range("BC5").MergeArea.UnMerge()
      $ws.Range($ws.Cells.Item(5, 55), $ws.Cells.Item(5, $nc + 2 * $L - 1)).Merge(); $ws.Range("BC5").Value2 = $t5
      $add = ''; $ifs = ''
      for ($i = 0; $i -lt $L; $i++) {
        $vi = ($ws.Cells.Item($first, $nc + 2 * $i + 1).Address($false, $false)) -replace '[0-9]', ''
        $add += ",$vi$first"
        $ifs += ',IF($' + $vi + $first + '=@@$8,"ردم' + $job.new_layers[$i] + '","")'
      }
      for ($c = $nc + 2 * $L; $c -le $nc + 2 * $L + 12; $c++) {
        $cell = $ws.Cells.Item($first, $c); $f = [string]$cell.Formula
        if ($f -match '^=IF\(COUNT\(' -and $f.Contains("CH$first")) { $cell.Formula = $f.Replace("CH$first", "CH$first$add") }
        elseif ($f -match 'TEXTJOIN' -and $f -match '([A-Z]+)\$8') {
          $hc = $matches[1]; $cell.Formula = $f.Substring(0, $f.Length - 1) + $ifs.Replace('@@', $hc) + ')'
        }
      }
    }
    # 0) كل صف ليه أرقام طلباته: نفك أي دمج في صفوف الداتا (الدمج بيبوّظ النسخ وربط الكمية برقم الطلب)
    $ws.Range($ws.Cells.Item($first, 1), $ws.Cells.Item($tot - 1, 140)).UnMerge() | Out-Null
    # 1) معادلات صف 11: كل كمية تتحسب بس لو فيه رقم طلب
    foreach ($g in $job.gates.PSObject.Properties) {
      $cell = $ws.Range("$($g.Name)$first"); $f = [string]$cell.Formula
      $pre = '=IF($' + $g.Value + $first + '="",0,'
      if ($f.StartsWith('=') -and -not $f.StartsWith($pre)) {
        if ($f -match '^=IF\(\$[A-Z]+\d+="",0,(.*)\)$') { $f = '=' + $Matches[1] }    # ربط قديم بعمود تاني – يتشال الأول
        $cell.Formula = $pre + $f.Substring(1) + ')' }
    }
    # 2) اللي مكتوب يدوي قبل كده (الرخصة، الغرفة، أرقام المستخلصات) محفوظ بالمقطع
    $old = @{}
    for ($r = $first; $r -lt $tot; $r++) {
      if (-not $ws.Cells.Item($r, 3).Value2) { continue }
      $m = @{}
      foreach ($c in $job.keep_cols) { $m[$c] = $ws.Range("$c$r").Value2 }
      foreach ($c in $job.un_cols) { $m["INV$c"] = $ws.Cells.Item($r, (ColN $ws $c) + 1).Value2 }
      $old[(Key $ws $r)] = $m
    }
    # 3) صفوف زيادة جوه الجدول (فوق الإجمالي) عشان SUBTOTAL والملخصات تكبر
    $n = $job.rows.Count; $cap = $tot - $first
    if ($n -gt $cap) {
      $k = $n - $cap
      $ws.Rows("$($tot - 1):$($tot - 2 + $k)").Insert() | Out-Null
      $tot += $k
    }
    # 4) كل الصفوف تاخد تنسيق ومعادلات صف 11
    if ($tot - 1 -gt $first) { $ws.Rows($first).Copy($ws.Range("A$($first + 1):A$($tot - 1)").EntireRow) | Out-Null }
    # 5) الخانات اليدوي
    for ($i = 0; $i -lt $n; $i++) {
      $r = $first + $i; $row = $job.rows[$i]
      foreach ($p in $row.v.PSObject.Properties) { Put ($ws.Range("$($p.Name)$r")) $p.Value }
      $o = $old[$row.key]
      foreach ($c in $job.keep_cols) { if ($o) { Put ($ws.Range("$c$r")) $o[$c] } else { $ws.Range("$c$r").ClearContents() | Out-Null } }
      foreach ($c in $job.un_cols) {
        $u = [string]$row.un.$c; $ci = ColN $ws $c
        Put ($ws.Cells.Item($r, $ci)) $u
        $inv = $null
        if ($u) { if ($o -and $o["INV$c"]) { $inv = $o["INV$c"] } else { $inv = [double]$job.no } }
        Put ($ws.Cells.Item($r, $ci + 1)) $inv
      }
    }
    # 6) الصفوف الفاضية: من غير داتا (معادلاتها بتطلع صفر)
    $cols = @($job.input_cols) + @($job.keep_cols)
    foreach ($c in $job.un_cols) { $cols += $c; $cols += $ws.Cells.Item(1, (ColN $ws $c) + 1).Address($false, $false) -replace '\d', '' }
    for ($r = $first + $n; $r -lt $tot; $r++) { foreach ($c in $cols) { $ws.Range("$c$r").ClearContents() | Out-Null } }
    foreach ($tt in $job.titles) { $s = Find-Sheet $wb $tt.sheet; $s.Range($tt.cell).Value2 = $tt.text }
    # 7) خلاصة حصر المستخلص: السابق (F) وأرقام الطلبات (M) بكود البند (B)
    $sm = Find-Sheet $wb $job.sum_sheet
    for ($r = 9; $r -le 400; $r++) {
      $code = [string]$sm.Cells.Item($r, 2).Value2
      if ($code -and $job.items.PSObject.Properties[$code]) {
        $it = $job.items.$code
        $sm.Cells.Item($r, 6).Formula = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:R}', [double]$it.prev)
        if ($it.wirs) { $sm.Cells.Item($r, 13).Formula = [string]$it.wirs }
      }
    }
    $xl.CalculateFull()
    $wb.Save()
    Write-Output "OK|$($t.name)|$n|$tot"
  } finally { $wb.Close($false) }
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.AskToUpdateLinks = $false
try { foreach ($t in $job.targets) { Do-Book $xl $t } }
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
'''


def build_job(project, calc, no, date_to, targets, backup_dir=None, tag=None):
    old_mode, old_cut = project.mode, project.cutoff
    project.mode = 'approved'; project.cutoff = date_to
    try:
        items, _ = project.build(date_to)
        rows, notes = input_rows(project, items)
    finally:
        project.mode, project.cutoff = old_mode, old_cut
    tg = []
    for name, path in targets:
        b = ''
        if backup_dir:
            base = os.path.splitext(path.replace('\\', '/').split('/')[-1])[0]
            b = os.path.join(backup_dir, f'{base} - before IPC-{no:02d} {tag or format(dt.datetime.now(), "%Y-%m-%d %H%M")}.xlsx')
            if tag and os.path.exists(b): b = ''          # التحديث الأوتوماتيك: نسخة واحدة في اليوم (أول تشغيلة)
        tg.append(dict(name=name, path=path, backup=b))
    job = dict(targets=tg, no=no, qs_sheet=QS_SHEET, sum_sheet=SUM_SHEET, first_row=FIRST_ROW,
               input_cols=INPUT_COLS, keep_cols=KEEP_COLS, un_cols=UN_COLS, gates=GATES, rows=rows, new_layers=NEW_LAYERS,
               items={str(x['code']): dict(prev=x['prev'] or 0, wirs=x.get('wirs') or '') for x in calc['rows']},
               titles=[dict(sheet=QS_SHEET, cell='A5', text=f'حصر  مستخلص رقم {no}'),
                       dict(sheet=SUM_SHEET, cell='A5', text=f'ملخص الحصر - مستخلص رقم {no:02d}')])
    return job, notes


def update_ho(project, calc, no, date_from, date_to, backup_dir=None, local=None, ho=None, include_ho=None, tag=None, skip_hash=None):
    """يحدّث شيتنا (OUR_SHEET) وبعده شيت الاستشاري. calc = ipc_core.compute(...) بتاع نفس المستخلص.
    بيرجّع (عدد المقاطع، ملاحظات)."""
    if os.name != 'nt': raise RuntimeError('Works on Windows only (needs Excel)')
    st = load_settings()
    local = full_path(local or st['LOCAL_SHEET']); ho = ho or st['HO_SHEET']
    targets = [('OUR_SHEET', local)]
    if include_ho is None: include_ho = HO_ENABLED
    if include_ho and ho: targets.append(('HO', ho))
    for _, p in targets:
        if not is_url(p) and not os.path.exists(p): raise FileNotFoundError(p)
    if backup_dir: os.makedirs(backup_dir, exist_ok=True)
    job, notes = build_job(project, calc, no, date_to, targets, backup_dir, tag)
    if skip_hash is not None:                       # التحديث الأوتوماتيك: لو مفيش جديد منفتحش الملفات
        import hashlib
        h = hashlib.md5(json.dumps([job['rows'], job['items'], [t['path'] for t in job['targets']], no],
                                   ensure_ascii=False, sort_keys=True, default=str).encode('utf-8')).hexdigest()
        if h == skip_hash: return 0, ['مفيش جديد'], h
    tmp = tempfile.mkdtemp()
    jf = os.path.join(tmp, 'job.json'); ps = os.path.join(tmp, 'ho.ps1')
    with open(jf, 'w', encoding='utf-8') as f: json.dump(job, f, ensure_ascii=False, default=str)
    with open(ps, 'w', encoding='utf-8-sig') as f: f.write(PS)
    r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps], env=dict(os.environ, HO_JOB=jf),
                       capture_output=True, text=True, timeout=900, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    shutil.rmtree(tmp, ignore_errors=True)
    ok = [ln for ln in (r.stdout or '').splitlines() if ln.startswith('OK|')]
    if r.returncode != 0 or len(ok) != len(targets):
        done = ', '.join(x.split('|')[1] for x in ok) or 'none'
        raise RuntimeError(f'Updated: {done}\n' + (r.stderr or r.stdout or 'Excel could not update the sheet').strip()[-500:])
    if skip_hash is not None: return len(job['rows']), notes, h
    return len(job['rows']), notes
