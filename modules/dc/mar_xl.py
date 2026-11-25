# -*- coding: utf-8 -*-
"""MAR ← سجل الإكسيل (طلب المستخدم 21-11-2026): الداتا كلها في إكسيلات المشروع، والبرنامج واجهة عليها.

make(): يسجّل سطر الـ MAR في «MAR LOG» (02-LOG MAR+MIR.xlsm)، يكتب رقم التقديم في شيت «MAR» (الفورم بتتملى من السجل
لوحدها) ويطلّع الشيت PDF (طلب اعتماد مواد + CLIENT Material Approval Form). كله من خلال إكسيل نفسه (زي الـ MIR في mat_core).
من غير Qt.
"""
import os, re, json, shutil, tempfile, subprocess, datetime as dt

NOWIN = 0x08000000 if os.name == 'nt' else 0
S_LOG, S_FORM = 'MAR LOG', 'MAR'
FIRST = 9                       # أول سطر داتا في MAR LOG
PICK, NAME_CELL, WARN_CELL = 'L6', 'L2', 'L7'     # شيت MAR: خانة رقم التقديم · اسم الملف · تنبيه «مش موجود في السجل»
DISCIPLINES = ('Civil', 'Mechanical', 'Electrical', 'Architectural', 'Other')
NO_BOQ = '—'                    # البند فاضي والمواصفة مكتوبة ← السطر الأول «—»
# شيت MAR: «رقم البنود» = السطر الأول من عمود G، «مواصفات العقد» = السطر التاني (G فيه سطر واحد ← زي الأول في الخانتين)
V = "IFERROR(INDEX('MAR LOG'!$G${first}:$G${end},MATCH($L$6,'MAR LOG'!$B${first}:$B${end},0)),\"\")"
F_BOQ = ('="رقم ووصف البنود بجدول الكميات "&IFERROR(LEFT({v},FIND(CHAR(10),{v})-1),{v})&""')
F_SPEC = ('=IF($L$6="","",IF(ISNUMBER(FIND(CHAR(10),{v})),"مواصفات العقد: "&MID({v},FIND(CHAR(10),{v})+1,999),'
          '"مواصفات العقد - بند رقم: "&{v}))')
CELL_BOQ, CELL_SPEC = 'B12', 'B15'


def split_spec(raw):
    """عمود G ← (البند، المواصفة). سطر واحد ← كله بند."""
    lines = [x.strip() for x in str(raw or '').replace('\r', '').split('\n')]
    lines = [x for x in lines if x]
    if not lines: return '', ''
    boq = '' if lines[0] == NO_BOQ else lines[0]
    return boq, ' / '.join(lines[1:])


def join_spec(boq, spec):
    """(البند، المواصفة) ← نص عمود G (سطرين)."""
    boq, spec = ' '.join(str(boq or '').split()), ' '.join(str(spec or '').split())
    if not spec: return boq
    return f'{boq or NO_BOQ}\n{spec}'


def form_formulas(end):
    v = V.format(first=FIRST, end=end)
    return {CELL_BOQ: F_BOQ.format(v=v), CELL_SPEC: F_SPEC.format(v=v)}


class XlError(RuntimeError):
    pass


PS = r"""
$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -Encoding UTF8 $env:MAR_JOB | ConvertFrom-Json
$res = @{ row = 0; isnew = $false; name = ''; pdf = ''; error = ''; saved = $false; warn = ''; grew = '' }
$xl = $null; $wb = $null; $mine = $false; $openedHere = $false
function Norm($s) { return ([string]$s -replace '\s', '').ToUpper() }
function Sheet($name) {
  foreach ($s in $wb.Worksheets) { if ((Norm ($s.Name)) -eq (Norm ($name))) { return $s } }
  throw ('Sheet not found: ' + $name)
}
function LastRow($ws, $col) { return $ws.Range($col + '1048576').End(-4162).Row }
function Put($ws, $addr, $c) {
  $cell = $ws.Range($addr)
  if ($c.kind -eq 'date') {
    $cell.Value2 = [DateTime]::ParseExact([string]$c.value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).ToOADate()
    if ([string]$cell.NumberFormat -eq 'General') { $cell.NumberFormat = 'dd/mm/yyyy' }
  }
  elseif ($c.kind -eq 'num') { $cell.Value2 = [double]$c.value }
  else { $cell.Value2 = [string]$c.value }
}
function FormEnd($form, $probe, $first) {
  $f = [string]$form.Range($probe).Formula
  $m = [regex]::Match($f, '\$A\$' + $first + ':\$A\$(\d+)')
  if ($m.Success) { return [int]$m.Groups[1].Value }
  return 0
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
  $ws = Sheet $job.log; $form = Sheet $job.form
  $first = [int]$job.first
  $no = Norm $job.no; $ref = Norm $job.ref
  if ($no -eq '' -or $ref -eq '') { throw 'The MAR number and the DC number are both needed.' }
  $last = LastRow $ws 'A'; $lb = LastRow $ws 'B'; if ($lb -gt $last) { $last = $lb }
  # 1) the line of this MAR (by its M number) - or the first empty line
  $row = 0; $other = ''
  for ($r = $first; $r -le $last; $r++) {
    $a = Norm $ws.Range('A' + $r).Text; $b = Norm $ws.Range('B' + $r).Text
    if ($a -eq $no) {
      if ($b -ne '' -and $b -ne $ref) { throw ('SAMENO ' + [string]$ws.Range('B' + $r).Text) }
      $row = $r; break
    }
    if ($b -eq $ref -and $other -eq '') { $other = [string]$ws.Range('A' + $r).Text }
  }
  if ($row -eq 0 -and $other -ne '') { throw ('SAMEREF ' + $other) }
  if ($row -eq 0) {
    $res.isnew = $true
    $row = $last + 1
    for ($r = $first; $r -le $last; $r++) {
      if ([string]$ws.Range('A' + $r).Formula -eq '' -and [string]$ws.Range('B' + $r).Formula -eq '' -and [string]$ws.Range('D' + $r).Formula -eq '') { $row = $r; break }
    }
    # the form reads the log down to a fixed line: on its last line a new line is inserted, so Excel makes the ranges longer by itself
    $end = FormEnd $form $job.name_cell $first
    if ($end -gt 0 -and $row -eq $end) { $null = $ws.Rows.Item($row).Insert(); $res.grew = 'insert' }
    if ($row -gt $first -and [string]$ws.Range('K' + ($row - 1)).Formula -like '=*' -and [string]$ws.Range('K' + $row).Formula -eq '') {
      $ws.Range('K' + $row).FormulaR1C1 = $ws.Range('K' + ($row - 1)).FormulaR1C1
    }
  }
  # 2) write: a new line takes everything; an old line only its empty cells (and the cells marked force)
  foreach ($c in $job.cells) {
    if ([string]$c.value -eq '') { continue }
    $addr = $c.col + $row
    if (-not $res.isnew -and [string]$ws.Range($addr).Formula -ne '' -and -not $c.force) { continue }
    Put $ws $addr $c
  }
  # 3a) the two boxes of the form: BOQ = first line of column G, spec = second line (once - then they stay)
  $end0 = FormEnd $form $job.name_cell $first
  if ($end0 -gt 0) {
    foreach ($p in $job.formulas.PSObject.Properties) {
      if ([string]$form.Range($p.Name).Formula -notlike '*CHAR(10)*') {
        $form.Range($p.Name).Formula = ([string]$p.Value).Replace('{END}', [string]$end0)
        $res.grew = ($res.grew + ' formulas').Trim()
      }
    }
  }
  # 3) is the line inside what the form reads? if not, make the form ranges longer
  $end = FormEnd $form $job.name_cell $first
  if ($end -gt 0 -and $row -gt $end) {
    $newEnd = $row + 200
    foreach ($col in @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O')) {
      $old = '$' + $col + '$' + $first + ':$' + $col + '$' + $end
      $new = '$' + $col + '$' + $first + ':$' + $col + '$' + $newEnd
      $null = $form.Cells.Replace($old, $new, 2, 1, $false)
    }
    $res.grew = 'ranges ' + $end + ' > ' + $newEnd
  }
  $res.row = $row
  $xl.Calculate()
  $wb.Save(); $res.saved = $true
  # 4) the form: choose this submittal, then PDF of the sheet
  if ($job.pdf) {
    $form.Range($job.pick).Value2 = [string]$job.ref
    $xl.Calculate()
    $res.warn = [string]$form.Range($job.warn_cell).Text
    $name = [string]$form.Range($job.name_cell).Text
    if ($name -like '#*') { $name = '' }
    $res.name = $name
    try { if ($form.PageSetup.PaperSize -ne 9) { $form.PageSetup.PaperSize = 9 } } catch { $res.warn = ($res.warn + ' (A4 could not be set: ' + $_.Exception.Message + ')').Trim() }
    $form.ExportAsFixedFormat(0, $job.pdf, 0, $true, $false)
    $res.pdf = $job.pdf
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
$res | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $env:MAR_OUT
"""


def rev_num(rev):
    """«R0» / «0» / «Rev 2» ← 0 / 0 / 2"""
    m = re.search(r'\d+', str(rev or ''))
    return int(m.group()) if m else 0


def parse_date(text):
    """«21-11-2026» / «21/11/2026» / «2026-11-21» ← date، أو None."""
    t = str(text or '').strip()
    for f in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%d.%m.%Y'):
        try: return dt.datetime.strptime(t, f).date()
        except ValueError: pass
    return None


def supplier_name(label):
    """اسم المورد اللي يتكتب في السجل من اختيار القايمة:
    «SUP · 18-AL_RAJHI_STEEL · الراجحي للصناعات الحديدية» ← «الراجحي للصناعات الحديدية»؛ من غير عربي ← «AL RAJHI STEEL»."""
    parts = [p.strip() for p in str(label or '').split('·') if p.strip()]
    if not parts: return ''
    if len(parts) >= 3: return parts[-1]
    p = parts[-1] if len(parts) == 2 else parts[0]
    m = re.match(r'^\d+\s*-\s*(.+)$', p)
    return m.group(1).replace('_', ' ').strip() if m else p


def cells(mar, edited=()):
    """خانات سطر MAR LOG. mar: no, ref, rev, desc, supplier, disc, spec, date (date), factory, prev.
    edited = الخانات اللي المستخدم غيّرها بإيده في سطر موجود (تتكتب فوق القديم)."""
    def c(col, key, kind='text'):
        v = mar.get(key)
        if kind == 'date': v = v.isoformat() if isinstance(v, dt.date) else ''
        elif kind == 'num': v = '' if v in (None, '') else str(int(v))
        else: v = str(v or '').strip()
        return dict(col=col, kind=kind, value=v, force=key in edited)
    return [c('A', 'no'), c('B', 'ref'), c('C', 'rev', 'num'), c('D', 'desc'), c('E', 'supplier'), c('F', 'disc'),
            c('G', 'spec'), c('H', 'date', 'date'), c('N', 'factory'), c('O', 'prev')]


def explain(err, base, mar):
    e = str(err or '')
    if e == 'READONLY':
        return f'{base} is open somewhere else (another Excel window or computer) – close it there and try again.'
    if e.startswith('SAMENO '):
        return (f"{mar.get('no')} is already in MAR LOG with another submittal number:  {e[7:]}\n\n"
                'Change the M number here, or correct that line in the log.')
    if e.startswith('SAMEREF '):
        return (f"{mar.get('ref')} is already written in MAR LOG for  {e[8:]}.\n\n"
                'One DC number = one MAR. Take the next DC number.')
    return f'{base}: {e[-300:]}'


def make(path, mar, pdf=None, edited=(), timeout=300):
    """يسجّل/يكمّل سطر الـ MAR ويطلّع الفورم PDF. يرجّع dict(row, isnew, name, pdf, warn, grew). أي مشكلة ← XlError."""
    path = os.path.abspath(path); base = os.path.basename(path)
    if os.name != 'nt': raise XlError('Writing to the Excel registers needs Excel on Windows')
    if not os.path.isfile(path): raise XlError(f'Not found: {path}')
    tmpd = tempfile.mkdtemp(prefix='mar_xl_')
    job, out, ps = (os.path.join(tmpd, n) for n in ('job.json', 'out.json', 'mar_xl.ps1'))
    with open(job, 'w', encoding='utf-8') as f:
        json.dump(dict(path=path, log=S_LOG, form=S_FORM, first=FIRST, pick=PICK, name_cell=NAME_CELL, warn_cell=WARN_CELL,
                       no=str(mar.get('no') or '').strip(), ref=str(mar.get('ref') or '').strip(),
                       cells=cells(mar, edited), pdf=pdf or '', formulas=form_formulas('{END}')), f, ensure_ascii=False)
    with open(ps, 'w', encoding='utf-8-sig') as f: f.write(PS)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps],
                           env=dict(os.environ, MAR_JOB=job, MAR_OUT=out), capture_output=True, timeout=timeout, creationflags=NOWIN)
        try:
            with open(out, encoding='utf-8-sig') as f: res = json.load(f)
        except Exception:
            err = (r.stderr or b'').decode('utf-8', 'replace').strip() or (r.stdout or b'').decode('utf-8', 'replace').strip()
            raise XlError(f'Excel did not answer for {base}: {err[-300:] or "no reply"}')
        if res.get('error'): raise XlError(explain(res['error'], base, mar))
        if pdf and not os.path.isfile(pdf): res['pdf'] = ''
        return res
    except subprocess.TimeoutExpired:
        raise XlError(f'Excel took too long with {base} – check that Excel is not waiting for an answer (a message), then try again.')
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)
