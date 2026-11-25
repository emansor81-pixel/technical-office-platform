# -*- coding: utf-8 -*-
"""DC Control – شيت DC.xlsx المشترك على OneDrive (الصادر والوارد).

القراية: من نسخة مؤقتة (من غير ما نلمس الأصل).
الكتابة: من خلال إكسيل نفسه (PowerShell + Excel COM) عشان OneDrive يدمج تعديلاتنا مع تعديلات الباقيين،
وبنملا الخانات الفاضية بس – عمرنا ما نكتب فوق حاجة حد كاتبها.
لو إكسيل ما اشتغلش، الكتابة تستنى في dc_sheet_pending.json لحد زرار Update DC sheet.
"""
import os, re, json, ntpath, shutil, tempfile, subprocess, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = os.path.join(HERE, 'dc_settings.json')
PENDING = os.path.join(HERE, 'dc_sheet_pending.json')
NOWIN = 0x08000000 if os.name == 'nt' else 0
DEFAULT = os.path.join(os.path.expanduser('~'), 'OneDrive', '01-N.W.C PROJECT', 'CO WATER', 'DC.xlsx')

# kind: (اسم التبويب من غير مسافات، أعمدة)
KINDS = {
    'YA':  ('DC-YA-OUT',  dict(no='A', date='B', desc='C', file='D')),
    'CLIENT': ('DC-CLIENT-OUT', dict(seq='A', date='B', no='C', desc='D', file='E')),
    'SUP': ('DC-SUP-OUT', dict(seq='A', date='B', no='C', desc='D', file='E')),
    'OTH': ('DC-OTH',     dict(no='A', date='B', to='C', desc='D', file='F')),
    'IN':  ('DC-IN-ALL',  dict(seq='A', date='B', frm='C', no='D', desc='E', file='F')),
}
LABEL = {'YA': 'Outgoing – the Consultant', 'CLIENT': 'Outgoing – CLIENT', 'SUP': 'Outgoing – Suppliers', 'OTH': 'Outgoing – Other', 'IN': 'Incoming'}
CODE2KIND = {'OUT_YA': 'YA', 'OUT_CLIENT': 'CLIENT', 'OUT_SUB': 'SUP', 'OUT_OTH': 'OTH',
             'IN_YA': 'IN', 'IN_CLIENT': 'IN', 'IN_SUB': 'IN', 'IN_OTH': 'IN'}
SENDER = {'IN_YA': 'الاستشاري', 'IN_CLIENT': 'الجهة المالكة', 'IN_SUB': 'مورد / مقاول', 'IN_OTH': ''}
FIRST = {'SUP': 3}                                # أول صف بيانات (الباقي بيبدأ من 4)
# تاريخ تسليم الخطاب (طلب المستخدم 21-11-2026: الداتا في الإكسيل مش في البرنامج) – عمود بعنوان ثابت بعد أعمدة التبويب
DELIV_HEAD, HEAD_ROW, MAX_COL = 'تاريخ التسليم', 2, 10
DELIV_FROM = {'YA': 'F', 'CLIENT': 'F', 'SUP': 'G', 'OTH': 'G'}      # أول عمود ممكن يتحط فيه (لو مليان بحاجة تانية ← اللي بعده)


def sheet_path():
    try:
        with open(CFG, encoding='utf-8') as fh: p = json.load(fh).get('DC_SHEET')
        if p: return p
    except Exception:
        pass
    return DEFAULT


def _norm(n): return re.sub(r'\s+', '', n or '').upper()


def _ws(wb, kind):
    want = _norm(KINDS[kind][0])
    return next((ws for ws in wb.worksheets if _norm(ws.title) == want), None)


def _ci(col):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(col)


def read(kind):
    """صفوف التبويب: dict(row, no, date, desc, …) – من نسخة مؤقتة."""
    import openpyxl
    src = sheet_path()
    fd, tmp = tempfile.mkstemp(prefix='dc_sheet_', suffix='.xlsx'); os.close(fd)   # اسم جديد كل مرة
    try:
        shutil.copy2(src, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
        return _rows(wb, kind, src)
    finally:
        try: os.remove(tmp)
        except OSError: pass


def _rows(wb, kind, src):
    ws = _ws(wb, kind)
    if ws is None: wb.close(); raise KeyError(f'Tab {KINDS[kind][0]} not found in {os.path.basename(src)}')
    cols = KINDS[kind][1]
    out = []
    f0 = FIRST.get(kind, 4)
    dcol = _deliv_col(ws, kind)[0] if kind in DELIV_FROM else None
    for i, r in enumerate(ws.iter_rows(min_row=f0, max_col=MAX_COL, values_only=True), f0):
        r = list(r) + [None] * MAX_COL
        d = {k: r[_ci(c) - 1] for k, c in cols.items()}
        d['row'] = i
        d['deliv'] = r[dcol - 1] if dcol else None       # تاريخ التسليم (لو العمود موجود)
        out.append(d)
    wb.close()
    return out


def _empty(v): return v is None or (isinstance(v, str) and not v.strip())


def _deliv_col(ws, kind):
    """عمود «تاريخ التسليم» في التبويب: (رقم العمود الموجود أو None، رقم أول عمود عنوانه فاضي ينفع يتحط فيه أو None)."""
    head = next(ws.iter_rows(min_row=HEAD_ROW, max_row=HEAD_ROW, max_col=MAX_COL, values_only=True), ())
    head = list(head) + [None] * MAX_COL
    have = next((i + 1 for i, v in enumerate(head[:MAX_COL]) if isinstance(v, str) and _norm(v) == _norm(DELIV_HEAD)), None)
    free = next((i for i in range(_ci(DELIV_FROM[kind]), MAX_COL + 1) if _empty(head[i - 1])), None)
    return have, free


def set_delivered(kind, no, date=None):
    """يكتب تاريخ تسليم الخطاب في سطره في الشيت (date=None ← يمسحه = الخطاب اتفتح تاني). يرجّع (ok, رسالة).
    مابيتحفظش في الانتظار: لو إكسيل ماكتبش، مفيش حاجة اتغيرت."""
    import openpyxl
    from openpyxl.utils import get_column_letter
    if kind not in DELIV_FROM: return False, f'No delivery column for {kind}'
    if os.name != 'nt': return False, 'Excel is only available on Windows'
    n = _tail(no)
    if n is None: return False, 'The letter number has no digits'
    src = sheet_path()
    fd, tmp = tempfile.mkstemp(prefix='dc_sheet_', suffix='.xlsx'); os.close(fd)
    try:
        shutil.copy2(src, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
        ws = _ws(wb, kind)
        if ws is None: wb.close(); return False, f'Tab {KINDS[kind][0]} not found'
        have, free = _deliv_col(ws, kind)
        row = next((d['row'] for d in _rows(wb, kind, src) if _tail(d.get('no')) == int(n)), None)
    except Exception as ex:
        return False, f'Could not read the DC sheet: {ex}'
    finally:
        try: os.remove(tmp)
        except OSError: pass
    if not row: return False, f'{no} is not in tab {KINDS[kind][0]}'
    col = have or free
    if not col: return False, 'No empty column for the delivery date in the DC sheet'
    L = get_column_letter(col); tab = KINDS[kind][0]
    entries = []
    if not have: entries.append(dict(tab=tab, row=HEAD_ROW, first=HEAD_ROW, check_cols=[], cells=[dict(col=L, kind='text', value=DELIV_HEAD)]))
    cell = dict(col=L, kind='date', value=date.isoformat(), force=True) if date else dict(col=L, kind='text', value='', force=True)
    entries.append(dict(tab=tab, row=row, first=row, check_cols=[], cells=[cell]))
    fd, job = tempfile.mkstemp(prefix='dc_job_', suffix='.json'); os.close(fd)
    with open(job, 'w', encoding='utf-8') as fh: json.dump(dict(path=src, entries=entries), fh, ensure_ascii=False)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PS_WRITE],
                           capture_output=True, timeout=180, env=dict(os.environ, DC_JOB=job), creationflags=NOWIN)
        out = (r.stdout or b'').decode('utf-8', 'replace').strip()
        err = (r.stderr or b'').decode('utf-8', 'replace').strip()
        if r.returncode != 0 or f'row {row} (1 cells)' not in out:
            return False, (err[-300:] or out[-300:] or 'Excel did not answer')
        return True, out
    except Exception as ex:
        return False, str(ex)
    finally:
        try: os.remove(job)
        except OSError: pass


def next_number(kind):
    """أول رقم صادر لسه مالوش وصف = الرقم اللي عليه الدور."""
    if kind == 'IN': return None
    rows = [d for d in read(kind) if not _empty(d.get('no'))]
    last = max([i for i, d in enumerate(rows) if not _empty(d.get('desc'))], default=-1)   # بعد آخر خطاب متسجّل
    for d in rows[last + 1:]:
        if _empty(d.get('desc')): return str(d['no']).strip()
    if rows and last >= 0:                                 # الأرقام المكتوبة خلصت – الرقم اللي بعده
        n = _tail(rows[last]['no'])
        if n is not None: return re.sub(r'\d+\s*$', lambda m: str(n + 1).zfill(len(m.group(0).strip())), str(rows[last]['no']).strip())
    return None


def subject(text):
    t = text or ''
    m = re.search(r'الموضوع\s*[:：]?\s*([^\n]{4,160})', t) or re.search(r'Subject\s*[:：]\s*([^\n]{4,160})', t, re.I)
    return ' '.join(m.group(1).split()).strip(' .:-') if m else ''


def _tail(v):
    m = re.search(r'(\d+)\s*$', str(v or ''))
    return int(m.group(1)) if m else None


def find_row(kind, n):
    """سطر الخطاب الصادر رقم n في تبويبه (بنقارن آخر رقم بس – كل تبويب مكتوب بشكل مختلف)."""
    for d in read(kind):
        if _tail(d.get('no')) == int(n): return d
    return None


def make_entry(kind, no='', date=None, desc='', file='', frm='', to=''):
    """سطر واحد للشيت. الصادر بيتحط في سطر رقمه، والوارد في أول سطر فاضي."""
    return dict(kind=kind, no=str(no or '').strip(), date=(date or dt.date.today()).isoformat(), desc=desc or '',
                file=file or '', frm=frm or '', to=to or '', added=f'{dt.datetime.now():%Y-%m-%d %H:%M}')


def entry_from_move(code, fields, text, dst):
    kind = CODE2KIND.get(code)
    if not kind: return None
    no = ''
    if kind != 'IN':
        n = fields.get('no')
        if n: no = {'YA': f'CTR-YA-WA-26-{int(n):04d}', 'CLIENT': f'CTR-CLIENT-WA-{int(n):04d}',
                    'SUP': f'CTR-SUP C - WA-{int(n):04d}', 'OTH': f'CTR-OTH- WA-{int(n):03d}'}[kind]
    else:
        n = fields.get('no')
        no = f'cons-26-{n}' if (code == 'IN_YA' and n) else (str(n) if n else '')
    d = fields.get('date')
    try: d = dt.date.fromisoformat(d) if d else None
    except ValueError: d = None
    e = make_entry(kind, no, d if kind != 'IN' else dt.date.today(), subject(text), dst, frm=SENDER.get(code, ''))
    if kind == 'IN' and d: e['letter_date'] = d.isoformat()
    return e


# ---------------------------------------------------------------- الكتابة من خلال إكسيل
PS_WRITE = r"""
$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -Encoding UTF8 $env:DC_JOB | ConvertFrom-Json
$path = $job.path
$xl = $null; $wb = $null; $mine = $false; $openedHere = $false
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
if ($xl) { foreach ($w in $xl.Workbooks) { if ($w.FullName -eq $path -or $w.Name -eq [IO.Path]::GetFileName($path)) { $wb = $w } } }
if (-not $wb) {
  if (-not $xl) { $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $mine = $true }
  $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open($path); $openedHere = $true
}
function Norm($s) { return ($s -replace '\s', '').ToUpper() }
$res = @()
foreach ($e in $job.entries) {
  $ws = $null
  foreach ($s in $wb.Worksheets) { if ((Norm $s.Name) -eq (Norm $e.tab)) { $ws = $s } }
  if (-not $ws) { $res += "NO TAB $($e.tab)"; continue }
  $row = [int]$e.row
  $last = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row + 1
  if ($row -eq 0) {                                 # وارد أو رقم مش موجود: أول سطر الخانات فيه فاضية
    for ($r = [int]$e.first; $r -le $last + 200; $r++) {
      $busy = $false
      foreach ($c in $e.check_cols) { if ([string]$ws.Range("$c$r").Formula -ne '') { $busy = $true } }
      if (-not $busy) { $row = $r; break }
    }
  }
  $n = 0
  foreach ($c in $e.cells) {
    $cell = $ws.Range("$($c.col)$row")
    if ([string]$cell.Formula -ne '' -and -not $c.force) { continue }      # مش هنكتب فوق حاجة موجودة
    if ($c.kind -eq 'date') {
      if ($c.value -is [DateTime]) { $d = $c.value }
      else { try { $d = [DateTime]::ParseExact([string]$c.value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
             catch { $d = [DateTime]::Parse([string]$c.value, [Globalization.CultureInfo]::InvariantCulture) } }
      try { $cell.Value2 = [double]$d.ToOADate() }
      catch { $cell.Formula = "=DATE($($d.Year),$($d.Month),$($d.Day))" }    # خانة مش بتقبل رقم (قايمة/جدول) ← معادلة
      try { $cell.NumberFormat = 'dd/mm/yyyy' } catch { }
    }
    elseif ($c.kind -eq 'link') { $ws.Hyperlinks.Add($cell, $c.value, '', $c.value, $c.text) | Out-Null }
    else { $cell.Value2 = $c.value }
    $n++
  }
  $res += "OK $($e.tab) row $row ($n cells)"
}
$wb.Save()
if ($openedHere) { $wb.Close($true) }
if ($mine) { $xl.Quit() }
$res -join "`n"
"""


def _cells(e):
    cols = KINDS[e['kind']][1]
    cells = []
    if e['kind'] == 'IN' and e.get('no'): cells.append(dict(col=cols['no'], kind='text', value=e['no']))
    if e['kind'] != 'IN' and e.get('no') and 'no' in cols:
        cells.append(dict(col=cols['no'], kind='text', value=e['no']))
    cells.append(dict(col=cols['date'], kind='date', value=e['date']))
    if e.get('desc'): cells.append(dict(col=cols['desc'], kind='text', value=e['desc']))
    if e.get('frm') and 'frm' in cols: cells.append(dict(col=cols['frm'], kind='text', value=e['frm']))
    if e.get('to') and 'to' in cols: cells.append(dict(col=cols['to'], kind='text', value=e['to']))
    if e.get('file'):
        cells.append(dict(col=cols['file'], kind='link', value=e['file'], text=ntpath.basename(e['file'])))
    return cells


def _job(entries):
    out = []
    for e in entries:
        tab, cols = KINDS[e['kind']]
        cells = _cells(e)
        m = dict(tab=tab, cells=cells, check_cols=[cols['date'], cols['desc']], row=0, first=FIRST.get(e['kind'], 4))
        if e['kind'] != 'IN' and _tail(e.get('no')) is not None:
            try: hit = find_row(e['kind'], _tail(e['no']))
            except Exception: hit = None
            if hit:                                      # الرقم مكتوب في الشيت – نسيبه بشكله
                m['row'] = hit['row']; m['cells'] = [c for c in cells if c['col'] != cols['no']]
        out.append(m)
    return dict(path=sheet_path(), entries=out)


def load_pending():
    from dc_io import load_json
    L = load_json(PENDING, [])
    return L if isinstance(L, list) else []


def _save_pending(L):
    from dc_io import save_json
    save_json(PENDING, L)


def write(entries):
    """يكتب في الشيت من خلال إكسيل. لو فشل، يحفظ السطور في الانتظار. يرجّع (ok, رسالة)."""
    entries = [e for e in entries if e]
    if not entries: return True, ''
    if os.name != 'nt':
        _save_pending(load_pending() + entries); return False, 'Excel is only available on Windows – saved for later'
    fd, job = tempfile.mkstemp(prefix='dc_job_', suffix='.json'); os.close(fd)
    with open(job, 'w', encoding='utf-8') as fh: json.dump(_job(entries), fh, ensure_ascii=False)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PS_WRITE],
                           capture_output=True, timeout=180, env=dict(os.environ, DC_JOB=job), creationflags=NOWIN)
        out = (r.stdout or b'').decode('utf-8', 'replace').strip()
        err = (r.stderr or b'').decode('utf-8', 'replace').strip()
        if r.returncode != 0 or 'OK ' not in out:
            raise RuntimeError(err[-400:] or out[-400:] or 'Excel did not answer')
        return True, out
    except Exception as ex:
        from dc_io import LOG
        LOG.error('DC sheet write failed (%d entries kept for later): %s', len(entries), ex)
        _save_pending(load_pending() + entries)
        return False, f'Saved for later (Update DC sheet): {ex}'
    finally:
        try: os.remove(job)
        except OSError: pass


def flush_pending():
    L = load_pending()
    if not L: return True, 'Nothing waiting'
    _save_pending([])
    ok, msg = write(L)
    return ok, msg
