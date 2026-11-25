# -*- coding: utf-8 -*-
"""DC Control – قراية الورق اللي بيتحط في 00_INBOX.

PDF من سيستم   ← الكلام بيتقري على طول (pypdf).
PDF سكانر      ← صورة كل صفحة تتقري بـ Tesseract (عربي + إنجليزي).
صورة موبايل    ← Tesseract يقراها ويطلّع منها PDF بكلام بيتعلّم عليه.

مش محتاج غير pypdf (موجودة) و Tesseract متسطّب. مفيش Pillow.
"""
import os, re, sys, shutil, struct, zlib, tempfile, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
PLATFORM = os.path.dirname(os.path.dirname(HERE))
IMG_EXT = ('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp')
LANGS = 'ara+eng'
NOWIN = 0x08000000 if os.name == 'nt' else 0          # CREATE_NO_WINDOW – من غير شاشة سودا


class ReadError(Exception):
    pass


# ---------------------------------------------------------------- Tesseract
def tesseract_exe():
    for p in (os.environ.get('TESSERACT_EXE'),
              r'C:\Program Files\Tesseract-OCR\tesseract.exe',
              r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
              os.path.expanduser(r'~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
              shutil.which('tesseract')):
        if p and os.path.isfile(p): return p
    return None


def tessdata_dir(exe):
    """ملف العربي: من فولدر البرنامج (<PLATFORM>\\tessdata) لو موجود، غير كده من فولدر Tesseract."""
    own = os.path.join(PLATFORM, 'tessdata')
    if os.path.isfile(os.path.join(own, 'ara.traineddata')): return own
    if exe:
        d = os.path.join(os.path.dirname(exe), 'tessdata')
        if os.path.isfile(os.path.join(d, 'ara.traineddata')): return d
        try:                                              # مكان تاني افتراضي (TESSDATA_PREFIX / لينكس)
            r = subprocess.run([exe, '--list-langs'], capture_output=True, timeout=30, creationflags=NOWIN)
            if b'\nara' in r.stdout.replace(b'\r', b''): return ''
        except Exception:
            pass
    return None


def ocr_status():
    exe = tesseract_exe()
    if not exe: return False, 'Tesseract is not installed'
    if tessdata_dir(exe) is None: return False, 'Arabic language data (ara.traineddata) is missing'
    return True, exe


def _tess(img, out_base=None, pdf=False, psm=3):
    exe = tesseract_exe()
    if not exe: raise ReadError('Tesseract is not installed')
    td = tessdata_dir(exe)
    cmd = [exe, img, out_base or 'stdout', '-l', LANGS if td is not None else 'eng', '--psm', str(psm)]
    if td: cmd[1:1] = ['--tessdata-dir', td]
    if pdf: cmd += ['txt', 'pdf']
    r = subprocess.run(cmd, capture_output=True, timeout=180, creationflags=NOWIN)
    if r.returncode != 0:
        raise ReadError((r.stderr or b'').decode('utf-8', 'replace').strip()[-300:] or 'OCR failed')
    if out_base:
        with open(out_base + '.txt', encoding='utf-8', errors='replace') as f: return f.read()
    return r.stdout.decode('utf-8', 'replace')


# ---------------------------------------------------------------- صور الـ PDF من غير Pillow
def _tiff_ccitt(data, w, h, k, black1):
    """ورق السكانر الأبيض والأسود (CCITT G3/G4) ← TIFF يفهمه Tesseract."""
    comp = 4 if k < 0 else 3
    tags = [(256, 4, w), (257, 4, h), (258, 3, 1), (259, 3, comp), (262, 3, 1 if black1 else 0),
            (273, 4, 0), (277, 3, 1), (278, 4, h), (279, 4, len(data))]
    off_data = 8 + 2 + len(tags) * 12 + 4
    out = b'II*\x00' + struct.pack('<I', 8) + struct.pack('<H', len(tags))
    for tag, typ, val in tags:
        if tag == 273: val = off_data
        out += struct.pack('<HHI', tag, typ, 1) + (struct.pack('<HH', val, 0) if typ == 3 else struct.pack('<I', val))
    return out + struct.pack('<I', 0) + data


def _png(raw, w, h, bpc, ncomp):
    """بيانات بكسل خام (بعد فك الضغط) ← PNG."""
    color = {1: 0, 3: 2}.get(ncomp)
    if color is None: return None
    stride = (w * ncomp * bpc + 7) // 8
    if len(raw) < stride * h: return None
    rows = b''.join(b'\x00' + raw[i * stride:(i + 1) * stride] for i in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, bpc, color, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(rows, 6)) + chunk(b'IEND', b''))


def _page_images(page):
    """أكبر صورة في الصفحة (صفحة السكانر) ← (امتداد، بايتات)."""
    found = []
    def walk(res, depth=0):
        if depth > 3 or not res: return
        xo = res.get('/XObject')
        if not xo: return
        xo = xo.get_object()
        for k in xo:
            o = xo[k].get_object()
            st = o.get('/Subtype')
            if st == '/Image':
                found.append(o)
            elif st == '/Form':
                walk(o.get('/Resources'), depth + 1)
    try:
        walk(page.get('/Resources').get_object() if page.get('/Resources') else None)
    except Exception:
        return []
    out = []
    for o in sorted(found, key=lambda o: -(int(o.get('/Width', 0)) * int(o.get('/Height', 0)))):
        w, h = int(o.get('/Width', 0)), int(o.get('/Height', 0))
        if w * h < 200 * 200: continue
        f = o.get('/Filter')
        f = [x for x in (f if isinstance(f, list) else [f]) if x] if f else []
        f = [str(x) for x in f]
        try:
            enc = o._data
            if f and f[-1] == '/DCTDecode':
                data = enc if len(f) == 1 else o.get_data()
                out.append(('.jpg', data))
            elif f and f[-1] == '/JPXDecode':
                out.append(('.jp2', enc))
            elif f and f[-1] == '/CCITTFaxDecode':
                p = o.get('/DecodeParms') or {}
                if isinstance(p, list): p = p[-1] or {}
                p = p.get_object() if hasattr(p, 'get_object') else p
                out.append(('.tif', _tiff_ccitt(enc, w, h, int(p.get('/K', 0)), bool(p.get('/BlackIs1', False)))))
            else:
                raw = o.get_data()
                cs = o.get('/ColorSpace')
                cs = cs.get_object() if hasattr(cs, 'get_object') else cs
                n = 3 if cs == '/DeviceRGB' else 1
                bpc = int(o.get('/BitsPerComponent', 8))
                png = _png(raw, w, h, bpc, n)
                if png: out.append(('.png', png))
        except Exception:
            continue
        if out: break
    return out


# ---------------------------------------------------------------- رسم صفحة الـ PDF كصورة
# ويندوز نفسه بيرسم الـ PDF (Windows.Data.Pdf) – مفيش برنامج زيادة. ده بيحل ورق يونيفاير
# اللي الكلام فيه مرسوم خطوط مش حروف، وأي سكانر.
PS_RENDER = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$m = [System.WindowsRuntimeSystemExtensions].GetMethods()
$gen = ($m | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
$act = ($m | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function Await($op, $t) { $x = $gen.MakeGenericMethod($t).Invoke($null, @($op)); $x.Wait(-1) | Out-Null; $x.Result }
function AwaitA($op) { $x = $act.Invoke($null, @($op)); $x.Wait(-1) | Out-Null }
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFolder, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime] | Out-Null
$f = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($env:DC_PDF)) ([Windows.Storage.StorageFile])
$d = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($f)) ([Windows.Data.Pdf.PdfDocument])
$dir = Await ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($env:DC_OUT)) ([Windows.Storage.StorageFolder])
$n = [Math]::Min([int]$d.PageCount, [int]$env:DC_PAGES)
for ($i = 0; $i -lt $n; $i++) {
  $p = $d.GetPage([uint32]$i)
  $o = Await ($dir.CreateFileAsync("r$i.png", [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
  $s = Await ($o.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
  $r = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $r.DestinationWidth = [uint32]([Math]::Min(3300, $p.Size.Width * 300 / 96))
  AwaitA ($p.RenderToStreamAsync($s, $r))
  $s.Dispose(); $p.Dispose()
}
"""


def render_pages(pdf, pages, work):
    """صور الصفحات الأولى (PNG ~300dpi) – ويندوز: Windows.Data.Pdf، غير كده pdftoppm لو موجود."""
    out = []
    try:
        if os.name == 'nt':
            env = dict(os.environ, DC_PDF=os.path.abspath(pdf), DC_OUT=os.path.abspath(work), DC_PAGES=str(pages))
            subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', PS_RENDER],
                           capture_output=True, timeout=180, env=env, creationflags=NOWIN)
            out = [os.path.join(work, f'r{i}.png') for i in range(pages)]
        elif shutil.which('pdftoppm'):
            subprocess.run(['pdftoppm', '-r', '300', '-png', '-l', str(pages), pdf, os.path.join(work, 'r')],
                           capture_output=True, timeout=180)
            out = sorted(os.path.join(work, x) for x in os.listdir(work) if x.startswith('r') and x.endswith('.png'))
    except Exception:
        return []
    return [x for x in out if os.path.isfile(x) and os.path.getsize(x) > 0]


# ---------------------------------------------------------------- القراية
# أرقام عربية/فارسية ← 0-9، والياء والكاف والهاء الفارسي ← العربي
AR_FIX = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹یکھۀ', '01234567890123456789يكهه')


def clean(t):
    """الحروف العربية المتشكّلة (ﺷﺮﻛﺔ) ← حروف عادية، ومسافات موحّدة."""
    import unicodedata
    t = unicodedata.normalize('NFKC', t or '').replace('\u200f', '').replace('\u200e', '')
    t = t.translate(AR_FIX)
    return re.sub(r'[ \t]+', ' ', t)


def _garbled(t):
    """خطوط PDF متلخبطة بتطلع حروف يوناني/رموز بدل العربي."""
    odd = len(re.findall(r'[\u0370-\u03FF\u02B0-\u02FF\uFB50-\uFDFF]', t))
    return odd > 0.08 * max(_meaningful(t), 1)


def _meaningful(t):
    return len(re.findall(r'[A-Za-z\u0600-\u06FF0-9]', t or ''))


def read_file(path, pages=2, work=None, force_ocr=False):
    """يرجّع dict: text, method (text/ocr/image), pages, ocr_pdf (للصور بس), warn."""
    import pypdf
    ext = os.path.splitext(path)[1].lower()
    work = work or tempfile.mkdtemp(prefix='dc_')
    res = dict(path=path, text='', method='', pages=0, ocr_pdf=None, warn='')
    if ext in IMG_EXT:
        base = os.path.join(work, 'img')
        res['text'] = clean(_tess(path, base, pdf=True))
        res['method'], res['pages'] = 'image', 1
        if os.path.isfile(base + '.pdf'): res['ocr_pdf'] = base + '.pdf'
        return res
    if ext != '.pdf':
        raise ReadError(f'Unsupported file type: {ext}')
    try:
        rd = pypdf.PdfReader(path)
        n = len(rd.pages)
    except Exception as e:
        raise ReadError(f'Cannot open PDF: {e}')
    res['pages'] = n
    txt = []
    for i in range(min(n, pages)):
        try: txt.append(rd.pages[i].extract_text() or '')
        except Exception: txt.append('')
    t = clean('\n'.join(txt))
    if _meaningful(t) >= 80 and not force_ocr and not _garbled(t):
        res['text'], res['method'] = t, 'text'
        return res
    # سكانر
    ok, why = ocr_status()
    if not ok:
        res['text'], res['method'], res['warn'] = t, 'text', why
        return res
    parts = []
    imgs = render_pages(path, min(n, pages), work)       # الأفضل: الصفحة كلها كصورة
    if not imgs:                                          # احتياطي: صورة السكانر اللي جوه الـ PDF
        for i in range(min(n, pages)):
            for ext_i, data in _page_images(rd.pages[i])[:1]:
                f = os.path.join(work, f'p{i}{ext_i}')
                with open(f, 'wb') as fh: fh.write(data)
                imgs.append(f)
    for f in imgs:
        try: parts.append(_tess(f))
        except ReadError as e: res['warn'] = str(e)
    res['text'], res['method'] = clean('\n'.join(parts)) or t, 'ocr'
    if not parts and not res['warn']: res['warn'] = 'No scanned image found in the PDF'
    return res


if __name__ == '__main__':                               # تجربة: python dc_read.py ملف1 ملف2 ...
    for p in sys.argv[1:]:
        try:
            r = read_file(p)
            print('=' * 80); print(os.path.basename(p), '|', r['method'], '| pages', r['pages'], '|', r['warn'])
            print(r['text'][:1500])
        except ReadError as e:
            print('ERR', p, e)
