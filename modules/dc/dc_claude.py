# -*- coding: utf-8 -*-
"""DC Control – مساعدة Claude (عن طريق Claude API).

لما القواعد مش متأكدة من ورقة، أو السكان طويل وممكن يكون فيه كذا طلب، البرنامج يبعت الصفحات لـ Claude:
يقرا الورقة كصورة، ويقول نوعها ورقمها وتاريخها، وكل مستند من صفحة كام لصفحة كام، وهل عليها توقيع/ختم.

المفتاح بيتحفظ على الجهاز بس (فولدر المستخدم – برا G:)، أو من متغير البيئة ANTHROPIC_API_KEY.
من غير مفتاح المساعدة مقفولة، والـ DC بيشتغل بالقواعد بس زي الأول.
"""
import os, re, io, json, base64, urllib.request, urllib.error
from dc_io import LOG, save_json, load_json

API = 'https://api.anthropic.com/v1/messages'
CFG_DIR = os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), 'CTR_PLATFORM')
CFG = os.path.join(CFG_DIR, 'claude.json')
MODELS = ['claude-sonnet-5-5', 'claude-haiku-5-5', 'claude-opus-5-5']
DEFAULT_MODEL = MODELS[0]
CLASSIFY_PAGES = 4            # الورقة اللي القواعد مش متأكدة منها: أول 4 صفحات بس
MAX_PAGES = 100               # السكان الطويل (ممكن يكون فيه كذا طلب): لحد 100 صفحة
BUNDLE_PAGES = 13             # سكان أطول من كده (طلب الموقع بمرفقاته ≈ 13 صفحة) ← نسأل عن تقسيمه
BUNDLE_TYPES = ('SITE_WIR', 'UNF_WIR', 'MIR', 'TEST', 'IPC_SIGNED', 'IPC')
IMG_TYPES = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif', '.webp': 'image/webp'}


class ClaudeError(Exception):
    pass


# ---------------------------------------------------------------- الإعدادات (على الجهاز بس)
def settings():
    s = load_json(CFG, {})
    return s if isinstance(s, dict) else {}


def save_settings(**kw):
    s = settings(); s.update(kw)
    os.makedirs(CFG_DIR, exist_ok=True)
    save_json(CFG, s)


def forget_key():
    s = settings(); s.pop('key', None)
    os.makedirs(CFG_DIR, exist_ok=True)
    save_json(CFG, s)


def key():
    return (os.environ.get('ANTHROPIC_API_KEY') or settings().get('key') or '').strip()


def model():
    m = settings().get('model')
    return m if m in MODELS else DEFAULT_MODEL


def enabled():
    """فيه مفتاح، والإرسال الأوتوماتيك مش مقفول."""
    return bool(key()) and settings().get('auto', True)


# ---------------------------------------------------------------- الاتصال
def _post(body, timeout=300):
    req = urllib.request.Request(API, data=json.dumps(body).encode('utf-8'), method='POST', headers={
        'x-api-key': key(), 'anthropic-version': '2023-06-01', 'content-type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        raw = e.read().decode('utf-8', 'ignore')
        try: raw = json.loads(raw)['error']['message']
        except Exception: pass
        why = {400: 'Claude could not read the request', 401: 'the API key is not valid', 403: 'the API key has no permission',
               413: 'the file is too big', 429: 'too many requests or no credit left', 529: 'Claude is busy – try again later'}
        raise ClaudeError(f"{why.get(e.code, f'HTTP {e.code}')} ({str(raw)[:160]})")
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise ClaudeError(f'no connection to Claude ({getattr(e, "reason", e)})')


def test():
    """تجربة المفتاح برسالة صغيرة. يرجّع اسم الموديل اللي رد."""
    if not key(): raise ClaudeError('no API key')
    r = _post({'model': model(), 'max_tokens': 5, 'messages': [{'role': 'user', 'content': 'ping'}]}, timeout=60)
    return r.get('model', model())


# ---------------------------------------------------------------- الملف ← محتوى الرسالة
def _pdf_bytes(path, pages):
    import pypdf
    rd = pypdf.PdfReader(path)
    n = len(rd.pages)
    take = min(n, pages or MAX_PAGES)
    if take == n and os.path.getsize(path) < 40_000_000:
        with open(path, 'rb') as fh: return fh.read(), n, n
    w = pypdf.PdfWriter()
    for i in range(take): w.add_page(rd.pages[i])
    buf = io.BytesIO(); w.write(buf)
    return buf.getvalue(), take, n


def _block(path, pages, alt_pdf=None):
    """(بلوك المحتوى، عدد الصفحات اللي اتبعتت، عدد صفحات الملف)"""
    ext = os.path.splitext(path)[1].lower()
    if ext in IMG_TYPES and os.path.getsize(path) <= 4_500_000:
        with open(path, 'rb') as fh: data = fh.read()
        return {'type': 'image', 'source': {'type': 'base64', 'media_type': IMG_TYPES[ext], 'data': base64.b64encode(data).decode()}}, 1, 1
    if ext != '.pdf':
        if alt_pdf and os.path.isfile(alt_pdf): path = alt_pdf          # صورة كبيرة/TIFF ← نسخة الـ PDF اللي الـ DC عملها
        else: raise ClaudeError('only PDF files and photos can be sent')
    data, sent, total = _pdf_bytes(path, pages)
    return {'type': 'document', 'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': base64.b64encode(data).decode()}}, sent, total


# ---------------------------------------------------------------- السؤال
HINTS = {
    'SITE_WIR': 'our Works Inspection Request form (طلب استلام أعمال, FORM-010-QUA-WIR) – number = "WIR Ser No."; its checklist, lab test reports, sketch and survey levels that follow belong to it',
    'UNF_WIR': 'CLIENT Unifier WIR printout (WIR-AP3500-Water…) – number = the 5-digit Unifier record number',
    'MIR': 'Material Inspection Request (MIR) – number = MIR no.',
    'MAR': 'Material Approval Request (MAR / M-no.) – number = M no.',
    'MS': 'Method Statement', 'ITP': 'Inspection and Test Plan', 'SD': 'Shop Drawing submittal',
    'SUB': 'subcontractor / supplier papers (profile, commercial registration, VAT certificate, national address, approval request) – company = their name',
    'NCR': 'Non-Conformance Report', 'TEST': 'a laboratory test report on its own (not attached to a WIR)',
    'CAL': 'calibration certificate', 'RFI': 'Request for Information – number = RFI no.',
    'OUT_YA': 'our letter to the Consultant (CTR-YA-WA-26-NNNN) – number = NNNN', 'OUT_CLIENT': 'our letter to CLIENT (CTR-CLIENT-WA-NNNN) – number = NNNN',
    'OUT_SUB': 'our letter to a supplier/subcontractor – number = letter no.', 'OUT_OTH': 'our letter to another party – number = letter no.',
    'IN_YA': 'letter from the Consultant (cons-26-N) – number = N', 'IN_CLIENT': 'letter / minutes from CLIENT', 'IN_SUB': 'letter from a supplier/subcontractor',
    'IN_OTH': 'letter from another party', 'PERMIT': 'excavation permit / municipality licence',
    'IPC': 'our Interim Payment Certificate (المستخلص الجاري) WITHOUT signatures – number = IPC no.',
    'IPC_SIGNED': 'our Interim Payment Certificate WITH the consultant/CLIENT signatures or stamps – number = IPC no.',
    'IPC_INV': 'CLIENT electronic tax invoice (فاتورة ضريبية) for an IPC (مستخلص رقم N) – number = IPC no.',
    'SUB_IPC': 'a subcontractor payment certificate / invoice to us – company = their name',
    'QUOTE': 'a quotation / price offer (عرض سعر) to us from a subcontractor, supplier, plant or laboratory – company = their name',
    'DN': 'a supplier delivery note (إذن توريد / Delivery Note) for materials delivered to us, alone or with its tax invoice – '
          'company = the supplier; also give "dn" (the delivery note no. exactly as written), "invoice" (invoice no.) and '
          '"items" [{"desc", "code", "qty", "unit"}] for every delivered item',
    'SUB_CONTRACT': 'a subcontract / labour or equipment hire contract between us (الطرف الأول) and a subcontractor – '
                    'number_raw = the contract no. exactly as written (e.g. 2610/0424); company = the subcontractor; '
                    'also give "approvals": which of the 3 approvals you can see – "legal" (legal affairs stamp / الشؤون القانونية), '
                    '"sub" (the subcontractor\'s own signature or stamp), "chair" (the chairman of the board\'s signature / رئيس مجلس الإدارة)',
    'PAYROLL': 'a payroll sheet (مسير رواتب) of our workers or equipment operators – it goes with the IPC; '
               'give "payroll": "Workers" or "Operators", and date = the payroll month',
}
SYSTEM = ('You are the document controller of the Contractor (CTR), the contractor of the Client Authority (CLIENT) '
          'project "Sample Water Networks Project", PO 000000000. The consultant is the Consultant (مكتب الاستشاري). '
          'You receive scanned or photographed papers in Arabic and English and you sort them. '
          'Report every separate document in the file, in page order, with its first and last page. A document starts at its own '
          'form, cover or letter page; the attachments after it (checklists, lab reports, sketches, survey levels, photos) belong to it. '
          '"signed" is true only if you can see handwritten signatures or stamps of the consultant or CLIENT on the main page. '
          'Dates as YYYY-MM-DD. Lower the confidence when you are not sure. Types:\n'
          + '\n'.join(f'- {c}: {h}' for c, h in HINTS.items()) + '\n- OTHER: anything else')


def _tool():
    doc = {'type': 'object', 'properties': {
        'first_page': {'type': 'integer'}, 'last_page': {'type': 'integer'},
        'type': {'type': 'string', 'enum': list(HINTS) + ['OTHER']},
        'number': {'type': 'string', 'description': 'the number for this type (see the list), digits only, or empty'},
        'number_raw': {'type': 'string', 'description': 'exactly as written: SUB_CONTRACT – the contract no. (e.g. 2610/0424); '
                       'a paper we submitted to the Consultant – our outgoing serial CTR-YA-WA-26-NNNN (not the previous submittal no.)'},
        'date': {'type': 'string', 'description': 'YYYY-MM-DD or empty'},
        'company': {'type': 'string', 'description': 'company name for SUB / SUB_IPC / DN, else empty'},
        'signed': {'type': 'boolean'}, 'title': {'type': 'string', 'description': 'short title as written on the paper'},
        'dn': {'type': 'string', 'description': 'DN only: delivery note no. as written'},
        'invoice': {'type': 'string', 'description': 'DN only: invoice no. as written'},
        'items': {'type': 'array', 'description': 'DN only: delivered items', 'items': {'type': 'object', 'properties': {
            'desc': {'type': 'string'}, 'code': {'type': 'string'}, 'qty': {'type': 'number'}, 'unit': {'type': 'string'}}}},
        'approvals': {'type': 'array', 'description': 'SUB_CONTRACT only: approvals you can see on it',
                      'items': {'type': 'string', 'enum': ['legal', 'sub', 'chair']}},
        'payroll': {'type': 'string', 'description': 'PAYROLL only: Workers or Operators'},
        'confidence': {'type': 'number', 'description': '0 to 1'}},
        'required': ['first_page', 'last_page', 'type', 'confidence']}
    return {'name': 'report_documents', 'description': 'Report every separate document found in the file, in page order.',
            'input_schema': {'type': 'object', 'properties': {'documents': {'type': 'array', 'items': doc}}, 'required': ['documents']}}


def clean_docs(raw, pages=None):
    """قايمة المستندات (من الـ API أو من نتيجة الشات) ← صفحات أرقام صحيحة جوه الملف، ورقم من أرقام بس."""
    docs = []
    for d in raw or []:
        if not isinstance(d, dict): continue
        try:
            a, b = int(d.get('first_page', 1)), int(d.get('last_page', d.get('first_page', 1)))
        except (TypeError, ValueError):
            continue
        if pages: a, b = max(1, min(a, pages)), max(1, min(b, pages))
        try: conf = float(d.get('confidence') or 0)
        except (TypeError, ValueError): conf = 0.0
        docs.append(dict(d, first_page=min(a, b), last_page=max(a, b), type=str(d.get('type', 'OTHER')).strip(),
                         number=re.sub(r'\D', '', str(d.get('number') or '')), confidence=conf,
                         number_raw=str(d.get('number_raw') or (d.get('number') if str(d.get('type')).strip() == 'SUB_CONTRACT' else '')
                                        or '').strip()))
    return docs


def ask(path, bundle=False, alt_pdf=None):
    """يبعت الملف لـ Claude ويرجّع {'documents': [...], 'pages_sent': n, 'pages': total, 'model': …}.
    bundle=True ← الملف كله (لحد 100 صفحة) عشان يحدد كل مستند فيه؛ غير كده أول 4 صفحات بس."""
    if not key(): raise ClaudeError('no API key')
    blk, sent, total = _block(path, MAX_PAGES if bundle else CLASSIFY_PAGES, alt_pdf)
    what = (f'This scan has {sent} pages and may contain several documents scanned together. List each one with its page range, covering all pages.'
            if bundle else f'These are the first {sent} page(s) of the file. Report the main document (first_page 1).')
    body = {'model': model(), 'max_tokens': 4000, 'system': SYSTEM, 'tools': [_tool()],
            'tool_choice': {'type': 'tool', 'name': 'report_documents'},
            'messages': [{'role': 'user', 'content': [blk, {'type': 'text', 'text': what}]}]}
    r = _post(body)
    out = next((c.get('input') for c in r.get('content', []) if c.get('type') == 'tool_use'), None)
    if not isinstance(out, dict) or not isinstance(out.get('documents'), list):
        raise ClaudeError('Claude did not return the documents list')
    docs = clean_docs(out['documents'], sent)
    u = r.get('usage', {})
    LOG.info('Claude %s: %s – %d page(s) sent, %d document(s), tokens in %s / out %s', r.get('model'), os.path.basename(path),
             sent, len(docs), u.get('input_tokens'), u.get('output_tokens'))
    return {'documents': docs, 'pages_sent': sent, 'pages': total, 'model': r.get('model'), 'bundle': bundle}
