"""
T3 Pass Template — HTML-based rendering via WeasyPrint.
Produces pixel-perfect passes matching the Dhwajarohan reference design.
"""

import io
import os
import base64
from jinja2 import Template
from weasyprint import HTML

try:
    import qrcode
    _HAS_QRCODE = True
except ImportError:
    _HAS_QRCODE = False

_FONTS_DIR = os.path.join(os.path.dirname(__file__), 'fonts')

def _font_url(filename):
    path = os.path.join(_FONTS_DIR, filename)
    return f"file://{path}"

MONTHS_HI = {
    '01': 'जनवरी', '02': 'फ़रवरी', '03': 'मार्च', '04': 'अप्रैल',
    '05': 'मई', '06': 'जून', '07': 'जुलाई', '08': 'अगस्त',
    '09': 'सितंबर', '10': 'अक्टूबर', '11': 'नवंबर', '12': 'दिसंबर',
}


def fix_image_url(url):
    """Convert Google Drive share links to direct image URLs."""
    if not url:
        return url
    # Handle Google Drive /file/d/{id}/view links
    import re
    m = re.search(r'drive[.]google[.]com/file/d/([a-zA-Z0-9_-]+)', url)
    if m:
        return f"https://drive.google.com/uc?export=view&id={m.group(1)}"
    # Handle Google Drive open?id= links
    m2 = re.search(r'drive[.]google[.]com/open[?]id=([a-zA-Z0-9_-]+)', url)
    if m2:
        return f"https://drive.google.com/uc?export=view&id={m2.group(1)}"
    return url


def logo_file_as_dataurl(filename):
    """Read a logo file from the same directory and return as base64 data URL."""
    try:
        filepath = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(filepath):
            return ''
        with open(filepath, 'rb') as f:
            data = f.read()
        ext = filename.rsplit('.', 1)[-1].lower()
        ct = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
              'gif': 'image/gif', 'webp': 'image/webp'}.get(ext, 'image/png')
        return f"data:{ct};base64,{base64.b64encode(data).decode()}"
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning(f"logo_file_as_dataurl failed: {e}")
        return ''


def fetch_image_as_dataurl(url):
    """Fetch image URL → base64 data URL. Uses requests for redirect/auth handling."""
    if not url:
        return url
    try:
        import requests as _req
        resp = _req.get(url, timeout=10, allow_redirects=True, headers={
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120.0',
            'Accept': 'image/*,*/*;q=0.8',
        })
        resp.raise_for_status()
        ct = resp.headers.get('Content-Type', 'image/png').split(';')[0].strip()
        # If Google Drive returns HTML (login page), bail out
        if 'text/html' in ct:
            import logging
            logging.getLogger(__name__).warning(f"fetch_image_as_dataurl: got HTML for {url} — image may be private")
            return url
        b64 = base64.b64encode(resp.content).decode()
        return f"data:{ct};base64,{b64}"
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning(f"fetch_image_as_dataurl failed for {url}: {e}")
        return url


def format_date_hi(date_str):
    if not date_str:
        return ''
    parts = str(date_str).split('-')
    if len(parts) != 3:
        return date_str
    dd, mm, yyyy = parts
    if mm in MONTHS_HI:
        try:
            return f"{int(dd)} {MONTHS_HI[mm]} {yyyy}"
        except ValueError:
            return date_str
    return date_str


def make_qr_dataurl(text, box_size=8):
    if not _HAS_QRCODE or not text:
        return ''
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=1,
    )
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()


T3_CSS = """
@page {
  size: 210mm 148mm;
  margin: 0;
}

@font-face {
  font-family: 'NotoDeva';
  src: url('""" + _font_url('NotoSansDevanagari-Regular.ttf') + """') format('truetype');
  font-weight: 400;
}
@font-face {
  font-family: 'NotoDeva';
  src: url('""" + _font_url('NotoSansDevanagari-Bold.ttf') + """') format('truetype');
  font-weight: 700;
}
@font-face {
  font-family: 'Poppins';
  src: url('""" + _font_url('Poppins-Regular.ttf') + """') format('truetype');
  font-weight: 400;
}
@font-face {
  font-family: 'Poppins';
  src: url('""" + _font_url('Poppins-Bold.ttf') + """') format('truetype');
  font-weight: 700;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

html, body {
  font-family: 'NotoDeva', 'Poppins', sans-serif;
  background: #F4F0E5;
}

.pass {
  position: relative;
  width: 210mm;
  height: 148mm;
  overflow: hidden;
  background: #F4F0E5;
  page-break-after: always;
}
.pass:last-child { page-break-after: auto; }

.bg-temple {
  position: absolute;
  inset: 0;
  background-size: cover;
  background-position: center;
  opacity: 0.08;
  z-index: 0;
}

.header {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 50mm;
  background: linear-gradient(180deg, #1a6fd4 0%, #0f52ba 100%);
  border-radius: 0 0 50% 50% / 0 0 20% 20%;
  z-index: 1;
  padding: 5mm 10mm 0 10mm;
  text-align: center;
}

.header::before {
  content: '';
  position: absolute;
  inset: 0;
  background-image: repeating-linear-gradient(
    90deg,
    rgba(255,255,255,0.04) 0,
    rgba(255,255,255,0.04) 2px,
    transparent 2px,
    transparent 10px
  );
  border-radius: inherit;
  pointer-events: none;
}

.kartakar-label {
  font-family: 'NotoDeva', serif;
  font-size: 10pt;
  font-weight: 500;
  color: rgba(255,255,255,0.85);
  letter-spacing: 1pt;
  margin-top: 1mm;
}

.org-name {
  font-family: 'NotoDeva', serif;
  font-size: 26pt;
  font-weight: 700;
  color: white;
  line-height: 1.0;
  margin-top: 1mm;
}

.event-name {
  font-family: 'NotoDeva', serif;
  font-size: 15pt;
  font-weight: 700;
  color: #FFDD96;
  margin-top: 3mm;
  display: inline-block;
  padding: 0.5mm 3mm 1mm 3mm;
  border-bottom: 0.4mm solid #C8A04A;
}

.event-date {
  font-family: 'NotoDeva', serif;
  font-size: 11pt;
  color: white;
  margin-top: 3mm;
  display: inline-block;
  padding: 0.3mm 4mm 0.8mm 4mm;
  border-bottom: 0.3mm solid #C8A04A;
}

.body {
  position: absolute;
  top: 50mm;
  left: 0;
  right: 0;
  bottom: 28mm;
  padding: 2mm 8mm 0 8mm;
  z-index: 2;
  display: flex;
  align-items: center;
  gap: 5mm;
}

.logo-wrap {
  flex: 0 0 auto;
  width: 40mm;
  height: 40mm;
  border-radius: 50%;
  background: white;
  padding: 1.5mm;
}

.logo-inner {
  width: 100%;
  height: 100%;
  border-radius: 50%;
  background-color: #F4F0E5;
  background-size: cover;
  background-position: center;
  background-repeat: no-repeat;
}

.fields {
  flex: 1;
  display: block;
  padding-right: 2mm;
  padding-top: 4mm;
}

.field-row {
  display: flex;
  align-items: baseline;
  font-family: 'NotoDeva', sans-serif;
  margin-bottom: 2.5mm;
}

.field-label {
  flex: 0 0 28mm;
  font-size: 10pt;
  font-weight: 600;
  color: #404040;
}

.field-colon {
  flex: 0 0 3mm;
  font-size: 13pt;
  font-weight: 700;
  color: #606060;
}

.field-value {
  flex: 1;
  font-size: 13pt;
  font-weight: 700;
  color: #1A1A1A;
}

.qr-col {
  flex: 0 0 auto;
  text-align: center;
  padding-top: 1mm;
}

.qr {
  width: 34mm;
  height: 34mm;
  background: white;
  padding: 1mm;
  border: 0.2mm solid #DDD;
  border-radius: 1.5mm;
}

.qr img {
  width: 100%;
  height: 100%;
  display: block;
}

.qr-label {
  font-family: 'Poppins', sans-serif;
  font-size: 5pt;
  color: #999;
  letter-spacing: 0.5pt;
  margin-top: 0.8mm;
}

.footer {
  position: absolute;
  bottom: 0;
  left: 0;
  right: 0;
  height: 28mm;
  background: linear-gradient(180deg, #1a6fd4 0%, #0f52ba 100%);
  z-index: 2;
  padding: 4mm 10mm 3mm 10mm;
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
}

.footer::before {
  content: '';
  position: absolute;
  inset: 0;
  background-image: repeating-linear-gradient(
    90deg,
    rgba(255,255,255,0.04) 0,
    rgba(255,255,255,0.04) 2px,
    transparent 2px,
    transparent 10px
  );
  pointer-events: none;
}

.notes {
  flex: 1;
  z-index: 1;
  padding-top: 1mm;
}

.note {
  font-family: 'NotoDeva', sans-serif;
  font-size: 9pt;
  color: white;
  line-height: 1.5;
  font-weight: 500;
}

.note::before {
  content: '* ';
  color: #FFDD96;
  font-weight: 700;
}

.authority {
  flex: 0 0 55mm;
  text-align: center;
  z-index: 1;
  padding-top: 0.5mm;
}

.sign-image {
  max-height: 11mm;
  max-width: 40mm;
  display: block;
  margin: 0 auto;
  filter: brightness(0) invert(1);
}

.sign-name {
  font-family: 'NotoDeva', serif;
  font-size: 11pt;
  font-weight: 700;
  color: white;
  margin-top: 0.5mm;
  line-height: 1.2;
}

.sign-title {
  font-family: 'NotoDeva', sans-serif;
  font-size: 9pt;
  color: #E8E8E8;
  margin-top: 0.3mm;
}

.sign-issuing {
  font-family: 'Poppins', sans-serif;
  font-size: 8pt;
  font-weight: 700;
  color: #FFDD96;
  letter-spacing: 1pt;
  margin-top: 1.5mm;
}
"""


PASS_DIV_TEMPLATE = Template(r"""
<div class="pass">
  {% if bg_image_url %}<div class="bg-temple" style="background-image: url('{{ bg_image_url }}')"></div>{% endif %}

  <div class="header">
    <div class="org-name">{{ org }}</div>
    {% if event %}<div><span class="event-name">{{ event }}</span></div>{% endif %}
    {% if date_hi %}<div><span class="event-date">कार्यकर्ता पास- {{ date_hi }} तक मान्य</span></div>{% endif %}
  </div>

  <div class="body">
    <div class="logo-wrap">
      <div class="logo-inner"{% if logo_url %} style="background-image: url('{{ logo_url }}')"{% endif %}></div>
    </div>

    <div class="fields">
      <div class="field-row">
        <span class="field-label">आई. डी. कोड</span>
        <span class="field-colon">:</span>
        <span class="field-value">{{ vol_id }}</span>
      </div>
      <div class="field-row">
        <span class="field-label">नाम</span>
        <span class="field-colon">:</span>
        <span class="field-value">{{ name }}</span>
      </div>
      <div class="field-row">
        <span class="field-label">आधार</span>
        <span class="field-colon">:</span>
        <span class="field-value">{{ aadhaar }}</span>
      </div>
      <div class="field-row">
        <span class="field-label">दायित्व</span>
        <span class="field-colon">:</span>
        <span class="field-value">{{ role }}</span>
      </div>
      {% if permission %}
      <div class="field-row">
        <span class="field-label">अनुमति</span>
        <span class="field-colon">:</span>
        <span class="field-value">{{ permission }}</span>
      </div>
      {% endif %}
    </div>

    <div class="qr-col">
      <div class="qr"><img src="{{ qr_dataurl }}" alt="QR"/></div>
      <div class="qr-label">SCAN TO VERIFY</div>
    </div>
  </div>

  <div class="footer">
    <div class="notes">
      {% if note1 %}<div class="note">{{ note1 }}</div>{% endif %}
      {% if note2 %}<div class="note">{{ note2 }}</div>{% endif %}
    </div>

    <div class="authority">
      {% if signing_image %}<img src="{{ signing_image }}" class="sign-image" alt="Signature"/>{% endif %}
      {% if signing_name %}<div class="sign-name">{{ signing_name }}</div>{% endif %}
      {% if signing_title %}<div class="sign-title">{{ signing_title }}</div>{% endif %}
      <div class="sign-issuing">Issuing Authority</div>
    </div>
  </div>
</div>
""")


def _pass_context(vol, event=None):
    event = event or {}
    org        = str(vol.get('org') or event.get('org_name') or 'श्री राम जन्मभूमि तीर्थ क्षेत्र').strip()
    event_name = str(vol.get('event_label') or event.get('name') or '').strip()
    date_raw   = str(vol.get('expiry') or event.get('expiry_date') or '').strip()
    date_hi    = format_date_hi(date_raw)
    vol_id     = str(vol.get('id') or '').strip()
    name       = str(vol.get('name_hi') or vol.get('name') or '').strip()
    role       = str(vol.get('role') or '').strip()
    permission = str(vol.get('permission') or '').strip()
    aadhaar    = str(vol.get('aadhaar') or '').strip()
    # Mask aadhaar — show only last 4
    if aadhaar and len(aadhaar.replace(' ','')) >= 4:
        digits = aadhaar.replace(' ','')
        aadhaar = 'XXXX XXXX ' + digits[-4:]
    _raw_logo = fix_image_url(str(vol.get('logo_url') or event.get('logo_url') or '').strip())
    # If it's the local Render-hosted logo, read from disk directly (faster + reliable)
    if 'passflow-pass-generator.onrender.com/static/logo/' in _raw_logo:
        _fname = _raw_logo.rsplit('/', 1)[-1]
        logo_url = logo_file_as_dataurl(_fname) or fetch_image_as_dataurl(_raw_logo)
    else:
        logo_url = fetch_image_as_dataurl(_raw_logo)
    bg_image   = fetch_image_as_dataurl(fix_image_url(str(vol.get('bg_image') or event.get('bg_image') or '').strip()))
    sig_img    = fetch_image_as_dataurl(fix_image_url(str(vol.get('signing_image') or event.get('signing_image') or '').strip()))
    sig_name   = str(vol.get('signing_authority') or '').strip()
    sig_title  = str(vol.get('signing_title') or '').strip()

    notes_raw = str(vol.get('backside_text') or event.get('backside_text') or '').strip()
    if notes_raw:
        lines = notes_raw.split('\n')
        note1 = lines[0].strip() if len(lines) > 0 else ''
        note2 = lines[1].strip() if len(lines) > 1 else ''
    else:
        note1 = 'यह प्रवेश-पत्र आधार कार्ड के साथ ही मान्य है।'
        note2 = 'मंदिर परिसर में मोबाइल/कैमरा इत्यादि पूर्णतः प्रतिबंधित है।'

    verify_url = str(vol.get('verify_url') or '').strip()
    qr_data    = f"{verify_url}/{vol_id}" if verify_url else vol_id
    qr_dataurl = make_qr_dataurl(qr_data, box_size=8)

    return dict(
        org=org, event=event_name, date_hi=date_hi,
        vol_id=vol_id, name=name, role=role, permission=permission, aadhaar=aadhaar,
        logo_url=logo_url, bg_image_url=bg_image,
        qr_dataurl=qr_dataurl,
        note1=note1, note2=note2,
        signing_image=sig_img, signing_name=sig_name, signing_title=sig_title,
    )


def _build_full_html(volunteers, event=None):
    passes_html = ''.join(PASS_DIV_TEMPLATE.render(**_pass_context(v, event)) for v in volunteers)
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{T3_CSS}</style></head>
<body>{passes_html}</body></html>"""


def render_t3_pdf(vol, event=None):
    """Render a single T3 pass as PDF bytes."""
    html_str = _build_full_html([vol], event)
    return HTML(string=html_str, base_url=os.path.dirname(__file__)).write_pdf()


def render_t3_multi_pdf(volunteers, event=None):
    """Render multiple T3 passes (one per page) as a single PDF."""
    if not volunteers:
        raise ValueError("No volunteers provided")
    html_str = _build_full_html(volunteers, event)
    return HTML(string=html_str, base_url=os.path.dirname(__file__)).write_pdf()
