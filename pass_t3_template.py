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


# Pass type → color + label mapping
PASS_TYPE_CONFIG = {
    'karyakarta':    {'color': '#0f52ba', 'label': 'कार्यकर्ता पास',    'accent': '#FFDD96'},
    'vishesh_atithi':{'color': '#8B0000', 'label': 'विशेष अतिथि पास',  'accent': '#FFD700'},
    'vip':           {'color': '#0A0A0A', 'label': 'VIP पास',           'accent': '#C8A04A'},
    'press':         {'color': '#1A5C2A', 'label': 'प्रेस पास',         'accent': '#FFFFFF'},
    'seva':          {'color': '#0F5C4A', 'label': 'सेवा पास',          'accent': '#FFDD96'},
    'staff':         {'color': '#1A2C5C', 'label': 'स्टाफ पास',         'accent': '#FFFFFF'},
}

def get_pass_type_style(pass_type):
    return PASS_TYPE_CONFIG.get(pass_type or 'karyakarta', PASS_TYPE_CONFIG['karyakarta'])

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
  border-bottom: 0.4mm solid {{ accent_color }};
}

.event-date {
  font-family: 'NotoDeva', serif;
  font-size: 12pt;
  font-weight: 700;
  color: white;
  margin-top: 3mm;
  display: inline-block;
  padding: 0.3mm 4mm 0.8mm 4mm;
  border-bottom: 0.3mm solid #C8A04A;
}

.body-table {
  position: absolute;
  top: 50mm;
  left: 0;
  right: 0;
  bottom: 28mm;
  width: 100%;
  padding: 0 6mm;
  z-index: 2;
  border-collapse: collapse;
}

.logo-cell {
  width: 44mm;
  vertical-align: middle;
  text-align: center;
  padding: 2mm 4mm 2mm 2mm;
}

.fields-cell {
  vertical-align: middle;
  padding: 0;
}

.qr-cell {
  width: 40mm;
  vertical-align: middle;
  text-align: center;
  padding: 2mm;
}

.fields-table {
  border-collapse: collapse;
  width: 100%;
}

.fields-table tr td {
  padding: 2mm 1mm;
  vertical-align: middle;
}

.fl {
  font-family: 'NotoDeva', sans-serif;
  font-size: 10pt;
  font-weight: 600;
  color: #404040;
  white-space: nowrap;
  width: 28mm;
}

.fc {
  font-size: 12pt;
  font-weight: 700;
  color: #606060;
  width: 4mm;
  text-align: center;
}

.fv {
  font-family: 'NotoDeva', sans-serif;
  font-size: 13pt;
  font-weight: 700;
  color: #1A1A1A;
}

.id-val {
  font-family: 'Poppins', sans-serif;
  font-size: 14pt;
  letter-spacing: 0.5pt;
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


.logo-wrap {
  width: 40mm;
  height: 40mm;
  border-radius: 50%;
  background: white;
  padding: 1.5mm;
  display: inline-block;
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

.qr {
  width: 34mm;
  height: 34mm;
  background: white;
  padding: 1mm;
  border: 0.2mm solid #DDD;
  border-radius: 1.5mm;
  display: inline-block;
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
  font-size: 12pt;
  font-weight: 700;
  color: white;
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
  letter-spacing: 0.5pt;
  margin-bottom: 1mm;
}
"""


PASS_DIV_TEMPLATE = Template(r"""
<div class="pass">
  {% if bg_image_url %}<div class="bg-temple" style="background-image: url('{{ bg_image_url }}')"></div>{% endif %}

  <div class="header" style="background: linear-gradient(180deg, {{ header_color_light }} 0%, {{ header_color }} 100%) !important;">
    <div class="org-name">{{ org }}</div>
    {% if event %}<div><span class="event-name" style="color: {{ accent_color }} !important; border-bottom-color: {{ accent_color }} !important;">{{ event }}</span></div>{% endif %}
    {% if date_hi %}<div><span class="event-date">{{ pass_label }}- {{ date_hi }} तक मान्य</span></div>{% endif %}
  </div>

  <table class="body-table">
    <tr>
      <td class="logo-cell">
        <div class="logo-wrap">
          <div class="logo-inner"{% if logo_url %} style="background-image: url('{{ logo_url }}')"{% endif %}></div>
        </div>
      </td>
      <td class="fields-cell">
        <table class="fields-table">
          <tr><td class="fl">आई. डी. कोड</td><td class="fc">:</td><td class="fv id-val">{{ vol_id }}</td></tr>
          <tr><td class="fl">नाम</td><td class="fc">:</td><td class="fv">{{ name }}</td></tr>
          <tr><td class="fl">आधार</td><td class="fc">:</td><td class="fv">{{ aadhaar }}</td></tr>
          <tr><td class="fl">दायित्व</td><td class="fc">:</td><td class="fv">{{ role }}</td></tr>
          {% if permission %}<tr><td class="fl">अनुमति</td><td class="fc">:</td><td class="fv">{{ permission }}</td></tr>{% endif %}
        </table>
      </td>
      <td class="qr-cell">
        <div class="qr"><img src="{{ qr_dataurl }}" alt="QR"/></div>
        <div class="qr-label">SCAN TO VERIFY</div>
      </td>
    </tr>
  </table>

  <div class="footer" style="background: linear-gradient(180deg, {{ header_color_light }} 0%, {{ header_color }} 100%) !important;">
    <div class="notes">
      {% if note1 %}<div class="note">{{ note1 }}</div>{% endif %}
      {% if note2 %}<div class="note">{{ note2 }}</div>{% endif %}
    </div>

    <div class="authority">
      {% if signing_image %}<img src="{{ signing_image }}" class="sign-image" alt="Signature"/>{% endif %}
      <div class="sign-issuing">Issuing Authority :</div>
      {% if signing_name %}<div class="sign-name">{{ signing_name }}</div>{% endif %}
      {% if signing_title %}<div class="sign-title">{{ signing_title }}</div>{% endif %}
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

    # Pass type → dynamic color + label
    pass_type  = str(vol.get('pass_type') or 'karyakarta').strip().lower()
    pt_style   = get_pass_type_style(pass_type)
    header_color = pt_style['color']
    # Lighter version of header for gradient top
    header_color_light = {
        '#0f52ba': '#1a6fd4',
        '#8B0000': '#C0392B',
        '#0A0A0A': '#2A2A2A',
        '#0F5C4A': '#1A9A7A',
        '#1A5C2A': '#247A38',
        '#1A2C5C': '#243D7A',
    }.get(header_color, header_color)
    pass_label   = pt_style['label']
    accent_color = pt_style['accent']
    qr_data    = f"{verify_url}/{vol_id}" if verify_url else vol_id
    qr_dataurl = make_qr_dataurl(qr_data, box_size=8)

    return dict(
        org=org, event=event_name, date_hi=date_hi,
        vol_id=vol_id, name=name, role=role, permission=permission, aadhaar=aadhaar,
        logo_url=logo_url, bg_image_url=bg_image,
        qr_dataurl=qr_dataurl,
        note1=note1, note2=note2,
        signing_image=sig_img, signing_name=sig_name, signing_title=sig_title,
        header_color=header_color, header_color_light=header_color_light, pass_label=pass_label, accent_color=accent_color,
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
