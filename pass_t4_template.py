"""
T4 Pass Template — Saffron Rally
Target: Political events, party worker passes
Orientation: A6 Landscape (148mm × 105mm)
Visual: Saffron/orange gradient header, tricolor accent strip, bold Hindi name
"""

import io, os, base64, re
from jinja2 import Template
from weasyprint import HTML

try:
    import qrcode as _qrcode
    _HAS_QRCODE = True
except ImportError:
    _HAS_QRCODE = False

_FONTS_DIR = os.path.join(os.path.dirname(__file__), 'fonts')

def _font_url(f): return f"file://{os.path.join(_FONTS_DIR, f)}"

MONTHS_HI = {
    '01':'जनवरी','02':'फ़रवरी','03':'मार्च','04':'अप्रैल',
    '05':'मई','06':'जून','07':'जुलाई','08':'अगस्त',
    '09':'सितंबर','10':'अक्टूबर','11':'नवंबर','12':'दिसंबर',
}

def format_date_hi(date_str):
    if not date_str: return ''
    for sep in ['-','/']:
        parts = date_str.split(sep)
        if len(parts) == 3:
            dd,mm,yyyy = parts
            if mm in MONTHS_HI:
                return f"{int(dd)} {MONTHS_HI[mm]} {yyyy}"
    return date_str

def mask_aadhaar(aadhaar):
    if not aadhaar: return ''
    digits = re.sub(r'\D','',str(aadhaar))
    if len(digits) >= 4:
        return 'XXXX XXXX ' + digits[-4:]
    return aadhaar

def fix_image_url(url):
    if not url: return url
    m = re.search(r'drive[.]google[.]com/file/d/([a-zA-Z0-9_-]+)', url)
    if m: return f"https://drive.google.com/uc?export=view&id={m.group(1)}"
    m2 = re.search(r'drive[.]google[.]com/open[?]id=([a-zA-Z0-9_-]+)', url)
    if m2: return f"https://drive.google.com/uc?export=view&id={m2.group(1)}"
    return url

def fetch_image_as_dataurl(url):
    if not url: return ''
    try:
        import requests as _req
        resp = _req.get(fix_image_url(url), timeout=10, allow_redirects=True,
            headers={'User-Agent':'Mozilla/5.0','Accept':'image/*'})
        resp.raise_for_status()
        ct = resp.headers.get('Content-Type','image/png').split(';')[0].strip()
        if 'text/html' in ct: return ''
        return f"data:{ct};base64,{base64.b64encode(resp.content).decode()}"
    except: return ''

def logo_file_as_dataurl(filename):
    try:
        path = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(path): return ''
        with open(path,'rb') as f: data = f.read()
        ext = filename.rsplit('.',1)[-1].lower()
        ct = {'png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg'}.get(ext,'image/png')
        return f"data:{ct};base64,{base64.b64encode(data).decode()}"
    except: return ''

def make_qr_dataurl(data):
    if not _HAS_QRCODE or not data: return ''
    try:
        qr = _qrcode.QRCode(version=1, error_correction=_qrcode.constants.ERROR_CORRECT_M,
            box_size=6, border=2)
        qr.add_data(data); qr.make(fit=True)
        img = qr.make_image(fill_color='#1A1A1A', back_color='white')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"
    except: return ''

CSS = """
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_b}') format('truetype'); font-weight:700; }}
@font-face {{ font-family:'Inter'; src: url('{inter_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'Inter'; src: url('{inter_b}') format('truetype'); font-weight:700; }}

@page {{
  size: 148mm 105mm;
  margin: 0;
}}

* {{ box-sizing: border-box; margin: 0; padding: 0; }}

body {{
  width: 148mm;
  height: 105mm;
  overflow: hidden;
  font-family: 'NotoSansDevanagari', 'Inter', sans-serif;
  background: #FFFDF8;
  position: relative;
}}

/* Tricolor accent strip at very top */
.tricolor {{
  width: 100%;
  height: 2.5mm;
  display: flex;
}}
.tc-s {{ flex: 1; background: #FF9933; }}
.tc-w {{ flex: 1; background: #FFFFFF; border-top: 0.3mm solid #eee; border-bottom: 0.3mm solid #eee; }}
.tc-g {{ flex: 1; background: #138808; }}

/* Saffron header */
.header {{
  background: linear-gradient(135deg, #FF6B00 0%, #FF9933 50%, #FFB347 100%);
  padding: 3mm 5mm 3mm 5mm;
  display: flex;
  align-items: center;
  gap: 3mm;
  min-height: 22mm;
}}

.header-logo {{
  width: 16mm;
  height: 16mm;
  border-radius: 50%;
  background: white;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  border: 0.8mm solid rgba(255,255,255,0.6);
}}
.header-logo img {{ width: 100%; height: 100%; object-fit: cover; border-radius: 50%; }}
.header-logo-placeholder {{
  width: 100%; height: 100%;
  background: linear-gradient(135deg, #FF6B00, #CC4400);
  border-radius: 50%;
}}

.header-text {{ flex: 1; }}
.org-name {{
  font-family: 'NotoSansDevanagari', sans-serif;
  font-size: 11pt;
  font-weight: 700;
  color: #fff;
  line-height: 1.2;
  text-shadow: 0 1px 2px rgba(0,0,0,0.2);
}}
.event-name {{
  font-family: 'NotoSansDevanagari', sans-serif;
  font-size: 8.5pt;
  font-weight: 400;
  color: rgba(255,255,255,0.92);
  margin-top: 1mm;
  line-height: 1.2;
}}
.event-date {{
  font-size: 7.5pt;
  color: rgba(255,255,255,0.8);
  margin-top: 1.5mm;
  font-family: 'Inter', sans-serif;
}}

.header-badge {{
  background: rgba(255,255,255,0.2);
  border: 0.5mm solid rgba(255,255,255,0.5);
  border-radius: 2mm;
  padding: 1.5mm 3mm;
  font-size: 7pt;
  font-weight: 700;
  color: #fff;
  text-align: center;
  font-family: 'Inter', sans-serif;
  letter-spacing: 0.5pt;
  white-space: nowrap;
  flex-shrink: 0;
}}

/* Body */
.body {{
  display: flex;
  gap: 0;
  padding: 3mm 4mm 2mm 4mm;
  flex: 1;
}}

.fields {{ flex: 1; }}

.field-row {{
  display: flex;
  align-items: baseline;
  padding: 1.2mm 0;
  border-bottom: 0.2mm solid #F0E8D8;
  gap: 2mm;
}}
.field-row:last-child {{ border-bottom: none; }}

.field-label {{
  font-size: 6.5pt;
  color: #888;
  font-weight: 400;
  min-width: 22mm;
  flex-shrink: 0;
  font-family: 'Inter', sans-serif;
  text-transform: uppercase;
  letter-spacing: 0.3pt;
}}
.field-value {{
  font-family: 'NotoSansDevanagari', 'Inter', sans-serif;
  font-size: 8pt;
  font-weight: 700;
  color: #1A1A1A;
  line-height: 1.2;
  flex: 1;
}}
.field-value.large {{
  font-size: 10.5pt;
  color: #CC4400;
}}
.field-value.en {{
  font-family: 'Inter', sans-serif;
}}

/* QR block */
.qr-block {{
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 1mm;
  padding-left: 3mm;
  border-left: 0.3mm solid #F0E8D8;
  min-width: 22mm;
}}
.qr-img {{ width: 20mm; height: 20mm; border: 0.5mm solid #eee; }}
.qr-id {{
  font-size: 6pt;
  color: #888;
  text-align: center;
  font-family: 'Inter', sans-serif;
  word-break: break-all;
}}

/* Footer */
.footer {{
  background: linear-gradient(90deg, #FF6B00, #CC4400);
  padding: 1.5mm 4mm;
  display: flex;
  justify-content: space-between;
  align-items: center;
}}
.footer-sign {{
  font-family: 'NotoSansDevanagari', sans-serif;
  font-size: 6.5pt;
  color: rgba(255,255,255,0.9);
  line-height: 1.3;
}}
.footer-sign strong {{ color: #fff; font-weight: 700; }}
.footer-right {{
  font-size: 6pt;
  color: rgba(255,255,255,0.7);
  font-family: 'Inter', sans-serif;
  text-align: right;
}}
""".format(
    deva_r=_font_url('NotoSansDevanagari-Regular.ttf'),
    deva_b=_font_url('NotoSansDevanagari-Bold.ttf'),
    inter_r=_font_url('Inter-Regular.ttf'),
    inter_b=_font_url('Inter-Bold.ttf'),
)

TEMPLATE = Template("""
<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>{{ css }}</style>
</head><body>

<div class="tricolor"><div class="tc-s"></div><div class="tc-w"></div><div class="tc-g"></div></div>

<div class="header">
  <div class="header-logo">
    {% if logo_data %}<img src="{{ logo_data }}">{% else %}<div class="header-logo-placeholder"></div>{% endif %}
  </div>
  <div class="header-text">
    <div class="org-name">{{ org_name }}</div>
    <div class="event-name">{{ event_name }}</div>
    {% if event_date %}<div class="event-date">📅 {{ event_date }}</div>{% endif %}
  </div>
  <div class="header-badge">कार्यकर्ता<br>पास</div>
</div>

<div class="body">
  <div class="fields">
    <div class="field-row">
      <div class="field-label">पास आई.डी.</div>
      <div class="field-value en">{{ vol_id }}</div>
    </div>
    <div class="field-row">
      <div class="field-label">नाम</div>
      <div class="field-value large">{{ name }}</div>
    </div>
    {% if role %}
    <div class="field-row">
      <div class="field-label">दायित्व</div>
      <div class="field-value">{{ role }}</div>
    </div>
    {% endif %}
    {% if mobile %}
    <div class="field-row">
      <div class="field-label">मोबाइल</div>
      <div class="field-value en">{{ mobile }}</div>
    </div>
    {% endif %}
    {% if aadhaar %}
    <div class="field-row">
      <div class="field-label">आधार</div>
      <div class="field-value en">{{ aadhaar }}</div>
    </div>
    {% endif %}
    {% if permission %}
    <div class="field-row">
      <div class="field-label">अनुमति</div>
      <div class="field-value">{{ permission }}</div>
    </div>
    {% endif %}
    {% if expiry %}
    <div class="field-row">
      <div class="field-label">वैधता</div>
      <div class="field-value en">{{ expiry }} तक</div>
    </div>
    {% endif %}
  </div>
  <div class="qr-block">
    {% if qr_data %}<img class="qr-img" src="{{ qr_data }}">{% endif %}
    <div class="qr-id">{{ vol_id }}</div>
  </div>
</div>

<div class="footer">
  <div class="footer-sign">
    जारीकर्ता: <strong>{{ signing_authority }}</strong><br>
    {{ signing_title }}
  </div>
  <div class="footer-right">मान्यता: {{ expiry }}<br>यह पास हस्तांतरणीय नहीं है</div>
</div>

</body></html>
""")


def generate_pass_t4(volunteer: dict, event: dict, qr_url: str = '') -> bytes:
    name = volunteer.get('name') or volunteer.get('name_hi') or ''
    vol_id = volunteer.get('id','')
    mobile = str(volunteer.get('mobile','')).replace('+91','').strip()
    aadhaar = mask_aadhaar(volunteer.get('aadhaar',''))
    role = volunteer.get('role','') or ''
    permission = volunteer.get('permission','') or ''
    expiry = volunteer.get('expiry','') or event.get('expiry_date','')

    org_name = event.get('org_name') or event.get('org_name_hi') or ''
    event_name = event.get('name') or event.get('event_name','')
    event_date = format_date_hi(event.get('start_date',''))
    signing_authority = event.get('signing_authority','') or ''
    signing_title = event.get('signing_title','') or ''

    # Logo
    logo_data = ''
    logo_url = event.get('logo_url','')
    if logo_url:
        logo_data = fetch_image_as_dataurl(logo_url)

    # QR
    qr_data = make_qr_dataurl(qr_url or vol_id)

    html = TEMPLATE.render(
        css=CSS,
        org_name=org_name, event_name=event_name, event_date=event_date,
        name=name, vol_id=vol_id, mobile=mobile, aadhaar=aadhaar,
        role=role, permission=permission, expiry=expiry,
        logo_data=logo_data, qr_data=qr_data,
        signing_authority=signing_authority, signing_title=signing_title,
    )

    buf = io.BytesIO()
    HTML(string=html, base_url=os.path.dirname(__file__)).write_pdf(buf)
    return buf.getvalue()
