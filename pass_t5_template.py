"""
T5 Pass Template — Academic
Target: Universities, schools, convocations, seminars
Orientation: A6 Portrait (105mm × 148mm)
Visual: Maroon + gold, bilingual (Hindi + English), crest space, formal
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
    '01':'January','02':'February','03':'March','04':'April',
    '05':'May','06':'June','07':'July','08':'August',
    '09':'September','10':'October','11':'November','12':'December',
}

def format_date_en(date_str):
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
    if len(digits) >= 4: return 'XXXX XXXX ' + digits[-4:]
    return aadhaar

def fetch_image_as_dataurl(url):
    if not url: return ''
    try:
        import re as re2
        m = re2.search(r'drive[.]google[.]com/file/d/([a-zA-Z0-9_-]+)', url)
        if m: url = f"https://drive.google.com/uc?export=view&id={m.group(1)}"
        import requests as _req
        resp = _req.get(url, timeout=10, allow_redirects=True,
            headers={'User-Agent':'Mozilla/5.0','Accept':'image/*'})
        resp.raise_for_status()
        ct = resp.headers.get('Content-Type','image/png').split(';')[0].strip()
        if 'text/html' in ct: return ''
        return f"data:{ct};base64,{base64.b64encode(resp.content).decode()}"
    except: return ''

def make_qr_dataurl(data):
    if not _HAS_QRCODE or not data: return ''
    try:
        qr = _qrcode.QRCode(version=1, error_correction=_qrcode.constants.ERROR_CORRECT_M, box_size=6, border=2)
        qr.add_data(data); qr.make(fit=True)
        img = qr.make_image(fill_color='#6B0000', back_color='white')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"
    except: return ''

CSS = """
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_b}') format('truetype'); font-weight:700; }}
@font-face {{ font-family:'Inter'; src: url('{inter_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'Inter'; src: url('{inter_b}') format('truetype'); font-weight:700; }}
@font-face {{ font-family:'Cormorant'; src: url('{corm_r}') format('truetype'); font-weight:600; }}

@page {{
  size: 105mm 148mm;
  margin: 0;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  width: 105mm;
  height: 148mm;
  overflow: hidden;
  font-family: 'Inter', sans-serif;
  background: #FDFAF5;
  position: relative;
}}

/* Gold top bar */
.gold-bar {{
  width: 100%;
  height: 3mm;
  background: linear-gradient(90deg, #8B0000, #C8A04A, #8B0000);
}}

/* Header */
.header {{
  background: #6B0000;
  padding: 4mm 5mm 3mm;
  text-align: center;
  position: relative;
}}
.header-logo-wrap {{
  display: flex;
  justify-content: center;
  margin-bottom: 2mm;
}}
.header-logo {{
  width: 18mm;
  height: 18mm;
  border-radius: 50%;
  background: #fff;
  border: 1mm solid #C8A04A;
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
}}
.header-logo img {{ width: 100%; height: 100%; object-fit: cover; }}
.header-logo-ph {{
  width: 100%; height: 100%;
  background: linear-gradient(135deg, #8B0000, #4A0000);
  border-radius: 50%;
}}

.org-name {{
  font-family: 'Inter', sans-serif;
  font-size: 9pt;
  font-weight: 700;
  color: #fff;
  line-height: 1.25;
  letter-spacing: 0.3pt;
}}
.org-name-hi {{
  font-family: 'NotoSansDevanagari', sans-serif;
  font-size: 8pt;
  color: #C8A04A;
  margin-top: 1mm;
  line-height: 1.2;
}}

/* Gold divider */
.gold-divider {{
  height: 0.4mm;
  background: linear-gradient(90deg, transparent, #C8A04A, transparent);
  margin: 2mm 0;
}}

/* Pass type banner */
.pass-banner {{
  background: #C8A04A;
  text-align: center;
  padding: 1.5mm 0;
  font-family: 'Inter', sans-serif;
  font-size: 8pt;
  font-weight: 700;
  color: #4A0000;
  letter-spacing: 2pt;
  text-transform: uppercase;
}}

/* Event info */
.event-block {{
  padding: 3mm 5mm 2mm;
  text-align: center;
  border-bottom: 0.3mm solid #E8D8B0;
}}
.event-name-en {{
  font-family: 'Inter', sans-serif;
  font-size: 9pt;
  font-weight: 700;
  color: #4A0000;
  letter-spacing: 0.2pt;
  line-height: 1.3;
}}
.event-name-hi {{
  font-family: 'NotoSansDevanagari', sans-serif;
  font-size: 8.5pt;
  color: #6B0000;
  margin-top: 1mm;
  line-height: 1.2;
}}
.event-date {{
  font-size: 7.5pt;
  color: #888;
  margin-top: 1.5mm;
}}

/* Fields */
.fields-block {{
  padding: 3mm 5mm;
  flex: 1;
}}
.field-row {{
  display: flex;
  padding: 1.5mm 0;
  border-bottom: 0.2mm solid #F0E8D8;
  align-items: baseline;
  gap: 2mm;
}}
.field-row:last-child {{ border-bottom: none; }}
.field-label {{
  font-size: 6pt;
  font-weight: 700;
  color: #8B0000;
  text-transform: uppercase;
  letter-spacing: 0.5pt;
  min-width: 22mm;
  flex-shrink: 0;
}}
.field-value {{
  font-size: 8.5pt;
  font-weight: 700;
  color: #1A1A1A;
  font-family: 'NotoSansDevanagari', 'Inter', sans-serif;
  flex: 1;
  line-height: 1.3;
}}
.field-value.en {{
  font-family: 'Inter', sans-serif;
}}
.field-value.name-field {{
  font-size: 11pt;
  color: #4A0000;
}}

/* QR + sign */
.bottom-block {{
  display: flex;
  align-items: center;
  padding: 2mm 5mm 2mm;
  border-top: 0.3mm solid #E8D8B0;
  gap: 3mm;
}}
.qr-wrap {{
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.5mm;
}}
.qr-img {{ width: 22mm; height: 22mm; border: 0.5mm solid #C8A04A; }}
.qr-id {{ font-size: 5.5pt; color: #888; text-align: center; }}
.sign-wrap {{
  flex: 1;
  text-align: center;
}}
.sign-line {{
  width: 100%;
  height: 0.3mm;
  background: #8B0000;
  margin-bottom: 1mm;
}}
.sign-name {{
  font-family: 'Inter', sans-serif;
  font-size: 6.5pt;
  font-weight: 700;
  color: #4A0000;
  line-height: 1.3;
}}
.sign-title {{
  font-size: 6pt;
  color: #888;
}}

/* Footer gold bar */
.gold-bar-bottom {{
  width: 100%;
  height: 3mm;
  background: linear-gradient(90deg, #8B0000, #C8A04A, #8B0000);
  position: absolute;
  bottom: 0;
}}
""".format(
    deva_r=_font_url('NotoSansDevanagari-Regular.ttf'),
    deva_b=_font_url('NotoSansDevanagari-Bold.ttf'),
    inter_r=_font_url('Inter-Regular.ttf'),
    inter_b=_font_url('Inter-Bold.ttf'),
    corm_r=_font_url('Inter-Bold.ttf'),
)

TEMPLATE = Template("""
<!DOCTYPE html><html><head><meta charset="UTF-8"><style>{{ css }}</style></head>
<body>

<div class="gold-bar"></div>
<div class="header">
  <div class="header-logo-wrap">
    <div class="header-logo">
      {% if logo_data %}<img src="{{ logo_data }}">{% else %}<div class="header-logo-ph"></div>{% endif %}
    </div>
  </div>
  <div class="org-name">{{ org_name }}</div>
  {% if org_name_hi %}<div class="org-name-hi">{{ org_name_hi }}</div>{% endif %}
</div>

<div class="pass-banner">ENTRY PASS &nbsp;|&nbsp; प्रवेश पत्र</div>

<div class="event-block">
  <div class="event-name-en">{{ event_name }}</div>
  {% if event_date %}<div class="event-date">📅 {{ event_date }}</div>{% endif %}
</div>

<div class="fields-block">
  <div class="field-row">
    <div class="field-label">Pass ID</div>
    <div class="field-value en">{{ vol_id }}</div>
  </div>
  <div class="field-row">
    <div class="field-label">Name / नाम</div>
    <div class="field-value name-field">{{ name }}</div>
  </div>
  {% if role %}
  <div class="field-row">
    <div class="field-label">Role / दायित्व</div>
    <div class="field-value">{{ role }}</div>
  </div>
  {% endif %}
  {% if mobile %}
  <div class="field-row">
    <div class="field-label">Mobile</div>
    <div class="field-value en">{{ mobile }}</div>
  </div>
  {% endif %}
  {% if aadhaar %}
  <div class="field-row">
    <div class="field-label">Aadhaar</div>
    <div class="field-value en">{{ aadhaar }}</div>
  </div>
  {% endif %}
  {% if permission %}
  <div class="field-row">
    <div class="field-label">Access / अनुमति</div>
    <div class="field-value">{{ permission }}</div>
  </div>
  {% endif %}
  {% if expiry %}
  <div class="field-row">
    <div class="field-label">Valid Until</div>
    <div class="field-value en">{{ expiry }}</div>
  </div>
  {% endif %}
</div>

<div class="bottom-block">
  <div class="qr-wrap">
    {% if qr_data %}<img class="qr-img" src="{{ qr_data }}">{% endif %}
    <div class="qr-id">{{ vol_id }}</div>
  </div>
  <div class="sign-wrap">
    <div class="sign-line"></div>
    <div class="sign-name">{{ signing_authority }}</div>
    <div class="sign-title">{{ signing_title }}</div>
  </div>
</div>

<div class="gold-bar-bottom"></div>
</body></html>
""")


def generate_pass_t5(volunteer: dict, event: dict, qr_url: str = '') -> bytes:
    name = volunteer.get('name') or volunteer.get('name_hi') or ''
    vol_id = volunteer.get('id','')
    mobile = str(volunteer.get('mobile','')).replace('+91','').strip()
    aadhaar = mask_aadhaar(volunteer.get('aadhaar',''))
    role = volunteer.get('role','') or ''
    permission = volunteer.get('permission','') or ''
    expiry = volunteer.get('expiry','') or event.get('expiry_date','')

    org_name = event.get('org_name','') or ''
    org_name_hi = event.get('org_name_hi','') or ''
    event_name = event.get('name','') or ''
    event_date = format_date_en(event.get('start_date','') or expiry)
    signing_authority = event.get('signing_authority','') or ''
    signing_title = event.get('signing_title','') or ''

    logo_data = ''
    if event.get('logo_url'):
        logo_data = fetch_image_as_dataurl(event['logo_url'])

    qr_data = make_qr_dataurl(qr_url or vol_id)

    html = TEMPLATE.render(
        css=CSS,
        org_name=org_name, org_name_hi=org_name_hi,
        event_name=event_name, event_date=event_date,
        name=name, vol_id=vol_id, mobile=mobile, aadhaar=aadhaar,
        role=role, permission=permission, expiry=expiry,
        logo_data=logo_data, qr_data=qr_data,
        signing_authority=signing_authority, signing_title=signing_title,
    )

    buf = io.BytesIO()
    HTML(string=html, base_url=os.path.dirname(__file__)).write_pdf(buf)
    return buf.getvalue()
