"""
T6 Pass Template — Mahotsav / Festival
Target: Cultural events, melas, temple festivals, religious gatherings
Orientation: A6 Portrait (105mm × 148mm)
Visual: Warm red-gold, ornate border, festive feel, Devanagari dominant
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

def mask_aadhaar(a):
    if not a: return ''
    d = re.sub(r'\D','',str(a))
    return 'XXXX XXXX ' + d[-4:] if len(d) >= 4 else a

def fetch_image_as_dataurl(url):
    if not url: return ''
    try:
        m = re.search(r'drive[.]google[.]com/file/d/([a-zA-Z0-9_-]+)', url)
        if m: url = f"https://drive.google.com/uc?export=view&id={m.group(1)}"
        import requests as _req
        r = _req.get(url, timeout=10, allow_redirects=True, headers={'User-Agent':'Mozilla/5.0','Accept':'image/*'})
        r.raise_for_status()
        ct = r.headers.get('Content-Type','image/png').split(';')[0].strip()
        if 'text/html' in ct: return ''
        return f"data:{ct};base64,{base64.b64encode(r.content).decode()}"
    except: return ''

def make_qr_dataurl(data):
    if not _HAS_QRCODE or not data: return ''
    try:
        qr = _qrcode.QRCode(version=1, error_correction=_qrcode.constants.ERROR_CORRECT_M, box_size=6, border=2)
        qr.add_data(data); qr.make(fit=True)
        img = qr.make_image(fill_color='#8B2800', back_color='#FFFDF0')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"
    except: return ''

CSS = """
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_b}') format('truetype'); font-weight:700; }}
@font-face {{ font-family:'Inter'; src: url('{inter_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'Inter'; src: url('{inter_b}') format('truetype'); font-weight:700; }}

@page {{ size: 105mm 148mm; margin: 0; }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}

body {{
  width: 105mm;
  height: 148mm;
  overflow: hidden;
  background: #FFFDF0;
  font-family: 'NotoSansDevanagari', sans-serif;
}}

/* Decorative outer border */
.outer-border {{
  position: absolute;
  inset: 1.5mm;
  border: 0.8mm solid #C8A04A;
  pointer-events: none;
  z-index: 10;
}}
.inner-border {{
  position: absolute;
  inset: 2.5mm;
  border: 0.3mm solid #E8C870;
  pointer-events: none;
  z-index: 10;
}}

/* Corner ornaments */
.corner {{
  position: absolute;
  width: 6mm;
  height: 6mm;
  z-index: 11;
}}
.corner-tl {{ top: 1mm; left: 1mm; border-top: 1.2mm solid #C8102E; border-left: 1.2mm solid #C8102E; }}
.corner-tr {{ top: 1mm; right: 1mm; border-top: 1.2mm solid #C8102E; border-right: 1.2mm solid #C8102E; }}
.corner-bl {{ bottom: 1mm; left: 1mm; border-bottom: 1.2mm solid #C8102E; border-left: 1.2mm solid #C8102E; }}
.corner-br {{ bottom: 1mm; right: 1mm; border-bottom: 1.2mm solid #C8102E; border-right: 1.2mm solid #C8102E; }}

/* Header */
.header {{
  padding: 5mm 6mm 3mm;
  text-align: center;
  background: linear-gradient(180deg, #FFF8E8 0%, #FFFDF0 100%);
  border-bottom: 0.5mm solid #E8C870;
}}

.logo-wrap {{
  display: flex;
  justify-content: center;
  margin-bottom: 2mm;
}}
.logo-circle {{
  width: 20mm;
  height: 20mm;
  border-radius: 50%;
  border: 1mm solid #C8A04A;
  background: white;
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 0 4mm rgba(200,160,74,0.3);
}}
.logo-circle img {{ width:100%; height:100%; object-fit:cover; }}
.logo-ph {{ width:100%;height:100%;background:linear-gradient(135deg,#C8102E,#8B0000);border-radius:50%; }}

.org-name {{
  font-size: 10.5pt;
  font-weight: 700;
  color: #4A0000;
  line-height: 1.25;
  letter-spacing: 0.3pt;
}}
.event-name {{
  font-size: 9pt;
  font-weight: 700;
  color: #C8102E;
  margin-top: 1.5mm;
  line-height: 1.3;
}}
.event-date {{
  font-family: 'Inter', sans-serif;
  font-size: 7pt;
  color: #888;
  margin-top: 1.5mm;
}}

/* Decorative divider */
.deco-divider {{
  text-align: center;
  font-size: 10pt;
  color: #C8A04A;
  letter-spacing: 3pt;
  padding: 1mm 0;
  background: linear-gradient(90deg, transparent, #FFF8E8, transparent);
}}

/* Pass label */
.pass-label {{
  text-align: center;
  background: #C8102E;
  color: white;
  font-size: 8pt;
  font-weight: 700;
  padding: 1.5mm 0;
  letter-spacing: 2pt;
}}

/* Name block — prominent */
.name-block {{
  text-align: center;
  padding: 3mm 6mm 2mm;
  border-bottom: 0.3mm dashed #E8C870;
}}
.name-main {{
  font-size: 14pt;
  font-weight: 700;
  color: #1A1A1A;
  line-height: 1.2;
}}
.name-id {{
  font-family: 'Inter', sans-serif;
  font-size: 7pt;
  color: #888;
  margin-top: 1mm;
  letter-spacing: 0.5pt;
}}

/* Fields in 2 columns */
.fields-grid {{
  padding: 2mm 6mm;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0;
}}
.field-cell {{
  padding: 1.5mm 1mm;
  border-bottom: 0.2mm solid #F0E8D0;
}}
.field-cell.full {{ grid-column: 1 / -1; }}
.f-label {{
  font-family: 'Inter', sans-serif;
  font-size: 6pt;
  color: #C8102E;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.5pt;
  margin-bottom: 0.5mm;
}}
.f-val {{
  font-size: 8pt;
  font-weight: 700;
  color: #1A1A1A;
  line-height: 1.3;
}}
.f-val.en {{ font-family: 'Inter', sans-serif; }}

/* Bottom: QR + sign */
.bottom {{
  display: flex;
  align-items: center;
  padding: 2mm 6mm 4mm;
  gap: 3mm;
  margin-top: auto;
}}
.qr-col {{
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.5mm;
}}
.qr-img {{ width: 20mm; height: 20mm; border: 0.5mm solid #C8A04A; }}
.qr-id {{ font-family:'Inter',sans-serif; font-size:5.5pt; color:#888; }}
.sign-col {{
  flex: 1;
  text-align: center;
}}
.sign-line {{ height:0.3mm; background:#C8A04A; width:80%; margin:0 auto 1mm; }}
.sign-name {{ font-size:7pt; font-weight:700; color:#4A0000; }}
.sign-title {{ font-size:6pt; color:#888; }}
""".format(
    deva_r=_font_url('NotoSansDevanagari-Regular.ttf'),
    deva_b=_font_url('NotoSansDevanagari-Bold.ttf'),
    inter_r=_font_url('Inter-Regular.ttf'),
    inter_b=_font_url('Inter-Bold.ttf'),
)

TEMPLATE = Template("""
<!DOCTYPE html><html><head><meta charset="UTF-8"><style>{{ css }}</style></head>
<body>
<div class="outer-border"></div>
<div class="inner-border"></div>
<div class="corner corner-tl"></div>
<div class="corner corner-tr"></div>
<div class="corner corner-bl"></div>
<div class="corner corner-br"></div>

<div class="header">
  <div class="logo-wrap">
    <div class="logo-circle">
      {% if logo_data %}<img src="{{ logo_data }}">{% else %}<div class="logo-ph"></div>{% endif %}
    </div>
  </div>
  <div class="org-name">{{ org_name }}</div>
  <div class="event-name">{{ event_name }}</div>
  {% if event_date %}<div class="event-date">&#128197; {{ event_date }}</div>{% endif %}
</div>

<div class="deco-divider">&#x2605; &#x2605; &#x2605;</div>
<div class="pass-label">&#x0915;&#x093E;&#x0930;&#x094D;&#x092F;&#x0915;&#x0930;&#x094D;&#x0924;&#x093E; &#x092A;&#x094D;&#x0930;&#x0935;&#x0947;&#x0936; &#x092A;&#x0924;&#x094D;&#x0930;</div>

<div class="name-block">
  <div class="name-main">{{ name }}</div>
  <div class="name-id">{{ vol_id }}</div>
</div>

<div class="fields-grid">
  {% if role %}
  <div class="field-cell">
    <div class="f-label">&#x0926;&#x093E;&#x092F;&#x093F;&#x0924;&#x094D;&#x0935;</div>
    <div class="f-val">{{ role }}</div>
  </div>
  {% endif %}
  {% if mobile %}
  <div class="field-cell">
    <div class="f-label">&#x092E;&#x094B;&#x092C;&#x093E;&#x0907;&#x0932;</div>
    <div class="f-val en">{{ mobile }}</div>
  </div>
  {% endif %}
  {% if aadhaar %}
  <div class="field-cell">
    <div class="f-label">&#x0906;&#x0927;&#x093E;&#x0930;</div>
    <div class="f-val en">{{ aadhaar }}</div>
  </div>
  {% endif %}
  {% if permission %}
  <div class="field-cell">
    <div class="f-label">&#x0905;&#x0928;&#x0941;&#x092E;&#x0924;&#x093F;</div>
    <div class="f-val">{{ permission }}</div>
  </div>
  {% endif %}
  {% if expiry %}
  <div class="field-cell full">
    <div class="f-label">&#x092E;&#x093E;&#x0928;&#x094D;&#x092F; &#x0924;&#x093E;&#x0930;&#x0940;&#x0916; &#x0924;&#x0915;</div>
    <div class="f-val">{{ expiry_hi }}</div>
  </div>
  {% endif %}
</div>

<div class="bottom">
  <div class="qr-col">
    {% if qr_data %}<img class="qr-img" src="{{ qr_data }}">{% endif %}
    <div class="qr-id">{{ vol_id }}</div>
  </div>
  <div class="sign-col">
    <div class="sign-line"></div>
    <div class="sign-name">{{ signing_authority }}</div>
    <div class="sign-title">{{ signing_title }}</div>
  </div>
</div>

</body></html>
""")


def generate_pass_t6(volunteer: dict, event: dict, qr_url: str = '') -> bytes:
    name = volunteer.get('name') or volunteer.get('name_hi') or ''
    vol_id = volunteer.get('id','')
    mobile = str(volunteer.get('mobile','')).replace('+91','').strip()
    aadhaar = mask_aadhaar(volunteer.get('aadhaar',''))
    role = volunteer.get('role','') or ''
    permission = volunteer.get('permission','') or ''
    expiry = volunteer.get('expiry','') or event.get('expiry_date','')
    expiry_hi = format_date_hi(expiry)

    org_name = event.get('org_name','') or ''
    event_name = event.get('name','') or ''
    event_date = format_date_hi(event.get('start_date','') or expiry)
    signing_authority = event.get('signing_authority','') or ''
    signing_title = event.get('signing_title','') or ''

    logo_data = ''
    if event.get('logo_url'):
        logo_data = fetch_image_as_dataurl(event['logo_url'])

    qr_data = make_qr_dataurl(qr_url or vol_id)

    html = TEMPLATE.render(
        css=CSS,
        org_name=org_name, event_name=event_name, event_date=event_date,
        name=name, vol_id=vol_id, mobile=mobile, aadhaar=aadhaar,
        role=role, permission=permission, expiry=expiry, expiry_hi=expiry_hi,
        logo_data=logo_data, qr_data=qr_data,
        signing_authority=signing_authority, signing_title=signing_title,
    )

    buf = io.BytesIO()
    HTML(string=html, base_url=os.path.dirname(__file__)).write_pdf(buf)
    return buf.getvalue()
