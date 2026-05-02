"""
T8 Pass Template — VIP Premium
Target: VVIP, donors, platform guests, senior officials
Orientation: A6 Portrait (105mm × 148mm)
Visual: Black + gold, premium minimal, large name, foil-style border
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

def format_date_hi(d):
    if not d: return ''
    for sep in ['-','/']:
        p = d.split(sep)
        if len(p)==3 and p[1] in MONTHS_HI:
            return f"{int(p[0])} {MONTHS_HI[p[1]]} {p[2]}"
    return d

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
        img = qr.make_image(fill_color='#C8A04A', back_color='#0A0A0A')
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
  background: #0A0A0A;
  font-family: 'Inter', sans-serif;
  color: white;
}}

/* Gold outer border effect */
.gold-border {{
  position: absolute;
  inset: 1.5mm;
  border: 0.6mm solid #C8A04A;
  pointer-events: none;
  z-index: 10;
}}
.gold-border-inner {{
  position: absolute;
  inset: 3mm;
  border: 0.2mm solid rgba(200,160,74,0.3);
  pointer-events: none;
  z-index: 10;
}}

/* Top gold shimmer bar */
.gold-shimmer {{
  height: 1.5mm;
  background: linear-gradient(90deg, #8B6914, #C8A04A, #FFD700, #C8A04A, #8B6914);
}}

/* Header */
.header {{
  padding: 5mm 6mm 4mm;
  text-align: center;
  border-bottom: 0.3mm solid rgba(200,160,74,0.25);
}}

.vip-tag {{
  display: inline-block;
  background: linear-gradient(135deg, #C8A04A, #FFD700);
  color: #0A0A0A;
  font-size: 7pt;
  font-weight: 700;
  letter-spacing: 3pt;
  text-transform: uppercase;
  padding: 1.5mm 6mm;
  border-radius: 1mm;
  margin-bottom: 3mm;
}}

.logo-wrap {{
  display: flex;
  justify-content: center;
  margin-bottom: 2.5mm;
}}
.logo-circle {{
  width: 18mm;
  height: 18mm;
  border-radius: 50%;
  border: 0.8mm solid #C8A04A;
  background: rgba(255,255,255,0.05);
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
}}
.logo-circle img {{ width:100%;height:100%;object-fit:cover;border-radius:50%; }}
.logo-ph {{ width:100%;height:100%;background:rgba(200,160,74,0.15);border-radius:50%; }}

.org-name {{
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
  font-size: 9.5pt;
  font-weight: 700;
  color: #fff;
  line-height: 1.25;
}}
.event-name {{
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
  font-size: 7.5pt;
  color: #C8A04A;
  margin-top: 1.5mm;
  line-height: 1.3;
}}

/* Gold divider */
.gold-divider {{
  height: 0.3mm;
  background: linear-gradient(90deg, transparent, #C8A04A, transparent);
  margin: 0;
}}

/* Name block — the hero */
.name-block {{
  padding: 5mm 6mm 4mm;
  text-align: center;
  border-bottom: 0.3mm solid rgba(200,160,74,0.2);
}}
.name-label {{
  font-size: 6pt;
  font-weight: 700;
  color: rgba(200,160,74,0.7);
  text-transform: uppercase;
  letter-spacing: 2pt;
  margin-bottom: 2mm;
}}
.holder-name {{
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
  font-size: 16pt;
  font-weight: 700;
  color: #fff;
  line-height: 1.15;
}}
.holder-role {{
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
  font-size: 8pt;
  color: #C8A04A;
  margin-top: 1.5mm;
  letter-spacing: 0.5pt;
}}
.holder-id {{
  font-size: 6.5pt;
  color: rgba(255,255,255,0.4);
  margin-top: 1mm;
  letter-spacing: 1pt;
}}

/* Fields — minimal on dark */
.fields {{
  padding: 3mm 7mm;
  display: flex;
  flex-direction: column;
  gap: 0;
}}
.f-row {{
  display: flex;
  align-items: baseline;
  padding: 1.5mm 0;
  border-bottom: 0.2mm solid rgba(255,255,255,0.06);
  gap: 2mm;
}}
.f-row:last-child {{ border-bottom: none; }}
.f-label {{
  font-size: 5.5pt;
  font-weight: 700;
  color: rgba(200,160,74,0.6);
  text-transform: uppercase;
  letter-spacing: 0.8pt;
  min-width: 20mm;
  flex-shrink: 0;
}}
.f-val {{
  font-size: 8pt;
  font-weight: 600;
  color: rgba(255,255,255,0.9);
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
  flex: 1;
}}

/* Bottom: QR + sign */
.bottom {{
  display: flex;
  align-items: center;
  padding: 2mm 6mm 4mm;
  border-top: 0.3mm solid rgba(200,160,74,0.2);
  gap: 3mm;
  margin-top: auto;
}}
.qr-wrap {{
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 1mm;
}}
.qr-img {{ width: 20mm; height: 20mm; border: 0.5mm solid rgba(200,160,74,0.4); }}
.qr-id {{ font-size:5.5pt; color:rgba(255,255,255,0.35); }}
.sign-wrap {{
  flex: 1;
  text-align: center;
}}
.sign-line {{ height:0.3mm;background:rgba(200,160,74,0.5);width:80%;margin:0 auto 1.5mm; }}
.sign-name {{ font-size:7pt;font-weight:700;color:#C8A04A; }}
.sign-title {{ font-size:6pt;color:rgba(255,255,255,0.45); }}

/* Bottom gold bar */
.gold-shimmer-bottom {{
  position: absolute;
  bottom: 0;
  left: 0;
  right: 0;
  height: 1.5mm;
  background: linear-gradient(90deg, #8B6914, #C8A04A, #FFD700, #C8A04A, #8B6914);
}}
""".format(
    deva_r=_font_url('NotoSansDevanagari-Regular.ttf'),
    deva_b=_font_url('NotoSansDevanagari-Bold.ttf'),
    inter_r=_font_url('Inter-Regular.ttf'),
    inter_b=_font_url('Inter-Bold.ttf'),
)

TEMPLATE = Template("""
<!DOCTYPE html><html><head><meta charset="UTF-8"><style>{{ css }}</style></head>
<body>
<div class="gold-border"></div>
<div class="gold-border-inner"></div>
<div class="gold-shimmer"></div>

<div class="header">
  <div class="vip-tag">VIP PASS</div>
  <div class="logo-wrap">
    <div class="logo-circle">
      {% if logo_data %}<img src="{{ logo_data }}">{% else %}<div class="logo-ph"></div>{% endif %}
    </div>
  </div>
  <div class="org-name">{{ org_name }}</div>
  <div class="event-name">{{ event_name }}</div>
</div>

<div class="gold-divider"></div>

<div class="name-block">
  <div class="name-label">Pass Holder</div>
  <div class="holder-name">{{ name }}</div>
  {% if role %}<div class="holder-role">{{ role }}</div>{% endif %}
  <div class="holder-id">ID: {{ vol_id }}</div>
</div>

<div class="fields">
  {% if mobile %}
  <div class="f-row">
    <div class="f-label">Mobile</div>
    <div class="f-val">{{ mobile }}</div>
  </div>
  {% endif %}
  {% if aadhaar %}
  <div class="f-row">
    <div class="f-label">Aadhaar</div>
    <div class="f-val">{{ aadhaar }}</div>
  </div>
  {% endif %}
  {% if permission %}
  <div class="f-row">
    <div class="f-label">Access</div>
    <div class="f-val">{{ permission }}</div>
  </div>
  {% endif %}
  {% if expiry %}
  <div class="f-row">
    <div class="f-label">Valid Until</div>
    <div class="f-val">{{ expiry_hi }}</div>
  </div>
  {% endif %}
</div>

<div class="bottom">
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

<div class="gold-shimmer-bottom"></div>
</body></html>
""")


def generate_pass_t8(volunteer: dict, event: dict, qr_url: str = '') -> bytes:
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
    signing_authority = event.get('signing_authority','') or ''
    signing_title = event.get('signing_title','') or ''

    logo_data = ''
    if event.get('logo_url'):
        logo_data = fetch_image_as_dataurl(event['logo_url'])

    qr_data = make_qr_dataurl(qr_url or vol_id)

    html = TEMPLATE.render(
        css=CSS,
        org_name=org_name, event_name=event_name,
        name=name, vol_id=vol_id, mobile=mobile, aadhaar=aadhaar,
        role=role, permission=permission, expiry=expiry, expiry_hi=expiry_hi,
        logo_data=logo_data, qr_data=qr_data,
        signing_authority=signing_authority, signing_title=signing_title,
    )

    buf = io.BytesIO()
    HTML(string=html, base_url=os.path.dirname(__file__)).write_pdf(buf)
    return buf.getvalue()
