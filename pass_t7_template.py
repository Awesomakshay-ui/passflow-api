"""
T7 Pass Template — Corporate / NGO
Target: NGO events, conferences, corporate gatherings, seminars
Orientation: A6 Landscape (148mm × 105mm)
Visual: Dark navy + white, clean minimal, accent stripe, professional
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
        img = qr.make_image(fill_color='#0A1628', back_color='white')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}"
    except: return ''

CSS = """
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'NotoSansDevanagari'; src: url('{deva_b}') format('truetype'); font-weight:700; }}
@font-face {{ font-family:'Inter'; src: url('{inter_r}') format('truetype'); font-weight:400; }}
@font-face {{ font-family:'Inter'; src: url('{inter_b}') format('truetype'); font-weight:700; }}

@page {{ size: 148mm 105mm; margin: 0; }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}

body {{
  width: 148mm;
  height: 105mm;
  overflow: hidden;
  background: #fff;
  font-family: 'Inter', sans-serif;
  display: flex;
  flex-direction: row;
}}

/* Left sidebar — dark navy */
.sidebar {{
  width: 38mm;
  background: #0A1628;
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 5mm 3mm;
  gap: 3mm;
  flex-shrink: 0;
}}

.logo-wrap {{
  width: 20mm;
  height: 20mm;
  border-radius: 50%;
  background: rgba(255,255,255,0.1);
  border: 0.5mm solid rgba(255,255,255,0.2);
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
}}
.logo-wrap img {{ width:100%; height:100%; object-fit:cover; border-radius:50%; }}
.logo-ph {{ width:100%;height:100%;background:rgba(255,255,255,0.15);border-radius:50%; }}

.org-side {{
  font-size: 7pt;
  font-weight: 700;
  color: rgba(255,255,255,0.85);
  text-align: center;
  line-height: 1.4;
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
}}

.access-badge {{
  margin-top: auto;
  background: #1D4ED8;
  color: white;
  font-size: 6.5pt;
  font-weight: 700;
  padding: 1.5mm 3mm;
  border-radius: 1.5mm;
  text-align: center;
  letter-spacing: 0.5pt;
  width: 100%;
}}

.qr-side {{
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.5mm;
}}
.qr-img {{ width: 22mm; height: 22mm; border: 0.5mm solid rgba(255,255,255,0.2); }}
.qr-id {{ font-size: 5.5pt; color: rgba(255,255,255,0.5); text-align:center; }}

/* Main content */
.main {{
  flex: 1;
  display: flex;
  flex-direction: column;
}}

/* Top accent stripe */
.accent-stripe {{
  height: 2mm;
  background: linear-gradient(90deg, #1D4ED8, #3B82F6, #60A5FA);
}}

/* Event header */
.event-header {{
  padding: 3mm 5mm 2mm;
  border-bottom: 0.3mm solid #F0F0F0;
}}
.event-name {{
  font-size: 9.5pt;
  font-weight: 700;
  color: #0A1628;
  line-height: 1.25;
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
}}
.event-meta {{
  display: flex;
  gap: 3mm;
  margin-top: 1mm;
}}
.event-meta span {{
  font-size: 6.5pt;
  color: #888;
}}
.event-meta .dot {{ color: #ddd; }}

/* Name section */
.name-section {{
  padding: 2.5mm 5mm 2mm;
  border-bottom: 0.3mm solid #F0F0F0;
}}
.pass-type-tag {{
  font-size: 6pt;
  font-weight: 700;
  color: #1D4ED8;
  text-transform: uppercase;
  letter-spacing: 1.5pt;
  margin-bottom: 1mm;
}}
.holder-name {{
  font-size: 14pt;
  font-weight: 700;
  color: #0A1628;
  line-height: 1.1;
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
}}
.holder-role {{
  font-size: 8pt;
  color: #555;
  margin-top: 0.8mm;
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
}}

/* Fields */
.fields {{
  padding: 2mm 5mm;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0;
  flex: 1;
}}
.f-cell {{
  padding: 1.2mm 1mm;
  border-bottom: 0.2mm solid #F5F5F5;
}}
.f-cell.full {{ grid-column: 1/-1; }}
.f-label {{
  font-size: 5.5pt;
  font-weight: 700;
  color: #1D4ED8;
  text-transform: uppercase;
  letter-spacing: 0.5pt;
  margin-bottom: 0.3mm;
}}
.f-val {{
  font-size: 7.5pt;
  font-weight: 600;
  color: #1A1A1A;
  line-height: 1.3;
  font-family: 'NotoSansDevanagari','Inter',sans-serif;
}}

/* Footer */
.footer {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 1.5mm 5mm;
  background: #F8F9FA;
  border-top: 0.3mm solid #eee;
}}
.footer-sign {{
  font-size: 6pt;
  color: #555;
  line-height: 1.4;
}}
.footer-sign strong {{ color: #0A1628; font-weight:700; font-size:7pt; }}
.footer-meta {{
  font-size: 5.5pt;
  color: #aaa;
  text-align: right;
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

<div class="sidebar">
  <div class="logo-wrap">
    {% if logo_data %}<img src="{{ logo_data }}">{% else %}<div class="logo-ph"></div>{% endif %}
  </div>
  <div class="org-side">{{ org_name }}</div>
  <div class="qr-side">
    {% if qr_data %}<img class="qr-img" src="{{ qr_data }}">{% endif %}
    <div class="qr-id">{{ vol_id }}</div>
  </div>
  <div class="access-badge">{{ access_label }}</div>
</div>

<div class="main">
  <div class="accent-stripe"></div>
  <div class="event-header">
    <div class="event-name">{{ event_name }}</div>
    <div class="event-meta">
      {% if event_date %}<span>&#128197; {{ event_date }}</span>{% endif %}
    </div>
  </div>
  <div class="name-section">
    <div class="pass-type-tag">ENTRY PASS</div>
    <div class="holder-name">{{ name }}</div>
    {% if role %}<div class="holder-role">{{ role }}</div>{% endif %}
  </div>
  <div class="fields">
    <div class="f-cell">
      <div class="f-label">Pass ID</div>
      <div class="f-val">{{ vol_id }}</div>
    </div>
    {% if mobile %}
    <div class="f-cell">
      <div class="f-label">Mobile</div>
      <div class="f-val">{{ mobile }}</div>
    </div>
    {% endif %}
    {% if aadhaar %}
    <div class="f-cell">
      <div class="f-label">Aadhaar</div>
      <div class="f-val">{{ aadhaar }}</div>
    </div>
    {% endif %}
    {% if permission %}
    <div class="f-cell">
      <div class="f-label">Access Zone</div>
      <div class="f-val">{{ permission }}</div>
    </div>
    {% endif %}
    {% if expiry %}
    <div class="f-cell full">
      <div class="f-label">Valid Until</div>
      <div class="f-val">{{ expiry }}</div>
    </div>
    {% endif %}
  </div>
  <div class="footer">
    <div class="footer-sign">
      <strong>{{ signing_authority }}</strong><br>{{ signing_title }}
    </div>
    <div class="footer-meta">Powered by Pravesh<br>thepravesh.in</div>
  </div>
</div>

</body></html>
""")


def generate_pass_t7(volunteer: dict, event: dict, qr_url: str = '') -> bytes:
    name = volunteer.get('name') or volunteer.get('name_hi') or ''
    vol_id = volunteer.get('id','')
    mobile = str(volunteer.get('mobile','')).replace('+91','').strip()
    aadhaar = mask_aadhaar(volunteer.get('aadhaar',''))
    role = volunteer.get('role','') or ''
    permission = volunteer.get('permission','') or ''
    expiry = volunteer.get('expiry','') or event.get('expiry_date','')

    org_name = event.get('org_name','') or ''
    event_name = event.get('name','') or ''
    event_date = expiry
    signing_authority = event.get('signing_authority','') or ''
    signing_title = event.get('signing_title','') or ''
    access_label = permission or role or 'GENERAL'

    logo_data = ''
    if event.get('logo_url'):
        logo_data = fetch_image_as_dataurl(event['logo_url'])

    qr_data = make_qr_dataurl(qr_url or vol_id)

    html = TEMPLATE.render(
        css=CSS,
        org_name=org_name, event_name=event_name, event_date=event_date,
        name=name, vol_id=vol_id, mobile=mobile, aadhaar=aadhaar,
        role=role, permission=permission, expiry=expiry,
        logo_data=logo_data, qr_data=qr_data, access_label=access_label,
        signing_authority=signing_authority, signing_title=signing_title,
    )

    buf = io.BytesIO()
    HTML(string=html, base_url=os.path.dirname(__file__)).write_pdf(buf)
    return buf.getvalue()
