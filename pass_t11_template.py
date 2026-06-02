"""
T11 Pass Template — Saint Pass
Target: Spiritual organizations with saint's photo
Format: A5 Landscape (210mm × 148mm) — identical to T3
Visual: Exact T3 design + saint photo top-left header + volunteer photo above QR
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

PASS_TYPE_CONFIG = {
    'karyakarta':    {'color': '#0f52ba', 'label': 'कार्यकर्ता पास',   'accent': '#FFD700', 'validity': 'तक मान्य'},
    'vishesh_atithi':{'color': '#8B0000', 'label': 'विशेष अतिथि पास', 'accent': '#FFD700', 'validity': 'को मान्य'},
    'vip':           {'color': '#0A0A0A', 'label': 'VIP पास',          'accent': '#C8A04A', 'validity': 'तक मान्य'},
    'press':         {'color': '#1A5C2A', 'label': 'प्रेस पास',        'accent': '#FFFFFF', 'validity': 'को मान्य'},
    'seva':          {'color': '#0F5C4A', 'label': 'सेवा पास',         'accent': '#FFD700', 'validity': 'तक मान्य'},
    'staff':         {'color': '#1A2C5C', 'label': 'स्टाफ पास',        'accent': '#FFFFFF', 'validity': 'तक मान्य'},
}

def get_pass_type_style(pass_type):
    return PASS_TYPE_CONFIG.get(pass_type or 'karyakarta', PASS_TYPE_CONFIG['karyakarta'])

def format_date_hi(d):
    if not d: return ''
    for sep in ['-', '/']:
        p = d.split(sep)
        if len(p) == 3 and p[1] in MONTHS_HI:
            return f"{int(p[0])} {MONTHS_HI[p[1]]} {p[2]}"
    return d

def mask_aadhaar(a):
    if not a: return ''
    d = re.sub(r'\D', '', str(a))
    return 'XXXX XXXX ' + d[-4:] if len(d) >= 4 else a

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
            headers={'User-Agent': 'Mozilla/5.0', 'Accept': 'image/*'})
        resp.raise_for_status()
        ct = resp.headers.get('Content-Type', 'image/png').split(';')[0].strip()
        if 'text/html' in ct: return ''
        return f"data:{ct};base64,{base64.b64encode(resp.content).decode()}"
    except: return ''

def logo_file_as_dataurl(filename):
    try:
        path = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(path): return ''
        with open(path, 'rb') as f: data = f.read()
        ext = filename.rsplit('.', 1)[-1].lower()
        ct = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg'}.get(ext, 'image/png')
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


T11_CSS = """
@page {{
  size: 210mm 148mm;
  margin: 0;
}}

@font-face {{
  font-family: 'NotoDeva';
  src: url('{deva_b}') format('truetype');
  font-weight: bold;
}}
@font-face {{
  font-family: 'NotoDeva';
  src: url('{deva_r}') format('truetype');
  font-weight: normal;
}}
@font-face {{
  font-family: 'Poppins';
  src: url('{poppins_b}') format('truetype');
  font-weight: bold;
}}
@font-face {{
  font-family: 'Poppins';
  src: url('{poppins_r}') format('truetype');
  font-weight: normal;
}}

* {{
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}}

body {{
  width: 210mm;
  height: 148mm;
  position: relative;
  overflow: hidden;
  background: #F4F0E5;
  font-family: 'Poppins', sans-serif;
}}

/* ── Header ── */
.header {{
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 50mm;
  background: linear-gradient(180deg, #1a6fd4 0%, #0f52ba 100%);
  border-radius: 0 0 50% 50% / 0 0 20% 20%;
  z-index: 1;
  padding: 8mm 10mm 0 10mm;
  text-align: center;
  display: flex;
  flex-direction: column;
  justify-content: center;
}}

.header::before {{
  content: '';
  position: absolute;
  inset: 0;
  background-image: repeating-linear-gradient(
    90deg,
    rgba(255,255,255,0.04) 0px,
    rgba(255,255,255,0.04) 1px,
    transparent 1px,
    transparent 8mm
  );
  pointer-events: none;
  border-radius: inherit;
}}

/* Saint photo — scaled up to be larger than logo */
.saint-photo-wrap {{
  position: absolute;
  left: 8mm;
  top: 50%;
  transform: translateY(-50%);
  z-index: 2;
}}
.saint-photo-circle {{
  width: 45mm;
  height: 45mm;
  border-radius: 50%;
  border: 1mm solid #FFDD96;
  background: #fff;
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 0 0 1.5mm rgba(200,160,74,0.3);
}}
.saint-photo-circle img {{
  width: 100%;
  height: 100%;
  object-fit: cover;
  border-radius: 50%;
}}
.saint-placeholder {{
  font-family: 'Poppins', sans-serif;
  font-size: 14pt;
  font-weight: bold;
  color: #0F52BA;
  text-align: center;
}}

.org-name {{
  font-family: 'NotoDeva', serif;
  font-size: 26pt;
  font-weight: 700;
  color: white;
  line-height: 1.0;
  margin-top: 0mm;
}}

.event-name {{
  font-family: 'NotoDeva', serif;
  font-size: 15pt;
  font-weight: 700;
  color: #FFD700;
  margin-top: 4mm;
  display: inline-block;
  padding: 0.5mm 3mm 1mm 3mm;
  border-bottom: 0.4mm solid #C8A04A;
}}

.event-date {{
  font-family: 'NotoDeva', serif;
  font-size: 12pt;
  font-weight: 700;
  color: white;
  margin-top: 4mm;
  display: inline-block;
  padding: 0.3mm 4mm 0.8mm 4mm;
  border-bottom: 0.3mm solid #C8A04A;
}}

/* ── Body ── */
.body-table {{
  position: absolute;
  top: 50mm;
  left: 0;
  bottom: 28mm;
  width: 100%;
  padding: 0 6mm;
  display: table;
}}

/* Logo cell — scaled down down so it is smaller than saint photo */
.logo-cell {{
  display: table-cell;
  width: 36mm;
  vertical-align: middle;
  text-align: center;
  padding: 2mm 4mm 2mm 2mm;
}}
.logo-circle {{
  width: 30mm;
  height: 30mm;
  border-radius: 50%;
  background: #F4F0E5;
  border: 0.5mm solid #ccc;
  overflow: hidden;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}}
.logo-circle img {{
  width: 100%;
  height: 100%;
  object-fit: cover;
  border-radius: 50%;
}}

/* Fields cell */
.fields-cell {{
  display: table-cell;
  vertical-align: middle;
  padding: 2mm 4mm;
}}
.fields-table {{
  border-collapse: collapse;
  width: 100%;
}}
.fields-table tr td {{
  padding: 2mm 1mm;
  vertical-align: middle;
}}
.fl {{
  font-family: 'Poppins', sans-serif;
  font-size: 11pt;
  font-weight: 600;
  color: #404040;
  white-space: nowrap;
  padding-right: 2mm;
  width: 34mm;
}}
.fc {{
  font-size: 12pt;
  font-weight: 700;
  color: #606060;
  padding: 0 1mm;
  width: 4mm;
}}
.fv {{
  font-family: 'Poppins', sans-serif;
  font-size: 12pt;
  font-weight: 700;
  color: #1A1A1A;
}}

/* QR cell — extended for volunteer photo */
.qr-cell {{
  width: 40mm;
  vertical-align: middle;
  text-align: center;
  padding: 2mm;
  display: table-cell;
}}
.volunteer-photo-circle {{
  width: 28mm;
  height: 28mm;
  border-radius: 50%;
  border: 0.5mm solid #C8A04A;
  background: #F4F0E5;
  overflow: hidden;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-bottom: 2mm;
}}
.volunteer-photo-circle img {{
  width: 100%;
  height: 100%;
  object-fit: cover;
  border-radius: 50%;
}}
.vol-photo-placeholder {{
  font-family: 'Poppins', sans-serif;
  font-size: 12pt;
  font-weight: bold;
  color: #888;
}}
.qr-img {{
  width: 30mm;
  height: 30mm;
}}
.scan-label {{
  font-family: 'Poppins', sans-serif;
  font-size: 6pt;
  color: #888;
  letter-spacing: 1.5pt;
  margin-top: 1mm;
}}

/* ── Footer ── */
.footer {{
  position: absolute;
  left: 0;
  right: 0;
  bottom: 0;
  height: 28mm;
  background: linear-gradient(180deg, #1a6fd4 0%, #0f52ba 100%);
  z-index: 2;
  padding: 4mm 10mm 3mm 10mm;
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
}}

.notes {{
  display: flex;
  flex-direction: column;
  gap: 1.5mm;
}}
.note {{
  font-family: 'NotoDeva', sans-serif;
  font-size: 9pt;
  color: white;
  line-height: 1.5;
}}

.authority {{
  text-align: right;
}}
.sign-issuing {{
  font-family: 'Poppins', sans-serif;
  font-size: 8pt;
  font-weight: normal;
  color: rgba(255,255,255,0.8);
  margin-bottom: 1mm;
}}
.sign-image {{
  max-height: 10mm;
  max-width: 40mm;
  display: block;
  margin-left: auto;
  margin-bottom: 1mm;
}}
.sign-name {{
  font-family: 'Poppins', sans-serif;
  font-size: 10pt;
  font-weight: bold;
  color: #fff;
}}
.sign-title {{
  font-family: 'NotoDeva', sans-serif;
  font-size: 10pt;
  font-weight: 700;
  color: #FFDD96;
  margin-top: 0.5mm;
  letter-spacing: 0.3pt;
}}
""".format(
    deva_r=_font_url('NotoSansDevanagari-Regular.ttf'),
    deva_b=_font_url('NotoSansDevanagari-Bold.ttf'),
    poppins_r=_font_url('Poppins-Regular.ttf'),
    poppins_b=_font_url('Poppins-Bold.ttf'),
)

PASS_DIV = Template("""
<div style="position:relative;width:210mm;height:148mm;overflow:hidden;background:#F4F0E5;page-break-after:always;">
<style>{{ css }}</style>

<div class="header" style="background: linear-gradient(180deg, {{ header_color_light }} 0%, {{ header_color }} 100%) !important;">

  {%- if saint_photo %}
  <div class="saint-photo-wrap">
    <div class="saint-photo-circle"><img src="{{ saint_photo }}"></div>
  </div>
  {%- elif saint_initials %}
  <div class="saint-photo-wrap">
    <div class="saint-photo-circle"><div class="saint-placeholder">{{ saint_initials }}</div></div>
  </div>
  {%- endif %}

  <div class="org-name">{{ org }}</div>
  {% if event %}<div><span class="event-name" style="color: {{ accent_color }} !important; border-bottom-color: {{ accent_color }} !important;">{{ event }}</span></div>{% endif %}
  {% if date_hi %}<div><span class="event-date">{{ pass_label }}- {{ date_hi }} {{ validity_sfx }}</span></div>{% endif %}
</div>

<div class="body-table">
  <table style="width:100%;height:100%;border-collapse:collapse;">
    <tr>
      <!-- Logo -->
      <td class="logo-cell">
        <div class="logo-circle">
          {%- if logo_data %}<img src="{{ logo_data }}">
          {%- elif bg_image_url %}<img src="{{ bg_image_url }}">
          {%- endif %}
        </div>
      </td>

      <!-- Fields (English labels applied) -->
      <td class="fields-cell">
        <table class="fields-table">
          <tr><td class="fl">ID Code</td><td class="fc">:</td><td class="fv">{{ vol_id }}</td></tr>
          <tr><td class="fl">Name</td><td class="fc">:</td><td class="fv">{{ name }}</td></tr>
          {% if aadhaar %}<tr><td class="fl">Aadhaar</td><td class="fc">:</td><td class="fv">{{ aadhaar }}</td></tr>{% endif %}
          {% if mobile %}<tr><td class="fl">Mobile Number</td><td class="fc">:</td><td class="fv">{{ mobile }}</td></tr>{% endif %}
          {% if permission %}<tr><td class="fl">Address</td><td class="fc">:</td><td class="fv">{{ permission }}</td></tr>{% endif %}
        </table>
      </td>

      <!-- Volunteer photo + QR -->
      <td class="qr-cell">
        {% if vol_photo %}
        <div class="volunteer-photo-circle"><img src="{{ vol_photo }}"></div>
        {% else %}
        <div class="volunteer-photo-circle"><div class="vol-photo-placeholder">{{ name_initial }}</div></div>
        {% endif %}
        {% if qr_dataurl %}<img class="qr-img" src="{{ qr_dataurl }}">{% endif %}
        <div class="scan-label">SCAN TO VERIFY</div>
      </td>
    </tr>
  </table>
</div>

<div class="footer" style="background: linear-gradient(180deg, {{ header_color_light }} 0%, {{ header_color }} 100%) !important;">
  <div class="notes">
    {% if note1 %}<div class="note">{{ note1 }}</div>{% endif %}
    {% if note2 %}<div class="note">{{ note2 }}</div>{% endif %}
  </div>
  <div class="authority">
    <div class="sign-issuing">Issuing Authority :</div>
    {% if signing_image %}<img src="{{ signing_image }}" class="sign-image" alt="Signature">{% endif %}
    <div class="sign-name">{{ signing_name }}</div>
    {% if signing_title %}<div class="sign-title">{{ signing_title }}</div>{% endif %}
  </div>
</div>
</div>
""")


def _pass_context_t11(vol, event=None):
    event = event or {}
    org        = str(vol.get('org') or event.get('org_name') or '').strip()
    event_name = str(vol.get('event_label') or event.get('name') or '').strip()
    date_raw   = str(vol.get('expiry') or event.get('expiry_date') or '').strip()
    date_hi    = format_date_hi(date_raw)
    vol_id     = str(vol.get('id') or '').strip()
    name       = str(vol.get('name_hi') or vol.get('name') or '').strip()
    permission = str(vol.get('permission') or '').strip()
    aadhaar    = str(vol.get('aadhaar') or '').strip()
    mobile     = str(vol.get('mobile') or '').replace('+91', '').strip()

    if aadhaar and len(aadhaar.replace(' ', '')) >= 4:
        digits = aadhaar.replace(' ', '')
        aadhaar = 'XXXX XXXX ' + digits[-4:]

    # Pass type
    pass_type    = str(vol.get('pass_type') or 'karyakarta').strip().lower()
    pt_style     = get_pass_type_style(pass_type)
    header_color = pt_style['color']
    header_color_light = {
        '#0f52ba': '#1a6fd4', '#8B0000': '#C0392B', '#0A0A0A': '#2A2A2A',
        '#0F5C4A': '#1A9A7A', '#1A5C2A': '#247A38', '#1A2C5C': '#243D7A',
    }.get(header_color, header_color)
    pass_label   = pt_style['label']
    accent_color = pt_style['accent']
    validity_sfx = pt_style.get('validity', 'तक मान्य')

    # Logos / images
    logo_data = ''
    logo_url = event.get('logo_url', '')
    if logo_url:
        logo_data = fetch_image_as_dataurl(logo_url)
    if not logo_data:
        logo_data = logo_file_as_dataurl('srjbtk_logo_official.png')

    # Saint photo (from event)
    saint_photo = ''
    saint_url = event.get('saint_photo_url', '')
    if saint_url:
        saint_photo = fetch_image_as_dataurl(saint_url)
    # Fallback initials from signing name
    signing_name = str(event.get('signing_authority') or '').strip()
    saint_initials = ''.join(w[0] for w in signing_name.split()[:2]).upper() if signing_name else ''

    # Volunteer photo
    vol_photo = ''
    photo_url = vol.get('photo_url', '')
    if photo_url:
        vol_photo = fetch_image_as_dataurl(photo_url)

    # Name initial for placeholder
    name_initial = name[0] if name else ''

    bg_image = ''
    note1 = '* यह प्रवेश-पत्र आधार कार्ड के साथ ही मान्य है'
    note2 = '* मंदिर परिसर में मोबाइल/कैमरा इत्यादि पूर्णतः प्रतिबंधित है'

    verify_url = str(vol.get('verify_url') or '').strip()
    qr_data    = f"{verify_url}/{vol_id}" if verify_url else vol_id
    qr_dataurl = make_qr_dataurl(qr_data)

    sig_img   = ''
    sig_img_url = event.get('signing_image', '')
    if sig_img_url:
        sig_img = fetch_image_as_dataurl(sig_img_url)

    sig_title = str(event.get('signing_title') or '').strip()

    return dict(
        org=org, event=event_name, date_hi=date_hi,
        vol_id=vol_id, name=name, permission=permission,
        aadhaar=aadhaar, mobile=mobile,
        logo_data=logo_data, bg_image_url=bg_image,
        saint_photo=saint_photo, saint_initials=saint_initials,
        vol_photo=vol_photo, name_initial=name_initial,
        qr_dataurl=qr_dataurl,
        note1=note1, note2=note2,
        signing_image=sig_img, signing_name=signing_name, signing_title=sig_title,
        header_color=header_color, header_color_light=header_color_light,
        pass_label=pass_label, accent_color=accent_color,
        validity_sfx=validity_sfx, pass_type=pass_type,
    )


def render_t11_pdf(vol, event=None):
    ctx = _pass_context_t11(vol, event)
    html = PASS_DIV.render(css=T11_CSS, **ctx)
    full = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{{margin:0;padding:0;}}</style></head>
<body>{html}</body></html>"""
    return HTML(string=full, base_url=os.path.dirname(__file__)).write_pdf()


def render_t11_multi_pdf(volunteers, event=None):
    if not volunteers:
        raise ValueError("No volunteers provided")
    pages = []
    for vol in volunteers:
        ctx = _pass_context_t11(vol, event)
        pages.append(PASS_DIV.render(css=T11_CSS, **ctx))
    full = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{{margin:0;padding:0;}}</style></head>
<body>{''.join(pages)}</body></html>"""
    return HTML(string=full, base_url=os.path.dirname(__file__)).write_pdf()


# Alias for app.py routing consistency
def generate_pass_t11(volunteer: dict, event: dict, qr_url: str = '') -> bytes:
    if qr_url:
        volunteer = {**volunteer, 'verify_url': qr_url.rsplit('/', 1)[0]}
    return render_t11_pdf(volunteer, event)