import io, os, sys, logging
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

sys.path.insert(0, os.path.dirname(__file__))
app = Flask(__name__)
CORS(app)
logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

_generator = None
def get_generator():
    global _generator
    if _generator is None:
        import pass_generator as pg
        _generator = pg
    return _generator

# T3 HTML-based renderer (WeasyPrint)
_t3_renderer = None
def get_t3_renderer():
    global _t3_renderer
    if _t3_renderer is None:
        import pass_t3_template as t3
        _t3_renderer = t3
    return _t3_renderer

_renderers = {}
def get_renderer(tmpl):
    if tmpl not in _renderers:
        mod = __import__(f'pass_{tmpl}_template')
        _renderers[tmpl] = mod
    return _renderers[tmpl]

# Page sizes in mm (width x height for landscape)
PAGE_SIZES = {
    'a6':  (148, 105),
    'a5':  (210, 148),
    'a7':  (105,  74),
    'a4':  (297, 210),  # 2-up side by side
}

def build_pdf_bytes(vols, size='a6', backside=False, template='t1'):
    pg  = get_generator()
    MM  = pg.MM
    w_mm, h_mm = PAGE_SIZES.get(size, PAGE_SIZES['a6'])
    CW  = w_mm * MM
    CH  = h_mm * MM

    two_up = (size == 'a4')
    if two_up:
        pass_w = CW / 2
        pass_h = CH
    else:
        pass_w = CW
        pass_h = CH

    orig_CW, orig_CH = pg.CW, pg.CH
    pg.CW = pass_w
    pg.CH = pass_h

    buf = io.BytesIO()
    from reportlab.pdfgen import canvas as rl_canvas
    c = rl_canvas.Canvas(buf, pagesize=(CW, CH))

    if two_up:
        for i in range(0, len(vols), 2):
            c.saveState()
            c.translate(0, 0)
            pg.draw_pass(c, vols[i], template)
            c.restoreState()
            if i + 1 < len(vols):
                c.saveState()
                c.translate(pass_w, 0)
                pg.draw_pass(c, vols[i+1], template)
                c.restoreState()
            c.showPage()
            if backside:
                c.saveState(); c.translate(0, 0)
                pg.draw_backside(c, vols[i], pass_w, pass_h)
                c.restoreState()
                if i + 1 < len(vols):
                    c.saveState(); c.translate(pass_w, 0)
                    pg.draw_backside(c, vols[i+1], pass_w, pass_h)
                    c.restoreState()
                c.showPage()
    else:
        for vol in vols:
            pg.draw_pass(c, vol, template)
            c.showPage()
            if backside:
                pg.draw_backside(c, vol, pass_w, pass_h)
                c.showPage()

    c.save()
    pg.CW = orig_CW
    pg.CH = orig_CH
    buf.seek(0)
    return buf


def build_t3_buf(vols, event, backside=False):
    """Render T3 via HTML/WeasyPrint. Returns BytesIO ready for send_file."""
    t3 = get_t3_renderer()
    pdf_bytes = t3.render_t3_multi_pdf(vols, event)
    # Optional: merge backside PDF if requested (T3 backside still uses ReportLab)
    if backside:
        from pypdf import PdfReader, PdfWriter
        # Build backside using existing ReportLab code
        backside_buf = build_pdf_bytes(vols, size='a5', backside=True, template='t1')
        # For T3 we only want the backside pages, not the T1 front. Simplest: render
        # just the backside by calling draw_backside on a fresh canvas.
        pg = get_generator()
        MM = pg.MM
        CW = 210 * MM; CH = 148 * MM
        bb = io.BytesIO()
        from reportlab.pdfgen import canvas as rl_canvas
        orig_CW, orig_CH = pg.CW, pg.CH
        pg.CW = CW; pg.CH = CH
        c = rl_canvas.Canvas(bb, pagesize=(CW, CH))
        for v in vols:
            pg.draw_backside(c, v, CW, CH)
            c.showPage()
        c.save()
        pg.CW = orig_CW; pg.CH = orig_CH
        bb.seek(0)

        # Interleave: front, back, front, back, ...
        front = PdfReader(io.BytesIO(pdf_bytes))
        back  = PdfReader(bb)
        writer = PdfWriter()
        for i in range(len(vols)):
            if i < len(front.pages): writer.add_page(front.pages[i])
            if i < len(back.pages):  writer.add_page(back.pages[i])
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return out

    return io.BytesIO(pdf_bytes)


def enrich(vol, event):
    v = dict(vol)
    pass_type = str(v.get('pass_type') or '').strip().lower().replace(' ', '_').replace('-', '_')
    if pass_type == 'vishesh_atithi':
        mobile = str(v.get('mobile') or '').strip()
        v['role'] = ''
        v['daayitva'] = ''
        v['dayitva'] = ''
        v['designation'] = ''
        v['mobile'] = mobile
        v['display_label'] = 'Mobile'
        v['display_value'] = mobile
    if not v.get('event_label')   and event.get('name'):          v['event_label']   = event['name']
    if not v.get('expiry')        and event.get('expiry_date'):   v['expiry']        = event['expiry_date']
    if not v.get('org')           and event.get('org_name'):      v['org']           = event['org_name']
    if not v.get('logo_url')      and event.get('logo_url'):      v['logo_url']      = event['logo_url']
    # If no logo set, use the default SRJBTK logo served locally
    if not v.get('logo_url'):
        v['logo_url'] = 'https://passflow-pass-generator.onrender.com/static/logo/srjbtk_logo_official.png'
    if not v.get('backside_lang')  and event.get('backside_lang'):  v['backside_lang']  = event['backside_lang']
    if not v.get('signing_image')     and event.get('signing_image'):     v['signing_image']     = event['signing_image']
    if not v.get('signing_authority') and event.get('signing_authority'): v['signing_authority'] = event['signing_authority']
    if not v.get('signing_title')     and event.get('signing_title'):     v['signing_title']     = event['signing_title']
    if not v.get('bg_image')      and event.get('bg_image'):        v['bg_image']       = event['bg_image']
    if not v.get('backside_text') and event.get('backside_text'): v['backside_text'] = event['backside_text']
    event_id = event.get('id', '')
    if event_id:
        v['verify_url'] = f'https://passflow-api.caakshayshukla.workers.dev/v/{event_id}'
    return v

@app.route('/static/logo/<filename>', methods=['GET'])
def serve_logo(filename):
    """Serve static logo/image files from the app directory."""
    from flask import send_from_directory
    import re
    # Security: only allow safe filenames
    if not re.match(r'^[a-zA-Z0-9_\-\.]+$', filename):
        return 'Not found', 404
    return send_from_directory(os.path.dirname(__file__), filename)


@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "service": "passflow-pass-generator"})

@app.route('/test-deva', methods=['GET'])
def test_deva():
    import base64, io as _io
    pg = get_generator()
    results = {}
    text = 'श्री राम जन्मभूमि'
    for color, name in [((255,255,255), 'white'), ((26,26,26), 'dark')]:
        try:
            img = pg.deva(text, pt=13, bold=True, color=color)
            if img is None:
                results[name] = 'deva() returned None'
            else:
                buf = _io.BytesIO()
                img.save(buf, 'PNG')
                results[name] = f'OK: {img.width}x{img.height}px, {buf.tell()} bytes'
        except Exception as e:
            results[name] = f'ERROR: {e}'
    return jsonify(results)


@app.route('/test-t3', methods=['GET'])
def test_t3():
    """Quick test endpoint: generates a sample T3 pass and returns it."""
    sample_vol = {
        'id': 'SMDR0010',
        'name': 'Amod Kumar Singh',
        'name_hi': 'आमोद कुमार सिंह',
        'role': 'संगठन मंत्री',
        'expiry': '29-04-2026',
        'signing_authority': 'Champat Rai',
        'signing_title': 'General Secretary',
    }
    sample_event = {
        'id': 'test-event',
        'name': 'श्री शिव मन्दिर ध्वजारोहण, प्रवेश पत्र',
        'org_name': 'श्री राम जन्मभूमि तीर्थ क्षेत्र',
        'expiry_date': '29-04-2026',
    }
    vol = enrich(sample_vol, sample_event)
    t3 = get_t3_renderer()
    pdf_bytes = t3.render_t3_pdf(vol, sample_event)
    return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                     as_attachment=False, download_name='test-t3.pdf')


@app.route('/debug-hb', methods=['GET'])
def debug_hb():
    result = {}
    font_dir  = os.path.join(os.path.dirname(__file__), 'fonts')
    font_path = os.path.join(font_dir, 'NotoSansDevanagari-Bold.ttf')
    result['font_dir_exists'] = os.path.exists(font_dir)
    result['fonts']           = os.listdir(font_dir) if os.path.exists(font_dir) else []
    result['noto_exists']     = os.path.exists(font_path)
    for lib in ['uharfbuzz', 'freetype', 'numpy', 'weasyprint', 'jinja2']:
        try:    __import__(lib); result[lib] = 'OK'
        except Exception as e: result[lib] = 'ERROR: ' + str(e)
    if result.get('uharfbuzz') == 'OK' and result.get('noto_exists'):
        try:
            import uharfbuzz as hb
            px = 60
            with open(font_path, 'rb') as f: fd = f.read()
            hf = hb.Font(hb.Face(hb.Blob(fd))); hf.scale = (px*64, px*64)
            buf2 = hb.Buffer(); buf2.add_str('अनूप'); buf2.guess_segment_properties(); hb.shape(hf, buf2, {})
            result['shaping'] = f'OK — {len(buf2.glyph_infos)} glyphs'
        except Exception as e: result['shaping'] = 'ERROR: ' + str(e)
    return jsonify(result)

@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    try:
        data = request.get_json(force=True)
        if not data: return jsonify({"error": "No JSON body"}), 400
        vols  = data.get('volunteers', [])
        event = data.get('event', {})
        if not vols: return jsonify({"error": "No volunteers"}), 400
        if len(vols) > 3000: return jsonify({"error": "Max 3000"}), 400
        pass_type_override = data.get('pass_type_override', '').strip().lower()
        # Apply override before enrichment so pass-type-specific cleanup runs.
        if pass_type_override:
            vols = [dict(v, pass_type=pass_type_override) for v in vols]
        enriched = [enrich(v, event) for v in vols]
        size     = data.get('size', 'a6').lower()
        backside = bool(data.get('backside', False))
        template = data.get('template', 't1').lower()
        log.info(f"Generating PDF for {len(enriched)} volunteers size={size} backside={backside} template={template}")

        if template == 't3':
            buf = build_t3_buf(enriched, event, backside=backside)
        else:
            buf = build_pdf_bytes(enriched, size=size, backside=backside, template=template)

        fn  = f"passes_{(event.get('name') or 'event').replace(' ','_')[:40]}_{len(enriched)}.pdf"
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=fn)
    except Exception as e:
        log.error(f"generate-pdf error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route('/generate-single', methods=['POST'])
def generate_single():
    try:
        data = request.get_json(force=True)
        if not data: return jsonify({"error": "No JSON body"}), 400
        vol   = data.get('volunteer', {})
        event = data.get('event', {})
        if not vol: return jsonify({"error": "No volunteer"}), 400
        pass_type_override = data.get('pass_type_override', '').strip().lower()
        if pass_type_override:
            vol = dict(vol, pass_type=pass_type_override)
        vol = enrich(vol, event)
        size     = data.get('size', 'a6').lower()
        backside = bool(data.get('backside', False))
        template = data.get('template', 't1').lower()
        log.info(f"Generating single pass for {vol.get('id','unknown')} size={size} backside={backside} template={template}")

        if template == 't3':
            buf = build_t3_buf([vol], event, backside=backside)
        elif template in ('t4', 't5', 't6', 't7', 't8', 't9', 't10', 't11'):
            mod = get_renderer(template)
            fn_name = f'generate_pass_{template}'
            qr_url = vol.get('qr_url') or vol.get('id','')
            pdf_bytes = getattr(mod, fn_name)(vol, event, qr_url)
            buf = io.BytesIO(pdf_bytes)
        else:
            buf = build_pdf_bytes([vol], size=size, backside=backside, template=template)

        fn  = f"pass_{str(vol.get('id') or vol.get('name') or 'pass').replace(' ','_')[:30]}.pdf"
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=fn)
    except Exception as e:
        log.error(f"generate-single error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
