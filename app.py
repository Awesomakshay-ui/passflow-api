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

    # For A4 2-up: two passes side by side on one page
    two_up = (size == 'a4')
    # Single pass dimensions for 2-up
    if two_up:
        pass_w = CW / 2
        pass_h = CH
    else:
        pass_w = CW
        pass_h = CH

    # Monkeypatch dimensions if different from default A6
    orig_CW, orig_CH = pg.CW, pg.CH
    pg.CW = pass_w
    pg.CH = pass_h

    buf = io.BytesIO()
    from reportlab.pdfgen import canvas as rl_canvas
    c = rl_canvas.Canvas(buf, pagesize=(CW, CH))

    if two_up:
        # Pair up volunteers, 2 per page
        for i in range(0, len(vols), 2):
            # Left pass
            c.saveState()
            c.translate(0, 0)
            pg.draw_pass(c, vols[i], template)
            c.restoreState()
            # Right pass (if exists)
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
    # Restore original dimensions
    pg.CW = orig_CW
    pg.CH = orig_CH
    buf.seek(0)
    return buf

def enrich(vol, event):
    v = dict(vol)
    if not v.get('event_label')   and event.get('name'):          v['event_label']   = event['name']
    if not v.get('expiry')        and event.get('expiry_date'):   v['expiry']        = event['expiry_date']
    if not v.get('org')           and event.get('org_name'):      v['org']           = event['org_name']
    if not v.get('logo_url')      and event.get('logo_url'):      v['logo_url']      = event['logo_url']
    if not v.get('backside_lang')  and event.get('backside_lang'):  v['backside_lang']  = event['backside_lang']
    if not v.get('backside_text') and event.get('backside_text'): v['backside_text'] = event['backside_text']
    # Always set verify_url from event id so QR points to correct endpoint
    event_id = event.get('id', '')
    if event_id:
        v['verify_url'] = f'https://passflow-api.caakshayshukla.workers.dev/v/{event_id}'
    return v

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


@app.route('/debug-hb', methods=['GET'])
def debug_hb():
    result = {}
    font_dir  = os.path.join(os.path.dirname(__file__), 'fonts')
    font_path = os.path.join(font_dir, 'NotoSansDevanagari-Bold.ttf')
    result['font_dir_exists'] = os.path.exists(font_dir)
    result['fonts']           = os.listdir(font_dir) if os.path.exists(font_dir) else []
    result['noto_exists']     = os.path.exists(font_path)
    for lib in ['uharfbuzz', 'freetype', 'numpy']:
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
        if result.get('freetype') == 'OK':
            try:
                import freetype
                face = freetype.Face(font_path); face.set_pixel_sizes(0, px)
                face.load_glyph(buf2.glyph_infos[0].codepoint, freetype.FT_LOAD_RENDER)
                result['freetype_render'] = f'OK — {face.glyph.bitmap.width}x{face.glyph.bitmap.rows}'
            except Exception as e: result['freetype_render'] = 'ERROR: ' + str(e)
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
        enriched = [enrich(v, event) for v in vols]
        size     = data.get('size', 'a6').lower()
        backside = bool(data.get('backside', False))
        template = data.get('template', 't1').lower()
        log.info(f"Generating PDF for {len(enriched)} volunteers size={size} backside={backside} template={template}")
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
        vol = enrich(vol, event)
        size     = data.get('size', 'a6').lower()
        backside = bool(data.get('backside', False))
        template = data.get('template', 't1').lower()
        log.info(f"Generating single pass for {vol.get('id','unknown')} size={size} backside={backside} template={template}")
        buf = build_pdf_bytes([vol], size=size, backside=backside, template=template)
        fn  = f"pass_{str(vol.get('id') or vol.get('name') or 'pass').replace(' ','_')[:30]}.pdf"
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=fn)
    except Exception as e:
        log.error(f"generate-single error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
