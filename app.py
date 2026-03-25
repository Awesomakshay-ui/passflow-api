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

def build_pdf_bytes(vols):
    pg = get_generator()
    buf = io.BytesIO()
    from reportlab.pdfgen import canvas as rl_canvas
    c = rl_canvas.Canvas(buf, pagesize=(pg.CW, pg.CH))
    for vol in vols:
        pg.draw_pass(c, vol)
        c.showPage()
    c.save()
    buf.seek(0)
    return buf

def enrich(vol, event):
    v = dict(vol)
    if not v.get('event_label') and event.get('name'):        v['event_label'] = event['name']
    if not v.get('expiry')      and event.get('expiry_date'): v['expiry']      = event['expiry_date']
    if not v.get('org')         and event.get('org_name'):    v['org']         = event['org_name']
    return v

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "service": "passflow-pass-generator"})

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
        log.info(f"Generating PDF for {len(enriched)} volunteers")
        buf = build_pdf_bytes(enriched)
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
        log.info(f"Generating single pass for {vol.get('id','unknown')}")
        buf = build_pdf_bytes([vol])
        fn  = f"pass_{str(vol.get('id') or vol.get('name') or 'pass').replace(' ','_')[:30]}.pdf"
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=fn)
    except Exception as e:
        log.error(f"generate-single error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
