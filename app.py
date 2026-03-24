"""
PassFlow Pass Generator — Flask Microservice
=============================================
Accepts volunteer data as JSON, returns a PDF using the
exact same Python pass generator used locally.

POST /generate-pdf
  Body: { "volunteers": [...], "event": {...} }
  Returns: application/pdf

POST /generate-single
  Body: { "volunteer": {...}, "event": {...} }
  Returns: application/pdf

GET /health
  Returns: { "status": "ok" }
"""

import io
import os
import sys
import logging

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# ── Import the pass generator ──────────────────────────────────────
# pass_generator.py sits in the same folder as this file
sys.path.insert(0, os.path.dirname(__file__))

app = Flask(__name__)
CORS(app)  # Allow requests from passflow.pages.dev and any origin

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)


# ── Lazy-load the generator (avoids font registration errors on import) ──
_generator = None

def get_generator():
    global _generator
    if _generator is None:
        import pass_generator as pg
        _generator = pg
    return _generator


def build_pdf_bytes(vols):
    """Run the pass generator on a list of volunteer dicts, return PDF bytes."""
    pg = get_generator()

    # Write to an in-memory buffer
    buf = io.BytesIO()

    from reportlab.pdfgen import canvas as rl_canvas
    c = rl_canvas.Canvas(buf, pagesize=(pg.CW, pg.CH))
    for vol in vols:
        pg.draw_pass(c, vol)
        c.showPage()
    c.save()

    buf.seek(0)
    return buf


@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "service": "passflow-pass-generator"})


@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    """
    Generate a multi-page PDF for a list of volunteers.
    Body JSON:
    {
        "volunteers": [
            {
                "id": "VPS0001",
                "name": "Rajendra Kumar",
                "name_hi": "राजेंद्र कुमार",
                "role": "Security",
                "aadhaar": "XXXX XXXX 4257",
                "mobile": "7080118720",
                "permission": "Mobile Only",
                "pass_type": "standard",
                "expiry": "31-03-2026",
                "org": "Shri Ram Janmbhoomi Teerth Kshetra",
                "event_label": "Ram Navami 2026"
            },
            ...
        ],
        "event": {
            "name": "Ram Navami 2026",
            "expiry_date": "31-03-2026"
        }
    }
    """
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "No JSON body"}), 400

        volunteers = data.get('volunteers', [])
        event      = data.get('event', {})

        if not volunteers:
            return jsonify({"error": "No volunteers provided"}), 400

        if len(volunteers) > 3000:
            return jsonify({"error": "Maximum 3000 passes per request"}), 400

        # Enrich each volunteer with event info if not already set
        enriched = []
        for v in volunteers:
            vol = dict(v)
            if not vol.get('event_label') and event.get('name'):
                vol['event_label'] = event['name']
            if not vol.get('expiry') and event.get('expiry_date'):
                vol['expiry'] = event['expiry_date']
            if not vol.get('org') and event.get('org_name'):
                vol['org'] = event['org_name']
            enriched.append(vol)

        log.info(f"Generating PDF for {len(enriched)} volunteers")
        buf = build_pdf_bytes(enriched)

        event_name = (event.get('name') or 'passes').replace(' ', '_')[:40]
        filename   = f"passes_{event_name}_{len(enriched)}.pdf"

        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        log.error(f"generate-pdf error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route('/generate-single', methods=['POST'])
def generate_single():
    """
    Generate a single-page PDF for one volunteer.
    Body JSON: { "volunteer": {...}, "event": {...} }
    """
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "No JSON body"}), 400

        vol   = data.get('volunteer', {})
        event = data.get('event', {})

        if not vol:
            return jsonify({"error": "No volunteer provided"}), 400

        # Enrich with event info
        if not vol.get('event_label') and event.get('name'):
            vol['event_label'] = event['name']
        if not vol.get('expiry') and event.get('expiry_date'):
            vol['expiry'] = event['expiry_date']
        if not vol.get('org') and event.get('org_name'):
            vol['org'] = event['org_name']

        log.info(f"Generating single pass for {vol.get('id', 'unknown')}")
        buf = build_pdf_bytes([vol])

        vol_id   = str(vol.get('id') or vol.get('name') or 'pass').replace(' ', '_')[:30]
        filename = f"pass_{vol_id}.pdf"

        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        log.error(f"generate-single error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
