# PassFlow Pass Generator — Microservice

Python Flask API that generates volunteer passes as print-ready PDFs.
Deployed on Render.com, called by the PassFlow web app.

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Health check |
| POST | `/generate-pdf` | Generate PDF for multiple volunteers |
| POST | `/generate-single` | Generate single-page PDF for one volunteer |

---

## Setup & Deployment (Render.com)

### Step 1 — Add font files
Copy these files into a `fonts/` folder inside this repo:
```
fonts/
  Poppins-Bold.ttf
  Poppins-Regular.ttf
  Poppins-Light.ttf
  Poppins-Medium.ttf
  NotoSansDevanagari-Bold.ttf
  NotoSansDevanagari-Regular.ttf
  FreeSerifBold.ttf
  FreeSerif.ttf
```
These are already on your machine at:
`C:\Users\Aksha\Downloads\Volunteer passes\fonts\`

### Step 2 — Add logo
Copy `srjbtk_logo_official.png` to the root of this repo.
It's at: `C:\Users\Aksha\Downloads\Volunteer passes\srjbtk_logo_official.png`

### Step 3 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial passflow microservice"
git remote add origin https://github.com/YOUR_USERNAME/passflow-api.git
git push -u origin main
```

### Step 4 — Deploy on Render.com
1. Go to https://render.com → New → Web Service
2. Connect your GitHub repo
3. Render auto-detects `render.yaml` and configures everything
4. Click Deploy
5. Your service URL will be: `https://passflow-pass-generator.onrender.com`

### Step 5 — Update PassFlow frontend
In `index.html`, set:
```javascript
const PASS_API = 'https://passflow-pass-generator.onrender.com';
```

---

## API Reference

### POST /generate-pdf
```json
{
  "volunteers": [
    {
      "id": "VPS0001",
      "name": "Rajendra Kumar",
      "name_hi": "राजेंद्र कुमार",
      "role": "Security",
      "aadhaar": "123412341234",
      "mobile": "9876543210",
      "permission": "Mobile Only",
      "pass_type": "standard",
      "expiry": "31-03-2026",
      "org": "Shri Ram Janmbhoomi Teerth Kshetra",
      "event_label": "Ram Navami 2026"
    }
  ],
  "event": {
    "name": "Ram Navami 2026",
    "expiry_date": "31-03-2026"
  }
}
```
Returns: `application/pdf`

### POST /generate-single
```json
{
  "volunteer": { ...same fields as above... },
  "event": { "name": "Ram Navami 2026" }
}
```
Returns: `application/pdf`

---

## Notes
- Free Render tier spins down after 15 min inactivity — first request takes ~30 sec
- Upgrade to Render Starter ($7/month) for always-on
- Max 3000 passes per request
- Fonts and logo must be present for full quality output
