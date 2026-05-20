:: ═══════════════════════════════════════════════════════════════
:: Pravesh Pass Generator — Google Cloud Run Deployment
:: Run this once to set up. After that, just run STEP 4 to redeploy.
:: ═══════════════════════════════════════════════════════════════

:: ── STEP 1: Install gcloud CLI (do once) ─────────────────────
:: Download from: https://cloud.google.com/sdk/docs/install
:: Run the installer, then restart CMD

:: ── STEP 2: Login and set project (do once) ─────────────────
gcloud auth login
gcloud projects create pravesh-passflow --name="Pravesh PassFlow"
gcloud config set project pravesh-passflow

:: Enable required APIs
gcloud services enable run.googleapis.com
gcloud services enable cloudbuild.googleapis.com
gcloud services enable containerregistry.googleapis.com

:: ── STEP 3: Build and deploy (first time) ───────────────────
cd "C:\Users\Aksha\Downloads\passflow-api"

:: Copy Dockerfile and .dockerignore into your passflow-api folder
:: (download from outputs, place them in passflow-api folder)

:: Build and deploy in one command — Cloud Build handles Docker
gcloud run deploy pravesh-pass-generator \
    --source . \
    --region asia-south1 \
    --platform managed \
    --allow-unauthenticated \
    --memory 2Gi \
    --cpu 2 \
    --timeout 300 \
    --max-instances 10 \
    --min-instances 0 \
    --port 8080

:: ── STEP 4: Redeploy after any code change ───────────────────
:: Just run this one command from passflow-api folder:
:: gcloud run deploy pravesh-pass-generator --source . --region asia-south1

:: ── STEP 5: Get your service URL ─────────────────────────────
:: After deploy you'll see:
:: Service URL: https://pravesh-pass-generator-XXXXXXXX-el.a.run.app
:: Copy that URL and update PASS_API in whatsapp.js

:: ── STEP 6: Update Worker with new URL ──────────────────────
:: In whatsapp.js change:
:: const PASS_API = 'https://passflow-pass-generator.onrender.com';
:: TO:
:: const PASS_API = 'https://pravesh-pass-generator-XXXXXXXX-el.a.run.app';
:: Then: wrangler deploy

:: ── REGION CHOICE ────────────────────────────────────────────
:: asia-south1 = Mumbai — closest to Prayagraj, lowest latency
:: Expected cold start: 2-3 seconds (vs 30-50s on Render)
:: Expected T3 pass time: 4-5 seconds (vs 17s on Render)

:: ── FREE TIER LIMITS ─────────────────────────────────────────
:: 2 million requests/month FREE
:: 360,000 GB-seconds compute FREE
:: At 2GB RAM × 5s per pass = 10 GB-seconds per pass
:: Free tier covers: 36,000 passes/month
:: Way more than you'll ever need on free tier

:: ── TROUBLESHOOTING ──────────────────────────────────────────
:: Build fails — check Dockerfile is in passflow-api folder
:: 502 errors — increase --timeout (WeasyPrint can be slow)
:: OOM errors — increase --memory (try 2Gi)
:: Cold start slow — set --min-instances 1 (costs ~$5/month)
