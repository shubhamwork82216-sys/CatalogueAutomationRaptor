# Raptor QC Suite — Streamlit Cloud Deployment Guide
© Raptor Supplies | Created by Shubham Sisodia

---

## Files in This Package

| File | Purpose | Upload to GitHub? |
|------|---------|-------------------|
| `app.py` | Main Streamlit app | ✅ Yes |
| `requirements.txt` | Python packages | ✅ Yes |
| `packages.txt` | System packages (Tesseract) | ✅ Yes |
| `.streamlit/config.toml` | App theme & settings | ✅ Yes |
| `.streamlit/secrets.toml` | API keys (**local only**) | ❌ NEVER |
| `.gitignore` | Protects secrets from GitHub | ✅ Yes |
| `README.md` | This guide | ✅ Yes |

---

## Step-by-Step Deployment

### Step 1 — Create GitHub Repository

1. Go to **github.com** → Sign in
2. Click **"New repository"**
3. Name: `raptor-qc-suite`
4. Set to **Private** (recommended for internal tools)
5. Click **"Create repository"**

### Step 2 — Upload Files to GitHub

Upload these files maintaining the folder structure:
```
raptor-qc-suite/          ← GitHub repo root
├── app.py
├── requirements.txt
├── packages.txt
├── .gitignore
├── README.md
└── .streamlit/
    └── config.toml       ← Upload this (NOT secrets.toml)
```

**⚠️ DO NOT upload `.streamlit/secrets.toml` — it contains API keys!**

### Step 3 — Deploy on Streamlit Cloud

1. Go to **share.streamlit.io**
2. Sign in with your GitHub account
3. Click **"New app"**
4. Fill in:
   - **Repository:** `your-username/raptor-qc-suite`
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Click **"Advanced settings"** → Add secrets (see Step 4)
6. Click **"Deploy!"**

### Step 4 — Add API Keys as Secrets

In Streamlit Cloud app settings → **"Secrets"** tab, paste:

```toml
MISTRAL_API_KEY = "your_actual_mistral_key"
GEMINI_API_KEY = "your_actual_gemini_key"
ADMIN_PASSWORD = "raptor_admin_2026"
```

---

## Streamlit Cloud Limits (Free Tier)

| Resource | Free Limit | Your App Needs |
|----------|-----------|----------------|
| RAM | ~1 GB | ~1.5–2 GB |
| CPU | Shared | Shared |
| Storage | 1 GB | ~500 MB |
| Apps | 3 public | 1 needed |
| Sleep after inactivity | 7 days | Wakes on visit |

**⚠️ RAM Warning:**
Your app uses PyTorch + BLIP (~1GB) + OpenCV + Pandas.
This may exceed the free tier RAM limit.

**Solutions if it crashes:**
1. Upgrade to **Streamlit Teams** ($20/month) — 2.7GB RAM
2. Use **Google Cloud Run** instead (more RAM, cheaper)
3. Disable BLIP checkbox in Advanced AI QC tab

---

## What Works on Streamlit Cloud

| Feature | Status | Notes |
|---------|--------|-------|
| Step-1 Catalogue Mapping | ✅ Works | |
| Step-2 Title Automation (Mistral) | ✅ Works | Needs API key |
| Step-3 L3 Mapping (Mistral API) | ✅ Works | Needs API key |
| Step-3 L3 Mapping (Ollama) | ❌ Won't work | Ollama is local only |
| Resolution Check | ✅ Works | |
| Reference Image Check | ✅ Works | |
| Advanced AI QC (CV only) | ✅ Works | Disable BLIP checkbox |
| Advanced AI QC (BLIP) | ⚠️ May crash | RAM dependent |
| SRGAN Image Enhancer | ⚠️ May crash | Needs PyTorch RAM |
| Product Data QC | ✅ Works | |
| A/B Testing QC | ✅ Works | |
| MRO Price Sheet Mapper | ✅ Works | |

---

## Updating the App

1. Edit `app.py` on your computer
2. Upload the new version to GitHub (replace old file)
3. Streamlit Cloud auto-detects the change and redeploys in ~2 minutes

---

## Troubleshooting

### App crashes with "Memory limit exceeded"
→ Disable BLIP checkbox in Advanced AI QC tab
→ Or upgrade to Streamlit Teams plan

### "ModuleNotFoundError" on startup
→ Check `requirements.txt` has all packages listed
→ Click "Reboot app" in Streamlit Cloud dashboard

### Tesseract not found
→ Confirm `packages.txt` is uploaded to GitHub root
→ Reboot the app from dashboard

### API key not working
→ Go to App Settings → Secrets → verify key is pasted correctly
→ No quotes needed around keys in Streamlit secrets

### App is sleeping
→ Just visit the URL — it wakes up in ~30 seconds
→ First visit after sleep triggers a cold start

---

## Comparison: Streamlit Cloud vs Google Cloud Run

| | Streamlit Cloud (Free) | Google Cloud Run |
|---|---|---|
| Setup time | 10 minutes | 20 minutes |
| Cost | Free | Free tier available |
| RAM | ~1 GB ⚠️ | 4–8 GB ✅ |
| BLIP/PyTorch | May crash | Works perfectly |
| Always on | No (sleeps) | Yes (optional) |
| Best for | Quick sharing | Production use |
