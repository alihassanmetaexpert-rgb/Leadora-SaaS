# Leadora — Google Maps Lead Generation SaaS

<div align="center">

**Scrape 500 verified business leads with emails in 15 minutes. Auto-sync to Google Sheets. One click.**

[![Live Demo](https://img.shields.io/badge/Live%20Demo-leadoraleads.lovable.app-blue?style=for-the-badge)](https://leadoraleads.lovable.app)
[![Backend](https://img.shields.io/badge/Backend-Railway-purple?style=for-the-badge)](https://leadora-saas-production.up.railway.app)
[![Python](https://img.shields.io/badge/Python-3.11-green?style=for-the-badge&logo=python)](https://python.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.100+-teal?style=for-the-badge&logo=fastapi)](https://fastapi.tiangolo.com)

</div>

---

## What is Leadora?

Leadora is a SaaS tool that scrapes **Google Maps** for local business leads — names, phone numbers, emails, websites, and ratings — and automatically exports them to **Google Sheets** in real time.

Built as a direct alternative to Apollo.io ($149/month) and Outscraper, but at **$19/month** with live data and zero technical setup required.

---

## Live Demo

🌐 **Frontend:** https://leadoraleads.lovable.app  
⚙️ **Backend API:** https://leadora-saas-production.up.railway.app  
📖 **API Docs:** https://leadora-saas-production.up.railway.app/docs

---

## Features

- 🗺️ **Live Google Maps scraping** — real-time data, not a static database
- 📧 **Automatic email finder** — checks 15 pages per website in parallel
- 📊 **Auto Google Sheets sync** — one click, leads sorted by email availability
- 🔍 **Smart sub-category search** — searches 20+ variants to maximize lead coverage
- 💥 **500 leads in 15 minutes** — confirmed in production
- 🔄 **Chrome crash recovery** — auto-restarts and continues if Chrome crashes mid-scrape
- 📁 **Excel export** — one-click .xlsx download
- 🔐 **Google OAuth** — users connect their own Google account securely
- 📋 **Load My Sheets** — dropdown shows user's existing spreadsheets automatically
- 🏷️ **15+ niche tags** — Marketing Agency, Dental Clinic, Law Firm, Restaurant, and more

---

## Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Python 3.11 + FastAPI + Uvicorn |
| **Scraping** | Selenium + Chromium (headless) |
| **Email Finding** | Custom multi-page parallel scraper |
| **Database** | Redis (OAuth tokens + job state) |
| **Auth** | Google OAuth 2.0 |
| **Deployment** | Railway (Docker) |
| **Frontend** | React + Tailwind CSS (Lovable.dev) |
| **Sheets** | Google Sheets API (OAuth) |

---

## Architecture

```
User (Browser)
      │
      ▼
React Frontend (Lovable → leadoraleads.lovable.app)
      │
      ▼
FastAPI Backend (Railway → leadora-saas-production.up.railway.app)
      │
      ├── Redis (token storage + job state)
      │
      ├── Selenium + Chromium (Google Maps scraping)
      │        └── Smart sub-category search
      │        └── Chrome crash recovery + auto-restart
      │
      ├── Email Finder (15-page parallel scraper per website)
      │
      └── Google Sheets API (OAuth auto-sync)
```

---

## API Endpoints

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/scrape` | Start a new scraping job |
| `GET` | `/job/{job_id}` | Get job status + leads |
| `GET` | `/job/{job_id}/logs` | Get full job logs |
| `GET` | `/jobs` | List recent jobs |
| `DELETE` | `/job/{job_id}` | Cancel a running job |
| `GET` | `/auth/login` | Get Google OAuth URL |
| `GET` | `/auth/status/{user_id}` | Check OAuth connection |
| `GET` | `/sheets/list` | List user's Google Sheets |
| `POST` | `/sheets/sync` | Sync leads to Google Sheet |
| `POST` | `/sheets/test` | Verify sheet connection |
| `GET` | `/user/{user_id}/plan` | Get user plan + usage |
| `POST` | `/user/{user_id}/reset-usage` | Reset lead usage |
| `GET` | `/plans` | Get all plan definitions |
| `POST` | `/webhook/paddle` | Payment webhook handler |

---

## Scraping Performance

| Test | Result |
|---|---|
| Marketing Agencies — Los Angeles | 100 leads ✅ |
| Restaurants — New York | 122 leads ✅ |
| Marketing Agencies — Houston TX | 200 leads in 15 min ✅ |
| **Maximum confirmed** | **500 leads in 15 min ✅** |

---

## Pricing

| Plan | Leads | Price |
|---|---|---|
| Free Trial | 50 lifetime | $0 |
| Basic | 300/month | $19/month |
| Pro | 1,000/month | $49/month |
| Agency | 5,000/month | $99/month |

---

## Competitors

| Tool | Price | Leadora Advantage |
|---|---|---|
| Apollo.io | $149/month | 8x cheaper, live data |
| Outscraper | $3/1,000 records | All-in-one, no technical setup |
| ZoomInfo | $299/month | Live Maps vs static database |
| Lead Scrape | $247/year | Auto Sheets sync included |

**Unique advantage:** Auto Google Sheets sync with leads sorted by email — not available in any competitor at this price.

---

## Project Structure

```
Leadora-SaaS/
├── api.py              ← FastAPI backend (subscription system, OAuth, job management)
├── scraper_app.py      ← Google Maps scraper (Selenium + email finder)
├── sheets_service.py   ← Google Sheets integration
├── merge_leads.py      ← Excel merger utility
├── Dockerfile          ← Railway Docker config (Chromium included)
├── nixpacks.toml       ← Railway build config
├── requirements.txt    ← Python dependencies
├── .env.txt            ← Environment variable template
└── .gitignore          ← Excludes secrets and cache
```

---

## Environment Variables

```env
GOOGLE_CLIENT_ID=
GOOGLE_CLIENT_SECRET=
GOOGLE_CREDENTIALS_JSON=
RAILWAY_PUBLIC_DOMAIN=leadora-saas-production.up.railway.app
REDIS_URL=
OWNER_USER_IDS=
LEMONSQUEEZY_WEBHOOK_SECRET=
```

---

## Local Development

```bash
# Clone the repo
git clone https://github.com/alihassanmetaexpert-rgb/Leadora-SaaS.git
cd Leadora-SaaS

# Install dependencies
pip install -r requirements.txt

# Set environment variables (copy .env.txt to .env and fill in)
cp .env.txt .env

# Run locally
python api.py
# API available at http://localhost:8000
# Docs at http://localhost:8000/docs
```

---

## Deployment

Backend is deployed on **Railway** using Docker.

Every `git push` to `main` triggers an automatic redeploy.

```bash
git add .
git commit -m "your change"
git push origin main
# Railway auto-deploys in ~3 minutes
```

---

## Built By

**Ali Hassan** — Full-stack founder and Meta Marketing Strategist  
🌐 https://leadoraleads.lovable.app  
💼 https://www.linkedin.com/in/ali-hassan-a14461278  
🐙 https://github.com/alihassanmetaexpert-rgb

---

## License

This project is proprietary software. All rights reserved.  
© 2026 Leadora
