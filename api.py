"""
Leadora — FastAPI Backend v2.0
=================================
Google Maps Lead Generation + Google Sheets OAuth + Paddle Subscriptions
Tokens and plans stored in Redis for persistence across restarts.
"""

import os
import uuid
import json
import hashlib
import hmac
import threading
import time
from datetime import datetime, timezone
from typing import Optional

from fastapi import FastAPI, BackgroundTasks, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, RedirectResponse
from pydantic import BaseModel

# ── Load env ──────────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Google OAuth config ───────────────────────────────────────────────────────
GOOGLE_CLIENT_ID     = os.getenv("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET", "")
RAILWAY_URL          = os.getenv("RAILWAY_PUBLIC_DOMAIN", "leadora-saas-production.up.railway.app")
REDIRECT_URI         = f"https://{RAILWAY_URL}/auth/callback"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ── Lemon Squeezy config ──────────────────────────────────────────────────────
LEMONSQUEEZY_WEBHOOK_SECRET = os.getenv("LEMONSQUEEZY_WEBHOOK_SECRET", "")

# Map Lemon Squeezy Variant IDs → plan keys (fill in after store approval)
LEMONSQUEEZY_VARIANT_MAP = {
    # "123456": "basic",
    # "123457": "pro",
    # "123458": "agency",
}

# ── Plan definitions ──────────────────────────────────────────────────────────
PLANS = {
    "free_trial": {"name": "Free Trial", "leads_limit": 100,  "monthly": False, "price": 0},
    "basic":      {"name": "Basic",      "leads_limit": 1000, "monthly": True,  "price": 19},
    "pro":        {"name": "Pro",        "leads_limit": 3000, "monthly": True,  "price": 49},
    "agency":     {"name": "Agency",     "leads_limit": 10000,"monthly": True,  "price": 99},
}

# ── Redis setup ───────────────────────────────────────────────────────────────
import redis as redis_lib

REDIS_URL    = os.getenv("REDIS_URL", "redis://localhost:6379")
redis_client = redis_lib.from_url(
    REDIS_URL,
    decode_responses=True,
    socket_connect_timeout=5,
    socket_timeout=5,
    retry_on_timeout=True,
    health_check_interval=30,
)

def save_token(user_id: str, token_data: dict):
    try:
        redis_client.set(f"token:{user_id}", json.dumps(token_data), ex=60*60*24*30)
    except Exception as e:
        print(f"[Redis] save_token failed: {e}")

def get_token(user_id: str) -> dict:
    try:
        data = redis_client.get(f"token:{user_id}")
        return json.loads(data) if data else None
    except Exception as e:
        print(f"[Redis] get_token failed: {e}")
        return None

def delete_token(user_id: str):
    try:
        redis_client.delete(f"token:{user_id}")
    except Exception as e:
        print(f"[Redis] delete_token failed: {e}")

def has_token(user_id: str) -> bool:
    try:
        return redis_client.exists(f"token:{user_id}") > 0
    except Exception as e:
        print(f"[Redis] has_token failed: {e}")
        return False

# ── Plan helpers ──────────────────────────────────────────────────────────────

def get_user_plan(user_id: str) -> dict:
    try:
        data = redis_client.get(f"plan:{user_id}")
        if data:
            return json.loads(data)
    except Exception as e:
        print(f"[Redis] get_user_plan failed: {e}")
    return {
        "plan": "free_trial", "leads_used": 0, "leads_limit": 50,
        "monthly": False, "subscription_id": None,
        "activated_at": datetime.now(timezone.utc).isoformat(), "status": "active",
    }

def save_user_plan(user_id: str, plan_data: dict):
    try:
        redis_client.set(f"plan:{user_id}", json.dumps(plan_data), ex=60*60*24*400)
    except Exception as e:
        print(f"[Redis] save_user_plan failed: {e}")

def get_month_key() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m")

def get_leads_used_this_month(user_id: str) -> int:
    try:
        val = redis_client.get(f"usage:{user_id}:{get_month_key()}")
        return int(val) if val else 0
    except Exception as e:
        print(f"[Redis] get_leads_used_this_month failed: {e}")
        return 0

def add_leads_used(user_id: str, count: int):
    try:
        plan_data = get_user_plan(user_id)
        if plan_data["plan"] == "free_trial":
            plan_data["leads_used"] = plan_data.get("leads_used", 0) + count
            save_user_plan(user_id, plan_data)
        else:
            key = f"usage:{user_id}:{get_month_key()}"
            redis_client.incr(key, count)
            redis_client.expire(key, 60*60*24*35)
    except Exception as e:
        print(f"[Redis] add_leads_used failed: {e}")

def get_leads_remaining(user_id: str) -> int:
    try:
        plan_data = get_user_plan(user_id)
        limit = PLANS[plan_data["plan"]]["leads_limit"]
        used  = plan_data.get("leads_used", 0) if plan_data["plan"] == "free_trial" else get_leads_used_this_month(user_id)
        return max(0, limit - used)
    except Exception as e:
        print(f"[Redis] get_leads_remaining failed: {e}")
        return 999  # fail open — don't block the user

# ── Owner bypass for testing ──────────────────────────────────────────────────
OWNER_USER_IDS = os.getenv("OWNER_USER_IDS", "").split(",")

def check_plan_limit(user_id: str, requested: int):
    if not user_id:
        return True, ""
    if user_id.strip() in [uid.strip() for uid in OWNER_USER_IDS if uid.strip()]:
        return True, ""
    plan_data  = get_user_plan(user_id)
    plan_key   = plan_data["plan"]
    limit      = PLANS[plan_key]["leads_limit"]
    remaining  = get_leads_remaining(user_id)
    if remaining <= 0:
        msg = f"Free trial limit reached ({limit} leads total). Please upgrade to continue." if plan_key == "free_trial" else f"Monthly limit reached ({limit} leads). Resets next month or upgrade your plan."
        return False, msg
    if requested > remaining:
        period = "trial" if plan_key == "free_trial" else "month"
        return False, f"You only have {remaining} leads remaining this {period}. Reduce your request or upgrade your plan."
    return True, ""

def activate_plan(user_id: str, plan_key: str, subscription_id: str = None):
    if plan_key not in PLANS:
        raise ValueError(f"Unknown plan: {plan_key}")
    plan_data = {
        "plan": plan_key, "leads_used": 0,
        "leads_limit": PLANS[plan_key]["leads_limit"],
        "monthly": PLANS[plan_key]["monthly"],
        "subscription_id": subscription_id,
        "activated_at": datetime.now(timezone.utc).isoformat(),
        "status": "active",
    }
    save_user_plan(user_id, plan_data)
    return plan_data

# ── In-memory job store ───────────────────────────────────────────────────────
jobs: dict = {}

# ── App ───────────────────────────────────────────────────────────────────────
app = FastAPI(title="Leadora API", description="Google Maps Lead Generation API", version="2.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# ── Health check ──────────────────────────────────────────────────────────────

@app.get("/health")
def health_check():
    try:
        redis_client.ping()
        redis_ok = True
    except Exception:
        redis_ok = False
    return {"status": "ok", "redis": "connected" if redis_ok else "disconnected"}

# ── Models ────────────────────────────────────────────────────────────────────

class ScrapeRequest(BaseModel):
    query: str
    city: str
    limit: int = 20
    find_emails: bool = True
    sheet_url: Optional[str] = None
    sheet_name: str = "Leads"
    user_id: Optional[str] = None

class SheetTestRequest(BaseModel):
    sheet_url: str
    user_id: str
    sheet_name: str = "Leads"

class SheetSyncRequest(BaseModel):
    sheet_url: str
    user_id: str
    job_id: str
    sheet_name: str = "Leads"

class ActivatePlanRequest(BaseModel):
    user_id: str
    plan: str
    subscription_id: Optional[str] = None

# ── Google OAuth ──────────────────────────────────────────────────────────────

def get_google_auth_url(user_id: str) -> str:
    from urllib.parse import urlencode
    params = {
        "client_id": GOOGLE_CLIENT_ID, "redirect_uri": REDIRECT_URI,
        "response_type": "code", "scope": " ".join(SCOPES),
        "access_type": "offline", "prompt": "consent", "state": user_id,
    }
    return "https://accounts.google.com/o/oauth2/v2/auth?" + urlencode(params)

def exchange_code_for_token(code: str) -> dict:
    import requests
    return requests.post("https://oauth2.googleapis.com/token", data={
        "code": code, "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI, "grant_type": "authorization_code",
    }).json()

def get_gspread_client(user_id: str):
    import gspread
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request as GRequest
    token_data = get_token(user_id)
    if not token_data:
        raise Exception("not_authenticated")
    creds = Credentials(
        token=token_data.get("access_token"), refresh_token=token_data.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=GOOGLE_CLIENT_ID, client_secret=GOOGLE_CLIENT_SECRET, scopes=SCOPES,
    )
    if not creds.valid and creds.refresh_token:
        creds.refresh(GRequest())
        token_data["access_token"] = creds.token
        save_token(user_id, token_data)
    return gspread.authorize(creds)

# ── Sheets ────────────────────────────────────────────────────────────────────

SHEET_HEADERS = ["Name","Category","City","Address","Phone","Email","Website","Rating","Source","Maps URL","Date Added"]

def extract_sheet_id(url: str) -> str:
    import re
    url = url.strip()
    if re.match(r'^[a-zA-Z0-9_-]{20,60}$', url):
        return url
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if match:
        return match.group(1)
    raise Exception("Invalid Google Sheet URL")

def write_leads_to_sheet(client, sheet_url: str, leads: list, sheet_name: str = "Leads"):
    sheet_id    = extract_sheet_id(sheet_url)
    spreadsheet = client.open_by_key(sheet_id)
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except Exception:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=5000, cols=15)
    first_row = ws.row_values(1)
    if not first_row or first_row[0] != "Name":
        ws.insert_row(SHEET_HEADERS, index=1)
        ws.format("A1:K1", {
            "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": {"red":1,"green":1,"blue":1}},
            "backgroundColor": {"red":0.102,"green":0.451,"blue":0.910},
            "horizontalAlignment": "CENTER"
        })
    existing      = ws.get_all_values()
    existing_keys = {(r[0].strip().lower(), r[4].strip()) for r in existing[1:] if len(r) >= 5}
    now  = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Sort: leads WITH email first, leads WITHOUT email at the bottom
    sorted_leads = sorted(leads, key=lambda l: (0 if str(l.get("email","")).strip() else 1))

    rows = []
    for lead in sorted_leads:
        name  = str(lead.get("name","")).strip()
        phone = str(lead.get("phone","")).strip()
        key   = (name.lower(), phone)
        if key in existing_keys:
            continue
        existing_keys.add(key)
        rows.append([name, str(lead.get("category","")), str(lead.get("city","")),
                     str(lead.get("address","")), phone, str(lead.get("email","")),
                     str(lead.get("website","")), str(lead.get("rating","")),
                     str(lead.get("source","Google Maps")), str(lead.get("maps_url","")), now])
    if rows:
        for i in range(0, len(rows), 50):
            ws.append_rows(rows[i:i+50], value_input_option="RAW")
            time.sleep(0.5)
    return len(rows), spreadsheet.title

# ── Log collector ─────────────────────────────────────────────────────────────

def make_log_collector(job_id: str):
    def log(msg: str):
        if job_id in jobs:
            jobs[job_id]["logs"].append(str(msg))
    return log

# ── Scrape job ────────────────────────────────────────────────────────────────

def run_scrape_job(job_id: str, request: ScrapeRequest):
    log = make_log_collector(job_id)
    try:
        jobs[job_id]["status"]     = "running"
        jobs[job_id]["started_at"] = datetime.now().isoformat()
        log(f"🚀 Starting Leadora for: {request.query} in {request.city}")
        log(f"🎯 Target: {request.limit} leads")
        log("=" * 50)

        from scraper_app import (
            build_driver, collect_urls_fast, scrape_listing_fast,
            find_emails_parallel, export_excel, get_related, get_nearby, SAVE_FOLDER,
        )

        all_leads   = []
        seen_names  = set()
        seen_phones = set()
        t0 = time.time()

        def add(new_leads, src_tag):
            added = 0
            for lead in new_leads:
                if len(all_leads) >= request.limit:
                    break
                name  = lead.get("name","").strip().lower()
                phone = lead.get("phone","").strip()
                if not name or name in seen_names:
                    continue
                if phone and phone in seen_phones:
                    continue
                seen_names.add(name)
                if phone:
                    seen_phones.add(phone)
                lead["source"] = src_tag
                lead.setdefault("email", "")  # ensure email key always exists for frontend
                all_leads.append(lead)
                added += 1
                log(f"  [{len(all_leads)}/{request.limit}] ✓ {lead['name']}" + (f"  |  {phone}" if phone else ""))
                jobs[job_id]["leads_found"] = len(all_leads)
                jobs[job_id]["leads"]       = all_leads.copy()
            return added

        driver = None
        try:
            log("\n🌐 Starting Chrome browser...")
            driver = build_driver()
            log("✓ Chrome ready (single-driver mode)\n")

            def run_source(q, c, needed, label, src_tag):
                nonlocal driver
                jobs[job_id]["current_source"] = f"{label}: {q} in {c}"
                log(f"{'─'*40}")
                log(f"🔍 {label}: '{q}' in '{c}'")
                urls = collect_urls_fast(driver, q, c, needed, log, label)
                if not urls:
                    log("   No results found")
                    return 0
                log(f"   Scraping {len(urls)} listings (sequential, batch-safe)...")
                batch = []
                # Process in batches of 25 — restart Chrome if it crashes
                BATCH = 25
                for chunk_start in range(0, len(urls), BATCH):
                    if jobs[job_id].get("cancelled"):
                        break
                    chunk = urls[chunk_start:chunk_start + BATCH]
                    log(f"   Batch {chunk_start//BATCH + 1}: {len(chunk)} listings...")
                    for i, url in enumerate(chunk, 1):
                        if jobs[job_id].get("cancelled"):
                            break
                        jobs[job_id]["status"] = f"Scraping {chunk_start+i}/{len(urls)} — {len(all_leads)} leads found"
                        try:
                            d = scrape_listing_fast(driver, url)
                            if d.get("name"):
                                batch.append(d)
                        except Exception as scrape_err:
                            log(f"   ⚠ Scrape error: {scrape_err} — restarting Chrome...")
                            try: driver.quit()
                            except: pass
                            try:
                                driver = build_driver()
                                log("   ✓ Chrome restarted, continuing...")
                            except Exception as restart_err:
                                log(f"   ❌ Chrome restart failed: {restart_err}")
                                return add(batch, src_tag)
                return add(batch, src_tag)

            consecutive_timeouts = [0]  # mutable counter for nested fn

            def restart_chrome():
                nonlocal driver
                log("   🔄 Restarting Chrome...")
                try: driver.quit()
                except: pass
                try:
                    driver = build_driver()
                    log("   ✓ Chrome restarted successfully")
                    consecutive_timeouts[0] = 0
                    return True
                except Exception as e:
                    log(f"   ❌ Chrome restart failed: {e}")
                    return False

            def safe_run(q, c, needed, label, src_tag):
                """run_source with auto Chrome restart on any crash or timeout."""
                nonlocal driver
                # If Chrome is already dead, restart before trying
                try:
                    driver.execute_script("return 1;")
                except Exception:
                    log(f"   ⚠ Chrome dead before {label} — restarting...")
                    if not restart_chrome():
                        return 0
                try:
                    result = run_source(q, c, needed, label, src_tag)
                    consecutive_timeouts[0] = 0
                    return result
                except Exception as e:
                    err = str(e).lower()
                    if any(x in err for x in ["tab crashed","no such window","no such session","disconnected","timeout"]):
                        log(f"   ⚠ Chrome crashed in {label} — restarting...")
                        if restart_chrome():
                            try:
                                return run_source(q, c, needed, label, src_tag)
                            except Exception as e2:
                                log(f"   ❌ Retry failed: {e2}")
                    else:
                        log(f"   ❌ Source error: {e}")
                    return 0

            n = safe_run(request.query, request.city, request.limit, "Primary", "Google Maps")
            log(f"✓ +{n} leads  |  total: {len(all_leads)}/{request.limit}  |  {time.time()-t0:.1f}s")

            if len(all_leads) < request.limit and not jobs[job_id].get("cancelled"):
                related = get_related(request.query)
                log(f"\n🔄 Trying {len(related)} related search terms...")
                for i, term in enumerate(related, 1):
                    if len(all_leads) >= request.limit or jobs[job_id].get("cancelled"):
                        break
                    n = safe_run(term, request.city, request.limit - len(all_leads), f"Related-{i}", f"Google Maps ({term})")
                    log(f"   +{n}  total: {len(all_leads)}/{request.limit}")

            if len(all_leads) < request.limit and not jobs[job_id].get("cancelled"):
                nearby = get_nearby(request.city)
                if nearby:
                    log(f"\n🏙️  Trying {len(nearby)} nearby cities...")
                    for i, nc in enumerate(nearby, 1):
                        if len(all_leads) >= request.limit or jobs[job_id].get("cancelled"):
                            break
                        n = safe_run(request.query, nc, request.limit - len(all_leads), f"Nearby-{i}", f"Google Maps ({nc})")
                        log(f"   +{n}  total: {len(all_leads)}/{request.limit}")
        finally:
            if driver:
                try: driver.quit()
                except: pass

        if not all_leads:
            log("\n❌ No leads found. Try a different search or city.")
            jobs[job_id]["status"] = "completed"
            jobs[job_id]["leads"]  = []
            return

        log(f"\n✓ Scraping done: {len(all_leads)} leads in {time.time()-t0:.1f}s")

        # Track usage (skip for owner accounts)
        if request.user_id:
            is_owner = request.user_id.strip() in [uid.strip() for uid in OWNER_USER_IDS if uid.strip()]
            if not is_owner:
                add_leads_used(request.user_id, len(all_leads))
            remaining = get_leads_remaining(request.user_id)
            log(f"📊 Plan usage updated — {remaining} leads remaining this period{' (owner bypass)' if is_owner else ''}")

        # Find emails
        if request.find_emails and not jobs[job_id].get("cancelled"):
            log(f"\n✉  Finding emails...")
            jobs[job_id]["current_source"] = "Finding emails..."
            def email_upd(name, email, done, total):
                if email:
                    log(f"  ✉ {name} → {email}")
                else:
                    log(f"  — {name} → no email")
                jobs[job_id]["leads"] = all_leads.copy()  # always sync so frontend gets emails live
            find_emails_parallel(all_leads, log, email_upd, lambda: jobs[job_id].get("cancelled", False))
            log(f"✓ Email search done")

        # Sync to sheets
        sheets_added = 0
        if request.sheet_url and request.user_id and has_token(request.user_id):
            log(f"\n📊 Syncing {len(all_leads)} leads to Google Sheets...")
            try:
                client = get_gspread_client(request.user_id)
                added, title = write_leads_to_sheet(client, request.sheet_url, all_leads, request.sheet_name)
                sheets_added = added
                log(f"✅ {added} leads added to '{title}'!")
            except Exception as e:
                log(f"⚠️  Sheets sync failed: {e}")

        jobs[job_id]["sheets_added"] = sheets_added
        elapsed = time.time() - t0
        log(f"\n{'='*50}")
        log(f"✅ DONE in {elapsed:.1f}s!")
        log(f"📊 Total leads  : {len(all_leads)}")
        log(f"📞 With phone   : {sum(1 for l in all_leads if l.get('phone'))}")
        log(f"✉  With email   : {sum(1 for l in all_leads if l.get('email'))}")
        log(f"🌐 With website : {sum(1 for l in all_leads if l.get('website'))}")
        if sheets_added:
            log(f"📊 Sheets rows  : {sheets_added}")

        jobs[job_id]["status"]         = "completed"
        jobs[job_id]["leads"]          = all_leads
        jobs[job_id]["leads_found"]    = len(all_leads)
        jobs[job_id]["completed_at"]   = datetime.now().isoformat()
        jobs[job_id]["current_source"] = ""

    except Exception as e:
        import traceback
        log(f"\n❌ Error: {e}")
        log(traceback.format_exc())
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"]  = str(e)

# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {"status":"online","app":"Leadora API","version":"2.0.0","message":"API is running!"}

@app.post("/scrape")
def start_scrape(request: ScrapeRequest, background_tasks: BackgroundTasks):
    if not request.query.strip():
        raise HTTPException(status_code=400, detail="query cannot be empty")
    if not request.city.strip():
        raise HTTPException(status_code=400, detail="city cannot be empty")
    if request.limit < 1 or request.limit > 500:
        raise HTTPException(status_code=400, detail="limit must be between 1 and 500")

    if request.user_id:
        allowed, error_msg = check_plan_limit(request.user_id, request.limit)
        if not allowed:
            raise HTTPException(status_code=403, detail=error_msg)

    job_id = str(uuid.uuid4())[:12]
    jobs[job_id] = {
        "job_id":job_id,"status":"queued","query":request.query,"city":request.city,
        "limit":request.limit,"leads_found":0,"leads":[],"logs":[],"current_source":"",
        "excel_file":None,"sheets_added":0,"cancelled":False,
        "created_at":datetime.now().isoformat(),"started_at":None,"completed_at":None,"error":None,
    }
    threading.Thread(target=run_scrape_job, args=(job_id, request), daemon=True).start()
    return {"job_id":job_id,"message":f"Job started! Searching for {request.query} in {request.city}","poll_url":f"/job/{job_id}"}

@app.get("/job/{job_id}")
def get_job(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    job = jobs[job_id]
    return {
        "job_id":job_id,"status":job["status"],"query":job["query"],"city":job["city"],
        "leads_found":job["leads_found"],"current_source":job.get("current_source",""),
        "logs":job["logs"][-100:],"leads":job["leads"],"excel_file":job.get("excel_file"),
        "sheets_added":job.get("sheets_added",0),"error":job.get("error"),
        "created_at":job["created_at"],"completed_at":job.get("completed_at"),
    }

@app.get("/job/{job_id}/logs")
def get_logs(job_id: str, from_line: int = 0):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    logs_list = jobs[job_id]["logs"]
    return {"logs":logs_list[from_line:],"total_lines":len(logs_list),"status":jobs[job_id]["status"]}

@app.post("/job/{job_id}/cancel")
def cancel_job(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    jobs[job_id]["cancelled"] = True
    jobs[job_id]["status"]    = "cancelled"
    return {"message":"Job cancelled"}

@app.delete("/job/{job_id}")
def delete_job(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    del jobs[job_id]
    return {"message":"Job deleted"}

@app.get("/jobs")
def list_jobs():
    return [{"job_id":jid,"status":j["status"],"query":j["query"],"city":j["city"],"leads_found":j["leads_found"],"created_at":j["created_at"]} for jid,j in jobs.items()]

# ── Plan routes ───────────────────────────────────────────────────────────────

@app.get("/plans")
def list_plans():
    return {k: {"name":v["name"],"leads_limit":v["leads_limit"],"price":v["price"],"monthly":v["monthly"]} for k,v in PLANS.items()}

@app.get("/user/{user_id}/plan")
def get_plan(user_id: str):
    plan_data = get_user_plan(user_id)
    plan_key  = plan_data["plan"]
    limit     = PLANS[plan_key]["leads_limit"]
    used      = plan_data.get("leads_used",0) if plan_key == "free_trial" else get_leads_used_this_month(user_id)
    remaining = max(0, limit - used)
    return {
        "user_id":user_id,"plan":plan_key,"plan_name":PLANS[plan_key]["name"],
        "leads_limit":limit,"leads_used":used,"leads_remaining":remaining,
        "monthly":PLANS[plan_key]["monthly"],"status":plan_data.get("status","active"),
        "activated_at":plan_data.get("activated_at"),"subscription_id":plan_data.get("subscription_id"),
        "reset_info":"Resets monthly" if plan_key != "free_trial" else "Lifetime limit",
    }

@app.post("/user/activate-plan")
def activate_user_plan(request: ActivatePlanRequest):
    try:
        plan_data = activate_plan(request.user_id, request.plan, request.subscription_id)
        return {"success":True,"message":f"Plan '{request.plan}' activated for {request.user_id}","plan_data":plan_data}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/user/{user_id}/reset-usage")
def reset_user_usage(user_id: str):
    """Reset leads_used to 0 for testing."""
    plan_data = get_user_plan(user_id)
    plan_data["leads_used"] = 0
    save_user_plan(user_id, plan_data)
    redis_client.delete(f"usage:{user_id}:{get_month_key()}")
    return {"success": True, "message": f"Usage reset for {user_id}", "plan": plan_data["plan"]}

# ── Paddle webhook ────────────────────────────────────────────────────────────

@app.post("/webhook/paddle")
async def paddle_webhook(request: Request):
    body = await request.body()

    if PADDLE_WEBHOOK_SECRET:
        sig_header = request.headers.get("Paddle-Signature","")
        try:
            parts    = dict(p.split("=",1) for p in sig_header.split(";"))
            ts       = parts.get("ts","")
            h1       = parts.get("h1","")
            signed   = f"{ts}:{body.decode()}"
            expected = hmac.new(PADDLE_WEBHOOK_SECRET.encode(), signed.encode(), hashlib.sha256).hexdigest()
            if not hmac.compare_digest(h1, expected):
                raise HTTPException(status_code=401, detail="Invalid webhook signature")
        except Exception:
            raise HTTPException(status_code=401, detail="Webhook verification failed")

    try:
        payload    = json.loads(body)
        event_type = payload.get("event_type","")
        data       = payload.get("data",{})
        custom_data     = data.get("custom_data",{})
        user_id         = custom_data.get("user_id","")
        subscription_id = data.get("id","")
        items    = data.get("items",[])
        price_id = items[0].get("price",{}).get("id","") if items else ""
        plan_key = PADDLE_PRICE_MAP.get(price_id,"")

        if not user_id:
            return {"status":"ignored","reason":"no user_id in custom_data"}

        if event_type in ("subscription.activated","subscription.updated"):
            if plan_key:
                activate_plan(user_id, plan_key, subscription_id)
                return {"status":"ok","action":f"activated {plan_key} for {user_id}"}
        elif event_type == "subscription.canceled":
            activate_plan(user_id, "free_trial", None)
            return {"status":"ok","action":f"downgraded {user_id} to free_trial"}
        elif event_type == "subscription.past_due":
            plan_data = get_user_plan(user_id)
            plan_data["status"] = "past_due"
            save_user_plan(user_id, plan_data)
            return {"status":"ok","action":"marked past_due"}

        return {"status":"ignored","event_type":event_type}
    except Exception as e:
        return JSONResponse({"error":str(e)}, status_code=500)

# ── Auth routes ───────────────────────────────────────────────────────────────

@app.get("/auth/login")
def auth_login(user_id: str):
    if not GOOGLE_CLIENT_ID:
        raise HTTPException(status_code=500, detail="Google OAuth not configured")
    return RedirectResponse(url=get_google_auth_url(user_id))

@app.get("/auth/callback")
def auth_callback(code: str, state: str):
    try:
        token_data = exchange_code_for_token(code)
        save_token(state, token_data)
        return RedirectResponse(url=f"https://leadoraleads.lovable.app/dashboard?sheets_connected=true&user_id={state}")
    except Exception as e:
        return JSONResponse({"error":str(e)}, status_code=500)

@app.get("/auth/status/{user_id}")
def auth_status(user_id: str):
    connected = has_token(user_id)
    return {"user_id":user_id,"authenticated":connected,"message":"Google account connected ✅" if connected else "Not connected"}

@app.post("/auth/revoke/{user_id}")
def auth_revoke(user_id: str):
    if has_token(user_id):
        delete_token(user_id)
        return {"success":True,"message":"Google account disconnected"}
    return {"success":False,"message":"No account was connected"}

# ── Sheets routes ─────────────────────────────────────────────────────────────

@app.post("/sheets/test")
def test_sheet(request: SheetTestRequest):
    if not has_token(request.user_id):
        return {"success":False,"message":"Please connect your Google account first"}
    try:
        client      = get_gspread_client(request.user_id)
        spreadsheet = client.open_by_key(extract_sheet_id(request.sheet_url))
        return {"success":True,"message":f"Connected to '{spreadsheet.title}' ✅","sheet_title":spreadsheet.title}
    except Exception as e:
        err = str(e)
        if "not_authenticated" in err:
            return {"success":False,"message":"Please connect your Google account first"}
        return {"success":False,"message":f"Error: {err}"}

@app.post("/sheets/sync")
def sync_to_sheet(request: SheetSyncRequest):
    if not has_token(request.user_id):
        return {"success":False,"message":"Please connect your Google account first"}
    if request.job_id not in jobs:
        return {"success":False,"message":"Job not found"}
    leads = jobs[request.job_id].get("leads",[])
    if not leads:
        return {"success":False,"message":"No leads to sync"}
    try:
        client       = get_gspread_client(request.user_id)
        added, title = write_leads_to_sheet(client, request.sheet_url, leads, request.sheet_name)
        return {"success":True,"added":added,"message":f"✅ {added} leads synced to '{title}'"}
    except Exception as e:
        return {"success":False,"message":f"Error: {e}"}

@app.get("/sheets/list")
def list_user_sheets(user_id: str):
    """
    Returns a list of all Google Sheets the user has access to.
    Used by frontend to show a dropdown instead of manual URL paste.
    """
    if not has_token(user_id):
        return {"success": False, "message": "Please connect your Google account first", "sheets": []}
    try:
        client = get_gspread_client(user_id)
        spreadsheets = client.list_spreadsheet_files()
        sheets = [
            {
                "id":    s["id"],
                "name":  s["name"],
                "url":   f"https://docs.google.com/spreadsheets/d/{s['id']}",
            }
            for s in spreadsheets
        ]
        return {"success": True, "sheets": sheets, "count": len(sheets)}
    except Exception as e:
        err = str(e)
        if "not_authenticated" in err:
            return {"success": False, "message": "Please connect your Google account first", "sheets": []}
        return {"success": False, "message": f"Error: {err}", "sheets": []}


# ── Debug endpoint ────────────────────────────────────────────────────────────

@app.get("/auth/debug-token")
async def debug_token(request: Request):
    import requests as req
    code = request.query_params.get("code", "")
    if not code:
        return {"error": "No code provided — usage: /auth/debug-token?code=YOUR_CODE"}
    response = req.post("https://oauth2.googleapis.com/token", data={
        "code": code,
        "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
    })
    result = response.json()
    if "access_token" in result:
        save_token("debug_test", result)
        return {"status": response.status_code, "success": True, "has_access_token": True, "has_refresh_token": "refresh_token" in result, "redirect_uri_used": REDIRECT_URI}
    return {"status": response.status_code, "success": False, "error": result, "redirect_uri_used": REDIRECT_URI}

# ── Run ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*55)
    print("  Leadora API v2.0 — Starting...")
    print("="*55)
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True)
