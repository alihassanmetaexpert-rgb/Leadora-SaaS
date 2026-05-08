"""
LeadHunter Pro — FastAPI Backend
=================================
This file wraps your existing scraper_app.py and sheets_service.py
into a web API so customers can use it from a browser.

Place this file at: C:\LeadScraper\api.py
Run with: uvicorn api:app --reload --host 0.0.0.0 --port 8000
"""

import os
import sys
import uuid
import json
import asyncio
import threading
import time
from datetime import datetime
from typing import Optional
from pathlib import Path

from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel

# ── Add LeadScraper folder to path ────────────────────────────────────────────
sys.path.insert(0, r"C:\LeadScraper")

# ── Import your existing sheets service ───────────────────────────────────────
from sheets_service import (
    append_leads_to_sheet,
    test_sheet_connection,
    extract_spreadsheet_id,
    authenticate_user,
    is_user_authenticated,
    revoke_user_token,
    InvalidSheetURLError,
)

# ── App setup ─────────────────────────────────────────────────────────────────
app = FastAPI(
    title="LeadHunter Pro API",
    description="Google Maps Lead Generation API",
    version="1.0.0"
)

# Allow frontend to talk to this backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # We'll restrict this later when we deploy
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── In-memory job store ───────────────────────────────────────────────────────
# Stores the status and results of each scraping job
jobs: dict = {}

# ── Request models ────────────────────────────────────────────────────────────

class ScrapeRequest(BaseModel):
    query: str          # e.g. "dental clinic"
    city: str           # e.g. "Lahore Pakistan"
    limit: int = 20     # how many leads to find
    find_emails: bool = True
    sheet_url: Optional[str] = None   # user's Google Sheet URL
    sheet_name: str = "Leads"
    user_id: Optional[str] = None     # user session ID

class SheetTestRequest(BaseModel):
    sheet_url: str
    user_id: str
    sheet_name: str = "Leads"

class AuthRequest(BaseModel):
    user_id: str

# ── Helper: safe log collector ────────────────────────────────────────────────

def make_log_collector(job_id: str):
    """Returns a log function that saves messages to the job's log list."""
    def log(msg: str):
        if job_id in jobs:
            jobs[job_id]["logs"].append(str(msg))
    return log

# ── The actual scraping function (runs in background thread) ──────────────────

def run_scrape_job(job_id: str, request: ScrapeRequest):
    """
    This runs your existing scraper logic in a background thread.
    It imports the core scraping functions from scraper_app.py.
    """
    log = make_log_collector(job_id)

    try:
        jobs[job_id]["status"] = "running"
        jobs[job_id]["started_at"] = datetime.now().isoformat()
        log(f"🚀 Starting LeadHunter for: {request.query} in {request.city}")
        log(f"🎯 Target: {request.limit} leads")
        log("=" * 50)

        # Import your scraper functions
        # We import here (inside the thread) to avoid tkinter GUI loading
        from scraper_app import (
            build_driver,
            collect_urls_fast,
            scrape_parallel,
            find_emails_parallel,
            export_excel,
            WorkerPool,
            get_related,
            get_nearby,
            DETAIL_WORKERS,
            SAVE_FOLDER,
        )

        all_leads = []
        seen_names = set()
        seen_phones = set()
        t0 = time.time()

        def add(new_leads, src_tag):
            added = 0
            for lead in new_leads:
                if len(all_leads) >= request.limit:
                    break
                name = lead.get("name", "").strip().lower()
                phone = lead.get("phone", "").strip()
                if not name or name in seen_names:
                    continue
                if phone and phone in seen_phones:
                    continue
                seen_names.add(name)
                if phone:
                    seen_phones.add(phone)
                lead["source"] = src_tag
                all_leads.append(lead)
                added += 1
                log(f"  [{len(all_leads)}/{request.limit}] ✓ {lead['name']}" +
                    (f"  |  {phone}" if phone else ""))
                # Update live progress
                jobs[job_id]["leads_found"] = len(all_leads)
                jobs[job_id]["leads"] = all_leads.copy()
            return added

        driver = None
        pool = None

        try:
            log("\n🌐 Starting Chrome browser...")
            driver = build_driver()
            log("✓ Chrome ready")

            pool = WorkerPool(size=DETAIL_WORKERS)
            pool.start(log)
            log(f"✓ {DETAIL_WORKERS} parallel workers ready\n")

            def run_source(q, c, needed, label, src_tag):
                jobs[job_id]["current_source"] = f"{label}: {q} in {c}"
                log(f"─" * 40)
                log(f"🔍 {label}: '{q}' in '{c}'")
                urls = collect_urls_fast(driver, q, c, needed, log, label)
                if not urls:
                    log(f"   No results found")
                    return 0
                log(f"   Scraping {len(urls)} listings ({DETAIL_WORKERS} parallel)...")

                def status_cb(done, total, found):
                    jobs[job_id]["status"] = (
                        f"Scraping {done}/{total} — {len(all_leads)+found} leads found"
                    )

                batch = scrape_parallel(
                    urls, pool, log, status_cb,
                    lambda: jobs[job_id].get("cancelled", False)
                )
                return add(batch, src_tag)

            # SOURCE 1: Primary search
            n = run_source(request.query, request.city, request.limit, "Primary", "Google Maps")
            log(f"✓ +{n} leads  |  total: {len(all_leads)}/{request.limit}  |  {time.time()-t0:.1f}s")

            # SOURCE 2: Related terms (if still need more)
            if len(all_leads) < request.limit and not jobs[job_id].get("cancelled"):
                related = get_related(request.query)
                log(f"\n🔄 Trying {len(related)} related search terms...")
                for i, term in enumerate(related, 1):
                    if len(all_leads) >= request.limit or jobs[job_id].get("cancelled"):
                        break
                    n = run_source(term, request.city, request.limit - len(all_leads),
                                   f"Related-{i}", f"Google Maps ({term})")
                    log(f"   +{n}  total: {len(all_leads)}/{request.limit}")

            # SOURCE 3: Nearby cities (if still need more)
            if len(all_leads) < request.limit and not jobs[job_id].get("cancelled"):
                nearby = get_nearby(request.city)
                if nearby:
                    log(f"\n🏙️  Trying {len(nearby)} nearby cities...")
                    for i, nc in enumerate(nearby, 1):
                        if len(all_leads) >= request.limit or jobs[job_id].get("cancelled"):
                            break
                        n = run_source(request.query, nc,
                                       request.limit - len(all_leads),
                                       f"Nearby-{i}", f"Google Maps ({nc})")
                        log(f"   +{n}  total: {len(all_leads)}/{request.limit}")

        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            if pool:
                try:
                    pool.quit()
                except:
                    pass

        if not all_leads:
            log("\n❌ No leads found. Try a different search or city.")
            jobs[job_id]["status"] = "completed"
            jobs[job_id]["leads"] = []
            return

        log(f"\n✓ Scraping done: {len(all_leads)} leads in {time.time()-t0:.1f}s")

        # Find emails
        if request.find_emails and not jobs[job_id].get("cancelled"):
            log(f"\n✉  Finding emails (parallel workers)...")
            jobs[job_id]["current_source"] = "Finding emails..."

            def email_upd(name, email, done, total):
                if email:
                    log(f"  ✉ {name} → {email}")
                jobs[job_id]["leads"] = all_leads.copy()

            find_emails_parallel(
                all_leads, log, email_upd,
                lambda: jobs[job_id].get("cancelled", False)
            )

        # Save Excel file
        fname = f"leads_{request.query.replace(' ','_')}_{request.city.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        fpath = os.path.join(SAVE_FOLDER, fname)
        export_excel(all_leads, fpath)
        log(f"\n📁 Excel saved: {fname}")
        jobs[job_id]["excel_file"] = fname

        # Sync to Google Sheets
        sheets_added = 0
        if request.sheet_url and request.user_id and not jobs[job_id].get("cancelled"):
            log(f"\n📊 Syncing to Google Sheets...")
            result = append_leads_to_sheet(
                request.sheet_url,
                all_leads,
                request.user_id,
                request.sheet_name,
                log
            )
            if result["success"]:
                sheets_added = result.get("added", 0)
                log(f"✅ {sheets_added} leads added to your Google Sheet!")
            else:
                log(f"⚠️  Sheets sync failed: {result['message']}")
            jobs[job_id]["sheets_added"] = sheets_added

        # Final summary
        elapsed = time.time() - t0
        log(f"\n{'=' * 50}")
        log(f"✅ DONE in {elapsed:.1f}s!")
        log(f"📊 Total leads  : {len(all_leads)}")
        log(f"📞 With phone   : {sum(1 for l in all_leads if l.get('phone'))}")
        log(f"✉  With email   : {sum(1 for l in all_leads if l.get('email'))}")
        log(f"🌐 With website : {sum(1 for l in all_leads if l.get('website'))}")
        if sheets_added:
            log(f"📊 Sheets rows  : {sheets_added}")

        jobs[job_id]["status"] = "completed"
        jobs[job_id]["leads"] = all_leads
        jobs[job_id]["leads_found"] = len(all_leads)
        jobs[job_id]["completed_at"] = datetime.now().isoformat()
        jobs[job_id]["current_source"] = ""

    except Exception as e:
        import traceback
        log(f"\n❌ Error: {e}")
        log(traceback.format_exc())
        jobs[job_id]["status"] = "error"
        jobs[job_id]["error"] = str(e)


# ── API Routes ────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    """Health check — confirms API is running."""
    return {
        "status": "online",
        "app": "LeadHunter Pro API",
        "version": "1.0.0",
        "message": "API is running! Send POST /scrape to start."
    }


@app.post("/scrape")
def start_scrape(request: ScrapeRequest, background_tasks: BackgroundTasks):
    """
    Start a new lead generation job.
    Returns a job_id immediately — poll /job/{job_id} for progress.
    """
    if not request.query.strip():
        raise HTTPException(status_code=400, detail="query cannot be empty")
    if not request.city.strip():
        raise HTTPException(status_code=400, detail="city cannot be empty")
    if request.limit < 1 or request.limit > 200:
        raise HTTPException(status_code=400, detail="limit must be between 1 and 200")

    job_id = str(uuid.uuid4())[:12]

    jobs[job_id] = {
        "job_id": job_id,
        "status": "queued",
        "query": request.query,
        "city": request.city,
        "limit": request.limit,
        "leads_found": 0,
        "leads": [],
        "logs": [],
        "current_source": "",
        "excel_file": None,
        "sheets_added": 0,
        "cancelled": False,
        "created_at": datetime.now().isoformat(),
        "started_at": None,
        "completed_at": None,
        "error": None,
    }

    # Run scraping in background thread (not blocking the API)
    thread = threading.Thread(
        target=run_scrape_job,
        args=(job_id, request),
        daemon=True
    )
    thread.start()

    return {
        "job_id": job_id,
        "message": f"Job started! Searching for {request.query} in {request.city}",
        "poll_url": f"/job/{job_id}"
    }


@app.get("/job/{job_id}")
def get_job(job_id: str):
    """
    Get the current status and results of a scraping job.
    Poll this every 2-3 seconds from the frontend.
    """
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")

    job = jobs[job_id]
    return {
        "job_id": job_id,
        "status": job["status"],             # queued / running / completed / error
        "query": job["query"],
        "city": job["city"],
        "leads_found": job["leads_found"],
        "current_source": job.get("current_source", ""),
        "logs": job["logs"][-30:],           # last 30 log lines
        "leads": job["leads"],               # all leads so far (live update)
        "excel_file": job.get("excel_file"),
        "sheets_added": job.get("sheets_added", 0),
        "error": job.get("error"),
        "created_at": job["created_at"],
        "completed_at": job.get("completed_at"),
    }


@app.get("/job/{job_id}/logs")
def get_logs(job_id: str, from_line: int = 0):
    """Get logs from a specific line (for incremental log streaming)."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    logs = jobs[job_id]["logs"]
    return {
        "logs": logs[from_line:],
        "total_lines": len(logs),
        "status": jobs[job_id]["status"]
    }


@app.post("/job/{job_id}/cancel")
def cancel_job(job_id: str):
    """Cancel a running scraping job."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    jobs[job_id]["cancelled"] = True
    jobs[job_id]["status"] = "cancelled"
    return {"message": "Job cancelled"}


@app.delete("/job/{job_id}")
def delete_job(job_id: str):
    """Remove a job from memory."""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    del jobs[job_id]
    return {"message": "Job deleted"}


@app.get("/jobs")
def list_jobs():
    """List all jobs (for debugging)."""
    return [
        {
            "job_id": jid,
            "status": j["status"],
            "query": j["query"],
            "city": j["city"],
            "leads_found": j["leads_found"],
            "created_at": j["created_at"],
        }
        for jid, j in jobs.items()
    ]


# ── Google Sheets Routes ───────────────────────────────────────────────────────

@app.post("/sheets/test")
def test_sheet(request: SheetTestRequest):
    """Test connection to a user's Google Sheet."""
    result = test_sheet_connection(
        request.sheet_url,
        request.user_id,
        request.sheet_name
    )
    return result


@app.post("/sheets/auth/start")
def start_auth(request: AuthRequest):
    """
    Start Google OAuth flow for a user.
    This opens a browser window on the SERVER — only works locally.
    For cloud deployment we'll handle this differently.
    """
    already = is_user_authenticated(request.user_id)
    if already:
        return {"success": True, "message": "Already authenticated with Google"}

    # Run auth in background (opens browser)
    def do_auth():
        authenticate_user(request.user_id)

    thread = threading.Thread(target=do_auth, daemon=True)
    thread.start()

    return {
        "success": True,
        "message": "Authentication started — check your browser to approve access"
    }


@app.get("/sheets/auth/status/{user_id}")
def auth_status(user_id: str):
    """Check if a user is authenticated with Google."""
    authenticated = is_user_authenticated(user_id)
    return {
        "user_id": user_id,
        "authenticated": authenticated,
        "message": "Connected to Google" if authenticated else "Not connected — please authenticate"
    }


@app.post("/sheets/auth/revoke")
def revoke_auth(request: AuthRequest):
    """Disconnect a user's Google account."""
    success = revoke_user_token(request.user_id)
    return {
        "success": success,
        "message": "Google account disconnected" if success else "No account was connected"
    }


# ── Run the server ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    print("\n" + "=" * 55)
    print("  LeadHunter Pro API — Starting...")
    print("=" * 55)
    print(f"  URL: http://localhost:8000")
    print(f"  Docs: http://localhost:8000/docs")
    print("=" * 55 + "\n")
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True)
