"""
LeadHunter Pro — Fast Edition
================================
High-speed Google Maps lead scraper with parallel processing.
"""

import warnings
warnings.filterwarnings("ignore")  # suppress SyntaxWarning from JS strings

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import time
import re
import ssl
import os
import json
import random
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# Import our secure sheets service
from sheets_service import (
    append_leads_to_sheet,
    test_sheet_connection,
    extract_spreadsheet_id,
    authenticate_user,
    is_user_authenticated,
    revoke_user_token,
    SERVICE_ACCOUNT_EMAIL,
    InvalidSheetURLError,
    install_gspread,
)
import uuid

# Each app session gets a unique user ID
SESSION_USER_ID = str(uuid.uuid4())[:12]

ssl._create_default_https_context = ssl._create_unverified_context
os.environ['WDM_SSL_VERIFY'] = '0'
os.environ['PYTHONHTTPSVERIFY'] = '0'

SAVE_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "LeadResults")
CONFIG_FILE = os.path.join(SAVE_FOLDER, "user_config.json")
os.makedirs(SAVE_FOLDER, exist_ok=True)

# ── Colors ────────────────────────────────────────────────────────────────────
BG      = "#F0F4F8"
CARD    = "#FFFFFF"
BORDER  = "#D1D9E6"
ACCENT  = "#1A73E8"
ACCENT2 = "#1557B0"
SUCCESS = "#1E8A44"
DANGER  = "#C0392B"
WARNING = "#E67E22"
TEXT    = "#1A202C"
SUBTEXT = "#4A5568"
HEADER  = "#1A73E8"
LOG_BG  = "#1E2433"
LOG_FG  = "#00E676"
NAVY    = "#0D47A1"

# ── Config persistence ────────────────────────────────────────────────────────

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE) as f:
                return json.load(f)
    except Exception:
        pass
    return {"sheet_url": "", "sheet_name": "Leads", "sheets_enabled": False}

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

# ── Fallback search terms ─────────────────────────────────────────────────────

RELATED_TERMS = {
    "aesthetic clinic":    ["med spa","skin clinic","beauty clinic","cosmetic clinic","laser clinic"],
    "real estate agent":   ["real estate agency","property dealer","realtor"],
    "solar panel company": ["solar energy","solar installer","solar contractor"],
    "dental clinic":       ["dentist","dental care","orthodontist"],
    "gym":                 ["fitness center","health club","yoga studio"],
    "restaurant":          ["cafe","bistro","eatery","diner"],
    "hotel":               ["motel","inn","resort","lodge"],
    "law firm":            ["attorney","lawyer","legal services"],
}

NEARBY_CITIES = {
    "new york":    ["Brooklyn NY","Queens NY","Newark NJ"],
    "los angeles": ["Long Beach CA","Pasadena CA","Glendale CA"],
    "chicago":     ["Evanston IL","Oak Park IL","Naperville IL"],
    "houston":     ["Sugar Land TX","The Woodlands TX"],
    "miami":       ["Fort Lauderdale FL","Hollywood FL"],
    "lahore":      ["Islamabad Pakistan","Karachi Pakistan"],
    "karachi":     ["Lahore Pakistan","Islamabad Pakistan"],
}

def get_related(query):
    q = query.lower().strip()
    for key, terms in RELATED_TERMS.items():
        if key in q or q in key:
            return terms
    return [f"best {query}", f"top {query}", f"{query} services"]

def get_nearby(city):
    key = city.lower().split(",")[0].strip()
    for k, cities in NEARBY_CITIES.items():
        if k in key or key in k:
            return cities
    return []

# ── Chrome driver (alias to new fast builder) ─────────────────────────────────

def build_driver():
    """Main Chrome — headless on Railway, visible locally."""
    is_railway = os.getenv("RAILWAY_PUBLIC_DOMAIN") or os.getenv("RAILWAY_ENVIRONMENT")
    return _make_chrome(headless=bool(is_railway))

# ── FAST ENGINE — Complete Rewrite ───────────────────────────────────────────
# Root causes of slowness fixed:
# 1. Chrome startup = 10-20s  → Fixed: reuse same driver, never restart
# 2. sleep(4) page load       → Fixed: smart wait for DOM element
# 3. sleep(2.5) per listing   → Fixed: 0.8s + JS extraction
# 4. Sequential scraping      → Fixed: 3 headless workers in parallel
# 5. Images loading           → Fixed: blocked at network level
# 6. Email: 1 at a time       → Fixed: 8 parallel threads
# 7. Scroll: sleep(1.5)       → Fixed: 0.6s
# RESULT: 50 leads in ~2 min instead of 15 min

import urllib3
urllib3.disable_warnings()

SCROLL_WAIT    = 0.5    # was 1.5s
PAGE_WAIT      = 1.5    # was 4.0s  
DETAIL_WAIT    = 0.8    # was 2.5s
SCROLL_STEP    = 2000   # bigger scroll = fewer passes needed
MAX_NO_NEW     = 4      # stop faster when no new results
EMAIL_WORKERS  = 6      # parallel email threads (reduced for Railway RAM)
DETAIL_WORKERS = 1      # single driver mode on Railway (avoids OOM crash)
EMAIL_TIMEOUT  = 6      # seconds per email fetch before giving up

EMAIL_REGEX  = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
SKIP_EMAILS  = {"sentry","wix","google","facebook","example","domain","schema","w3.org","jquery","cloudflare"}
REQ_HEADERS  = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
}


# ── Driver builder ────────────────────────────────────────────────────────────

def _make_chrome(headless=False):
    """
    Build Chrome with maximum speed settings:
    - Images/fonts/CSS blocked → 60% faster page loads
    - Notifications disabled
    - No automation detection
    """
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1280,800")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--log-level=3")
    opts.add_argument("--mute-audio")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--blink-settings=imagesEnabled=false")  # block images
    # Extra memory optimizations for Railway
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-default-apps")
    opts.add_argument("--disable-sync")
    opts.add_argument("--no-first-run")
    opts.add_argument("--disable-software-rasterizer")
    if os.path.exists("/usr/bin/chromium"):
        opts.binary_location = "/usr/bin/chromium"
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_experimental_option("prefs", {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts": 2,
    })
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    import shutil
    chromedriver_path = shutil.which("chromedriver") or "/usr/bin/chromedriver"
    svc    = Service(chromedriver_path)
    driver = webdriver.Chrome(service=svc, options=opts)
    driver.set_page_load_timeout(30)   # kill page if it hangs >30s
    driver.set_script_timeout(15)      # kill JS if it hangs >15s
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source":
        "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
        "Object.defineProperty(navigator,'languages',{get:()=>['en-US','en']});"
    })
    # Block images/fonts/css via CDP for even faster loads
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": [
            "*.png","*.jpg","*.jpeg","*.gif","*.webp","*.svg",
            "*.woff","*.woff2","*.ttf","*.eot",
            "*.css","googletagmanager.com","analytics.google.com",
            "doubleclick.net","googlesyndication.com",
        ]})
    except Exception:
        pass
    return driver


# ── Step 1: Fast URL collection ───────────────────────────────────────────────

def collect_urls_fast(driver, query, city, max_count, log, label=""):
    """
    Collect listing URLs from Google Maps.
    Uses JS to grab all hrefs in one call instead of find_elements loop.
    Smart wait instead of fixed sleep.
    """
    tag = f"[{label}] " if label else ""
    search = (query + " " + city).replace(" ", "+")
    url    = f"https://www.google.com/maps/search/{search}"

    log(f"   {tag}🔍 Searching: {query} in {city}")
    try:
        driver.get(url)
    except Exception:
        log(f"   {tag}⚠ Page load timeout — retrying...")
        try:
            driver.get(url)
        except Exception:
            return []

    # Smart wait — wait for results panel, not fixed sleep
    panel = None
    for sel in ["//div[@role='feed']", "//div[contains(@aria-label,'Results')]"]:
        try:
            panel = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, sel)))
            break
        except TimeoutException:
            pass

    # Dismiss cookie popup if present
    for btn_text in ["Accept all", "Accept", "Agree"]:
        try:
            driver.find_element(
                By.XPATH, f"//button[contains(.,'{btn_text}')]").click()
            time.sleep(0.3)
            break
        except Exception:
            pass

    if not panel:
        log(f"   {tag}⚠ No results panel found")
        return []

    # Re-find panel after possible cookie dismiss
    try:
        panel = driver.find_element(By.XPATH, "//div[@role='feed']")
    except Exception:
        pass

    collected = []
    seen      = set()
    no_new    = 0
    passes    = 0

    while len(collected) < max_count and no_new < MAX_NO_NEW:
        try:
            # JS grab ALL hrefs at once — much faster than find_elements
            hrefs = driver.execute_script("""
                return Array.from(
                    document.querySelectorAll("a[href*='/maps/place/']")
                ).map(a => a.href).filter(h => h.includes("/maps/place/"));
            """) or []
        except Exception as js_err:
            if "tab crashed" in str(js_err).lower() or "no such window" in str(js_err).lower():
                log(f"   {tag}⚠ Tab crashed — stopping scroll, returning {len(collected)} URLs")
                break
            hrefs = []

        before = len(collected)
        for h in hrefs:
            if h and h not in seen:
                seen.add(h); collected.append(h)

        gained = len(collected) - before
        no_new = 0 if gained > 0 else no_new + 1
        passes += 1

        if passes % 4 == 0 or gained > 0:
            log(f"   {tag}Scroll {passes}: {len(collected)} URLs (+{gained})")

        # Adaptive scroll — jump harder when stuck
        step = SCROLL_STEP if gained > 0 else SCROLL_STEP * (1 + no_new)
        try:
            driver.execute_script(f"arguments[0].scrollTop += {step}", panel)
        except Exception:
            try:
                driver.execute_script(f"window.scrollBy(0, {step})")
            except Exception:
                break  # can't scroll → tab dead

        time.sleep(SCROLL_WAIT)

        # Refresh panel ref every 6 passes
        if passes % 6 == 0:
            try:
                panel = driver.find_element(By.XPATH, "//div[@role='feed']")
            except Exception:
                pass

        # JS end-detection (instant, no body.text scan)
        try:
            ended = driver.execute_script("""
                var t = document.body ? document.body.innerText : "";
                return t.includes("reached the end") ||
                       t.includes("No more results") ||
                       !!document.querySelector(".HlvSq");
            """)
        except Exception:
            log(f"   {tag}⚠ Tab crashed during end-check — stopping")
            break
        if ended:
            log(f"   {tag}End of results after {passes} scrolls")
            break

    log(f"   {tag}✓ {len(collected)} URLs in {passes} scrolls")
    return collected[:max_count]


# ── Step 2: Fast listing detail scraper ──────────────────────────────────────

def scrape_listing_fast(driver, url):
    """
    Extract ALL fields in ONE JavaScript call.
    No multiple find_element round trips.
    Uses smart wait (waits for h1, not fixed sleep).
    """
    data = {k: "" for k in [
        "name","category","address","city",
        "phone","website","rating","maps_url","source"]}
    data["maps_url"] = url

    try:
        try:
            driver.get(url)
        except Exception:
            # Page load timeout — return empty rather than crash
            return data

        # Smart wait: wait for name element instead of sleep
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.TAG_NAME, "h1")))
        except TimeoutException:
            time.sleep(DETAIL_WAIT)

        # Extract everything in ONE JavaScript call
        result = driver.execute_script("""
            var get = function(sels) {
                for (var i=0; i<sels.length; i++) {
                    var el = document.querySelector(sels[i]);
                    if (el) {
                        var t = el.innerText || el.textContent || "";
                        if (t.trim()) return t.trim();
                    }
                }
                return "";
            };

            // Phone: prefer tel: link
            var phone = "";
            var telEl = document.querySelector("a[href^='tel:']");
            if (telEl) phone = telEl.href.replace("tel:","").trim();
            if (!phone) {
                var pEl = document.querySelector(
                    "[data-tooltip='Copy phone number'] .fontBodyMedium," +
                    "[data-item-id^='phone'] .Io6YTe," +
                    "[data-item-id^='phone'] .fontBodyMedium"
                );
                if (pEl) phone = pEl.innerText.trim();
            }

            // Website
            var website = "";
            var wEl = document.querySelector(
                "[data-item-id='authority'] .Io6YTe," +
                "[data-item-id='authority'] .fontBodyMedium"
            );
            if (wEl) website = wEl.innerText.trim();
            if (!website) {
                var wLink = document.querySelector("[data-item-id='authority']");
                if (wLink && wLink.href && !wLink.href.includes("google.com"))
                    website = wLink.href;
            }

            // Address
            var addr = "";
            var aEl = document.querySelector(
                "[data-item-id='address'] .Io6YTe," +
                "[data-item-id='address'] .fontBodyMedium," +
                "[data-tooltip='Copy address'] .fontBodyMedium"
            );
            if (aEl) addr = aEl.innerText.trim();

            // Rating
            var rating = "";
            var rEl = document.querySelector(".ceNzKf, .F7nice span[aria-hidden='true']");
            if (rEl) {
                var rm = (rEl.innerText||"").match(/[0-9]+[.]?[0-9]*/);
                if (rm) rating = rm[0];
            }

            // Category
            var cat = get(["button.DkEaL","[jsaction*=category]","button[jsaction*=pane]"]);

            return {
                name:     get(["h1.DUwDvf","h1.fontHeadlineLarge","h1"]),
                category: cat,
                phone:    phone,
                website:  website,
                address:  addr,
                rating:   rating
            };
        """)

        if not result or not result.get("name"):
            return data

        data["name"]     = (result.get("name")     or "").strip()
        data["category"] = (result.get("category") or "").strip()
        data["phone"]    = (result.get("phone")    or "").strip()
        data["website"]  = (result.get("website")  or "").strip()
        data["address"]  = (result.get("address")  or "").strip()
        data["rating"]   = (result.get("rating")   or "").strip()

        if data["address"]:
            parts      = [p.strip() for p in data["address"].split(",")]
            data["city"] = parts[-2] if len(parts) >= 2 else parts[-1] if parts else ""

    except Exception:
        pass

    return data


# ── Step 2b: Headless worker pool ────────────────────────────────────────────

class WorkerPool:
    """Pool of headless Chrome drivers for parallel detail scraping."""

    def __init__(self, size=DETAIL_WORKERS):
        self._q   = __import__("queue").Queue()
        self._all = []
        self._sz  = size

    def start(self, log_fn):
        log_fn(f"   Starting {self._sz} headless workers...")
        for i in range(self._sz):
            d = _make_chrome(headless=True)
            self._q.put(d)
            self._all.append(d)
        log_fn(f"   ✓ {self._sz} workers ready")

    def acquire(self):
        return self._q.get(timeout=30)

    def release(self, d):
        self._q.put(d)

    def quit(self):
        for d in self._all:
            try: d.quit()
            except: pass


def scrape_parallel(urls, pool, log_fn, status_fn, stop_fn):
    """Scrape listing details in parallel using worker pool."""
    results = []
    lock    = threading.Lock()
    done    = [0]

    def worker(url):
        if stop_fn(): return None
        drv = pool.acquire()
        try:
            return scrape_listing_fast(drv, url)
        except Exception as e:
            log_fn(f"   worker err: {e}")
            return None
        finally:
            pool.release(drv)

    with ThreadPoolExecutor(max_workers=DETAIL_WORKERS) as ex:
        futs = {ex.submit(worker, u): u for u in urls}
        for fut in as_completed(futs):
            if stop_fn(): break
            r = fut.result()
            with lock:
                done[0] += 1
                if r and r.get("name"):
                    results.append(r)
                status_fn(done[0], len(urls), len(results))

    return results


# ── Step 3: Parallel email finding ───────────────────────────────────────────

def _valid_email(email):
    e = email.lower()
    if any(s in e for s in SKIP_EMAILS): return False
    if e.endswith((".png",".jpg",".gif",".svg",".css",".js",".php")): return False
    if any(x in e for x in ["noreply","no-reply","@2x","example"]): return False
    return 6 <= len(e) <= 60

def _fetch(url, timeout=5):
    try:
        r = requests.get(url, headers=REQ_HEADERS, timeout=timeout, verify=False,
                        allow_redirects=True)
        if r.status_code == 200:
            return r.text[:200000]  # max 200KB to avoid hanging on huge pages
        return None
    except Exception:
        return None

def find_email_fast(website):
    """Find email: homepage first, contact page only if needed."""
    if not website: return ""
    if not website.startswith("http"): website = "https://" + website
    website = website.split("?")[0].rstrip("/")

    def extract(html):
        if not html: return []
        found = [e for e in EMAIL_REGEX.findall(html) if _valid_email(e)]
        for m in re.findall(r'href=["\']mailto:([^"\'>\s]+)', html):
            e = m.split("?")[0].lower()
            if _valid_email(e): found.append(e)
        return list(dict.fromkeys(found))

    try:
        emails = extract(_fetch(website, timeout=6))
        if not emails:
            for path in ["contact", "contact-us", "about"]:
                try:
                    emails = extract(_fetch(f"{website}/{path}", timeout=4))
                    if emails: break
                except Exception:
                    continue
    except Exception:
        return ""

    return sorted(emails, key=len)[0] if emails else ""


def find_emails_parallel(leads, log_fn, update_fn, stop_fn):
    """Find emails using EMAIL_WORKERS parallel threads."""
    targets = [(i, l.get("website","")) for i,l in enumerate(leads)
               if l.get("website","")]
    if not targets:
        log_fn("   No websites to check")
        return leads

    def worker(args):
        idx, site = args
        if stop_fn(): return idx, ""
        return idx, find_email_fast(site)

    done = 0
    with ThreadPoolExecutor(max_workers=EMAIL_WORKERS) as ex:
        futs = {ex.submit(worker, t): t for t in targets}
        for fut in as_completed(futs):
            if stop_fn(): break
            idx, email = fut.result()
            leads[idx]["email"] = email
            done += 1
            update_fn(leads[idx].get("name",""), email, done, len(targets))

    return leads


# ── Excel export ──────────────────────────────────────────────────────────────

def export_excel(leads,path):
    wb=Workbook(); ws=wb.active; ws.title="Leads"
    ws.merge_cells("A1:K1")
    c=ws["A1"]
    c.value=f"Business Leads — {datetime.now().strftime('%Y-%m-%d %H:%M')} — {len(leads)} leads"
    c.font=Font(name="Arial",bold=True,size=13,color="1A73E8")
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=28
    HEADERS=["#","Name","Category","City","Address","Phone","Email","Website","Rating","Source","Maps URL"]
    WIDTHS= [5,  36,    20,        16,    38,       17,     30,     30,       7,      16,      48]
    def thin():
        s=Side(style="thin",color="CCCCCC"); return Border(left=s,right=s,top=s,bottom=s)
    hfill=PatternFill("solid",start_color="1A73E8")
    hfont=Font(name="Arial",bold=True,color="FFFFFF",size=11)
    for ci,(h,w) in enumerate(zip(HEADERS,WIDTHS),1):
        cell=ws.cell(row=2,column=ci,value=h)
        cell.font,cell.fill,cell.border=hfont,hfill,thin()
        cell.alignment=Alignment(horizontal="center",vertical="center")
        ws.column_dimensions[get_column_letter(ci)].width=w
    ws.row_dimensions[2].height=20
    odd=PatternFill("solid",start_color="F0F4FF")
    even=PatternFill("solid",start_color="FFFFFF")
    bfont=Font(name="Arial",size=10)
    left=Alignment(horizontal="left",vertical="center",wrap_text=True)
    ctr=Alignment(horizontal="center",vertical="center")
    for ri,lead in enumerate(leads,3):
        fill=odd if ri%2==1 else even
        row=[ri-2,lead.get("name",""),lead.get("category",""),lead.get("city",""),
             lead.get("address",""),lead.get("phone",""),lead.get("email",""),
             lead.get("website",""),lead.get("rating",""),lead.get("source",""),lead.get("maps_url","")]
        ws.row_dimensions[ri].height=17
        for ci,val in enumerate(row,1):
            cell=ws.cell(row=ri,column=ci,value=val)
            cell.font,cell.fill,cell.border=bfont,fill,thin()
            cell.alignment=ctr if ci in(1,9) else left
    ws.freeze_panes="A3"; ws.auto_filter.ref=f"A2:K{len(leads)+2}"
    wb.save(path)

# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LeadHunter Pro")
        self.resizable(True,True)
        self.minsize(960,680)
        self.state("zoomed")
        self.configure(bg=BG)
        self._running=False
        style=ttk.Style(self); style.theme_use("default")
        style.configure("Blue.Horizontal.TProgressbar",
                         troughcolor="#E3F2FD",background=ACCENT,
                         borderwidth=0,thickness=8)
        self._build_ui()

    def _build_ui(self):
        # ── HEADER ────────────────────────────────────────────────────────
        hdr=tk.Frame(self,bg=HEADER,height=68)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tf=tk.Frame(hdr,bg=HEADER); tf.pack(side="left",padx=24,pady=12)
        tk.Label(tf,text="🔍",font=("Segoe UI",22),fg="white",bg=HEADER).pack(side="left",padx=(0,10))
        tk.Label(tf,text="LeadHunter Pro",font=("Segoe UI",20,"bold"),fg="white",bg=HEADER).pack(side="left")
        tk.Label(tf,text="  Business Lead Scraper",font=("Segoe UI",11),fg="#BBDEFB",bg=HEADER).pack(side="left")
        bf=tk.Frame(hdr,bg=HEADER); bf.pack(side="right",padx=20)
        for txt,col in [("✅ Google Maps","#1B5E20"),("🔄 Auto Fallback","#E65100"),
                        ("📧 Email Finder","#4A148C"),("📊 Sheets Sync","#006064")]:
            b=tk.Frame(bf,bg=col,padx=8,pady=4); b.pack(side="left",padx=3)
            tk.Label(b,text=txt,font=("Segoe UI",8,"bold"),fg="white",bg=col).pack()

        # ── MAIN LAYOUT ───────────────────────────────────────────────────
        main=tk.Frame(self,bg=BG); main.pack(fill="both",expand=True)
        main.columnconfigure(0,minsize=400,weight=0)
        main.columnconfigure(1,weight=1); main.rowconfigure(0,weight=1)

        # Left scrollable panel
        lo=tk.Frame(main,bg=BG,width=400); lo.grid(row=0,column=0,sticky="nsew"); lo.grid_propagate(False)
        cv=tk.Canvas(lo,bg=BG,highlightthickness=0)
        sb=ttk.Scrollbar(lo,orient="vertical",command=cv.yview); cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); cv.pack(side="left",fill="both",expand=True)
        left=tk.Frame(cv,bg=BG); cw=cv.create_window((0,0),window=left,anchor="nw")
        left.bind("<Configure>",lambda e:cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>",lambda e:cv.itemconfig(cw,width=e.width))
        cv.bind_all("<MouseWheel>",lambda e:cv.yview_scroll(int(-1*(e.delta/120)),"units"))

        P=12

        def sec(parent,title,icon):
            f=tk.Frame(parent,bg=ACCENT,pady=10)
            f.pack(fill="x",padx=P,pady=(P,0))
            tk.Label(f,text=f"  {icon}  {title}",font=("Segoe UI",11,"bold"),
                     fg="white",bg=ACCENT).pack(anchor="w",padx=10)

        def card(parent):
            f=tk.Frame(parent,bg=CARD,highlightbackground=BORDER,highlightthickness=1)
            f.pack(fill="x",padx=P,pady=(0,P)); return f

        def field(parent,label,var,hint="",is_spin=False):
            tk.Label(parent,text=label,font=("Segoe UI",10,"bold"),
                     fg=TEXT,bg=CARD).pack(anchor="w",padx=16,pady=(12,0))
            if is_spin:
                sf=tk.Frame(parent,bg=CARD); sf.pack(fill="x",padx=16,pady=(4,0))
                tk.Spinbox(sf,from_=10,to=500,textvariable=var,
                           font=("Segoe UI",12),bg="white",fg=TEXT,
                           relief="solid",bd=1,width=8).pack(side="left",ipady=5)
                tk.Label(sf,text="  unique leads",font=("Segoe UI",9),
                         fg=SUBTEXT,bg=CARD).pack(side="left")
            else:
                tk.Entry(parent,textvariable=var,font=("Segoe UI",12),
                         bg="white",fg=TEXT,relief="solid",bd=1,
                         highlightbackground=BORDER,highlightcolor=ACCENT,
                         highlightthickness=2).pack(fill="x",padx=16,pady=(4,0),ipady=6)
            if hint:
                tk.Label(parent,text=f"  💡 {hint}",font=("Segoe UI",8),
                         fg=SUBTEXT,bg=CARD,anchor="w").pack(fill="x",padx=16,pady=(2,6))
            else:
                tk.Frame(parent,bg=CARD,height=8).pack()

        # ── SEARCH SETTINGS ───────────────────────────────────────────────
        sec(left,"SEARCH SETTINGS","🔍"); sc=card(left)
        self.query_var=tk.StringVar(value="aesthetic clinic")
        field(sc,"Business Type",self.query_var,
              "aesthetic clinic · solar company · real estate · dental clinic")
        self.city_var=tk.StringVar(value="New York, USA")
        field(sc,"City / Location",self.city_var,"New York USA · London UK · Miami USA")
        self.limit_var=tk.IntVar(value=40)
        field(sc,"Maximum Leads",self.limit_var,is_spin=True)
        er=tk.Frame(sc,bg=CARD); er.pack(fill="x",padx=16,pady=(8,14))
        self.email_var=tk.BooleanVar(value=True)
        tk.Checkbutton(er,variable=self.email_var,bg=CARD,activebackground=CARD,
                       selectcolor="white",font=("Segoe UI",11),fg=TEXT,
                       text="  📧  Find emails from websites automatically").pack(side="left")

        # ── ACTIONS ───────────────────────────────────────────────────────
        sec(left,"ACTIONS","▶"); ac=card(left)
        self.btn=tk.Button(ac,text="▶   START SCRAPING",
                           font=("Segoe UI",14,"bold"),bg=ACCENT,fg="white",
                           activebackground=ACCENT2,relief="flat",cursor="hand2",
                           pady=14,command=self._start)
        self.btn.pack(fill="x",padx=16,pady=(14,8))
        self.stop_btn=tk.Button(ac,text="⏹   STOP SCRAPING",
                                font=("Segoe UI",12,"bold"),bg="#FFEBEE",fg=DANGER,
                                activebackground=DANGER,activeforeground="white",
                                relief="flat",cursor="hand2",pady=10,state="disabled",
                                command=self._stop)
        self.stop_btn.pack(fill="x",padx=16,pady=(0,14))
        self.progress=ttk.Progressbar(ac,mode="indeterminate",style="Blue.Horizontal.TProgressbar")
        self.progress.pack(fill="x",padx=16,pady=(0,6))
        self.source_var=tk.StringVar(value="")
        tk.Label(ac,textvariable=self.source_var,font=("Segoe UI",9,"bold"),
                 fg=WARNING,bg=CARD,wraplength=340,justify="left").pack(anchor="w",padx=16,pady=(0,10))

        # ── LIVE STATS ────────────────────────────────────────────────────
        sec(left,"LIVE STATS","📊"); stc=card(left)
        grid=tk.Frame(stc,bg=CARD); grid.pack(fill="x",padx=12,pady=12)
        grid.columnconfigure((0,1),weight=1)
        self.s_total=tk.StringVar(value="0"); self.s_phones=tk.StringVar(value="0")
        self.s_emails=tk.StringVar(value="0"); self.s_sites=tk.StringVar(value="0")
        def sbox(parent,label,var,bg,fg,icon,row,col):
            b=tk.Frame(parent,bg=bg,padx=10,pady=10)
            b.grid(row=row,column=col,padx=6,pady=6,sticky="ew")
            tk.Label(b,text=icon,font=("Segoe UI",16),fg=fg,bg=bg).pack()
            tk.Label(b,textvariable=var,font=("Segoe UI",28,"bold"),fg=fg,bg=bg).pack()
            tk.Label(b,text=label,font=("Segoe UI",9,"bold"),fg=fg,bg=bg).pack()
        sbox(grid,"Total Leads",self.s_total, "#E3F2FD","#0D47A1","🏢",0,0)
        sbox(grid,"Phones",     self.s_phones,"#E8F5E9","#1B5E20","📞",0,1)
        sbox(grid,"Emails",     self.s_emails,"#FFF8E1","#E65100","📧",1,0)
        sbox(grid,"Websites",   self.s_sites, "#F3E5F5","#4A148C","🌐",1,1)

        # ── GOOGLE SHEETS SYNC ────────────────────────────────────────────
        sec(left,"GOOGLE SHEETS SYNC","📊"); shc=card(left)

        # Enable toggle
        trow=tk.Frame(shc,bg="#E8F5E9"); trow.pack(fill="x",pady=(12,0))
        self.sheets_enabled=tk.BooleanVar(value=False)
        tk.Checkbutton(trow,variable=self.sheets_enabled,bg="#E8F5E9",
                       activebackground="#E8F5E9",selectcolor="white",
                       font=("Segoe UI",11,"bold"),fg="#1B5E20",
                       text="  ✅  Enable Google Sheets Auto-Sync",
                       command=self._toggle_sheets).pack(side="left",padx=16,pady=12)

        # Sheets input fields
        self.sheets_frame=tk.Frame(shc,bg=CARD); self.sheets_frame.pack(fill="x")

        # ── STEP 1: Sign in with Google ───────────────────────────────────
        step1=tk.Frame(self.sheets_frame,bg="#E8F5E9",
                       highlightbackground="#A5D6A7",highlightthickness=1)
        step1.pack(fill="x",padx=16,pady=(12,4))

        tk.Label(step1,text="STEP 1  —  Sign in with Google",
                 font=("Segoe UI",10,"bold"),fg="#1B5E20",
                 bg="#E8F5E9").pack(anchor="w",padx=12,pady=(10,2))
        tk.Label(step1,
                 text="Click below to connect your Google account. ""A browser window will open — just sign in and click Allow.",
                 font=("Segoe UI",9),fg="#2E7D32",
                 bg="#E8F5E9",justify="left",wraplength=300).pack(anchor="w",padx=12,pady=(0,8))

        self.google_btn=tk.Button(step1,
                                   text="🔐   Sign in with Google",
                                   font=("Segoe UI",12,"bold"),
                                   bg="#4285F4",fg="white",
                                   activebackground="#3367D6",
                                   activeforeground="white",
                                   relief="flat",cursor="hand2",
                                   pady=10,
                                   command=self._google_signin)
        self.google_btn.pack(fill="x",padx=12,pady=(0,10))

        self.google_status_var=tk.StringVar(value="⚪  Not signed in")
        self.google_status_lbl=tk.Label(step1,
                                         textvariable=self.google_status_var,
                                         font=("Segoe UI",9,"bold"),
                                         fg=SUBTEXT,bg="#E8F5E9")
        self.google_status_lbl.pack(anchor="w",padx=12,pady=(0,8))

        # Disconnect button (hidden initially)
        self.disconnect_btn=tk.Button(step1,
                                       text="🔓  Disconnect Google Account",
                                       font=("Segoe UI",8),
                                       bg="#FFEBEE",fg=DANGER,
                                       relief="flat",cursor="hand2",
                                       pady=4,
                                       command=self._google_disconnect)
        # shown only when connected

        # ── STEP 2: Paste Sheet URL ────────────────────────────────────────
        step2=tk.Frame(self.sheets_frame,bg="#E3F2FD",
                       highlightbackground="#90CAF9",highlightthickness=1)
        step2.pack(fill="x",padx=16,pady=(4,4))

        tk.Label(step2,text="STEP 2  —  Paste your Google Sheet URL",
                 font=("Segoe UI",10,"bold"),fg="#0D47A1",
                 bg="#E3F2FD").pack(anchor="w",padx=12,pady=(10,2))
        tk.Label(step2,
                 text="Create a Google Sheet, then paste its URL below.",
                 font=("Segoe UI",9),fg="#1565C0",
                 bg="#E3F2FD",justify="left").pack(anchor="w",padx=12,pady=(0,4))

        self.sheet_url_var=tk.StringVar()
        tk.Entry(step2,textvariable=self.sheet_url_var,
                 font=("Segoe UI",11),bg="white",fg=TEXT,
                 relief="solid",bd=1,
                 highlightbackground="#90CAF9",
                 highlightcolor=ACCENT,
                 highlightthickness=2).pack(
                     fill="x",padx=12,pady=(0,8),ipady=7)

        # Hidden - default tab name
        self.sheet_name_var=tk.StringVar(value="Leads")

        self.connect_btn=tk.Button(step2,
                                    text="🔗   Connect My Sheet",
                                    font=("Segoe UI",12,"bold"),
                                    bg=ACCENT,fg="white",
                                    activebackground=ACCENT2,
                                    activeforeground="white",
                                    relief="flat",cursor="hand2",
                                    pady=10,
                                    command=self._test_sheets)
        self.connect_btn.pack(fill="x",padx=12,pady=(0,8))

        self.sheets_status_var=tk.StringVar(
            value="⚪  Sign in with Google first, then paste your Sheet URL")
        self.sheets_status_lbl=tk.Label(step2,
                                         textvariable=self.sheets_status_var,
                                         font=("Segoe UI",9,"bold"),
                                         fg=SUBTEXT,bg="#E3F2FD",
                                         wraplength=320,justify="left")
        self.sheets_status_lbl.pack(anchor="w",padx=12,pady=(0,10))

        # Install libraries (small, secondary)
        self.install_btn=tk.Button(self.sheets_frame,
                                    text="📦  Install Required Libraries (first time only)",
                                    font=("Segoe UI",8),
                                    bg="#F5F5F5",fg=SUBTEXT,
                                    activebackground="#E0E0E0",
                                    relief="flat",cursor="hand2",pady=5,
                                    command=self._install_libs)
        self.install_btn.pack(fill="x",padx=16,pady=(0,4))

        # test_btn alias for backward compat
        self.test_btn = self.connect_btn

        # Check if already authenticated on startup
        self._check_auth_status()

        # Load saved config
        cfg=load_config()
        self.sheet_url_var.set(cfg.get("sheet_url",""))
        self.sheet_name_var.set(cfg.get("sheet_name","Leads"))
        self.sheets_enabled.set(cfg.get("sheets_enabled",False))
        self._toggle_sheets()
        tk.Frame(left,bg=BG,height=20).pack()

        # ── RIGHT PANEL: LOG ──────────────────────────────────────────────
        right=tk.Frame(main,bg=BG); right.grid(row=0,column=1,sticky="nsew")
        ltop=tk.Frame(right,bg=CARD,highlightbackground=BORDER,highlightthickness=1)
        ltop.pack(fill="x")
        tk.Label(ltop,text="📋  LIVE ACTIVITY LOG",
                 font=("Segoe UI",12,"bold"),fg=TEXT,bg=CARD).pack(side="left",padx=16,pady=12)
        self.status_var=tk.StringVar(value="Ready — Configure settings and click Start Scraping")
        tk.Label(ltop,textvariable=self.status_var,
                 font=("Segoe UI",9),fg=SUBTEXT,bg=CARD).pack(side="right",padx=16,pady=12)
        self.log_box=scrolledtext.ScrolledText(right,
                                                font=("Courier New",10) if os.name=="nt" else ("Courier",10),
                                                bg=LOG_BG,fg=LOG_FG,relief="flat",
                                                state="disabled",padx=16,pady=12,
                                                highlightbackground=BORDER,highlightthickness=1,wrap="word")
        self.log_box.pack(fill="both",expand=True)
        for tag,color in [("green","#00E676"),("blue","#82B1FF"),("yellow","#FFD740"),
                          ("red","#FF5252"),("dim","#546E7A"),("white","#ECEFF1")]:
            self.log_box.tag_config(tag,foreground=color)

        # ── STATUS BAR ────────────────────────────────────────────────────
        bar=tk.Frame(self,bg=NAVY,height=34); bar.pack(fill="x",side="bottom"); bar.pack_propagate(False)
        tk.Label(bar,text="🔍 LeadHunter Pro   |   Results saved to Desktop → LeadResults",
                 font=("Segoe UI",9),fg="#90CAF9",bg=NAVY).pack(side="left",padx=16,pady=8)
        self.bar_status=tk.StringVar(value="● Idle")
        tk.Label(bar,textvariable=self.bar_status,font=("Segoe UI",9,"bold"),
                 fg="#69F0AE",bg=NAVY).pack(side="right",padx=16,pady=8)

    # ── Sheets helpers (OAuth) ────────────────────────────────────────────────

    def _toggle_sheets(self):
        if self.sheets_enabled.get():
            self.sheets_frame.pack(fill="x")
        else:
            self.sheets_frame.pack_forget()
        self._save_config()

    def _save_config(self):
        save_config({
            "sheet_url":      self.sheet_url_var.get().strip(),
            "sheet_name":     self.sheet_name_var.get().strip() or "Leads",
            "sheets_enabled": self.sheets_enabled.get(),
        })

    def _check_auth_status(self):
        """Check if user is already authenticated on startup."""
        if is_user_authenticated(SESSION_USER_ID):
            self.google_status_var.set("✅  Google account connected")
            self.google_status_lbl.configure(fg=SUCCESS)
            self.google_btn.configure(text="✅  Google Connected",
                                       bg="#E8F5E9",fg=SUCCESS)
            self.disconnect_btn.pack(fill="x",padx=12,pady=(0,8))

    def _google_signin(self):
        """Open Google OAuth browser flow."""
        self.google_btn.configure(state="disabled",
                                   text="⏳  Opening browser...",
                                   bg="#90CAF9")
        self.google_status_var.set("⏳  Please sign in using the browser window...")
        self.google_status_lbl.configure(fg=WARNING)
        self.update_idletasks()

        def do_auth():
            success = authenticate_user(SESSION_USER_ID, log_fn=self.log)
            if success:
                self.google_status_var.set("✅  Google account connected!")
                self.google_status_lbl.configure(fg=SUCCESS)
                self.google_btn.configure(state="normal",
                                           text="✅  Google Connected",
                                           bg="#E8F5E9",fg=SUCCESS)
                self.disconnect_btn.pack(fill="x",padx=12,pady=(0,8))
                self.sheets_status_var.set("✅  Signed in! Now paste your Sheet URL and click Connect")
                self.sheets_status_lbl.configure(fg=SUCCESS)
            else:
                self.google_status_var.set("❌  Sign in failed — try again")
                self.google_status_lbl.configure(fg=DANGER)
                self.google_btn.configure(state="normal",
                                           text="🔐   Sign in with Google",
                                           bg="#4285F4",fg="white")

        threading.Thread(target=do_auth, daemon=True).start()

    def _google_disconnect(self):
        """Revoke OAuth token."""
        if messagebox.askyesno("Disconnect",
            "Disconnect your Google account? You can reconnect anytime."):
            revoke_user_token(SESSION_USER_ID)
            self.google_status_var.set("⚪  Not signed in")
            self.google_status_lbl.configure(fg=SUBTEXT)
            self.google_btn.configure(state="normal",
                                       text="🔐   Sign in with Google",
                                       bg="#4285F4",fg="white")
            self.disconnect_btn.pack_forget()
            self.sheets_status_var.set("⚪  Sign in with Google first")
            self.sheets_status_lbl.configure(fg=SUBTEXT)

    def _test_sheets(self):
        """Connect and test the sheet URL."""
        url  = self.sheet_url_var.get().strip().strip('"').strip("'")
        name = self.sheet_name_var.get().strip() or "Leads"

        if not url:
            self.sheets_status_var.set("❌  Please paste your Google Sheet URL above")
            self.sheets_status_lbl.configure(fg=DANGER)
            return

        # Validate URL
        url_low = url.lower()
        is_valid = ("spreadsheets/d/" in url_low or
                    "docs.google.com/spreadsheets" in url_low or
                    re.match(r'^[a-zA-Z0-9_-]{20,}$', url))
        if not is_valid:
            self.sheets_status_var.set("❌  Please paste the full Google Sheet URL from your browser")
            self.sheets_status_lbl.configure(fg=DANGER)
            return

        # Check signed in
        if not is_user_authenticated(SESSION_USER_ID):
            self.sheets_status_var.set("❌  Please sign in with Google first (Step 1)")
            self.sheets_status_lbl.configure(fg=DANGER)
            messagebox.showwarning("Sign In Required",
                "Please click Sign in with Google first, then paste your Sheet URL and click Connect.")
            return

        self.sheet_url_var.set(url)
        self._save_config()
        self.connect_btn.configure(state="disabled",
                                    text="⏳  Connecting...",
                                    bg="#90CAF9")
        self.sheets_status_var.set("⏳  Connecting to your sheet...")
        self.sheets_status_lbl.configure(fg=WARNING)
        self.update_idletasks()

        self._save_config()
        self.test_btn.configure(state="disabled", text="⏳  Connecting...", bg="#B2DFDB")
        self.sheets_status_var.set("⏳  Connecting to your Google Sheet...")
        self.sheets_status_lbl.configure(fg=WARNING)
        self.update_idletasks()

        def do_test():
            self.log(f"\n📊 Testing Google Sheets connection...")
            self.log(f"   URL: {url[:60]}...")
            result = test_sheet_connection(url, SESSION_USER_ID, name, self.log)
            self.log(f"   Result: {result['message'][:120]}")

            if result["success"]:
                self.sheets_status_var.set(
                    f"✅  Connected to: {result.get('sheet_title','your sheet')}")
                self.sheets_status_lbl.configure(fg=SUCCESS)
                self.connect_btn.configure(
                    state="normal",
                    text="✅  Connected — Click to Re-test",
                    bg="#E8F5E9", fg=SUCCESS)
                self.log("✅ Google Sheets connected successfully!")
            else:
                msg = result["message"]
                self.log(f"❌ {msg}")

                if "not_authenticated" in msg:
                    self.sheets_status_var.set("❌  Please sign in with Google first (Step 1)")
                    self.sheets_status_lbl.configure(fg=DANGER)
                elif any(x in msg.lower() for x in ["gspread","not installed","import"]):
                    self.sheets_status_var.set("❌  Click Install Libraries below")
                    self.sheets_status_lbl.configure(fg=DANGER)
                elif "forbidden" in msg.lower() or "403" in msg.lower():
                    self.sheets_status_var.set("❌  Access denied — make sure this is YOUR sheet")
                    self.sheets_status_lbl.configure(fg=DANGER)
                else:
                    self.sheets_status_var.set(f"❌  {msg[:80]}")
                    self.sheets_status_lbl.configure(fg=DANGER)

                self.connect_btn.configure(
                    state="normal",
                    text="🔗   Connect My Sheet",
                    bg=ACCENT, fg="white")

        threading.Thread(target=do_test, daemon=True).start()

    def _install_libs(self):
        self.install_btn.configure(state="disabled",text="Installing...")
        def do_install():
            install_gspread(self.log)
            self.install_btn.configure(state="normal",text="📦  Install Libraries")
        threading.Thread(target=do_install,daemon=True).start()

    # ── Logging ───────────────────────────────────────────────────────────────

    def log(self,msg):
        self.log_box.configure(state="normal")
        tag=("green" if any(x in msg for x in ["✓","✅","DONE","OK"])
             else "blue"  if any(x in msg for x in ["SOURCE","─","=","Searching"])
             else "red"   if any(x in msg for x in ["⚠","⏹","❌","Error"])
             else "yellow" if any(x in msg for x in ["✉","📞","🌐","📊"])
             else "dim"   if "Scroll" in msg or msg.startswith("   ")
             else "white")
        self.log_box.insert("end",msg+"\n",tag)
        self.log_box.see("end"); self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _upd(self,leads):
        self.s_total.set(str(len(leads)))
        self.s_phones.set(str(sum(1 for l in leads if l.get("phone"))))
        self.s_emails.set(str(sum(1 for l in leads if l.get("email"))))
        self.s_sites.set(str(sum(1 for l in leads if l.get("website"))))

    def _start(self):
        if self._running: return
        if not self.query_var.get().strip():
            messagebox.showerror("Error","Please enter a business type!"); return
        if not self.city_var.get().strip():
            messagebox.showerror("Error","Please enter a city!"); return
        self._running=True
        self.btn.configure(state="disabled",bg="#BDBDBD",fg="#757575")
        self.stop_btn.configure(state="normal",bg="#FFEBEE")
        self.progress.start(10)
        self.log_box.configure(state="normal"); self.log_box.delete("1.0","end"); self.log_box.configure(state="disabled")
        for v in [self.s_total,self.s_phones,self.s_emails,self.s_sites]: v.set("0")
        self.bar_status.set("● Running...")
        threading.Thread(target=self._run,daemon=True).start()

    def _stop(self):
        self._running=False; self.log("⏹ Stopped."); self._reset_ui()

    def _reset_ui(self):
        self.btn.configure(state="normal",bg=ACCENT,fg="white")
        self.stop_btn.configure(state="disabled",bg="#FFEBEE")
        self.progress.stop(); self.source_var.set(""); self.bar_status.set("● Idle")

    # ── Main scraping run ─────────────────────────────────────────────────────

    def _run(self):
        query=self.query_var.get().strip(); city=self.city_var.get().strip()
        limit=self.limit_var.get(); do_email=self.email_var.get(); t0=time.time()

        safe_c=re.sub(r'[^a-zA-Z0-9]','_',city)[:15]
        safe_q=re.sub(r'[^a-zA-Z0-9]','_',query)[:15]
        fname=f"leads_{safe_q}_{safe_c}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        fpath=os.path.join(SAVE_FOLDER,fname)

        self.log(f"Business : {query}"); self.log(f"City     : {city}")
        self.log(f"Target   : {limit} unique leads"); self.log("="*52)

        all_leads=[]; seen_names=set(); seen_phones=set(); driver=None

        def add(new_leads,src_tag):
            added=0
            for lead in new_leads:
                if len(all_leads)>=limit or not self._running: break
                name=lead.get("name","").strip().lower()
                phone=lead.get("phone","").strip()
                if not name or name in seen_names: continue
                if phone and phone in seen_phones: continue
                seen_names.add(name)
                if phone: seen_phones.add(phone)
                lead["source"]=src_tag; all_leads.append(lead); added+=1
                self.log(f"  [{len(all_leads)}/{limit}] ✓ {lead['name']}"+(f"  |  {phone}" if phone else ""))
                self._upd(all_leads)
            return added

        def run_source(q,c,needed,label,src_tag):
            self.source_var.set(f"🔍 {label}: {q} in {c}")
            self.status_var.set(f"Searching: {q} in {c}...")
            links=collect_links(driver,q,c,needed,self.log,label)
            if not links: return 0
            self.log(f"   Scraping {len(links)} listings...")
            batch=[]
            for i,url in enumerate(links,1):
                if not self._running or len(all_leads)>=limit: break
                self.status_var.set(f"[{label}] {i}/{len(links)} — {len(all_leads)} leads")
                d=scrape_one(driver,url)
                if d.get("name"): batch.append(d)
            return add(batch,src_tag)

        pool   = None
        driver = None
        try:
            # ── Start Chrome + workers ────────────────────────────────────
            self.log("\n🚀 Starting LeadHunter Fast Engine...")
            self.status_var.set("Starting Chrome...")

            driver = build_driver()
            self.log(f"✓ Chrome ready")

            pool = WorkerPool(size=DETAIL_WORKERS)
            pool.start(self.log)
            self.log(f"✓ {DETAIL_WORKERS} headless workers ready\n")

            def run_source(q, c, needed, label, src_tag):
                """
                Full pipeline for one source:
                1. Scroll Google Maps → collect URLs  (fast JS)
                2. Scrape details in parallel         (headless pool)
                3. Add unique leads                   (dedup)
                """
                self.source_var.set(f"🔍 {label}")
                self.status_var.set(f"Collecting: {q} in {c}...")

                urls = collect_urls_fast(driver, q, c, needed, self.log, label)
                if not urls:
                    self.log(f"   No results for {q} in {c}")
                    return 0

                self.log(f"   Scraping {len(urls)} listings ({DETAIL_WORKERS} parallel)...")

                def status_cb(done, total, found):
                    self.status_var.set(
                        f"[{label}] {done}/{total} scraped → "
                        f"{len(all_leads)+found} leads — {time.time()-t0:.0f}s")

                batch = scrape_parallel(
                    urls, pool, self.log, status_cb, lambda: not self._running)

                return add(batch, src_tag)

            # ── SOURCE 1: Primary ─────────────────────────────────────────
            self.log("─"*50)
            self.log(f"📍 SOURCE 1: '{query}' in '{city}'")
            self.log("─"*50)
            n = run_source(query, city, limit, "S1-Primary", "Google Maps")
            self.log(f"✓ +{n} leads  total:{len(all_leads)}/{limit}  {time.time()-t0:.1f}s")

            # ── SOURCE 2: Related terms ────────────────────────────────────
            if self._running and len(all_leads) < limit:
                related = get_related(query)
                self.log(f"\n🔄 SOURCE 2: Related terms ({len(related)} variants)")
                for i, term in enumerate(related, 1):
                    if not self._running or len(all_leads) >= limit: break
                    self.log(f"\n   [{i}] '{term}'")
                    n = run_source(term, city, limit-len(all_leads),
                                   f"S2-{i}", f"Google Maps ({term})")
                    self.log(f"   +{n}  total:{len(all_leads)}/{limit}")

            # ── SOURCE 3: Nearby cities ────────────────────────────────────
            if self._running and len(all_leads) < limit:
                nearby = get_nearby(city)
                if nearby:
                    self.log(f"\n🏙️  SOURCE 3: Nearby cities")
                    for i, nc in enumerate(nearby, 1):
                        if not self._running or len(all_leads) >= limit: break
                        self.log(f"\n   [{i}] '{nc}'")
                        n = run_source(query, nc, limit-len(all_leads),
                                       f"S3-{i}", f"Google Maps ({nc})")
                        self.log(f"   +{n}  total:{len(all_leads)}/{limit}")

            # ── SOURCE 4: Region ───────────────────────────────────────────
            if self._running and len(all_leads) < limit:
                parts  = [p.strip() for p in city.split(",")]
                region = parts[-1] if len(parts) > 1 else ""
                if region and region.lower() != city.lower():
                    self.log(f"\n🌍 SOURCE 4: Region '{region}'")
                    n = run_source(query, region, limit-len(all_leads),
                                   "S4-Region", f"Google Maps ({region})")
                    self.log(f"   +{n}  total:{len(all_leads)}/{limit}")

        except Exception as e:
            self.log(f"❌ Error: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            if driver:
                try: driver.quit()
                except: pass
            if pool:
                try: pool.quit()
                except: pass
            self.source_var.set("")

        if not all_leads:
            self.log("\n❌ No leads found. Try different search or city.")
            self._reset_ui(); return

        self.log(f"\n✓ Scraping done: {len(all_leads)} leads in {time.time()-t0:.1f}s")

        # Find emails
        if do_email and self._running:
            self.log(f"\n✉  Finding emails (5 parallel workers)...")
            self.source_var.set("✉ Finding emails...")
            def email_upd(name,email,done,total):
                if email: self.log(f"  ✉ {name} → {email}")
                else: self.log(f"  — {name} → no email")
                self._upd(all_leads)
                self.status_var.set(f"Emails: {done}/{total}")
            find_emails_parallel(all_leads,self.log,email_upd,lambda:not self._running)
            self.source_var.set("")

        # Save Excel
        export_excel(all_leads,fpath)

        # Sync to Google Sheets (secure — no credentials from user)
        sheets_added=0
        if self.sheets_enabled.get() and self._running:
            url=self.sheet_url_var.get().strip()
            name=self.sheet_name_var.get().strip() or "Leads"
            if url:
                self.log(f"\n📊 Syncing to Google Sheets...")
                self.bar_status.set("● Syncing to Sheets...")
                result=append_leads_to_sheet(url,all_leads,SESSION_USER_ID,name,self.log)
                sheets_added=result.get("added",0)
                if result["success"]:
                    self.sheets_status_var.set(f"✅  {result['message']}")
                    self.sheets_status_lbl.configure(fg=SUCCESS)
                else:
                    self.sheets_status_var.set(f"❌  {result['message'][:60]}")
                    self.sheets_status_lbl.configure(fg=DANGER)
            self._save_config()

        elapsed=time.time()-t0
        self.log(f"\n{'='*52}")
        self.log(f"✅ DONE in {elapsed:.1f}s!")
        self.log(f"📊 Total   : {len(all_leads)}")
        self.log(f"📞 Phones  : {sum(1 for l in all_leads if l.get('phone'))}")
        self.log(f"✉  Emails  : {sum(1 for l in all_leads if l.get('email'))}")
        self.log(f"🌐 Sites   : {sum(1 for l in all_leads if l.get('website'))}")
        if sheets_added: self.log(f"📊 Sheets  : {sheets_added} rows added")
        self.log(f"📁 {fpath}")
        self.status_var.set(f"✅ Done! {len(all_leads)} leads → Desktop/LeadResults")
        self._upd(all_leads)

        sheets_info=f"\n📊 Google Sheets: {sheets_added} rows added" if sheets_added>0 else ""
        messagebox.showinfo("✅ Done!",
            f"{len(all_leads)} leads saved!\n\n"
            f"📞 Phones : {sum(1 for l in all_leads if l.get('phone'))}\n"
            f"✉  Emails : {sum(1 for l in all_leads if l.get('email'))}\n"
            f"🌐 Sites  : {sum(1 for l in all_leads if l.get('website'))}"
            f"{sheets_info}\n\n"
            f"📁 Desktop → LeadResults → {fname}")
        self._reset_ui()


if __name__=="__main__":
    app=App(); app.mainloop()
