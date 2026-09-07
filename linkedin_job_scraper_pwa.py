"""
LinkedIn Job Scraper — Advanced & Robust Filtering Edition
==========================================================
Searches LinkedIn Jobs with granular URL-level filters and deep text-level
relevance filtering (keyword scoring, negative keyword exclusions, compound matching).
Drives Microsoft Edge with a stealth profile so your normal browser is untouched.

KEY IMPROVEMENTS & FILTERING ROBUSTNESS:
  1. Granular URL Filters:
     - Date Posted (Today / Past 24h / Past Week / Past Month / Any)
     - Sort Order (Most Recent / Most Relevant) — avoids stale jobs
     - Experience Levels (Internship, Entry, Associate, Mid-Senior, etc.)
     - Job Types (Internship, Full-time, Part-time, Contract, Temporary)
     - Workplace Types (On-site, Hybrid, Remote)
     - Exact GeoIDs (Germany, DACH, Switzerland, Austria, UK, US, etc.)
  2. Multi-Stage Intelligent Text Filtering:
     - Expand description ("See more" / "Mehr anzeigen") before evaluation.
     - Title Exclusion / Blacklist (e.g., Senior Manager, Sales, DevOps, Accounting).
     - Company Blacklist (optional).
     - Smart German/English compound matching (e.g., FEM-Berechnung, Crashsimulation).
     - Context-aware validation for ambiguous keywords (e.g., META, ANSA).
     - Configurable minimum keyword match threshold.
  3. Automatic Data Export:
     - Real-time deduplicated export to both CSV and formatted Excel (.xlsx).
     - Stores Job ID, Title, Company, Location, Workplace Type, Posted Date,
       Match Score, Matched Keywords, Key Context Lines, URL, and Timestamp.
  4. SDUI / Dynamic LinkedIn Layout Resilience:
     - Multi-layer fallback selectors for modern obfuscated CSS class layouts.
     - Human-like stealth timing, mouse jitter, and automated session recovery.

REQUIREMENTS:
  pip install selenium webdriver-manager pyperclip pandas openpyxl
"""

import os
import sys
import re
import time
import json
import random
import logging
import html
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    NoSuchWindowException,
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    WebDriverException,
    SessionNotCreatedException,
)

try:
    from urllib3.exceptions import ReadTimeoutError, MaxRetryError
except ImportError:
    ReadTimeoutError = type("ReadTimeoutError", (Exception,), {})
    MaxRetryError    = type("MaxRetryError",    (Exception,), {})

try:
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
except ImportError:
    EdgeChromiumDriverManager = None


# ===========================================================================
# 1. USER CONFIGURATION — TWEAK YOUR FILTERS HERE
# ===========================================================================

LINKEDIN_EMAIL    = ""   # Leave blank to log in manually on first run
LINKEDIN_PASSWORD = ""

# --- LOCATION FILTERS (EUROPE / EU / DACH / NORDIC SPECIALIZED) ---
# Choose a region, preset group, or country from the European options below.
# Examples:
#   "Germany" | "DACH" | "European Union" | "Europe" | "Nordics"
#   "Sweden" | "Denmark" | "Norway" | "Finland" | "Netherlands" | "Austria" | "Switzerland"
# Or provide a list to cycle through:
#   LOCATION_FILTER = ["Germany", "Austria", "Switzerland"]  # or "DACH" or "Nordics"
LOCATION_FILTER = "Germany"

# Accurate LinkedIn GeoIDs for European countries and economic zones:
GEO_ID_MAP = {
    # Pan-European & Regional Blocs
    "europe": "100506914",
    "european union": "105117694",
    "eu": "105117694",
    "dach": "101282230",             # Centers DACH on central European hub (Germany/DACH)
    "nordics": "102890719",          # Nordic / Scandinavian region
    "scandinavia": "102890719",
    "benelux": "102890719",

    # DACH & Central Europe
    "germany": "101282230",
    "deutschland": "101282230",
    "austria": "103883259",
    "österreich": "103883259",
    "switzerland": "106693273",
    "schweiz": "106693273",
    "liechtenstein": "105080838",

    # Nordics (Northern Europe)
    "sweden": "105117694",
    "sverige": "105117694",
    "denmark": "104514075",
    "danmark": "104514075",
    "norway": "103819153",
    "norge": "103819153",
    "finland": "100456013",
    "suomi": "100456013",
    "iceland": "100665061",

    # Western Europe
    "netherlands": "102890719",
    "nederland": "102890719",
    "belgium": "103291313",
    "luxembourg": "104035573",
    "france": "105015875",
    "united kingdom": "102257491",
    "uk": "102257491",
    "ireland": "104738515",

    # Southern Europe
    "italy": "103350119",
    "spain": "105646813",
    "portugal": "100364837",

    # Eastern & Central-Eastern Europe (EU members)
    "poland": "105072130",
    "czech republic": "104817264",
    "czechia": "104817264",
    "hungary": "100288700",
    "slovakia": "106155005",
    "slovenia": "106137682",
    "romania": "106670623",
}
USE_GEO_ID = True  # Automatically attach geoId if found in GEO_ID_MAP
# Options: "today" (or "24h") | "week" | "month" | "any"
DATE_FILTER = "today"

# Sort by: "recent" (Most Recent — RECOMMENDED to get newly posted jobs) | "relevance"
SORT_BY = "recent"

# --- AI SEARCH MODE ---
# When True: uses LinkedIn's AI natural-language search bar (types a full-sentence
# query and waits for the AI suggestions / results to load). This is more semantic
# than keyword-URL matching and picks up synonym matches LinkedIn's LLM understands.
# When False: falls back to direct URL-based keyword search (traditional mode).
USE_AI_SEARCH = True

# Natural language queries used in AI search mode. LinkedIn's AI interprets these
# as intent-based prompts and returns semantically-matched results across European hubs.
AI_SEARCH_QUERIES = [
    "Master thesis or internship in crash simulation FEM CAE LS-DYNA in Europe",
    "Masterarbeit Abschlussarbeit Struktursimulation FEM Crashsimulation Deutschland DACH",
    "thesis internship structural mechanics finite element vehicle safety automotive Europe",
    "Werkstudent Praktikum Simulation CAE Fahrzeugsicherheit Berechnung DACH",
    "internship master thesis computational mechanics ANSYS Abaqus digital twin Europe",
]

# --- EXTRA BOOLEAN URL FILTERS ---
# Only show jobs from companies flagged "Actively Hiring" on LinkedIn
ACTIVELY_HIRING_ONLY = True

# Only show "Easy Apply" jobs (LinkedIn-hosted application, no external redirect)
EASY_APPLY_ONLY = False

# --- EXPERIENCE LEVEL FILTER ---
# Options: "internship", "entry", "associate", "mid_senior", "director", "executive"
# Set to None or [] to search all levels.
EXPERIENCE_LEVEL = ["internship", "entry"]

# --- JOB TYPE FILTER (Employment Type) ---
# Options: "internship", "full_time", "part_time", "contract", "temporary"
# Set to None or [] to allow all job types.
JOB_TYPES = ["internship", "full_time", "contract"]

# --- WORKPLACE TYPE FILTER ---
# Options: "on_site", "remote", "hybrid"
# Set to None or [] to allow all workplace types.
WORKPLACE_TYPES = []

# --- MATCHING STRICTNESS & THRESHOLDS ---
# Minimum number of secondary keywords that must appear in the job description
MIN_SECONDARY_MATCHES = 1

# If True, the job title itself must contain at least one term matching thesis / simulation
REQUIRE_TITLE_KEYWORD_MATCH = False

# Automatically save matched jobs on LinkedIn profile (Saved Jobs)
SAVE_ON_LINKEDIN = True

# Max job cards to process per keyword search (None = unlimited)
MAX_JOBS_PER_KEYWORD = None

# Continuous mode (loop periodically)
CONTINUOUS_MODE   = False
LOOP_WAIT_MINUTES = 60
MAX_CYCLES        = None

# Output file paths
CSV_OUTPUT_FILE   = "linkedin_matched_jobs.csv"
EXCEL_OUTPUT_FILE = "linkedin_matched_jobs.xlsx"


# ===========================================================================
# 2. KEYWORDS & FILTER PATTERNS
# ===========================================================================

# TRADITIONAL keyword-per-search list (used when USE_AI_SEARCH=False)
PRIMARY_KEYWORDS = [
    "Abschlussarbeit Berechnung",
    "master thesis simulation",
    "master thesis finite element",
    "master thesis CAE",
    "master thesis LS-DYNA",
    "master thesis Abaqus",
    "master thesis vehicle safety",
    "master thesis digital twin",
    "thesis computational mechanics",
    "thesis nonlinear FEM",
    "Masterarbeit Struktursimulation",
    "Masterarbeit FEM",
    "Masterarbeit Crashsimulation",
    "Masterarbeit Fahrzeugsicherheit",
    "Masterarbeit Finite Elemente",
    "Masterarbeit CAE",
    "Masterarbeit Simulation Karosserie",
    "Masterarbeit Simulation",
    "Abschlussarbeit Strukturmechanik",
    "Abschlussarbeit Crashsimulation",
    "Praktikum Simulation FEM",
    "Werkstudent Simulation CAE",
]

SECONDARY_KEYWORDS = [
    "LS-DYNA",
    "Abaqus",
    "PERMAS",
    "ANSA",
    "META",
    "Finite Element Analysis",
    "FEM",
    "FEA",
    "CFD",
    "CAE",
    "nonlinear FEM",
    "explicit crash simulation",
    "XFEM",
    "Cohesive Zone Model",
    "fracture mechanics",
    "structural simulation",
    "Struktursimulation",
    "Crashsimulation",
    "simulation automation",
    "Python simulation automation",
    "CAE automation",
    "parametric FEA",
    "FEA automation",
    "parametric modelling",
    "FE-Modellierung",
    "CFD-Simulation",
    "vehicle safety simulation",
    "Fahrzeugsicherheit",
    "crashworthiness",
    "ECE R66",
    "chassis simulation",
    "digital twin automotive",
    "Durability and Crash",
    "ANSYS",
    "MATLAB Simulink",
    "computational mechanics",
    "Berechnungsingenieur",
    "Strukturmechanik",
]

# Ambiguous acronym context rules (must co-occur with domain context)
_CONTEXT_REQUIRED = {
    "META": re.compile(
        r"(?i)(?=.*\bMETA\b)"
        r"(?=.*(?:post.?processor|BETA\s*CAE|solver|simulation|CAE|FEM|crash|Abaqus|LS-DYNA|pre.?processor))",
        re.DOTALL,
    ),
    "ANSA": re.compile(
        r"(?i)(?=.*\bANSA\b)"
        r"(?=.*(?:pre.?processor|BETA\s*CAE|meshing|mesh|solver|CAE|FEM|simulation|crash|geometry|model))",
        re.DOTALL,
    ),
    "CFD": re.compile(
        r"(?i)(?=.*\bCFD\b)"
        r"(?=.*(?:fluid|flow|strmung|aerodynamic|thermal|thermo|openfoam|ansys\s*fluent|star-ccm|solver))",
        re.DOTALL,
    ),
}

# Negative Title Filters (drop jobs immediately if title matches these)
EXCLUDE_TITLE_PATTERNS = [
    re.compile(r"(?i)\b(?:Senior\s+Manager|Director|Vice\s+President|VP|Head\s+of\s+Sales|Account\s+Executive)\b"),
    re.compile(r"(?i)\b(?:DevOps|Full\s*Stack|Frontend|Backend\s+Web|Cloud\s+Architect|SAP\s+Consultant)\b"),
    re.compile(r"(?i)\b(?:Buchhalter|Accountant|HR\s+Manager|Recruiter|Office\s+Assistant|Legal\s+Counsel)\b"),
    re.compile(r"(?i)\b(?:Medical\s+Doctor|Nurse|Pflegekraft|Store\s+Manager|Filialleiter)\b"),
]

# Negative Line/Context Filters (ignore lines related to tariffs, salaries, or embedded microcontrollers)
EXCLUDE_LINE_PATTERNS = [
    re.compile(r"(?i)salary|Gehalt|remuneration|entlohnung|Vergtung"),
    re.compile(r"(?i)collective\s+agreement|Tarifvertrag|tarifgebunden"),
    re.compile(r"(?i)ERA\s+group|ERA-Gruppe|Entgeltrahmen"),
    re.compile(r"(?i)\bpension\b|betriebsrente|sozialleistung"),
    re.compile(r"(?i)\bstraße\b|\bstreet\b|postal|postleitzahl"),
    re.compile(r"(?i)embedded\s+C|bare.?metal|microcontroller|firmware|RTOS|AUTOSAR"),
]

# Company Blacklist (optional: add names of agencies/companies you want to ignore)
BLACKLISTED_COMPANIES = []


# ===========================================================================
# 3. INTERNAL MAPPINGS & REGEX BUILDERS
# ===========================================================================

_DATE_FILTER_MAP = {
    "today": "r86400",
    "24h": "r86400",
    "day": "r86400",
    "week": "r604800",
    "past_week": "r604800",
    "month": "r2592000",
    "past_month": "r2592000",
    "any": "",
}

_EXPERIENCE_LEVEL_MAP = {
    "internship": "1",
    "entry": "2",
    "associate": "3",
    "mid_senior": "4",
    "director": "5",
    "executive": "6",
}

_JOB_TYPE_MAP = {
    "full_time": "F",
    "part_time": "P",
    "contract": "C",
    "temporary": "T",
    "internship": "I",
    "volunteer": "V",
    "other": "O",
}

_WORKPLACE_TYPE_MAP = {
    "on_site": "1",
    "remote": "2",
    "hybrid": "3",
}

_SORT_MAP = {
    "recent": "DD",      # Date posted (Most recent)
    "relevance": "R",    # Default relevance
}

def _build_smart_pattern(keyword: str) -> re.Pattern:
    """Builds regex that handles German compound words, hyphens, and whitespace variations."""
    if keyword in _CONTEXT_REQUIRED:
        if keyword == "META":
            return re.compile(r"\bMETA(?!le?)\b", re.IGNORECASE)
    # Allow optional hyphens or compound words for simulation/mechanics keywords
    escaped = re.escape(keyword)
    escaped = escaped.replace(r"\ ", r"[\s\-_]?")
    return re.compile(rf"(?:\b|(?<=[a-z\-_])){escaped}(?:\b|(?=[a-z\-_]))", re.IGNORECASE)

SECONDARY_PATTERNS = [_build_smart_pattern(kw) for kw in SECONDARY_KEYWORDS]


# ===========================================================================
# 4. TIMING & RANDOMIZATION
# ===========================================================================

SHORT_PAUSE  = (1.0, 2.5)
LONG_PAUSE   = (3.5, 6.5)
READ_PAUSE   = (2.5, 5.0)
TYPE_DELAY   = (0.06, 0.15)

IDLE_ACTION_CHANCE  = 0.25
SCROLL_PAUSE_CHANCE = 0.35
_BETWEEN_KEYWORD_PAUSE = (20, 50)


# ===========================================================================
# 5. LOGGING
# ===========================================================================

LOG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    f"linkedin_scraper_{datetime.now().strftime('%Y-%m-%d')}.log",
)

log = logging.getLogger("LinkedInScraper")
log.setLevel(logging.INFO)
if log.hasHandlers():
    log.handlers.clear()

file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
file_fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%H:%M:%S")
file_handler.setFormatter(file_fmt)
log.addHandler(file_handler)

class ConsoleHandler(logging.StreamHandler):
    def __init__(self):
        super().__init__(sys.stdout)
        self.last_was_no_match = False

    def emit(self, record):
        try:
            msg = self.format(record)
            if "No match." in record.getMessage():
                sys.stdout.write(f"\r{msg}".ljust(80))
                sys.stdout.flush()
                self.last_was_no_match = True
            else:
                if self.last_was_no_match:
                    sys.stdout.write("\n")
                    self.last_was_no_match = False
                sys.stdout.write(f"{msg}\n")
                sys.stdout.flush()
        except Exception:
            self.handleError(record)

console_handler = ConsoleHandler()
console_handler.setFormatter(file_fmt)
log.addHandler(console_handler)

class NavigationTimeoutError(Exception): pass
class WebDriverConnectionError(Exception): pass


# ===========================================================================
# 6. EDGE DRIVER & STEALTH PROFILE SETUP
# ===========================================================================

_SCRAPER_PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".edge_scraper_profile")
_DRIVER_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".edge_driver_cache")

def _get_edge_version() -> str | None:
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Edge\BLBeacon")
        version, _ = winreg.QueryValueEx(key, "version")
        winreg.CloseKey(key)
        return version
    except Exception:
        pass
    for base in [r"C:\Program Files (x86)\Microsoft\Edge\Application", r"C:\Program Files\Microsoft\Edge\Application"]:
        if not os.path.isdir(base): continue
        for entry in os.listdir(base):
            if entry and entry[0].isdigit() and os.path.isdir(os.path.join(base, entry)):
                return entry
    return None

def _driver_major(exe_path: str) -> str | None:
    try:
        import subprocess
        out = subprocess.check_output([exe_path, "--version"], stderr=subprocess.DEVNULL, timeout=5).decode()
        m = re.search(r"(\d+)\.\d+\.\d+\.\d+", out)
        return m.group(1) if m else None
    except Exception:
        return None

def _find_edgedriver_on_system(edge_version: str | None) -> str | None:
    import shutil
    edge_major = int(edge_version.split(".")[0]) if edge_version else None

    def _check(path: str) -> str | None:
        if not os.path.isfile(path): return None
        drv_major_str = _driver_major(path)
        if edge_major is not None and drv_major_str:
            drv_major = int(drv_major_str)
            if abs(drv_major - edge_major) > 1: return None
            if drv_major != edge_major:
                log.warning(f"  ?? Using driver {drv_major}.x with Edge {edge_major}.x")
        return path

    if edge_version:
        for base in [r"C:\Program Files (x86)\Microsoft\Edge\Application", r"C:\Program Files\Microsoft\Edge\Application"]:
            found = _check(os.path.join(base, edge_version, "msedgedriver.exe"))
            if found: return found

    home = os.path.expanduser("~")
    winget_pkg_base = os.path.join(home, "AppData", "Local", "Microsoft", "WinGet", "Packages")
    if os.path.isdir(winget_pkg_base):
        for entry in os.listdir(winget_pkg_base):
            if "EdgeDriver" in entry or "edgedriver" in entry.lower():
                found = _check(os.path.join(winget_pkg_base, entry, "msedgedriver.exe"))
                if found: return found

    for d in [os.path.join(home, "Downloads", "edgedriver_win64"), os.path.join(home, "Downloads"), os.path.join(home, "Desktop"), os.path.dirname(os.path.abspath(__file__))]:
        found = _check(os.path.join(d, "msedgedriver.exe"))
        if found: return found

    on_path = shutil.which("msedgedriver")
    if on_path:
        found = _check(on_path)
        if found: return found
    return None

def _download_edgedriver(version: str) -> str | None:
    import urllib.request, zipfile, io
    os.makedirs(_DRIVER_CACHE_DIR, exist_ok=True)
    cached_exe   = os.path.join(_DRIVER_CACHE_DIR, "msedgedriver.exe")
    version_file = os.path.join(_DRIVER_CACHE_DIR, "version.txt")
    if os.path.isfile(cached_exe) and os.path.isfile(version_file):
        with open(version_file) as f:
            if f.read().strip() == version: return cached_exe
    urls = [
        f"https://msedgedriver.azureedge.net/{version}/edgedriver_win64.zip",
        f"https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/{version}/edgedriver_win64.zip",
        "https://aka.ms/msedgedriver-win64",
    ]
    edge_major = version.split(".")[0]
    for url in urls:
        try:
            log.info(f"  Downloading msedgedriver {version} from {url} ?")
            with urllib.request.urlopen(url, timeout=30) as resp:
                data = resp.read()
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for name in zf.namelist():
                    if name.lower().endswith("msedgedriver.exe"):
                        exe_bytes = zf.read(name)
                        tmp_exe = cached_exe + ".tmp"
                        with open(tmp_exe, "wb") as f: f.write(exe_bytes)
                        drv_major = _driver_major(tmp_exe)
                        if drv_major and drv_major != edge_major:
                            try: os.remove(tmp_exe)
                            except OSError: pass
                            break
                        os.replace(tmp_exe, cached_exe)
                        with open(version_file, "w") as f: f.write(version)
                        log.info(f"  ? Saved: {cached_exe}")
                        return cached_exe
        except Exception as exc:
            log.warning(f"  Download failed ({url}): {exc}")
    return None

def _resolve_edgedriver() -> str:
    cached_exe   = os.path.join(_DRIVER_CACHE_DIR, "msedgedriver.exe")
    version_file = os.path.join(_DRIVER_CACHE_DIR, "version.txt")
    edge_version = _get_edge_version()

    if os.path.isfile(cached_exe) and os.path.isfile(version_file):
        with open(version_file) as f: cached_ver = f.read().strip()
        if edge_version and cached_ver == edge_version:
            log.info(f"  Using cached msedgedriver {cached_ver}")
            return cached_exe

    system_path = _find_edgedriver_on_system(edge_version)
    if system_path:
        os.makedirs(_DRIVER_CACHE_DIR, exist_ok=True)
        try:
            import shutil
            shutil.copy2(system_path, cached_exe)
        except PermissionError:
            try:
                with open(system_path, "rb") as src, open(cached_exe, "wb") as dst:
                    dst.write(src.read())
            except Exception:
                return system_path
        if edge_version:
            with open(version_file, "w") as f: f.write(edge_version)
        return cached_exe

    if edge_version:
        path = _download_edgedriver(edge_version)
        if path: return path

    if EdgeChromiumDriverManager is not None:
        try: return EdgeChromiumDriverManager().install()
        except Exception: pass

    if os.path.isfile(cached_exe): return cached_exe

    log.error("Cannot find msedgedriver.exe")
    sys.exit(1)

_STEALTH_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
)

_STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
['cdc_adoQpoasnfa76pfcZLmcfl_Array','cdc_adoQpoasnfa76pfcZLmcfl_Promise','cdc_adoQpoasnfa76pfcZLmcfl_Symbol','__selenium_evaluate','__selenium_unwrapped','__webdriver_evaluate','__webdriver_script_fn','__webdriver_unwrapped','__fxdriver_evaluate','__driver_evaluate','__driver_unwrapped'].forEach(k => { try { delete window[k]; } catch(e) {} });
const pluginData = [{name:'Chrome PDF Plugin',filename:'internal-pdf-viewer',description:'Portable Document Format'},{name:'Chrome PDF Viewer',filename:'mhjfbmdgcfjbbpaeojofohoefgiehjai',description:''},{name:'Native Client',filename:'internal-nacl-plugin',description:''}];
const fakePlugins = Object.create(PluginArray.prototype);
pluginData.forEach((p, i) => { const plugin = Object.create(Plugin.prototype); Object.defineProperties(plugin, {name:{value:p.name,enumerable:true},filename:{value:p.filename,enumerable:true},description:{value:p.description,enumerable:true},length:{value:0,enumerable:true}}); Object.defineProperty(fakePlugins, i, {value: plugin, enumerable: true}); });
Object.defineProperty(fakePlugins, 'length', {value: pluginData.length});
Object.defineProperty(navigator, 'plugins', {get: () => fakePlugins});
Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en', 'de']});
Object.defineProperty(navigator, 'platform', {get: () => 'Win32'});
if (!window.chrome) { window.chrome = {app:{isInstalled:false,InstallState:{DISABLED:'d',INSTALLED:'i',NOT_INSTALLED:'n'},RunningState:{CANNOT_RUN:'c',READY_TO_RUN:'r',RUNNING:'r'}},runtime:{OnInstalledReason:{CHROME_UPDATE:'c',INSTALL:'i',SHARED_MODULE_UPDATE:'s',UPDATE:'u'},OnRestartRequiredReason:{APP_UPDATE:'a',OS_UPDATE:'o',PERIODIC:'p'},PlatformArch:{ARM:'arm',ARM64:'arm64',MIPS:'m',MIPS64:'m6',X86_32:'x',X86_64:'x6'},PlatformNaclArch:{ARM:'arm',MIPS:'m',MIPS64:'m6',X86_32:'x',X86_64:'x6'},PlatformOs:{ANDROID:'a',CROS:'c',LINUX:'l',MAC:'m',OPENBSD:'o',WIN:'w'},RequestUpdateCheckStatus:{NO_UPDATE:'n',THROTTLED:'t',UPDATE_AVAILABLE:'u'}}}; }
try { const origQuery = window.Notification ? Notification.requestPermission.bind(Notification) : null; Object.defineProperty(Notification, 'permission', {get: () => 'default'}); } catch(e) {}
const origToDataURL = HTMLCanvasElement.prototype.toDataURL;
HTMLCanvasElement.prototype.toDataURL = function(type) { const ctx = this.getContext('2d'); if (ctx) { const px = ctx.getImageData(0, 0, 1, 1); px.data[0] = px.data[0] ^ (Math.random() * 2 | 0); ctx.putImageData(px, 0, 0); } return origToDataURL.apply(this, arguments); };
"""

def build_driver() -> webdriver.Edge:
    os.makedirs(_SCRAPER_PROFILE_DIR, exist_ok=True)
    log.info(f"  Profile: {_SCRAPER_PROFILE_DIR}")

    for _lock in ("SingletonLock", "lockfile", "SingletonSocket", "SingletonCookie"):
        _lp = os.path.join(_SCRAPER_PROFILE_DIR, _lock)
        try: os.remove(_lp)
        except OSError: pass

    opts = EdgeOptions()
    opts.set_capability("pageLoadStrategy", "eager")
    opts.add_argument(f"--user-data-dir={_SCRAPER_PROFILE_DIR}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--no-first-run")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1920,1080")

    driver_path = _resolve_edgedriver()
    service = EdgeService(driver_path)

    try:
        driver = webdriver.Edge(service=service, options=opts)
    except SessionNotCreatedException as exc:
        msg = str(exc)
        if "DevToolsActivePort" in msg or "crashed" in msg.lower():
            log.error(
                "?  Edge failed to start (DevToolsActivePort).\n"
                "  Most likely causes:\n"
                "    1. Stale profile lock  delete the profile folder and retry:\n"
                f"       {_SCRAPER_PROFILE_DIR}\n"
                "    2. msedgedriver version mismatch:\n"
                f"       driver  : {driver_path}\n"
                f"       Edge    : {_get_edge_version()}\n"
            )
        else:
            log.error(f"? Edge failed to start: {exc}")
        raise

    try: driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": _STEALTH_JS})
    except Exception: pass
    try: driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": _STEALTH_USER_AGENT, "acceptLanguage": "en-US,en;q=0.9,de;q=0.8", "platform": "Win32"})
    except Exception: pass

    driver.set_page_load_timeout(60)
    log.info("? Edge launched with stealth profile!")
    return driver

def driver_is_alive(driver) -> bool:
    try:
        driver.execute_script("return 1;")
        return True
    except Exception:
        return False


# ===========================================================================
# 7. HUMAN-LIKE HELPERS
# ===========================================================================

def pause(range_=SHORT_PAUSE):
    lo, hi = range_
    duration = random.gauss((lo + hi) / 2, (hi - lo) / 4)
    time.sleep(max(0.0, min(hi * 1.5, duration)))

def micro_pause(): time.sleep(random.uniform(0.05, 0.25))

def human_type(element, text: str):
    for char in text:
        element.send_keys(char)
        delay = random.uniform(*TYPE_DELAY)
        if random.random() < 0.08: delay += random.uniform(0.1, 0.35)
        time.sleep(delay)

def human_scroll(driver, element=None, direction: str = "down", steps: int | None = None):
    if steps is None: steps = random.randint(3, 7)
    sign = 1 if direction == "down" else -1
    for _ in range(steps):
        px = random.randint(80, 350)
        if element: driver.execute_script("arguments[0].scrollTop += arguments[1];", element, sign * px)
        else: driver.execute_script(f"window.scrollBy(0, {sign * px});")
        time.sleep(random.uniform(0.06, 0.20))
        if random.random() < SCROLL_PAUSE_CHANCE: time.sleep(random.uniform(0.25, 0.75))

def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    micro_pause()
    try:
        ActionChains(driver).move_to_element_with_offset(element, random.randint(-4, 4), random.randint(-3, 3)).pause(random.uniform(0.08, 0.25)).click().perform()
    except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
        try: element.click()
        except (ElementClickInterceptedException, StaleElementReferenceException): driver.execute_script("arguments[0].click();", element)

def idle_action(driver):
    action = random.choice(["mouse_move", "scroll_up", "small_down", "nothing"])
    if action == "mouse_move":
        try:
            w, h = driver.execute_script("return [window.innerWidth, window.innerHeight];")
            ActionChains(driver).move_by_offset(random.randint(50, max(51, w // 4)), random.randint(50, max(51, h // 4))).perform()
        except Exception: pass
    elif action == "scroll_up":
        driver.execute_script(f"window.scrollBy(0, -{random.randint(50, 180)});")
        time.sleep(random.uniform(0.2, 0.5))
        driver.execute_script(f"window.scrollBy(0, {random.randint(50, 180)});")
    elif action == "small_down":
        driver.execute_script(f"window.scrollBy(0, {random.randint(30, 120)});")
    time.sleep(random.uniform(0.15, 0.6))

def human_mouse_jitter(driver):
    try:
        w, h = driver.execute_script("return [window.innerWidth, window.innerHeight];")
        steps = random.randint(4, 10)
        ac = ActionChains(driver)
        for _ in range(steps):
            dx = max(-(w or 1200) // 4, min((w or 1200) // 4, random.randint(-120, 120)))
            dy = max(-(h or 800) // 4, min((h or 800) // 4, random.randint(-80, 80)))
            ac.move_by_offset(dx, dy).pause(random.uniform(0.04, 0.15))
        ac.perform()
    except Exception: pass


# ===========================================================================
# 8. ADVANCED FILTERING & RELEVANCE SCORING
# ===========================================================================

def is_title_blacklisted(title: str) -> bool:
    if not title: return False
    return any(pat.search(title) for pat in EXCLUDE_TITLE_PATTERNS)

def is_company_blacklisted(company: str) -> bool:
    if not company or not BLACKLISTED_COMPANIES: return False
    c_low = company.lower()
    return any(b.lower() in c_low for b in BLACKLISTED_COMPANIES)

def _kw_matches(kw: str, pat: re.Pattern, text: str) -> bool:
    if not pat.search(text): return False
    ctx = _CONTEXT_REQUIRED.get(kw)
    if ctx is not None and not ctx.search(text): return False
    return True

def matching_keywords(text: str) -> list[str]:
    return [kw for kw, pat in zip(SECONDARY_KEYWORDS, SECONDARY_PATTERNS) if _kw_matches(kw, pat, text)]

def matching_lines(text: str, keywords: list[str]) -> list[str]:
    patterns = [_build_smart_pattern(kw) for kw in keywords]
    seen, result = set(), []
    for line in text.splitlines():
        s = line.strip()
        if len(s) < 10 or s in seen: continue
        if any(excl.search(s) for excl in EXCLUDE_LINE_PATTERNS): continue
        if any(pat.search(s) for pat in patterns):
            seen.add(s)
            result.append(s)
    return result

def evaluate_job(title: str, company: str, description: str) -> tuple[bool, list[str], list[str], int]:
    """
    Evaluates whether a job passes all filters and calculates a match relevance score.
    Returns: (passes_filters, matched_keywords, matched_lines, score)
    """
    if is_title_blacklisted(title):
        return False, [], [], 0

    if is_company_blacklisted(company):
        return False, [], [], 0

    found_kws = matching_keywords(description)
    if len(found_kws) < MIN_SECONDARY_MATCHES:
        return False, [], [], 0

    if REQUIRE_TITLE_KEYWORD_MATCH:
        title_matches = matching_keywords(title)
        if not title_matches and not any(k.lower() in title.lower() for k in ["thesis", "master", "abschlussarbeit", "fem", "cae", "simulation"]):
            return False, [], [], 0

    key_lines = matching_lines(description, found_kws)

    # Relevance Scoring:
    # 2 pts per unique keyword matched + 3 bonus pts if keyword in title
    score = len(found_kws) * 2
    for kw in found_kws:
        if re.search(rf"\b{re.escape(kw)}\b", title, re.IGNORECASE):
            score += 3

    return True, found_kws, key_lines, score


# ===========================================================================
# 9. PERSISTENCE (CSV & EXCEL EXPORT)
# ===========================================================================

def export_matched_job(record: dict):
    """Appends a matched job to CSV and updates Excel workbook in real time."""
    # 1. Append to CSV
    csv_exists = os.path.isfile(CSV_OUTPUT_FILE)
    df_new = pd.DataFrame([record])
    df_new.to_csv(
        CSV_OUTPUT_FILE,
        mode="a",
        header=not csv_exists,
        index=False,
        encoding="utf-8-sig"
    )

    # 2. Update Excel with deduplication
    try:
        if os.path.isfile(EXCEL_OUTPUT_FILE):
            df_existing = pd.read_excel(EXCEL_OUTPUT_FILE)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined = df_combined.drop_duplicates(subset=["Job ID"], keep="last")
        else:
            df_combined = df_new

        with pd.ExcelWriter(EXCEL_OUTPUT_FILE, engine="openpyxl") as writer:
            df_combined.to_excel(writer, index=False, sheet_name="Matched Jobs")
            ws = writer.sheets["Matched Jobs"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = min(max_len + 3, 50)
    except Exception as exc:
        log.warning(f"  Warning: Excel export error: {exc}")


# ===========================================================================
# 10. LOGIN & SEARCH NAVIGATION
# ===========================================================================

def login(driver: webdriver.Edge):
    try: driver.get("https://www.linkedin.com/feed/")
    except Exception: pass
    pause(LONG_PAUSE)
    if "/feed" in driver.current_url or "/mynetwork" in driver.current_url: return

    if LINKEDIN_EMAIL and LINKEDIN_PASSWORD:
        try:
            human_type(driver.find_element(By.ID, "username"), LINKEDIN_EMAIL)
            human_type(driver.find_element(By.ID, "password"), LINKEDIN_PASSWORD)
            driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            pause(LONG_PAUSE)
        except NoSuchElementException: pass
        if "/feed" not in driver.current_url and "/mynetwork" not in driver.current_url:
            try: WebDriverWait(driver, 60).until(lambda d: "/feed" in d.current_url or "/mynetwork" in d.current_url)
            except TimeoutException: pass
    else:
        log.info("??  Please log in manually in the Edge window.")
        try: WebDriverWait(driver, 300).until(lambda d: "/feed" in d.current_url or "/mynetwork" in d.current_url)
        except TimeoutException: pass


def get_current_location() -> str:
    """Returns primary active location string from LOCATION_FILTER."""
    if isinstance(LOCATION_FILTER, list) and LOCATION_FILTER:
        return str(LOCATION_FILTER[0]).strip()
    return str(LOCATION_FILTER).strip()


def build_search_url(keyword: str, location: str | None = None) -> str:
    """Builds a complete, robust LinkedIn Jobs URL with all configured filters."""
    loc = location or get_current_location()
    params = [
        f"keywords={quote_plus(keyword)}",
        f"location={quote_plus(loc)}"
    ]

    # Geo ID — exact country/region targeting bypasses ambiguous location popups
    if USE_GEO_ID:
        geo_id = GEO_ID_MAP.get(loc.strip().lower(), "")
        if geo_id:
            params.append(f"geoId={geo_id}")

    # Sort By (DD = Date Descending / Most Recent)
    sort_code = _SORT_MAP.get(SORT_BY.lower(), "DD")
    if sort_code:
        params.append(f"sortBy={sort_code}")

    # Date Filter (f_TPR)
    tpr = _DATE_FILTER_MAP.get(DATE_FILTER.lower(), "")
    if tpr:
        params.append(f"f_TPR={tpr}")

    # Actively Hiring (f_AL=true) — only companies currently recruiting
    if ACTIVELY_HIRING_ONLY:
        params.append("f_AL=true")

    # Easy Apply only (f_EA=true) — LinkedIn-hosted application form
    if EASY_APPLY_ONLY:
        params.append("f_EA=true")

    # Experience Level (f_E)
    if EXPERIENCE_LEVEL:
        codes = [
            _EXPERIENCE_LEVEL_MAP[lvl.lower()]
            for lvl in EXPERIENCE_LEVEL
            if lvl.lower() in _EXPERIENCE_LEVEL_MAP
        ]
        if codes:
            params.append(f"f_E={quote_plus(','.join(codes))}")

    # Job Type / Employment Type (f_JT)
    if JOB_TYPES:
        jt_codes = [
            _JOB_TYPE_MAP[jt.lower()]
            for jt in JOB_TYPES
            if jt.lower() in _JOB_TYPE_MAP
        ]
        if jt_codes:
            params.append(f"f_JT={quote_plus(','.join(jt_codes))}")

    # Workplace Type (f_WT)
    if WORKPLACE_TYPES:
        wt_codes = [
            _WORKPLACE_TYPE_MAP[wt.lower()]
            for wt in WORKPLACE_TYPES
            if wt.lower() in _WORKPLACE_TYPE_MAP
        ]
        if wt_codes:
            params.append(f"f_WT={quote_plus(','.join(wt_codes))}")

    return "https://www.linkedin.com/jobs/search/?" + "&".join(params)


# ---------------------------------------------------------------------------
# AI SEARCH — MULTI-STRATEGY ENGINE
# ---------------------------------------------------------------------------
# LinkedIn has two surfaces for AI-based job search:
#
#   Strategy A — LinkedIn AI Jobs Search endpoint  (/jobs/search/?useAiSearch=true)
#     Direct URL approach that appends `useAiSearch=true` to the search URL.
#     LinkedIn's server-side LLM parses the full natural-language query and
#     returns semantically-matched results rather than exact keyword hits.
#     All granular URL filters (date, experience level, job type, etc.) are
#     preserved because we build the URL ourselves.
#
#   Strategy B — AI Prompt / Chat Bar  (/jobs/ai-search/)
#     LinkedIn's experimental AI search page where users type a conversational
#     prompt. We navigate there and type into the AI prompt textarea.
#     Extra filters are injected post-navigation via URL parameter amendment.
#
#   Strategy C — Standard search bar with ENTER (no autocomplete selection)
#     Navigates to /jobs/, clears the search bar, types the full NL query, and
#     submits with Enter WITHOUT clicking any autocomplete dropdown suggestion.
#     Autocomplete suggestions reduce NL queries to a single keyword match;
#     bypassing them preserves the full query for LinkedIn's search ranker.
#     Location is set separately AFTER keyword submission to avoid query reset.
#
#   Fallback — Traditional URL search  (build_search_url)
#     Used only when all AI strategies fail.
# ---------------------------------------------------------------------------


def _build_extra_filter_params(current_url: str) -> list[str]:
    """Returns URL params for all configured filters that are not already in current_url."""
    extra: list[str] = []

    tpr = _DATE_FILTER_MAP.get(DATE_FILTER.lower(), "")
    if tpr and f"f_TPR={tpr}" not in current_url:
        extra.append(f"f_TPR={tpr}")

    sort_code = _SORT_MAP.get(SORT_BY.lower(), "DD")
    if sort_code and f"sortBy={sort_code}" not in current_url:
        extra.append(f"sortBy={sort_code}")

    if ACTIVELY_HIRING_ONLY and "f_AL=true" not in current_url:
        extra.append("f_AL=true")

    if EASY_APPLY_ONLY and "f_EA=true" not in current_url:
        extra.append("f_EA=true")

    if EXPERIENCE_LEVEL:
        codes = [
            _EXPERIENCE_LEVEL_MAP[lvl.lower()]
            for lvl in EXPERIENCE_LEVEL
            if lvl.lower() in _EXPERIENCE_LEVEL_MAP
        ]
        if codes and "f_E=" not in current_url:
            extra.append(f"f_E={quote_plus(','.join(codes))}")

    if JOB_TYPES:
        jt_codes = [
            _JOB_TYPE_MAP[jt.lower()]
            for jt in JOB_TYPES
            if jt.lower() in _JOB_TYPE_MAP
        ]
        if jt_codes and "f_JT=" not in current_url:
            extra.append(f"f_JT={quote_plus(','.join(jt_codes))}")

    if WORKPLACE_TYPES:
        wt_codes = [
            _WORKPLACE_TYPE_MAP[wt.lower()]
            for wt in WORKPLACE_TYPES
            if wt.lower() in _WORKPLACE_TYPE_MAP
        ]
        if wt_codes and "f_WT=" not in current_url:
            extra.append(f"f_WT={quote_plus(','.join(wt_codes))}")

    # Geo ID
    if USE_GEO_ID:
        loc = get_current_location()
        geo_id = GEO_ID_MAP.get(loc.strip().lower(), "")
        if geo_id and f"geoId={geo_id}" not in current_url:
            extra.append(f"geoId={geo_id}")

    return extra


def _append_filters_to_current_url(driver: webdriver.Edge) -> bool:
    """
    Reads driver.current_url and appends any missing filter params in-place.
    Only modifies the URL when there is something to add.  Returns True if
    the page was reloaded with new params, False otherwise.
    """
    current_url = driver.current_url
    if "linkedin.com/jobs" not in current_url:
        return False

    extra_params = _build_extra_filter_params(current_url)
    if not extra_params:
        return False

    sep = "&" if "?" in current_url else "?"
    new_url = current_url + sep + "&".join(extra_params)
    log.info(f"  [AI Search] Injecting filters: {', '.join(extra_params)}")
    try:
        driver.get(new_url)
        pause(LONG_PAUSE)
        return True
    except Exception as exc:
        log.warning(f"  [AI Search] Filter injection failed: {exc}")
        return False


def _results_are_loaded(driver: webdriver.Edge) -> bool:
    """Returns True if at least one job card is visible on the page."""
    try:
        return bool(driver.find_elements(By.CSS_SELECTOR, _CARD_COMBINED))
    except Exception:
        return False


# ── Strategy A: useAiSearch=true URL parameter ─────────────────────────────
def _ai_search_via_url_param(driver: webdriver.Edge, query: str, location: str | None = None) -> bool:
    """
    Builds a standard LinkedIn jobs search URL and appends `useAiSearch=true`.
    This signals LinkedIn's backend to process the query through its LLM
    pipeline, returning semantic matches instead of literal keyword results.

    All URL-level filters (date, experience, job type, geo, sort) are included
    directly in the URL so they are never lost after AI processing.
    """
    loc = location or get_current_location()
    params = [
        f"keywords={quote_plus(query)}",
        f"location={quote_plus(loc)}",
        "useAiSearch=true",   # <-- signals LinkedIn's AI semantic search
    ]

    if USE_GEO_ID:
        geo_id = GEO_ID_MAP.get(loc.strip().lower(), "")
        if geo_id:
            params.append(f"geoId={geo_id}")

    sort_code = _SORT_MAP.get(SORT_BY.lower(), "DD")
    if sort_code:
        params.append(f"sortBy={sort_code}")

    tpr = _DATE_FILTER_MAP.get(DATE_FILTER.lower(), "")
    if tpr:
        params.append(f"f_TPR={tpr}")

    if ACTIVELY_HIRING_ONLY:
        params.append("f_AL=true")
    if EASY_APPLY_ONLY:
        params.append("f_EA=true")

    if EXPERIENCE_LEVEL:
        codes = [
            _EXPERIENCE_LEVEL_MAP[lvl.lower()]
            for lvl in EXPERIENCE_LEVEL
            if lvl.lower() in _EXPERIENCE_LEVEL_MAP
        ]
        if codes:
            params.append(f"f_E={quote_plus(','.join(codes))}")

    if JOB_TYPES:
        jt_codes = [
            _JOB_TYPE_MAP[jt.lower()]
            for jt in JOB_TYPES
            if jt.lower() in _JOB_TYPE_MAP
        ]
        if jt_codes:
            params.append(f"f_JT={quote_plus(','.join(jt_codes))}")

    if WORKPLACE_TYPES:
        wt_codes = [
            _WORKPLACE_TYPE_MAP[wt.lower()]
            for wt in WORKPLACE_TYPES
            if wt.lower() in _WORKPLACE_TYPE_MAP
        ]
        if wt_codes:
            params.append(f"f_WT={quote_plus(','.join(wt_codes))}")

    url = "https://www.linkedin.com/jobs/search/?" + "&".join(params)
    log.info(f"  [AI Search · Strategy A] useAiSearch URL: {url[:120]}...")
    try:
        driver.get(url)
        pause(LONG_PAUSE)
        if _results_are_loaded(driver):
            log.info("  [AI Search · Strategy A] ✓ Results loaded.")
            return True
        log.warning("  [AI Search · Strategy A] No cards found after useAiSearch URL.")
        return False
    except Exception as exc:
        log.warning(f"  [AI Search · Strategy A] Failed: {exc}")
        return False


# ── Strategy B: LinkedIn /jobs/ai-search/ conversational prompt page ────────
def _ai_search_via_prompt_page(driver: webdriver.Edge, query: str) -> bool:
    """
    Navigates to LinkedIn's AI Jobs conversational search page (/jobs/ai-search/),
    waits for the prompt textarea to appear, types the natural-language query,
    and submits it. This is the most semantically rich surface but may not always
    be available (A/B test / premium feature). Falls back gracefully.
    """
    try:
        driver.get("https://www.linkedin.com/jobs/ai-search/")
        pause(LONG_PAUSE)

        # If LinkedIn redirected away, the feature is unavailable
        if "ai-search" not in driver.current_url:
            log.info("  [AI Search · Strategy B] /jobs/ai-search/ not available (redirected).")
            return False

        # Prompt textareas / inputs used by LinkedIn's AI search chat interface
        _PROMPT_SELECTORS = [
            "textarea[aria-label*='Search']",
            "textarea[placeholder*='Describe']",
            "textarea[placeholder*='Tell us']",
            "textarea[class*='ai-search']",
            "textarea[class*='prompt']",
            "div[contenteditable='true'][aria-label*='search' i]",
            "div[contenteditable='true'][aria-label*='prompt' i]",
            "textarea",   # last resort
        ]
        prompt_input = None
        for sel in _PROMPT_SELECTORS:
            try:
                prompt_input = WebDriverWait(driver, 6).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                )
                if prompt_input:
                    log.info(f"  [AI Search · Strategy B] Prompt input found: {sel}")
                    break
            except TimeoutException:
                continue

        if prompt_input is None:
            log.warning("  [AI Search · Strategy B] Prompt input not found.")
            return False

        # Type the natural-language query into the AI prompt box
        prompt_input.click()
        micro_pause()
        prompt_input.send_keys(Keys.CONTROL + "a")
        micro_pause()
        prompt_input.send_keys(Keys.DELETE)
        micro_pause()
        human_type(prompt_input, query)
        pause((0.8, 1.5))

        # Submit — try a send/submit button first, then Enter
        _SUBMIT_SELECTORS = [
            "button[aria-label*='Send']",
            "button[aria-label*='Search']",
            "button[aria-label*='Submit']",
            "button[type='submit']",
            "button[class*='send']",
        ]
        submitted = False
        for sel in _SUBMIT_SELECTORS:
            try:
                btn = driver.find_element(By.CSS_SELECTOR, sel)
                if btn.is_displayed() and btn.is_enabled():
                    safe_click(driver, btn)
                    submitted = True
                    log.info(f"  [AI Search · Strategy B] Submitted via button: {sel}")
                    break
            except Exception:
                continue
        if not submitted:
            prompt_input.send_keys(Keys.RETURN)
            log.info("  [AI Search · Strategy B] Submitted via Enter.")

        pause(LONG_PAUSE)

        # Append extra filters on top of whatever URL the AI search page landed on
        _append_filters_to_current_url(driver)

        if _results_are_loaded(driver):
            log.info("  [AI Search · Strategy B] ✓ Results loaded.")
            return True

        log.warning("  [AI Search · Strategy B] No job cards visible after AI prompt search.")
        return False

    except Exception as exc:
        log.warning(f"  [AI Search · Strategy B] Failed: {exc}")
        return False


# ── Strategy C: Standard search bar — full NL query, NO autocomplete pick ──
def _ai_search_via_searchbar(driver: webdriver.Edge, query: str, location: str | None = None) -> bool:
    """
    Navigates to /jobs/, finds the keyword search bar, types the FULL natural-
    language query, and presses Enter WITHOUT clicking any autocomplete suggestion.

    Clicking autocomplete reduces the NL query to a single keyword (e.g. typing
    'Master thesis FEM Germany' and clicking the first suggestion returns only
    jobs matching 'FEM'). By dismissing the dropdown and pressing Enter we send
    the full query string to LinkedIn's search ranker which still applies its
    semantic ranking even on the traditional /jobs/search/ endpoint.

    Location is set AFTER keyword submission so it cannot overwrite the query.
    Extra URL filters are injected via URL amendment once results are on-screen.
    """
    try:
        driver.get("https://www.linkedin.com/jobs/")
        pause(LONG_PAUSE)

        # --- Keyword input ---
        _KW_INPUT_SELECTORS = [
            "input#job-search-bar-keywords",
            "input[aria-label*='Search by title']",
            "input[aria-label*='Job title']",
            "input[aria-label*='title, keyword']",
            "input[placeholder*='Search by title']",
            "input[placeholder*='title']",
            "input[data-view-name*='job-search']",
            "input.jobs-search-box__text-input",
            "input[class*='jobs-search-box']",
            "input[id*='job'][type='text']",
        ]
        search_input = None
        for sel in _KW_INPUT_SELECTORS:
            try:
                search_input = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                )
                if search_input:
                    log.info(f"  [AI Search · Strategy C] Keyword input found: {sel}")
                    break
            except TimeoutException:
                continue

        if search_input is None:
            log.warning("  [AI Search · Strategy C] Keyword input not found.")
            return False

        # Clear field and type the full NL query
        safe_click(driver, search_input)
        micro_pause()
        search_input.send_keys(Keys.CONTROL + "a")
        micro_pause()
        search_input.send_keys(Keys.DELETE)
        micro_pause()
        human_type(search_input, query)

        # Wait for any autocomplete panel, then DISMISS it with Escape so the
        # full NL query is preserved — do NOT click any suggestion.
        pause((0.8, 1.5))
        try:
            # Check if a suggestion panel is visible
            dropdown_visible = driver.find_elements(
                By.CSS_SELECTOR,
                "div[role='listbox'], ul[role='listbox'], div[class*='typeahead'], "
                "div[class*='autocomplete'], div[class*='suggestion']"
            )
            if dropdown_visible and dropdown_visible[0].is_displayed():
                search_input.send_keys(Keys.ESCAPE)   # close dropdown
                micro_pause()
                log.info("  [AI Search · Strategy C] Autocomplete dismissed (Escape).")
        except Exception:
            pass

        # Submit the full NL query
        search_input.send_keys(Keys.RETURN)
        log.info("  [AI Search · Strategy C] Submitted full NL query via Enter.")
        pause(LONG_PAUSE)

        # --- Location field (set AFTER keyword submission) ---
        # Find and update the location box only if it still shows a stale location;
        # do NOT use the location box to trigger a search re-submit.
        active_loc = location or get_current_location()
        _LOC_INPUT_SELECTORS = [
            "input#job-search-bar-location",
            "input[aria-label*='City, state']",
            "input[aria-label*='ocation']",
            "input[placeholder*='City, state']",
            "input[placeholder*='Location']",
        ]
        for sel in _LOC_INPUT_SELECTORS:
            try:
                loc_el = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                )
                current_loc = loc_el.get_attribute("value") or ""
                if active_loc.lower() not in current_loc.lower():
                    safe_click(driver, loc_el)
                    micro_pause()
                    loc_el.send_keys(Keys.CONTROL + "a")
                    micro_pause()
                    human_type(loc_el, active_loc)
                    pause((0.6, 1.2))
                    # Dismiss autocomplete for location too, then submit
                    loc_el.send_keys(Keys.ESCAPE)
                    micro_pause()
                    loc_el.send_keys(Keys.RETURN)
                    log.info(f"  [AI Search · Strategy C] Location set: {active_loc}")
                    pause(LONG_PAUSE)
                break
            except (TimeoutException, NoSuchElementException):
                continue

        # Inject any missing URL-level filters WITHOUT overriding the AI search URL
        _append_filters_to_current_url(driver)

        if _results_are_loaded(driver):
            log.info("  [AI Search · Strategy C] ✓ Results loaded.")
            return True

        log.warning("  [AI Search · Strategy C] No job cards visible after search-bar submission.")
        return False

    except Exception as exc:
        log.warning(f"  [AI Search · Strategy C] Failed: {exc}")
        return False


def navigate_ai_search(driver: webdriver.Edge, query: str, location: str | None = None):
    """
    Master AI Search dispatcher.  Tries three strategies in priority order:

      A. useAiSearch=true URL parameter  — cleanest, most reliable
      B. /jobs/ai-search/ prompt page    — most semantic, but gated/A-B tested
      C. Search bar with full NL query   — fallback using standard search bar

    If all AI strategies fail, falls back to the traditional URL keyword search.
    All strategies preserve the full suite of URL-level filters.
    """
    active_loc = location or get_current_location()
    log.info(f"  [AI Search] Attempting AI search for: \"{query}\" (Region/Location: {active_loc})")

    # Strategy A — useAiSearch URL param (most reliable)
    try:
        if _ai_search_via_url_param(driver, query, active_loc):
            return
    except NavigationTimeoutError:
        raise
    except Exception as exc:
        log.warning(f"  [AI Search · Strategy A] Exception: {exc}")

    # Strategy B — /jobs/ai-search/ conversational prompt page
    try:
        if _ai_search_via_prompt_page(driver, query):
            return
    except NavigationTimeoutError:
        raise
    except Exception as exc:
        log.warning(f"  [AI Search · Strategy B] Exception: {exc}")

    # Strategy C — Standard search bar, full NL query, no autocomplete click
    try:
        if _ai_search_via_searchbar(driver, query, active_loc):
            return
    except NavigationTimeoutError:
        raise
    except Exception as exc:
        log.warning(f"  [AI Search · Strategy C] Exception: {exc}")

    # Final fallback — traditional URL keyword search
    log.warning("  [AI Search] All AI strategies exhausted. Falling back to URL keyword search.")
    try:
        url = build_search_url(query, active_loc)
        driver.get(url)
        pause(LONG_PAUSE)
    except Exception as exc:
        log.error(f"  [AI Search] URL fallback also failed: {exc}")
        raise


def navigate_to_jobs_search(driver: webdriver.Edge, keyword: str, location: str | None = None):
    """
    Dispatches to the AI search engine or the traditional URL keyword search
    depending on the USE_AI_SEARCH configuration flag.

    AI mode  : calls navigate_ai_search() which tries three AI strategies in
               priority order before falling back to the URL keyword search.
    URL mode : calls build_search_url() and navigates directly, with connection-
               error handling and retry logic.
    """
    active_loc = location or get_current_location()
    exp_label = ", ".join(EXPERIENCE_LEVEL) if EXPERIENCE_LEVEL else "any"
    jt_label  = ", ".join(JOB_TYPES)        if JOB_TYPES        else "all"
    mode      = "AI" if USE_AI_SEARCH else "URL"
    log.info(
        f"  [{mode}] Query: {keyword!r} | Location: {active_loc} | Date: {DATE_FILTER} | "
        f"Sort: {SORT_BY} | Levels: {exp_label} | Types: {jt_label}"
    )

    if USE_AI_SEARCH:
        navigate_ai_search(driver, keyword, active_loc)
    else:
        pause((0.8, 1.8))
        try:
            driver.get(build_search_url(keyword, active_loc))
        except Exception as exc:
            err = str(exc).lower()
            if any(t in err for t in [
                "readtimeouterror", "read timed out",
                "net::err_connection", "net::err_name",
                "connection refused", "connection reset",
            ]):
                time.sleep(3)
                if "linkedin.com/jobs" not in driver.current_url:
                    time.sleep(20)
                    raise NavigationTimeoutError(keyword) from exc
            else:
                raise
        pause(LONG_PAUSE)


# ===========================================================================
# 11. CARD DETECTION & EXPANSION
# ===========================================================================

_CARD_SELECTORS = [
    # Modern Authenticated SDUI view (2025+)
    "div[role='button'][componentkey^='job-card-component-ref-']",
    # Legacy authenticated view
    "li[data-occludable-job-id]", "li.jobs-search-results__list-item",
    "li.scaffold-layout__list-item", "div.job-card-container", "div[data-job-id]",
    # Public / Guest view
    "div.base-search-card[data-entity-urn]", "div.job-search-card[data-entity-urn]",
    "div.base-card[data-entity-urn]",
]

_CARD_COMBINED = ", ".join(_CARD_SELECTORS)
_CARD_WAIT_TIMEOUT = 12
STALE_RETRY_LIMIT  = 3

def wait_for_list_stable(driver: webdriver.Edge, timeout: int = _CARD_WAIT_TIMEOUT) -> str | None:
    try: WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, _CARD_COMBINED)))
    except TimeoutException: return None
    winning_sel = next((sel for sel in _CARD_SELECTORS if driver.find_elements(By.CSS_SELECTOR, sel)), None)
    if not winning_sel: return None
    prev = -1
    for _ in range(6):
        pause((0.3, 0.5))
        now = len(driver.find_elements(By.CSS_SELECTOR, winning_sel))
        if now > 0 and now == prev: break
        prev = now
    return winning_sel

def _extract_job_id(element) -> str | None:
    # 1. SDUI componentkey
    ck = element.get_attribute("componentkey") or ""
    m = re.search(r"job-card-component-ref-(\d+)", ck)
    if m: return m.group(1)

    # 2. Data attributes
    for attr in ("data-occludable-job-id", "data-job-id"):
        v = element.get_attribute(attr)
        if v and v.strip(): return v.strip()

    # 3. URN
    urn = element.get_attribute("data-entity-urn") or ""
    m = re.search(r"(\d{5,})", urn)
    if m: return m.group(1)

    # 4. Link href fallback
    try:
        link = element.find_element(By.CSS_SELECTOR, "a[href*='/jobs/view/']")
        href = link.get_attribute("href") or ""
        m = re.search(r"/jobs/view/[^?]*?(\d{5,})", href)
        if m: return m.group(1)
    except NoSuchElementException:
        pass
    return None

def collect_visible_job_ids(driver: webdriver.Edge, selector: str) -> list[str]:
    ids = []
    seen = set()
    for card in driver.find_elements(By.CSS_SELECTOR, selector):
        try:
            jid = _extract_job_id(card)
            if jid and jid not in seen:
                seen.add(jid)
                ids.append(jid)
        except StaleElementReferenceException: continue
    return ids

def fetch_card_by_id(driver: webdriver.Edge, job_id: str, wait: int = 5):
    safe_id = job_id.replace("'", "\\'")
    sdui_sel = f"div[componentkey='job-card-component-ref-{safe_id}']"
    auth_combined = ", ".join(
        f"{sel}[{attr}='{safe_id}']"
        for sel in _CARD_SELECTORS
        for attr in ("data-occludable-job-id", "data-job-id", "id")
    )
    guest_combined = f"div[data-entity-urn*='{safe_id}']"
    combined = f"{sdui_sel}, {auth_combined}, {guest_combined}"
    try: return WebDriverWait(driver, wait).until(EC.presence_of_element_located((By.CSS_SELECTOR, combined)))
    except TimeoutException: return None

def scroll_results_panel(driver: webdriver.Edge):
    try:
        panel = driver.find_element(By.CSS_SELECTOR,
            "div[data-component-type='LazyColumn'][componentkey='SearchResultsMainContent'], "
            "div.jobs-search-results-list, ul.jobs-search-results__list, "
            "div.scaffold-layout__list, ul[class*='jobs-search'], "
            "ul.jobs-search__results-list, section.two-pane-serp-page__results-list"
        )
        human_scroll(driver, element=panel, direction="down", steps=random.randint(2, 5))
    except NoSuchElementException: human_scroll(driver, direction="down", steps=random.randint(2, 4))


# ===========================================================================
# 12. DETAIL EXTRACTION & "SEE MORE" EXPANSION
# ===========================================================================

def expand_description_if_needed(driver: webdriver.Edge):
    """Clicks 'See more' / 'Mehr anzeigen' button to load the full job description."""
    expand_selectors = [
        "button[aria-label*='more' i]",
        "button[aria-label*='mehr' i]",
        "button.jobs-description__footer-button",
        "button.show-more-less-html__button",
        "button[data-tracking-control-name*='show_more']",
    ]
    for sel in expand_selectors:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                micro_pause()
                break
        except Exception:
            pass

def get_description_text(driver: webdriver.Edge) -> str:
    expand_description_if_needed(driver)
    for sel in [
        "div.jobs-description__content", "div.jobs-box__html-content",
        "article.jobs-description", "div#job-details",
        "section.jobs-description", "div.job-view-layout",
        "div.jobs-description", "div[class*='job-details']",
        "div.description__text", "section.description",
        "div.show-more-less-html__markup", "div.decorated-job-posting__details",
    ]:
        try:
            text = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, sel))).text.strip()
            if text: return text
        except TimeoutException: continue
    try: return driver.find_element(By.TAG_NAME, "body").text
    except Exception: return ""

def get_job_meta(driver: webdriver.Edge, job_id: str) -> dict:
    meta = {
        "Job ID": job_id,
        "Title": "",
        "Company": "",
        "Location": "",
        "Workplace Type": "",
        "Posted Date": "",
        "Job URL": f"https://www.linkedin.com/jobs/view/{job_id}/" if job_id else driver.current_url,
    }

    # 1. Job Title
    for sel in [
        "h1.t-24", "h2.t-24",
        "h1.jobs-unified-top-card__job-title", "h1.topcard__title",
        "h2.jobs-details-top-card__job-title",
        "h1.job-details-jobs-unified-top-card__job-title",
        "h1[class*='job-title']", "h2[class*='job-title']",
        "h1.top-card-layout__title", "h2.top-card-layout__title",
    ]:
        try:
            t = driver.find_element(By.CSS_SELECTOR, sel).text.strip()
            if t: meta["Title"] = t; break
        except NoSuchElementException: continue

    # 2. Company Name
    for sel in [
        "div.job-details-jobs-unified-top-card__company-name a",
        "a.ember-view.t-black.t-normal",
        "span.jobs-unified-top-card__company-name",
        "a.topcard__org-name-link",
        "div.job-details-jobs-unified-top-card__company-name",
        "span[class*='company-name']", "a[class*='company']",
        "h4.base-search-card__subtitle a",
    ]:
        try:
            t = driver.find_element(By.CSS_SELECTOR, sel).text.strip()
            if t: meta["Company"] = t; break
        except NoSuchElementException: continue

    # 3. Location & Workplace Type
    for sel in [
        "div.job-details-jobs-unified-top-card__primary-description-container span",
        "span.jobs-unified-top-card__bullet",
        "span.topcard__flavor--bullet",
        "span.job-search-card__location",
    ]:
        for el in driver.find_elements(By.CSS_SELECTOR, sel):
            try:
                t = el.text.strip()
                if not t: continue
                if any(w in t.lower() for w in ["hybrid", "remote", "vor ort", "on-site"]):
                    meta["Workplace Type"] = t
                elif any(c.isalpha() for c in t) and not meta["Location"]:
                    meta["Location"] = t
            except StaleElementReferenceException: continue
        if meta["Location"]: break

    # 4. Posted Date
    for sel in [
        "span.jobs-unified-top-card__posted-date",
        "span.posted-time-ago__text",
        "span[class*='posted-date']",
        "time",
    ]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            t = el.text.strip()
            if t: meta["Posted Date"] = t; break
        except NoSuchElementException: continue

    return meta

def try_save_job(driver: webdriver.Edge) -> bool:
    if not SAVE_ON_LINKEDIN:
        return False
    for sel in ["button.jobs-save-button", "button[aria-label*='Save']", "button[data-control-name='jobdetails_topcard_save']"]:
        try:
            btn = WebDriverWait(driver, 4).until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
            label = (btn.get_attribute("aria-label") or btn.text or "").lower()
            if "unsave" in label or "saved" in label or "gespeichert" in label:
                log.info("  ? Already saved on LinkedIn.")
                return False
            safe_click(driver, btn)
            pause(SHORT_PAUSE)
            log.info("  ? ? Saved on LinkedIn!")
            return True
        except (TimeoutException, NoSuchElementException): continue
    return False


# ===========================================================================
# 13. CARD PROCESSING PIPELINE
# ===========================================================================

def process_card_with_retry(
    driver: webdriver.Edge,
    job_id: str,
    job_number: int,
    stats: dict,
    keyword_searched: str = "",
) -> None:
    for attempt in range(1, STALE_RETRY_LIMIT + 1):
        try:
            card = fetch_card_by_id(driver, job_id, wait=5)
            if card is None:
                if attempt < STALE_RETRY_LIMIT:
                    time.sleep(attempt * 1.2)
                    continue
                return

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
            micro_pause()
            safe_click(driver, card)
            pause(SHORT_PAUSE)

            meta = get_job_meta(driver, job_id)
            description = get_description_text(driver)

            # Check blacklist and keyword criteria
            passes, matched_kws, key_lines, score = evaluate_job(
                meta["Title"], meta["Company"], description
            )

            if passes:
                log.info(f"  [#{job_number}] ? MATCH (Score: {score})  {meta['Title']} @ {meta['Company']} ({meta['Location']})")
                log.info(f"    ?? {meta['Job URL']}")
                log.info(f"    ?? {', '.join(matched_kws)}")

                record = {
                    "Job ID": job_id,
                    "Title": meta["Title"],
                    "Company": meta["Company"],
                    "Location": meta["Location"],
                    "Workplace Type": meta["Workplace Type"],
                    "Posted Date": meta["Posted Date"],
                    "Match Score": score,
                    "Matched Keywords": "; ".join(matched_kws),
                    "Primary Search Keyword": keyword_searched,
                    "Key Excerpts": "\n---\n".join(key_lines[:5]),
                    "Job URL": meta["Job URL"],
                    "Scraped At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }

                export_matched_job(record)
                stats["matched"] += 1

                if try_save_job(driver):
                    stats["saved"] += 1
            else:
                log.info(f"  [#{job_number}] No match ({meta.get('Title', '')[:40]}).")

            if random.random() < IDLE_ACTION_CHANCE: idle_action(driver)
            pause(SHORT_PAUSE)
            return

        except NoSuchWindowException:
            log.error(f"  [#{job_number}] Edge window closed!")
            raise
        except StaleElementReferenceException:
            if attempt < STALE_RETRY_LIMIT:
                time.sleep(attempt * 1.5)
            else:
                stats["errors"] += 1
        except TimeoutException:
            if attempt == STALE_RETRY_LIMIT: stats["errors"] += 1
            continue
        except (ReadTimeoutError, MaxRetryError, ConnectionError, ConnectionRefusedError, ConnectionResetError, OSError, WebDriverException) as exc:
            stats["errors"] += 1
            if not driver_is_alive(driver): raise WebDriverConnectionError(f"Driver died at #{job_number}") from exc
            cooldown = 20 + attempt * 10
            log.info(f"  [#{job_number}] Cooldown {cooldown}s ?")
            time.sleep(cooldown)
            continue


# ===========================================================================
# 14. PAGINATION & KEYWORD LOOP
# ===========================================================================

_NEXT_BTN_SELECTORS = [
    "button[aria-label='View next page']", "button[aria-label='Nchste Seite anzeigen']",
    "li.artdeco-pagination__indicator--number.active + li button",
    "li.artdeco-pagination__indicator--number.selected + li button",
    "button.jobs-search-pagination__button--next",
    "li[data-test-pagination-page-btn][aria-selected='true'] + li button",
    "div.jobs-search-pagination button[aria-label*='next']",
    "div.jobs-search-pagination button[aria-label*='Next']",
]

def _find_next_btn(driver: webdriver.Edge):
    for sel in _NEXT_BTN_SELECTORS:
        try:
            btn = driver.find_element(By.CSS_SELECTOR, sel)
            if btn.is_enabled() and btn.is_displayed(): return btn
        except NoSuchElementException: continue
    return None

def process_keyword(driver: webdriver.Edge, keyword: str, location: str | None = None) -> dict:
    active_loc = location or get_current_location()
    try: navigate_to_jobs_search(driver, keyword, active_loc)
    except NavigationTimeoutError: return {"searched": 0, "matched": 0, "saved": 0, "errors": 1}

    stats = {"searched": 0, "matched": 0, "saved": 0, "errors": 0}
    processed_ids: set[str] = set()
    page_no = 1
    jobs_this_keyword = 0

    while True:
        winning_sel = wait_for_list_stable(driver, timeout=_CARD_WAIT_TIMEOUT)
        if winning_sel is None:
            log.warning(f"  ??  No job cards found for {keyword}")
            break

        no_new_scrolls = 0
        while True:
            visible_ids = collect_visible_job_ids(driver, winning_sel)
            batch = [jid for jid in visible_ids if jid not in processed_ids]
            if batch:
                no_new_scrolls = 0
                time.sleep(random.uniform(0.8, 1.5))
                for job_id in batch:
                    if MAX_JOBS_PER_KEYWORD is not None and jobs_this_keyword >= MAX_JOBS_PER_KEYWORD: return stats
                    processed_ids.add(job_id)
                    jobs_this_keyword += 1
                    stats["searched"] += 1
                    try: process_card_with_retry(driver, job_id, jobs_this_keyword, stats, keyword)
                    except WebDriverConnectionError:
                        if driver_is_alive(driver): time.sleep(60); return stats
                        else: raise NoSuchWindowException("WebDriver connection lost")
                sel = wait_for_list_stable(driver, timeout=3)
                if sel: winning_sel = sel
            else:
                no_new_scrolls += 1
                if no_new_scrolls >= 3: break
                scroll_results_panel(driver)
                time.sleep(random.uniform(0.8, 1.5))
                sel = wait_for_list_stable(driver, timeout=3)
                if sel: winning_sel = sel

        scroll_results_panel(driver)
        time.sleep(random.uniform(0.5, 1.0))
        next_btn = _find_next_btn(driver)
        if next_btn is None: break
        safe_click(driver, next_btn)
        pause(LONG_PAUSE)
        page_no += 1

        deadline = time.time() + 15
        new_page_ready = False
        while time.time() < deadline:
            sel_check = wait_for_list_stable(driver, timeout=3)
            if sel_check and [jid for jid in collect_visible_job_ids(driver, sel_check) if jid not in processed_ids]:
                winning_sel = sel_check
                new_page_ready = True
                break
            time.sleep(0.5)
        if not new_page_ready: break

    return stats

def _run_one_cycle(driver: webdriver.Edge, cycle_no: int) -> dict:
    total = {"searched": 0, "matched": 0, "saved": 0, "errors": 0}
    # AI mode uses AI_SEARCH_QUERIES; traditional mode uses PRIMARY_KEYWORDS
    query_list = AI_SEARCH_QUERIES if USE_AI_SEARCH else PRIMARY_KEYWORDS
    mode_label = "AI" if USE_AI_SEARCH else "URL"

    # Support single location or list of European countries / regions
    locations = LOCATION_FILTER if isinstance(LOCATION_FILTER, list) else [LOCATION_FILTER]

    for loc in locations:
        loc_str = str(loc).strip()
        for i, kw in enumerate(query_list, 1):
            log.info(f"\n  [Cycle {cycle_no} · {mode_label} · {loc_str} · {i}/{len(query_list)}] {kw}\n  " + "-" * 40)
            stats = process_keyword(driver, kw, loc_str)
            for k in total: total[k] += stats.get(k, 0)
            log.info(f"  Done — searched: {stats['searched']}, matched: {stats['matched']}, saved: {stats['saved']}, errors: {stats['errors']}")
            if i < len(query_list) or loc != locations[-1]:
                kw_wait = random.uniform(*_BETWEEN_KEYWORD_PAUSE)
                time.sleep(kw_wait * 0.4)
                if driver_is_alive(driver): human_mouse_jitter(driver)
                time.sleep(kw_wait * 0.6)
    return total


# ===========================================================================
# 15. MAIN EXECUTION
# ===========================================================================

def main():
    mode_label = "AI Natural Language" if USE_AI_SEARCH else "Traditional URL Keyword"
    log.info("=" * 60)
    log.info(f"LinkedIn Job Scraper — Enhanced Filtering ({mode_label} Mode)")
    log.info("=" * 60)
    log.info(f"  Profile         : {_SCRAPER_PROFILE_DIR}")
    log.info(f"  Search Mode     : {mode_label}")
    log.info(f"  Queries loaded  : {len(AI_SEARCH_QUERIES) if USE_AI_SEARCH else len(PRIMARY_KEYWORDS)}")
    log.info(f"  Location        : {LOCATION_FILTER}")
    log.info(f"  Date Filter     : {DATE_FILTER} (Sort: {SORT_BY})")
    log.info(f"  Actively Hiring : {ACTIVELY_HIRING_ONLY}")
    log.info(f"  Easy Apply Only : {EASY_APPLY_ONLY}")
    log.info(f"  Experience      : {', '.join(EXPERIENCE_LEVEL) if EXPERIENCE_LEVEL else 'All'}")
    log.info(f"  Job Types       : {', '.join(JOB_TYPES) if JOB_TYPES else 'All'}")
    log.info(f"  Workplace Types : {', '.join(WORKPLACE_TYPES) if WORKPLACE_TYPES else 'All'}")
    log.info(f"  Min Matches    : {MIN_SECONDARY_MATCHES}")
    log.info(f"  Output Files   : {CSV_OUTPUT_FILE}, {EXCEL_OUTPUT_FILE}")
    log.info("")

    grand = {"searched": 0, "matched": 0, "saved": 0, "errors": 0}
    driver = None
    cycle_no = 0

    try:
        while True:
            cycle_no += 1
            if MAX_CYCLES is not None and cycle_no > MAX_CYCLES: break
            log.info(f"\n============================================================\n??  Cycle {cycle_no}  ({datetime.now().strftime('%H:%M:%S')})\n============================================================")

            if driver is None or not driver_is_alive(driver):
                if driver is not None:
                    try: driver.quit()
                    except Exception: pass
                try: driver = build_driver()
                except SessionNotCreatedException as exc:
                    log.error("Could not build driver. Retrying in 60s...")
                    time.sleep(60)
                    continue
                login(driver)

            try: cycle_stats = _run_one_cycle(driver, cycle_no)
            except NoSuchWindowException:
                if CONTINUOUS_MODE:
                    driver = None
                    time.sleep(60)
                    continue
                else: break
            except Exception as exc:
                if CONTINUOUS_MODE:
                    driver = None
                    time.sleep(90)
                    continue
                else: raise

            for k in grand: grand[k] += cycle_stats.get(k, 0)

            log.info(f"\n============================================================\n?  Cycle {cycle_no} done  searched: {cycle_stats['searched']}, matched: {cycle_stats['matched']}, saved: {cycle_stats['saved']}")
            log.info(f"  Grand total  searched: {grand['searched']}, matched: {grand['matched']}, saved: {grand['saved']}, errors: {grand['errors']}")
            log.info(f"  ?? Results saved to: {CSV_OUTPUT_FILE} and {EXCEL_OUTPUT_FILE}")
            log.info("============================================================")

            if not CONTINUOUS_MODE: break
            actual_wait = random.uniform(45, 90)
            log.info(f"\n  ?? Waiting {actual_wait:.0f} min before next cycle  (Ctrl+C to stop)\n")
            time.sleep(actual_wait * 60)

    except KeyboardInterrupt:
        log.info(f"\n============================================================\n??  Stopped by user (Ctrl+C).\n  Cycles   : {cycle_no}\n  Searched : {grand['searched']}\n  Matched  : {grand['matched']}\n  Saved    : {grand['saved']}\n  Errors   : {grand['errors']}\n============================================================")
    finally:
        if driver is not None:
            try: driver.quit()
            except Exception: pass
        log.info("Edge closed. Done.")


if __name__ == "__main__":
    main()
