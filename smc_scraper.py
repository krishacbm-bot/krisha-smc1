#!/usr/bin/env python3
"""
School SMC Scraper  (fetches ALL schools in a city)
---------------------------------------------------
  1. Finds every school in the city (OpenStreetMap, plus Google Places if a key is set).
  2. Finds each school's website (from the data, or via a web-search fallback).
  3. Crawls each website for the SMC (School Management Committee) page.
  4. Extracts SMC member details if published.
  5. Collects SMC-related PDFs and their download links.
  6. Saves everything to Excel (sheets: Schools, SMC Members, SMC PDFs).

Install:
    pip install requests beautifulsoup4 openpyxl lxml

Run:
    python smc_scraper.py --school "St. Xavier's High School" --city "Fort, Mumbai"
    python smc_scraper.py --url https://www.schoolwebsite.com
    python smc_scraper.py --input my_schools.xlsx     (columns: School Name, Location)
    python smc_scraper.py --city "Mumbai"            (ALL schools in the city)
    python smc_scraper.py                            (interactive menu)

Optional (better coverage): set GOOGLE_PLACES_API_KEY
    Windows:  set GOOGLE_PLACES_API_KEY=your_key
"""
import argparse
import base64
import os
import random
import re
import threading
import time
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, parse_qs, unquote

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# >>> Put your real email here. Overpass blocks anonymous/generic clients. <<<
CONTACT_EMAIL = "your-email@example.com"

HEADERS = {
    "User-Agent": f"Mozilla/5.0 (compatible; SchoolSMCScraper/1.0; {CONTACT_EMAIL})",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
OVERPASS_HEADERS = {
    "User-Agent": f"SchoolSMCScraper/1.0 (contact: {CONTACT_EMAIL})",
    "Accept": "*/*",
}
TIMEOUT = 15
MAX_PAGES_PER_SITE = 25
MAX_DEPTH = 2

OVERPASS_URLS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
]

SMC_RE = re.compile(
    r"\bsmc\b|school[\s_%20-]*management[\s_%20-]*committee|"
    r"management[\s_%20-]*committee|\bsdmc\b|\bsmdc\b|governing[\s_%20-]*body|"
    r"\bpta\b|parent[\s_-]*teacher[\s_-]*association",
    re.I,
)
STRONG_SMC_RE = re.compile(
    r"\bsmc\b|school[\s_%20-]*management[\s_%20-]*committee|management[\s_%20-]*committee",
    re.I,
)
ROLE_RE = re.compile(
    r"chair\s?(person|man|woman)?|president|vice[\s-]*president|secretary|treasurer|"
    r"principal|head\s?(master|mistress)|member|convenor|co-?ordinator|"
    r"parent\s+(rep|member)|teacher\s+(rep|member)|nominee|councillor|trustee",
    re.I,
)
COMMON_PATHS = [
    "/smc", "/smc/", "/smc.html", "/smc.php", "/school-management-committee",
    "/school-management-committee/", "/management-committee", "/sdmc",
    "/about-us/smc", "/about/smc", "/smc-members", "/governing-body",
]
SKIP_DOMAINS = (
    "facebook.", "instagram.", "twitter.", "x.com", "linkedin.", "youtube.",
    "wikipedia.", "justdial.", "schoolmykids.", "joonsquare.", "icbse.",
    "edustoke.", "uniapply.", "shiksha.", "careers360.", "collegedunia.",
    "google.", "bing.", "duckduckgo.", "maps.", "tripadvisor.", "sulekha.",
    "indiamart.", "wikimapia.", "yelp.", "mapcarta.",
)

session = requests.Session()
session.headers.update(HEADERS)


# ----------------------------------------------------------------------------
# 1. Find schools
# ----------------------------------------------------------------------------
def normalize_url(u):
    if not u:
        return ""
    u = u.strip()
    if not u.startswith(("http://", "https://")):
        u = "http://" + u
    return u


def parse_osm(elements):
    schools = []
    for el in elements:
        t = el.get("tags", {})
        name = t.get("name") or t.get("name:en")
        if not name:
            continue
        site = t.get("website") or t.get("contact:website") or t.get("url") or ""
        addr = ", ".join(filter(None, [
            t.get("addr:housenumber"), t.get("addr:street"),
            t.get("addr:suburb"), t.get("addr:city")]))
        schools.append({"name": name, "website": normalize_url(site),
                        "address": addr, "source": "OpenStreetMap"})
    return schools


def schools_from_osm(city):
    # Two query styles: strict admin boundary first, then any area with that name
    queries = [
        f"""[out:json][timeout:180];
        area["name"="{city}"]["boundary"="administrative"]->.a;
        (nwr["amenity"="school"](area.a););
        out tags center;""",
        f"""[out:json][timeout:180];
        area["name"="{city}"]->.a;
        (nwr["amenity"="school"](area.a););
        out tags center;""",
    ]
    for attempt in range(2):
        for q in queries:
            for url in OVERPASS_URLS:
                try:
                    r = requests.post(url, data={"data": q},
                                      headers=OVERPASS_HEADERS, timeout=240)
                    r.raise_for_status()
                    schools = parse_osm(r.json().get("elements", []))
                    if schools:
                        print(f"[OSM] {len(schools)} schools found")
                        return schools
                    print(f"[OSM] {url}: 0 results, trying next option")
                except Exception as e:
                    print(f"[OSM] {url} failed: {e}")
        print("Retrying in 10 seconds...")
        time.sleep(10)
    return []


def schools_from_google(city, api_key):
    """Google Places Text Search + Details (needs API key)."""
    schools = []
    # Google returns max ~60 per query, so use several query variants
    variants = [f"schools in {city}", f"CBSE schools in {city}", f"ICSE schools in {city}",
                f"high school in {city}", f"international school in {city}",
                f"primary school in {city}", f"government school in {city}"]
    for qtext in variants:
        token = None
        while True:
            params = {"query": qtext, "key": api_key}
            if token:
                params["pagetoken"] = token
                time.sleep(2.5)
            try:
                r = session.get("https://maps.googleapis.com/maps/api/place/textsearch/json",
                                params=params, timeout=TIMEOUT).json()
            except Exception as e:
                print(f"[Google] error: {e}")
                break
            for p in r.get("results", []):
                try:
                    d = session.get(
                        "https://maps.googleapis.com/maps/api/place/details/json",
                        params={"place_id": p["place_id"],
                                "fields": "name,website,formatted_address",
                                "key": api_key}, timeout=TIMEOUT).json().get("result", {})
                except Exception:
                    d = {}
                schools.append({"name": d.get("name", p["name"]),
                                "website": normalize_url(d.get("website", "")),
                                "address": d.get("formatted_address", ""),
                                "source": "Google Places"})
            token = r.get("next_page_token")
            if not token:
                break
    print(f"[Google] {len(schools)} schools found")
    return schools


# ---- website finder -------------------------------------------------------
SEARCH_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
SERPER_KEY = os.getenv("SERPER_API_KEY", "")      # optional: free key from serper.dev
ENGINE_HOSTS = ("bing.", "duckduckgo.", "mojeek.", "microsoft.", "google.", "gstatic.")
EXTRA_SKIP = ("schoolsindia.", "cbse.gov", "udiseplus", "getmyuni", "eduvidya", "indiacatalog",
              "yellowpages", "zaubacorp", "schooldekho", "schoolchalo", "edufever", "admissionsmumbai",
              "nobroker", "magicbricks", "99acres", "wikimedia", "quora.", "reddit.", "pinterest.",
              "olx.", "glassdoor.", "naukri.", "indeed.", "mapquest.", "foursquare.", "cybo.",
              "wikipedia.", "justdial.", "sulekha.", "tripadvisor.", "linkedin.", "youtube.", "twitter.",
              "facebook.", "instagram.", "x.com", "careers360.", "collegedunia.", "shiksha.", "edustoke.")
_search_lock = threading.Lock()
_last_search = [0.0]
_blocked_until = {}


class Blocked(Exception):
    pass


def _throttle():
    """Space out searches (shared by all threads) so search engines don't block us."""
    with _search_lock:
        wait = 2.0 + random.random() - (time.time() - _last_search[0])
        if wait > 0:
            time.sleep(wait)
        _last_search[0] = time.time()


def _clean_link(href):
    href = href.strip()
    if href.startswith("//"):
        href = "https:" + href
    if "uddg=" in href:                                   # DuckDuckGo redirect
        href = unquote(parse_qs(urlparse(href).query).get("uddg", [""])[0])
    elif "bing.com/ck/a" in href:                         # Bing redirect
        try:
            u = parse_qs(urlparse(href).query).get("u", [""])[0]
            if u.startswith("a1"):
                u = u[2:]
            u += "=" * (-len(u) % 4)
            href = base64.urlsafe_b64decode(u).decode("utf-8", "ignore")
        except Exception:
            return ""
    return href


def _links_from_html(html):
    soup = BeautifulSoup(html, "lxml")
    out = []
    for a in soup.find_all("a", href=True):
        href = _clean_link(a["href"])
        if href.startswith("http") and not any(e in urlparse(href).netloc.lower() for e in ENGINE_HOSTS):
            out.append(href)
    return out


def _engine_page(method, url, **kw):
    r = requests.request(method, url, headers={"User-Agent": SEARCH_UA, "Accept-Language": "en-IN,en;q=0.9"},
                         timeout=TIMEOUT, **kw)
    if r.status_code in (202, 403, 429) or re.search(r"captcha|unusual traffic|anomaly|are you a robot",
                                                      r.text[:4000], re.I):
        raise Blocked()
    return _links_from_html(r.text)


def _serper(q):
    r = requests.post("https://google.serper.dev/search", timeout=TIMEOUT,
                      headers={"X-API-KEY": SERPER_KEY, "Content-Type": "application/json"},
                      json={"q": q, "gl": "in", "num": 10})
    if r.status_code in (401, 403):
        raise Blocked()
    return [x["link"] for x in r.json().get("organic", []) if x.get("link")]


ENGINES = [
    ("bing", lambda q: _engine_page("get", "https://www.bing.com/search", params={"q": q, "setlang": "en"})),
    ("mojeek", lambda q: _engine_page("get", "https://www.mojeek.com/search", params={"q": q})),
    ("duckduckgo", lambda q: _engine_page("post", "https://html.duckduckgo.com/html/", data={"q": q})),
    ("ddg-lite", lambda q: _engine_page("post", "https://lite.duckduckgo.com/lite/", data={"q": q})),
]


def _candidates(q, notes):
    """Search several engines in turn; return a list of candidate website roots."""
    engines = ([("serper", _serper)] if SERPER_KEY else []) + ENGINES
    for name, fn in engines:
        if time.time() < _blocked_until.get(name, 0):
            continue
        if name != "serper":
            _throttle()
        try:
            links = fn(q)
        except Blocked:
            _blocked_until[name] = time.time() + 90
            notes.append(f"{name} blocked")
            continue
        except Exception:
            notes.append(f"{name} error")
            continue
        roots = []
        for href in links:
            host = urlparse(href).netloc.lower()
            if any(x in host for x in SKIP_DOMAINS + EXTRA_SKIP):
                continue
            root = f"{urlparse(href).scheme}://{urlparse(href).netloc}"
            if root not in roots:
                roots.append(root)
        if roots:
            return roots
        notes.append(f"{name} no results")
    return []


GENERIC_WORDS = {
    "school", "schools", "high", "the", "of", "and", "public", "english", "medium", "junior", "senior",
    "secondary", "higher", "primary", "college", "international", "model", "vidyalaya", "vidya", "mandir",
    "shala", "convent", "academy", "institute", "education", "educational", "society", "trust", "cbse",
    "icse", "board", "std", "pre", "for", "girls", "boys", "kids", "its"}
ADDR_SKIP = {"road", "street", "near", "opposite", "behind", "plot", "building", "nagar", "sector", "maharashtra",
             "india", "east", "west", "north", "south", "marg", "lane", "floor", "complex", "society", "chowk"}


def _norm(t):
    return re.sub(r"[^a-z0-9]+", " ", (t or "").lower().replace("'", "").replace("\u2019", ""))


def _has(tok, text):
    return tok in text or (tok.endswith("s") and tok[:-1] in text)


def name_tokens(name, location=""):
    """Distinctive words of the school name (generic words and the city/district are ignored)."""
    words = [w for w in _norm(name).split() if len(w) > 2]
    loc = set(_norm(location).split())
    toks = [w for w in words if w not in GENERIC_WORDS and w not in loc]
    return toks or [w for w in words if w not in GENERIC_WORDS] or words


def _page_info(url):
    final, page = fetch(url)
    if not page:
        return None
    soup = BeautifulSoup(page, "lxml")
    head = [soup.title.get_text(" ", strip=True) if soup.title else ""]
    head += [h.get_text(" ", strip=True) for h in soup.find_all(["h1", "h2"])[:4]]
    for m in soup.find_all("meta"):
        if (m.get("property") in ("og:site_name", "og:title") or m.get("name") in ("description", "application-name")) \
                and m.get("content"):
            head.append(m["content"])
    contact = ""
    for a in soup.find_all("a", href=True):
        if "contact" in (a["href"] + " " + a.get_text(" ")).lower():
            contact = urljoin(final or url, a["href"])
            break
    return {"final": final or url, "head": _norm(" ".join(head)),
            "body": _norm(soup.get_text(" ", strip=True)[:5000]),
            "raw": soup.get_text(" ", strip=True)[:5000], "contact": contact}


def verify_site(root, name, location="", address=""):
    """Score how well a website belongs to THIS school. Returns (score, location_ok, final_url) or None."""
    info = _page_info(root)
    if not info:
        return None
    toks = name_tokens(name, location)
    head_frac = sum(_has(t, info["head"]) for t in toks) / len(toks)
    body_frac = sum(_has(t, info["body"]) for t in toks) / len(toks)
    # The name must be in the page title / headings / meta. A name that is only mentioned somewhere in the
    # text (directory or news pages do that) can never reach the 0.5 needed to be returned.
    score = head_frac if head_frac > 0 else 0.4 * body_frac
    host = re.sub(r"[^a-z0-9]", "", urlparse(info["final"]).netloc.lower())
    if any(t in host for t in toks):
        score = min(1.0, score + 0.15)
        if body_frac >= 0.8:
            score = max(score, 0.8)

    text, raw = info["body"], info["raw"]
    loc_words = [w for w in _norm(address).split() if len(w) >= 5 and w not in ADDR_SKIP][:6]
    loc_words += [w for w in _norm(location).split() if len(w) >= 4]
    pins = re.findall(r"\b\d{6}\b", f"{address} {location}")

    def loc_found(t_norm, t_raw):
        return any(pn in t_raw for pn in pins) or any(_has(w, t_norm) for w in loc_words)

    loc_ok = loc_found(text, raw)
    if not loc_ok and info["contact"]:
        c = _page_info(info["contact"])
        if c:
            loc_ok = loc_found(c["body"], c["raw"])
    return score, loc_ok, info["final"], len(toks)


def search_website_ex(name, location="", address=""):
    """Find a school's own website. Returns (url, note).
    A website is only returned when its content really matches the school; otherwise the URL is left
    blank and the closest guess is mentioned in the note, so wrong sites never reach the Website column."""
    notes, cands = [], []
    for q in (f'"{name}" {location} school official website'.strip(), f"{name} {location}".strip()):
        cands = _candidates(q, notes)
        if cands:
            break
    if not cands:
        return "", ("Not found (" + ", ".join(dict.fromkeys(notes)) + ")") if notes else "Not found"

    best = None   # (rank, url, score, loc_ok, ntoks)
    for root in cands[:6]:
        v = verify_site(root, name, location, address)
        if not v:
            continue
        score, loc_ok, final, ntoks = v
        rank = score + (0.25 if loc_ok else 0)
        if not best or rank > best[0]:
            best = (rank, final, score, loc_ok, ntoks)
        if score >= 0.8 and loc_ok:
            break
    if not best:
        return "", f"Could not open the candidate sites (closest: {cands[0]})"
    _rank, url, score, loc_ok, ntoks = best
    if score >= 0.8 and (loc_ok or ntoks >= 2):
        return url, "Good match" if loc_ok else "Name matches, location not confirmed - please check"
    if score >= 0.5 and loc_ok:
        return url, "Low match - please check"
    return "", f"No confident match (closest guess: {url})"


def _clean_site(val):
    """Turn what a school typed in CBSE's 'Website' field into a usable URL (or '')."""
    for tok in re.split(r"[\s,;|]+", (val or "").strip()):
        t = tok.strip("()[]<>.:;,'\"")
        if "." in t and "@" not in t and len(t) > 4 and t.lower() not in ("n.a.", "n/a", "n.a"):
            if not t.lower().startswith(("http://", "https://")):
                t = "http://" + t
            return t
    return ""


def cbse_website(aff_no):
    """Read the school's own declared website from CBSE's official affiliation record (by affiliation number).
    Returns (url, note)."""
    aff = re.sub(r"\D", "", str(aff_no).split(".")[0])
    if not aff:
        return "", "No affiliation number in the file"
    url = f"https://saras.cbse.gov.in/saras/AffiliatedList/AfflicationDetails/{aff}"
    html = None
    for verify in (True, False):
        try:
            r = requests.get(url, headers={"User-Agent": SEARCH_UA}, timeout=25, verify=verify)
            if r.status_code == 200:
                html = r.text
            break
        except requests.exceptions.SSLError:
            continue
        except Exception:
            break
    if not html:
        return "", "Could not open the CBSE record"
    lines = [ln for ln in BeautifulSoup(html, "lxml").get_text("\n", strip=True).split("\n") if ln]
    val = ""
    for i, ln in enumerate(lines):
        if ln.lower().startswith("website"):
            rest = ln[7:].strip(" :-")
            nxt = lines[i + 1] if i + 1 < len(lines) else ""
            val = rest or ("" if nxt.lower().startswith("year of foundation") else nxt)
            break
    site = _clean_site(val)
    if site:
        return site, "From CBSE official record"
    return "", "CBSE record lists no website" if lines else "Could not read the CBSE record"


def search_website(name, city):
    return search_website_ex(name, city)[0]


def load_input_file(path):
    """Read schools from .xlsx or .csv. Needs a school-name column and
    (optionally) a location/city/address column."""
    import csv
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        rows = [list(r) for r in load_workbook(path, read_only=True).active.iter_rows(values_only=True)]
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    rows = [r for r in rows if r and any(c not in (None, "") for c in r)]
    header = [str(c or "").strip().lower() for c in rows[0]]
    name_i = next((i for i, h in enumerate(header) if "school" in h or "name" in h), 0)
    loc_i = next((i for i, h in enumerate(header)
                  if any(k in h for k in ("location", "city", "address", "area", "place"))), None)
    web_i = next((i for i, h in enumerate(header) if "website" in h or "url" in h), None)
    schools = []
    for r in rows[1:]:
        name = str(r[name_i] or "").strip() if len(r) > name_i else ""
        if not name:
            continue
        loc = str(r[loc_i] or "").strip() if loc_i is not None and len(r) > loc_i else ""
        web = str(r[web_i] or "").strip() if web_i is not None and len(r) > web_i else ""
        schools.append({"name": name, "location": loc, "address": loc,
                        "website": normalize_url(web), "source": "Input file"})
    return schools


def school_from_url(url, name=""):
    """Build a school record from just a website link (name is guessed from the page title)."""
    url = normalize_url(url)
    if not name:
        final, page = fetch(url)
        if page:
            t = BeautifulSoup(page, "lxml").title
            if t:
                name = re.split(r"\s[|\-–—:]\s", t.get_text(" ", strip=True))[0].strip()
        if not name or name.lower() in {"home", "welcome", "index", "homepage"}:
            name = urlparse(url).netloc
    return {"name": name, "website": url, "location": "", "address": "",
            "source": "Website link"}


def dedupe(schools):
    seen, out = {}, []
    for s in schools:
        key = re.sub(r"\W+", "", s["name"].lower())
        if key in seen:
            if not seen[key]["website"] and s["website"]:
                seen[key]["website"] = s["website"]
            continue
        seen[key] = s
        out.append(s)
    return out


# ----------------------------------------------------------------------------
# 2. Crawl one school website
# ----------------------------------------------------------------------------
def fetch(url):
    try:
        r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
        ctype = r.headers.get("Content-Type", "").lower()
        if r.status_code == 200 and "html" in ctype:
            return r.url, r.text
    except Exception:
        pass
    return None, None


def is_pdf(url):
    return urlparse(url).path.lower().endswith(".pdf")


def extract_members(soup):
    """Return list of (name, designation) from tables or 'Name - Role' lines."""
    members = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header = [c.get_text(" ", strip=True).lower() for c in rows[0].find_all(["th", "td"])]
        name_i = next((i for i, h in enumerate(header) if "name" in h or "member" in h), None)
        role_i = next((i for i, h in enumerate(header)
                       if any(k in h for k in ("designation", "post", "role", "position", "capacity"))), None)
        if name_i is None:
            continue
        for tr in rows[1:]:
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
            if len(cells) > name_i and cells[name_i]:
                role = cells[role_i] if role_i is not None and len(cells) > role_i else ""
                members.append((cells[name_i], role))
    if members:
        return members
    for el in soup.find_all(["li", "p", "div", "span"]):
        if el.find(["li", "p", "div", "table"]):
            continue
        text = el.get_text(" ", strip=True)
        if 5 < len(text) < 150 and ROLE_RE.search(text):
            parts = re.split(r"\s[-–—:|]\s|,\s", text, maxsplit=1)
            if len(parts) == 2:
                a, b = parts
                if ROLE_RE.search(b) and not ROLE_RE.search(a):
                    members.append((a.strip(), b.strip()))
                elif ROLE_RE.search(a):
                    members.append((b.strip(), a.strip()))
    return list(dict.fromkeys(members))[:60]


def crawl_school(school):
    result = {**school, "smc_pages": [], "members": [], "pdfs": [], "status": ""}
    base = school["website"]
    if not base:
        result["status"] = "No website found"
        return result

    root, html = fetch(base)
    if not root:
        result["status"] = "Website unreachable"
        return result
    host = urlparse(root).netloc.replace("www.", "")

    queue = deque([(root, 0)])
    for p in COMMON_PATHS:
        queue.append((urljoin(root, p), 1))
    seen, pages_done = set(), 0
    smc_pages, pdfs = {}, {}

    while queue and pages_done < MAX_PAGES_PER_SITE:
        url, depth = queue.popleft()
        url = url.split("#")[0]
        if url in seen:
            continue
        seen.add(url)
        final, page = (root, html) if url == root else fetch(url)
        if not page:
            continue
        pages_done += 1
        soup = BeautifulSoup(page, "lxml")
        title = (soup.title.get_text(" ", strip=True) if soup.title else "")
        body_text = soup.get_text(" ", strip=True)[:6000]

        page_is_smc = bool(STRONG_SMC_RE.search(urlparse(final).path)) or \
            bool(STRONG_SMC_RE.search(title)) or \
            len(STRONG_SMC_RE.findall(body_text)) >= 3
        if page_is_smc and final not in smc_pages:
            smc_pages[final] = {"title": title, "members": extract_members(soup)}

        for a in soup.find_all("a", href=True):
            href = urljoin(final, a["href"]).split("#")[0]
            text = a.get_text(" ", strip=True)
            if is_pdf(href):
                if SMC_RE.search(href) or SMC_RE.search(text) or page_is_smc:
                    pdfs[href] = {"text": text or os.path.basename(urlparse(href).path),
                                  "found_on": final}
                continue
            if urlparse(href).netloc.replace("www.", "") != host:
                continue
            if href.startswith("http") and depth < MAX_DEPTH:
                if SMC_RE.search(href) or SMC_RE.search(text):
                    queue.appendleft((href, depth + 1))
                else:
                    queue.append((href, depth + 1))

    result["smc_pages"] = [(u, d["title"]) for u, d in smc_pages.items()]
    result["members"] = [(u, n, r) for u, d in smc_pages.items() for n, r in d["members"]]
    result["pdfs"] = [(u, d["text"], d["found_on"]) for u, d in pdfs.items()]
    result["status"] = "SMC found" if (smc_pages or pdfs) else "No SMC info found on website"
    return result


SKIP_EXT = (".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".ico", ".zip", ".rar", ".mp4",
            ".mp3", ".avi", ".css", ".js", ".xml", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx")


def crawl_full(school, max_pages=300, max_depth=4, progress=None):
    """Scan the WHOLE website: every page (up to max_pages) and every PDF.
    SMC pages, SMC members and SMC-related PDFs are picked out of everything found."""
    result = {**school, "smc_pages": [], "members": [], "pdfs": [],
              "all_pages": [], "all_pdfs": [], "status": ""}
    base = school["website"]
    if not base:
        result["status"] = "No website found"
        return result
    root, html = fetch(base)
    if not root:
        result["status"] = "Website unreachable"
        return result
    host = urlparse(root).netloc.replace("www.", "")

    seen = {root}
    frontier = [root]
    for pth in COMMON_PATHS:
        u = urljoin(root, pth)
        if u not in seen:
            seen.add(u)
            frontier.append(u)
    pages, smc_pages, all_pdfs = {}, {}, {}
    depth = 0
    while frontier and len(pages) < max_pages and depth <= max_depth:
        todo = frontier[: max_pages - len(pages)]
        frontier = []
        with ThreadPoolExecutor(max_workers=5) as ex:
            fetched = list(ex.map(lambda u: (u, *((root, html) if u == root else fetch(u))), todo))
        for u, final, page in fetched:
            if not page:
                continue
            final = final or u
            if final in pages:
                continue
            soup = BeautifulSoup(page, "lxml")
            title = soup.title.get_text(" ", strip=True) if soup.title else ""
            pages[final] = title
            body_text = soup.get_text(" ", strip=True)[:6000]
            page_is_smc = bool(STRONG_SMC_RE.search(urlparse(final).path)) or \
                bool(STRONG_SMC_RE.search(title)) or \
                len(STRONG_SMC_RE.findall(body_text)) >= 3
            if page_is_smc:
                smc_pages[final] = {"title": title, "members": extract_members(soup)}
            for a in soup.find_all("a", href=True):
                href = urljoin(final, a["href"].strip()).split("#")[0]
                if not href.startswith("http"):
                    continue
                text = a.get_text(" ", strip=True)
                if is_pdf(href):
                    related = bool(SMC_RE.search(href) or SMC_RE.search(text) or page_is_smc)
                    if href not in all_pdfs:
                        all_pdfs[href] = {"title": text or os.path.basename(urlparse(href).path),
                                          "found_on": final, "smc": related}
                    elif related:
                        all_pdfs[href]["smc"] = True
                    continue
                if urlparse(href).netloc.replace("www.", "") != host:
                    continue
                if urlparse(href).path.lower().endswith(SKIP_EXT):
                    continue
                if href not in seen:
                    seen.add(href)
                    frontier.append(href)
            if progress:
                progress(len(pages), len(all_pdfs))
        frontier.sort(key=lambda x: 0 if SMC_RE.search(x) else 1)   # SMC-looking links first
        depth += 1

    result["all_pages"] = [(u, t) for u, t in pages.items()]
    result["all_pdfs"] = [(u, d["title"], d["found_on"], d["smc"]) for u, d in all_pdfs.items()]
    result["smc_pages"] = [(u, d["title"]) for u, d in smc_pages.items()]
    result["members"] = [(u, n, r) for u, d in smc_pages.items() for n, r in d["members"]]
    result["pdfs"] = [(u, d["title"], d["found_on"]) for u, d in all_pdfs.items() if d["smc"]]
    result["status"] = "SMC found" if (smc_pages or result["pdfs"]) else "No SMC info found on website"
    return result


# ----------------------------------------------------------------------------
# 3. Excel output
# ----------------------------------------------------------------------------
def style_sheet(ws, widths):
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.hyperlink:
                c.font = Font(name="Arial", color="0563C1", underline="single")
            else:
                c.font = Font(name="Arial")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def link(ws, row, col, url):
    if url:
        c = ws.cell(row=row, column=col, value=url)
        c.hyperlink = url


def save_excel(results, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Schools"
    ws.append(["School Name", "Address", "Website", "SMC Page URL(s)",
               "SMC Members Found", "SMC PDFs Found", "Status", "Data Source"])
    for r in results:
        ws.append([r["name"], r["address"], "", "", len(r["members"]), len(r["pdfs"]),
                   r["status"], r["source"]])
        row = ws.max_row
        link(ws, row, 3, r["website"])
        if r["smc_pages"]:
            ws.cell(row=row, column=4, value="\n".join(u for u, _ in r["smc_pages"]))
            ws.cell(row=row, column=4).hyperlink = r["smc_pages"][0][0]
            ws.cell(row=row, column=4).alignment = Alignment(wrap_text=True)
    style_sheet(ws, [40, 40, 38, 55, 14, 14, 28, 16])

    wm = wb.create_sheet("SMC Members")
    wm.append(["School Name", "Member Name", "Designation / Role", "Source Page"])
    for r in results:
        for url, name, role in r["members"]:
            wm.append([r["name"], name, role, ""])
            link(wm, wm.max_row, 4, url)
    style_sheet(wm, [40, 32, 32, 60])

    wp = wb.create_sheet("SMC PDFs")
    wp.append(["School Name", "PDF Title / Link Text", "PDF Download Link", "Found On Page"])
    for r in results:
        for url, text, found_on in r["pdfs"]:
            wp.append([r["name"], text, "", ""])
            link(wp, wp.max_row, 3, url)
            link(wp, wp.max_row, 4, found_on)
    style_sheet(wp, [40, 40, 70, 60])

    wb.save(path)
    print(f"Saved -> {path}")


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="School SMC scraper")
    ap.add_argument("--school", default=None, help='school name, e.g. "St. Xavier\'s High School"')
    ap.add_argument("--city", default=None, help="location / city of the school (or city to scan)")
    ap.add_argument("--url", default=None, help="school website link (finds SMC page + PDFs)")
    ap.add_argument("--input", default=None,
                    help="Excel/CSV file with columns: School Name, Location (and optional Website)")
    ap.add_argument("--out", default=None)
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--no-search", action="store_true",
                    help="city mode only: skip website search for schools without a website")
    args = ap.parse_args()

    # ---- decide mode -------------------------------------------------------
    if not (args.school or args.input or args.city or args.url):
        print("1) Look up a specific school (name + location)")
        print("2) Scan ALL schools in a city")
        mode = input("Choose 1 or 2: ").strip()
        if mode == "1":
            args.school = input("School name: ").strip()
            args.city = input("Location (area/city): ").strip()
        else:
            args.city = input("City name: ").strip()

    if args.url:                                     # MODE U: website link given
        schools = [school_from_url(args.url, args.school or "")]
        out = args.out or re.sub(r"\W+", "_", schools[0]["name"].lower()).strip("_")[:40] + "_smc.xlsx"
        city_mode = False
    elif args.input:                                 # MODE A: list of schools from a file
        schools = load_input_file(args.input)
        print(f"{len(schools)} schools loaded from {args.input}")
        out = args.out or "smc_results.xlsx"
        city_mode = False
    elif args.school:                                # MODE B: one school
        schools = [{"name": args.school, "location": args.city or "", "address": args.city or "",
                    "website": "", "source": "User input"}]
        out = args.out or re.sub(r"\W+", "_", args.school.lower()).strip("_")[:40] + "_smc.xlsx"
        city_mode = False
    else:                                            # MODE C: whole city
        schools = schools_from_osm(args.city)
        key = os.getenv("GOOGLE_PLACES_API_KEY")
        if key:
            schools += schools_from_google(args.city, key)
        schools = dedupe(schools)
        for s in schools:
            s["location"] = args.city
        out = args.out or f"{args.city.lower().replace(' ', '_')}_smc.xlsx"
        city_mode = True

    if not schools:
        print("No schools found. Check the spelling or try again in a few minutes.")
        return
    print(f"{len(schools)} schools to process")

    # ---- find websites -----------------------------------------------------
    if not (city_mode and args.no_search):
        missing = [s for s in schools if not s["website"]]
        print(f"Searching websites for {len(missing)} schools...")
        for i, s in enumerate(missing, 1):
            s["website"] = search_website(s["name"], s.get("location") or args.city or "")
            print(f"  [{i}/{len(missing)}] {s['name']} -> {s['website'] or 'not found'}")
            time.sleep(1.5)

    # ---- crawl for SMC page / members / PDFs -------------------------------
    results = []
    try:
        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            futs = {ex.submit(crawl_school, s): s for s in schools}
            for i, f in enumerate(as_completed(futs), 1):
                try:
                    r = f.result()
                except Exception as e:
                    r = {**futs[f], "smc_pages": [], "members": [], "pdfs": [],
                         "status": f"Error: {e}"}
                results.append(r)
                print(f"[{i}/{len(schools)}] {r['name']} -> {r['status']}")
                if i % 100 == 0:
                    save_excel(sorted(results, key=lambda x: x["name"].lower()), out)
    except KeyboardInterrupt:
        print("\nStopped by user - saving what we have so far...")
    results.sort(key=lambda r: r["name"].lower())
    save_excel(results, out)


if __name__ == "__main__":
    main()
