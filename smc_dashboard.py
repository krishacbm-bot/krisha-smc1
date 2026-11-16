#!/usr/bin/env python3
"""
School SMC Dashboard  (runs in your browser)

Start it (double-click the .exe, or run python smc_dashboard.py). No window stays open:
the page opens in your browser and Python runs quietly behind it.
Use the "Quit app" button in the page to stop it (it also stops by itself a few
minutes after the browser tab is closed).
"""
import io
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
import zipfile
import socket
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse


# When started without a console window, stdout/stderr do not exist: make printing safe.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")


SERVER_MODE = bool(os.environ.get("PORT") or os.environ.get("SMC_SERVER"))   # set on Render etc.


def ensure_packages():
    if getattr(sys, "frozen", False) or SERVER_MODE:      # packaged .exe or hosted server
        return
    import importlib.util
    need = {"flask": "flask", "requests": "requests", "bs4": "beautifulsoup4",
            "openpyxl": "openpyxl", "lxml": "lxml"}
    missing = [p for m, p in need.items() if importlib.util.find_spec(m) is None]
    if missing:
        print("Installing:", ", ".join(missing))
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *missing])


ensure_packages()

import requests  # noqa: E402
from flask import Flask, Response, jsonify, request, send_file  # noqa: E402

import smc_scraper as sc  # noqa: E402

for _n in ("school_from_url", "search_website", "crawl_school", "save_excel"):
    if not hasattr(sc, _n):
        sys.exit(f"smc_scraper.py is an old version (missing {_n}). Replace it with the latest one.")

import logging  # noqa: E402
logging.getLogger("werkzeug").setLevel(logging.ERROR)

app = Flask(__name__)
STATE = {"last_ping": time.time(), "seen": False}
JOBS = {}
ALLOWED_PDFS = set()          # only PDFs found by a scan can be downloaded through this app
OUT_DIR = tempfile.mkdtemp(prefix="smc_")


# ----------------------------------------------------------------------------
# Scan job
# ----------------------------------------------------------------------------
def run_scan(job_id, data):
    job = JOBS[job_id]
    try:
        email = (data.get("email") or "").strip()
        if "@" in email:
            sc.OVERPASS_HEADERS["User-Agent"] = f"SchoolSMCScraper/1.0 (contact: {email})"
            sc.session.headers["User-Agent"] = f"Mozilla/5.0 (compatible; SchoolSMCScraper/1.0; {email})"
        name = (data.get("name") or "").strip()
        if data.get("mode") == "url":
            job["step"] = "Opening the website ..."
            school = sc.school_from_url(data.get("url", ""), name)
        else:
            loc = (data.get("location") or "").strip()
            job["step"] = "Searching for the school's website ..."
            school = {"name": name, "location": loc, "address": loc, "website": "",
                      "source": "User input"}
            school["website"] = sc.search_website(name, loc)
        full = bool(data.get("full"))
        if full:
            def prog(n, p):
                job["step"] = f"Scanning the whole website ... {n} pages checked, {p} PDFs found"
            job["step"] = "Scanning the whole website ..."
            r = sc.crawl_full(school, progress=prog)
        else:
            job["step"] = "Looking for the SMC page, members and PDFs ..."
            r = sc.crawl_school(school)
        xlsx = os.path.join(OUT_DIR, f"{job_id}.xlsx")
        sc.save_excel([r], xlsx)
        for u, _t, _f in r["pdfs"]:
            ALLOWED_PDFS.add(u)
        if full:
            for u, _t, _f, _s in r["all_pdfs"]:
                ALLOWED_PDFS.add(u)
            from openpyxl import load_workbook
            wb = load_workbook(xlsx)
            w1 = wb.create_sheet("All Pages")
            w1.append(["Page title", "URL"])
            for u, t in r["all_pages"]:
                w1.append([t, u])
            w2 = wb.create_sheet("All PDFs")
            w2.append(["PDF title", "PDF link", "Found on page", "SMC related"])
            for u, t, f, smc in r["all_pdfs"]:
                w2.append([t, u, f, "Yes" if smc else "No"])
            for w, widths in ((w1, (60, 80)), (w2, (50, 80, 60, 12))):
                for i, wd in enumerate(widths):
                    w.column_dimensions["ABCD"[i]].width = wd
            wb.save(xlsx)
        job["result"] = {
            "name": r["name"], "website": r["website"], "status": r["status"],
            "smc_pages": [{"url": u, "title": t} for u, t in r["smc_pages"]],
            "members": [{"page": u, "name": n, "role": ro} for u, n, ro in r["members"]],
            "pdfs": [{"url": u, "title": t, "found_on": f} for u, t, f in r["pdfs"]],
        }
        if full:
            job["result"]["all_pages"] = [{"url": u, "title": t} for u, t in r["all_pages"]]
            job["result"]["all_pdfs"] = [{"url": u, "title": t, "found_on": f, "smc": smc}
                                         for u, t, f, smc in r["all_pdfs"]]
        job["state"] = "done"
    except Exception as e:  # show a readable message in the page
        job["state"] = "error"
        job["error"] = str(e)


@app.post("/api/scan")
def api_scan():
    data = request.get_json(force=True, silent=True) or {}
    if data.get("mode") == "url" and not (data.get("url") or "").strip():
        return jsonify(error="Paste the school website link first."), 400
    if data.get("mode") != "url" and not (data.get("name") or "").strip():
        return jsonify(error="Type the school name first."), 400
    job_id = uuid.uuid4().hex[:10]
    JOBS[job_id] = {"state": "running", "step": "Starting ..."}
    threading.Thread(target=run_scan, args=(job_id, data), daemon=True).start()
    return jsonify(id=job_id)


# ----------------------------------------------------------------------------
# Bulk scan (Excel / CSV list of schools)
# ----------------------------------------------------------------------------
def set_email(email):
    email = (email or "").strip()
    if "@" in email:
        sc.OVERPASS_HEADERS["User-Agent"] = f"SchoolSMCScraper/1.0 (contact: {email})"
        sc.session.headers["User-Agent"] = f"Mozilla/5.0 (compatible; SchoolSMCScraper/1.0; {email})"


def run_bulk(job_id, path, preset=None):
    job = JOBS[job_id]
    try:
        if preset:
            headers, schools = preset
        else:
            headers, rows_in = read_table(path)
            schools = to_schools(headers, rows_in)
        if not schools:
            raise ValueError("No schools found. The file needs a 'School Name' column and a 'Website' column.")
        total = len(schools)
        job.update(total=total, done=0, step=f"Checked 0 of {total} schools ...")

        def work(s):
            return sc.crawl_school(s)      # schools without a website are reported, never guessed

        results = []
        with ThreadPoolExecutor(max_workers=6) as ex:
            futs = {ex.submit(work, s): s for s in schools}
            for f in as_completed(futs):
                try:
                    r = f.result()
                except Exception as e:
                    r = {**futs[f], "smc_pages": [], "members": [], "pdfs": [], "status": f"Error: {e}"}
                results.append(r)
                job["done"] = len(results)
                job["step"] = f"Checked {len(results)} of {total} schools ..."
        results.sort(key=lambda r: r["name"].lower())
        xl = os.path.join(OUT_DIR, f"{job_id}.xlsx")
        sc.save_excel(results, xl)
        add_final_sheet(xl, results, headers)
        job["download_name"] = "SMC_results.xlsx"
        for r in results:
            for u, _t, _f in r["pdfs"]:
                ALLOWED_PDFS.add(u)
        job["results"] = [{
            "name": r["name"], "website": r["website"], "status": r["status"],
            "smc_pages": [{"url": u, "title": t} for u, t in r["smc_pages"]],
            "members": len(r["members"]),
            "pdfs": [{"url": u, "title": t} for u, t, _f in r["pdfs"]],
        } for r in results]
        job["state"] = "done"
    except Exception as e:
        job["state"] = "error"
        job["error"] = str(e)


@app.post("/api/bulk")
def api_bulk():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="Choose your Excel or CSV file first."), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return jsonify(error="The file must be an Excel (.xlsx) or CSV file."), 400
    set_email(request.form.get("email"))
    job_id = uuid.uuid4().hex[:10]
    path = os.path.join(OUT_DIR, job_id + ext)
    f.save(path)
    JOBS[job_id] = {"state": "running", "step": "Reading your file ...", "done": 0, "total": 0}
    threading.Thread(target=run_bulk, args=(job_id, path), daemon=True).start()
    return jsonify(id=job_id)


# ----------------------------------------------------------------------------
# Excel helpers shared by the 3 steps
# ----------------------------------------------------------------------------
def read_table(path):
    """Read an Excel/CSV file -> (headers, list of row-dicts)."""
    import csv
    if path.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        rows = [list(r) for r in wb.active.iter_rows(values_only=True)]
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    rows = [r for r in rows if r and any(c not in (None, "") for c in r)]
    if not rows:
        return [], []
    headers = [str(h).strip() if h not in (None, "") else f"Column {i + 1}" for i, h in enumerate(rows[0])]
    body = []
    for r in rows[1:]:
        body.append({h: ("" if (i >= len(r) or r[i] is None) else str(r[i]).strip())
                     for i, h in enumerate(headers)})
    return headers, body


def is_web_header(h):
    h = h.lower()
    return "website" in h or h in ("url", "site", "web", "link")


def to_schools(headers, rows):
    """Turn table rows into school records, keeping every original column in 'extra'."""
    low = {h: h.lower() for h in headers}
    name_h = next((h for h in headers if "school" in low[h] and "name" in low[h]), None) \
        or next((h for h in headers if "name" in low[h]), headers[0] if headers else None)
    web_h = next((h for h in headers if is_web_header(h)), None)
    aff_h = next((h for h in headers if "affiliation" in low[h]), None)
    loc_h = [h for h in headers if any(k in low[h] for k in ("district", "city"))] or \
            [h for h in headers if any(k in low[h] for k in ("location", "address", "area", "place"))]
    schools = []
    for r in rows:
        name = r.get(name_h, "") if name_h else ""
        if not name:
            continue
        loc = ", ".join(r[h] for h in loc_h if r.get(h))[:120]
        addr = ", ".join(r[h] for h in headers if any(k in low[h] for k in ("address", "location")) and r.get(h))
        schools.append({"name": name, "location": loc, "address": addr or loc,
                        "website": sc.normalize_url(r.get(web_h, "")) if web_h else "",
                        "aff": r.get(aff_h, "") if aff_h else "",
                        "source": "Uploaded file", "extra": dict(r)})
    return schools


def write_xlsx(path, title, headers, rows):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    sc.style_sheet(ws, [min(max(len(str(h)) + 6, 16), 50) for h in headers])
    wb.save(path)


def add_final_sheet(xlsx, results, headers):
    """First sheet of the step-3 Excel: the uploaded columns + website + SMC findings, one row per school."""
    from openpyxl import load_workbook
    base = [h for h in headers if not is_web_header(h)]
    wb = load_workbook(xlsx)
    ws = wb.create_sheet("Final (all data)", 0)
    ws.append(base + ["Website", "SMC Page(s)", "SMC Members Found", "SMC PDF Links", "Status"])
    for r in results:
        ex = r.get("extra", {})
        ws.append([ex.get(h, "") for h in base] + [
            r["website"], "\n".join(u for u, _ in r["smc_pages"]), len(r["members"]),
            "\n".join(u for u, _t, _f in r["pdfs"]), r["status"]])
    sc.style_sheet(ws, [min(max(len(h) + 6, 16), 40) for h in base] + [38, 55, 14, 60, 28])
    wb.active = 0
    wb.save(xlsx)


def table_result(job, kind, headers, rows):
    job["table"] = {"kind": kind, "headers": headers, "total": len(rows),
                    "rows": [[str(c) for c in r] for r in rows[:200]]}


def run_list(job_id, url, websites=False, smc=False):
    job = JOBS[job_id]
    try:
        try:
            import scrape_cbse_schools as lister
        except ImportError as e:
            if "playwright" in str(e).lower() or "pandas" in str(e).lower():
                raise ValueError("Step 1 needs one-time extras. Run:  pip install playwright pandas   "
                                 "and then:  playwright install chromium")
            raise ValueError("scrape_cbse_schools.py must be in the same folder as this app.")
        job["step"] = "Opening the school list ..."
        def prog(n, t, phase="list"):
            job["step"] = (f"Reading CBSE records ... {n} of {t} schools" if phase == "sites"
                           else f"Reading page {n} ... {t} schools so far")
        df = lister.scrape(url, progress=prog, websites=websites)
        if len(df) == 0:
            raise ValueError("No schools were found on that page. Step 1 understands schoolsindia.net list pages.")
        headers = [str(c) for c in df.columns]
        rows = [["" if v is None else v for v in r] for r in df.values.tolist()]
        write_xlsx(os.path.join(OUT_DIR, f"{job_id}.xlsx"), "Schools", headers, rows)
        job["download_name"] = "school_list.xlsx"
        if smc and "Website" in headers:
            recs = df.fillna("").astype(str).to_dict("records")
            run_bulk(job_id, None, preset=(headers, to_schools(headers, recs)))
            return
        table_result(job, "list", headers, rows)
        if "Website" in headers:
            w = headers.index("Website")
            n = sum(1 for r in rows if r[w])
            job["table"]["note"] = (f"{n} of {len(rows)} schools have a website in the official CBSE record "
                                    "(the others declared none). You can go straight to step 3.")
        job["state"] = "done"
    except Exception as e:
        msg = str(e)
        if "playwright install" in msg or "Executable doesn't exist" in msg:
            msg = "The browser part is missing. Run once:  playwright install chromium"
        job["state"] = "error"
        job["error"] = msg


def run_sites(job_id, path, key="", use_search=False):
    job = JOBS[job_id]
    try:
        sc.SERPER_KEY = key or os.getenv("SERPER_API_KEY", "")
        sc._blocked_until.clear()
        headers, rows = read_table(path)
        schools = to_schools(headers, rows)
        if not schools:
            raise ValueError("No schools found. The file needs a 'School Name' column.")
        total = len(schools)
        job.update(total=total, done=0, step=f"Finding websites ... 0 of {total}")

        def work(s):
            if s["website"]:
                s["note"] = "From your file"
                return
            s["website"], s["note"] = sc.cbse_website(s.get("aff", ""))      # official, most reliable
            if not s["website"] and use_search:
                site, note = sc.search_website_ex(s["name"], s.get("location", ""), s.get("address", ""))
                s["website"] = site
                s["note"] = f"{s['note']}; web search: {note}"

        done = 0
        with ThreadPoolExecutor(max_workers=3) as ex:
            for _ in as_completed([ex.submit(work, s) for s in schools]):
                done += 1
                job["done"] = done
                job["step"] = f"Finding websites ... {done} of {total}"
        base = [h for h in headers if not is_web_header(h)]
        out_headers = base + ["Website", "Website Check"]
        out_rows = [[s["extra"].get(h, "") for h in base] + [s["website"], s.get("note", "")] for s in schools]
        write_xlsx(os.path.join(OUT_DIR, f"{job_id}.xlsx"), "Schools with websites", out_headers, out_rows)
        job["download_name"] = "schools_with_websites.xlsx"
        table_result(job, "sites", out_headers, out_rows)
        found = sum(1 for s in schools if s["website"])
        affs = [s["aff"] for s in schools if s.get("aff")]
        dup_warn = ""
        if affs and len(affs) - len(set(affs)) > 0.3 * len(affs):
            dup_warn = ("WARNING: many schools in your file share the same Affiliation No. - the school list (step 1) "
                        "looks wrong, so these websites cannot be trusted. Run step 1 again with the newest files. ")
        job["table"]["note"] = dup_warn + f"Websites found for {found} of {total} schools. " + (
            "Websites marked 'From CBSE official record' were declared by the schools themselves. "
            "Blank rows have no website on record (or no confident web match), so nothing wrong was filled in." if found else
            "None were found. Check that the file has an 'Affiliation No.' column and that this PC can open saras.cbse.gov.in.")
        job["state"] = "done"
    except Exception as e:
        job["state"] = "error"
        job["error"] = str(e)


def save_upload(prefix_ok=(".xlsx", ".xlsm", ".csv")):
    f = request.files.get("file")
    if not f or not f.filename:
        return None, (jsonify(error="Choose your Excel or CSV file first."), 400)
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in prefix_ok:
        return None, (jsonify(error="The file must be an Excel (.xlsx) or CSV file."), 400)
    job_id = uuid.uuid4().hex[:10]
    path = os.path.join(OUT_DIR, job_id + ext)
    f.save(path)
    return (job_id, path), None


@app.post("/api/list")
def api_list():
    data = request.get_json(force=True, silent=True) or {}
    url = (data.get("url") or "").strip()
    if not url.startswith(("http://", "https://")):
        return jsonify(error="Paste the full link of the school list page (starting with https://)."), 400
    job_id = uuid.uuid4().hex[:10]
    JOBS[job_id] = {"state": "running", "step": "Starting ..."}
    smc = bool(data.get("smc"))
    threading.Thread(target=run_list, args=(job_id, url, bool(data.get("websites")) or smc, smc),
                     daemon=True).start()
    return jsonify(id=job_id)


@app.post("/api/sites")
def api_sites():
    up, err = save_upload()
    if err:
        return err
    job_id, path = up
    JOBS[job_id] = {"state": "running", "step": "Reading your file ...", "done": 0, "total": 0}
    key = (request.form.get("serper_key") or "").strip()
    use_search = request.form.get("use_search") == "1"
    threading.Thread(target=run_sites, args=(job_id, path, key, use_search), daemon=True).start()
    return jsonify(id=job_id)


@app.get("/api/job/<job_id>")
def api_job(job_id):
    return jsonify(JOBS.get(job_id) or {"state": "error", "error": "Unknown scan."})


def safe_filename(url):
    base = os.path.basename(urlparse(url).path) or "document.pdf"
    base = re.sub(r"[^\w.\- ]", "_", base)
    return base if base.lower().endswith(".pdf") else base + ".pdf"


@app.get("/download")
def download():
    url = request.args.get("url", "")
    if url not in ALLOWED_PDFS:
        return "This file was not found by a scan.", 403
    try:
        up = requests.get(url, headers=sc.HEADERS, stream=True, timeout=30)
        up.raise_for_status()
    except Exception as e:
        return f"Could not download the file: {e}", 502
    return Response(up.iter_content(65536), content_type="application/pdf",
                    headers={"Content-Disposition": f'attachment; filename="{safe_filename(url)}"'})


@app.get("/excel/<job_id>")
def excel(job_id):
    path = os.path.join(OUT_DIR, f"{re.sub(r'[^a-z0-9]', '', job_id)}.xlsx")
    if not os.path.exists(path):
        return "Not found", 404
    return send_file(path, as_attachment=True,
                     download_name=JOBS.get(job_id, {}).get("download_name", "SMC_results.xlsx"))


@app.get("/zip/<job_id>")
def zip_all(job_id):
    job = JOBS.get(job_id)
    if not job or "result" not in job:
        return "Not found", 404
    buf, used = io.BytesIO(), set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        lst = job["result"].get("all_pdfs") if request.args.get("all") else job["result"]["pdfs"]
        for p in (lst or [])[:100]:
            try:
                data = requests.get(p["url"], headers=sc.HEADERS, timeout=30).content
            except Exception:
                continue
            fn = safe_filename(p["url"])
            while fn in used:
                fn = "x_" + fn
            used.add(fn)
            z.writestr(fn, data)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="SMC_PDFs.zip", mimetype="application/zip")


@app.get("/api/ping")
def api_ping():
    STATE["last_ping"] = time.time()
    STATE["seen"] = True
    return jsonify(app="smc-finder")


_PIXEL = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000b49444154789c6360000200000500017a5eab3f0000000049454e44ae426082")


@app.get("/api/ping.png")
def api_ping_png():
    """Tiny image the start page loads to detect the app (images work even when fetch is blocked)."""
    STATE["last_ping"] = time.time()
    STATE["seen"] = True
    return Response(_PIXEL, mimetype="image/png", headers={"Cache-Control": "no-store"})


@app.after_request
def allow_start_page(resp):
    """Lets start.html (opened from a file) check whether the app is already running."""
    if request.path == "/api/ping":
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Private-Network"] = "true"
    return resp


@app.before_request
def require_password():
    """Optional: set APP_PASSWORD on the server to protect the site (any username)."""
    pw = os.environ.get("APP_PASSWORD")
    if not pw or request.path == "/healthz":
        return None
    a = request.authorization
    if a and a.password == pw:
        return None
    return Response("Password required", 401, {"WWW-Authenticate": 'Basic realm="SMC Finder"'})


@app.get("/healthz")
def healthz():
    return "ok"


@app.post("/api/quit")
def api_quit():
    if SERVER_MODE:
        return jsonify(ok=False, error="Quit is disabled on the hosted version.")
    threading.Timer(0.6, lambda: os._exit(0)).start()
    return jsonify(ok=True)


def watchdog():
    """Stop quietly when the browser tab has been closed (and no scan is running)."""
    while True:
        time.sleep(10)
        busy = any(j.get("state") == "running" for j in JOBS.values())
        if STATE["seen"] and not busy and time.time() - STATE["last_ping"] > 180:
            os._exit(0)


@app.get("/")
def index():
    page = PAGE
    if SERVER_MODE:     # no "Quit app" button on the hosted version
        page = page.replace('id="quit"', 'id="quit" style="display:none"', 1)
    return Response(page, mimetype="text/html")


# ----------------------------------------------------------------------------
# Front end (HTML + CSS + JS)
# ----------------------------------------------------------------------------
PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>School SMC Finder</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:wght@600;700&family=Source+Sans+3:wght@400;600&display=swap" rel="stylesheet">
<style>
:root{--bg:#f4f6f3;--card:#fff;--ink:#17261f;--muted:#5d6b64;--line:#d9e0db;--accent:#1d5a42;--accent-ink:#fff;--mark:#f0b43c;--bad:#a4332b;
--head:"Bricolage Grotesque","Segoe UI",system-ui,sans-serif;--body:"Source Sans 3","Segoe UI",system-ui,sans-serif}
@media(prefers-color-scheme:dark){:root{--bg:#101a15;--card:#18251e;--ink:#e7efe9;--muted:#9bada3;--line:#2c3d34;--accent:#6fd0a0;--accent-ink:#0b1711;--bad:#ff8f84}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 var(--body)}
main{max-width:980px;margin:0 auto;padding:32px 20px 64px}
.top{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}
h1{font:700 clamp(28px,4.5vw,42px)/1.1 var(--head);margin:0 0 8px;letter-spacing:-.01em}
h2{font:600 20px/1.2 var(--head);margin:0 0 12px}
.lead{color:var(--muted);margin:0 0 24px;max-width:60ch}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:20px;margin-bottom:20px}
.seg{display:inline-flex;flex-wrap:wrap;border:1px solid var(--line);border-radius:8px;overflow:hidden;margin-bottom:16px}
.seg button{border:0;background:none;color:var(--ink);padding:9px 16px;font:600 15px var(--body);cursor:pointer}
.seg button[aria-pressed=true]{background:var(--accent);color:var(--accent-ink)}
.row{display:flex;gap:12px;flex-wrap:wrap}
.field{flex:1 1 260px;display:flex;flex-direction:column;gap:4px}
label{font-weight:600;font-size:14px}
input{padding:11px 12px;border:1px solid var(--line);border-radius:7px;font:16px var(--body);background:var(--bg);color:var(--ink);width:100%}
input:focus-visible,button:focus-visible,a:focus-visible{outline:3px solid var(--mark);outline-offset:2px}
.btn{display:inline-block;border:0;border-radius:7px;background:var(--accent);color:var(--accent-ink);padding:11px 20px;font:600 16px var(--body);cursor:pointer;text-decoration:none}
.btn.alt{background:none;color:var(--accent);border:1px solid var(--accent)}
.btn.sm{padding:7px 13px;font-size:14px}
.btn:disabled{opacity:.55;cursor:wait}
.hint{color:var(--muted);font-size:14px;margin:8px 0 0}
.err{color:var(--bad);font-weight:600;margin:12px 0 0}
.status{display:flex;align-items:center;gap:12px;color:var(--muted)}
.spin{width:20px;height:20px;border:3px solid var(--line);border-top-color:var(--accent);border-radius:50%;animation:s 1s linear infinite}
@keyframes s{to{transform:rotate(360deg)}}
@media(prefers-reduced-motion:reduce){.spin{animation-duration:3s}}
.head{display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;align-items:flex-start}
.head a.site{color:var(--accent);word-break:break-all}
.badge{display:inline-block;padding:3px 10px;border-radius:99px;font-size:14px;font-weight:600;background:var(--mark);color:#2b2000}
.badge.ok{background:var(--accent);color:var(--accent-ink)}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:16px 0 0}
.stat{border-left:4px solid var(--mark);padding:2px 12px}
.stat b{display:block;font:700 28px/1.1 var(--head)}
.stat span{color:var(--muted);font-size:14px}
.list{list-style:none;margin:0;padding:0}
.list li{display:flex;justify-content:space-between;gap:12px;align-items:center;padding:12px 0;border-top:1px solid var(--line);flex-wrap:wrap}
.list li:first-child{border-top:0}
.list .t{flex:1 1 320px;min-width:0}
.list .t strong{display:block;overflow-wrap:anywhere}
.list .t small{color:var(--muted);overflow-wrap:anywhere}
.acts{display:flex;gap:8px}
.scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%}
th,td{text-align:left;padding:9px 10px;border-bottom:1px solid var(--line)}
th{font-size:14px;color:var(--muted)}
.empty{color:var(--muted);margin:0}
.chk{display:flex;gap:8px;align-items:center;margin:16px 0 0;font-weight:600;cursor:pointer}.chk input{width:auto}
.plain{list-style:none;margin:0;padding:0}.plain li{margin:0 0 6px}
td a{color:var(--accent);overflow-wrap:anywhere}
[hidden]{display:none!important}
</style></head><body><main>
<div class="top"><h1>School SMC Finder</h1><button class="btn alt sm" id="quit" title="Stop the app">Quit app</button></div>
<p class="lead">Paste a school's website link. We find its School Management Committee page, the members it lists, and any SMC PDF files you can download.</p>

<section class="card" aria-label="Search">
  <div class="seg" role="group" aria-label="Search by">
    <button id="m-list" aria-pressed="true">1. School list from a link</button>
    <button id="m-sites" aria-pressed="false">2. Find websites</button>
    <button id="m-file" aria-pressed="false">3. Find SMC pages &amp; PDFs</button>
    <button id="m-url" aria-pressed="false">One website link</button>
    <button id="m-name" aria-pressed="false">One school name</button>
  </div>
  <div class="row" id="f-list">
    <div class="field"><label for="listurl">Link of the school list page</label>
      <input id="listurl" type="text" placeholder="https://www.schoolsindia.net/cbse-schools-in-mumbai/district" autocomplete="off">
      <label class="chk"><input type="checkbox" id="getsites" checked> Also get each school's website from the official CBSE record (slower, but then you can skip step 2)</label>
      <label class="chk"><input type="checkbox" id="getsmc" checked> Also find the SMC page and PDFs of every school that has a website (much slower: a big list can take hours)</label>
      <p class="hint">Gets every school on that list (name, address, district, affiliation no.) into an Excel file. Works with schoolsindia.net list pages. Needs a one-time extra install (see the README).</p></div>
  </div>
  <div class="row" id="f-sites" hidden>
    <div class="field"><label for="sfile">Excel file from step 1</label>
      <input id="sfile" type="file" accept=".xlsx,.xlsm,.csv">
      <label class="chk"><input type="checkbox" id="usesearch"> Also search the web where CBSE has no website listed (less reliable - check those rows)</label>
      <label for="skey" style="margin-top:10px">Search API key (optional, only used with the web search above)</label>
      <input id="skey" type="text" placeholder="Free key from serper.dev - makes this step much more reliable" autocomplete="off">
      <p class="hint">Reads each school's <b>official CBSE record</b> (using the Affiliation No. in your file) and copies the website the school itself declared there. No guessing. Schools that never declared one stay blank.</p></div>
  </div>
  <div class="row" id="f-url" hidden>
    <div class="field"><label for="url">School website link</label>
      <input id="url" type="text" placeholder="https://www.myschool.edu.in" autocomplete="off"></div>
    <div class="field"><label for="uname">School name (optional)</label>
      <input id="uname" type="text" placeholder="Guessed from the site if empty"></div>
  </div>
  <div class="row" id="f-name" hidden>
    <div class="field"><label for="name">School name</label>
      <input id="name" type="text" placeholder="St. Xavier's High School"></div>
    <div class="field"><label for="loc">Location (area and city)</label>
      <input id="loc" type="text" placeholder="Andheri West, Mumbai"></div>
  </div>
  <div class="row" id="f-file" hidden>
    <div class="field"><label for="file">Excel or CSV file</label>
      <input id="file" type="file" accept=".xlsx,.xlsm,.csv">
      <p class="hint">Upload the Excel from step 2 (columns <b>School Name</b> and <b>Website</b>; your other columns are kept). You get one Excel with every school, its location, website, SMC pages and PDF links.</p></div>
  </div>
  <label class="chk" id="full-wrap"><input type="checkbox" id="full"> Scan the whole website (finds every page and every PDF - slower)</label>
  <p style="margin:16px 0 0"><button class="btn" id="go">Find SMC page and PDFs</button></p>
  <p class="err" id="err" hidden></p>
</section>

<section class="card" id="busy" hidden><div class="status"><div class="spin"></div><span id="step">Starting ...</span></div></section>

<div id="tbl" hidden>
  <section class="card">
    <h2 id="t-title"></h2>
    <p class="hint" id="t-note"></p>
    <p style="margin:12px 0 0"><a class="btn sm" id="t-xlsx">Download Excel</a></p>
  </section>
  <section class="card"><div class="scroll"><table id="t-table"></table></div><p class="hint" id="t-more"></p></section>
</div>

<div id="bulk" hidden>
  <section class="card">
    <h2>Results for <span id="b-n"></span> schools</h2>
    <div class="stats">
      <div class="stat"><b id="b-site">0</b><span>With a website</span></div>
      <div class="stat"><b id="b-smc">0</b><span>With SMC page</span></div>
      <div class="stat"><b id="b-pdf">0</b><span>PDF files</span></div>
    </div>
    <p style="margin:16px 0 0"><a class="btn alt sm" id="b-xlsx">Download Excel</a></p>
    <div class="row" style="margin-top:14px">
      <div class="field"><input id="b-q" type="search" placeholder="Search school or website..."></div>
      <select id="b-f" aria-label="Filter"><option value="all">All schools</option><option value="smc">Has SMC page</option><option value="pdf">Has PDF</option><option value="none">Nothing found</option></select>
    </div>
  </section>
  <section class="card"><div class="scroll"><table id="b-table"></table></div></section>
</div>

<div id="out" hidden>
  <section class="card">
    <div class="head">
      <div><h2 id="r-name"></h2><a class="site" id="r-site" target="_blank" rel="noopener"></a></div>
      <span class="badge" id="r-status"></span>
    </div>
    <div class="stats">
      <div class="stat"><b id="n-pages">0</b><span>SMC pages</span></div>
      <div class="stat"><b id="n-members">0</b><span>Members listed</span></div>
      <div class="stat"><b id="n-pdfs">0</b><span>PDF files</span></div>
    </div>
    <p style="margin:16px 0 0" class="row">
      <a class="btn alt sm" id="dl-xlsx">Download Excel</a>
      <a class="btn alt sm" id="dl-zip" hidden>Download all PDFs (zip)</a>
    </p>
  </section>
  <section class="card"><h2>SMC pages</h2><ul class="list" id="pages"></ul></section>
  <section class="card"><h2>SMC members</h2><div class="scroll" id="members"></div></section>
  <section class="card"><h2>PDF files</h2><ul class="list" id="pdfs"></ul></section>
  <section class="card" id="allf-card" hidden>
    <div class="head"><h2>All PDFs on the website (<span id="allf-n"></span>)</h2><a class="btn alt sm" id="dl-zip-all">Download PDFs (zip, first 100)</a></div>
    <ul class="list" id="allf"></ul></section>
  <section class="card" id="allp-card" hidden><h2>All pages scanned (<span id="allp-n"></span>)</h2><ul class="list" id="allp"></ul></section>
</div>

<script>
const $=id=>document.getElementById(id);
let mode="list";
const LABELS={list:"Get the school list",sites:"Find the websites",file:"Find SMC pages and PDFs",url:"Find SMC page and PDFs",name:"Find SMC page and PDFs"};
function setMode(m){mode=m;["list","sites","file","url","name"].forEach(k=>{$("m-"+k).setAttribute("aria-pressed",m==k);$("f-"+k).hidden=m!=k});
  $("full-wrap").hidden=!(m=="url"||m=="name");$("go").textContent=LABELS[m];}
["list","sites","file","url","name"].forEach(k=>$("m-"+k).onclick=()=>setMode(k));setMode("list");
function el(tag,attrs={},text){const e=document.createElement(tag);for(const k in attrs)e.setAttribute(k,attrs[k]);if(text!=null)e.textContent=text;return e}
function showErr(m){$("err").textContent=m;$("err").hidden=!m}
function emptyMsg(t){return el("p",{class:"empty"},t)}

$("go").onclick=async()=>{
  showErr("");$("out").hidden=true;$("bulk").hidden=true;$("tbl").hidden=true;
  $("go").disabled=true;$("busy").hidden=false;$("step").textContent="Starting ...";
  try{
    let r;
    if(mode=="list"){
      r=await fetch("/api/list",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({url:$("listurl").value,websites:$("getsites").checked,smc:$("getsmc").checked})});
    }else if(mode=="sites"){
      const fd=new FormData();if($("sfile").files[0])fd.append("file",$("sfile").files[0]);fd.append("serper_key",$("skey").value);fd.append("use_search",$("usesearch").checked?"1":"0");
      r=await fetch("/api/sites",{method:"POST",body:fd});
    }else if(mode=="file"){
      const fd=new FormData();if($("file").files[0])fd.append("file",$("file").files[0]);
      r=await fetch("/api/bulk",{method:"POST",body:fd});
    }else{
      const body=mode=="url"?{mode,url:$("url").value,name:$("uname").value,full:$("full").checked}:{mode,name:$("name").value,location:$("loc").value,full:$("full").checked};
      r=await fetch("/api/scan",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});
    }
    const j=await r.json();
    if(!r.ok){throw new Error(j.error||"Something went wrong.")}
    poll(j.id);
  }catch(e){showErr(e.message);done()}
};
function done(){$("go").disabled=false;$("busy").hidden=true}
async function poll(id){
  try{
    const j=await (await fetch("/api/job/"+id)).json();
    if(j.state=="running"){$("step").textContent=j.step||"Working ...";setTimeout(()=>poll(id),1000);return}
    done();
    if(j.state=="error"){showErr("Scan failed: "+(j.error||"unknown error")+". Check the link and try again.");return}
    if(j.table){renderTable(j.table,id)}else if(j.results){renderBulk(j.results,id)}else{render(j.result,id)}
  }catch(e){done();showErr("Lost connection to the app. Is it still running?")}
}
function render(r,id){
  $("out").hidden=false;
  $("r-name").textContent=r.name;
  const a=$("r-site");a.textContent=r.website||"No website found";a.href=r.website||"#";
  const st=$("r-status");st.textContent=r.status;st.className="badge"+(r.smc_pages.length||r.pdfs.length?" ok":"");
  $("n-pages").textContent=r.smc_pages.length;$("n-members").textContent=r.members.length;$("n-pdfs").textContent=r.pdfs.length;
  $("dl-xlsx").href="/excel/"+id;
  $("dl-zip").hidden=!r.pdfs.length;$("dl-zip").href="/zip/"+id;

  const pages=$("pages");pages.replaceChildren();
  if(!r.smc_pages.length)pages.append(el("li",{},"")),pages.firstChild.replaceChildren(emptyMsg("No SMC page was found on this website."));
  r.smc_pages.forEach(p=>{const li=el("li"),t=el("div",{class:"t"});
    t.append(el("strong",{},p.title||"SMC page"),el("small",{},p.url));
    const ac=el("div",{class:"acts"});const o=el("a",{class:"btn sm",href:p.url,target:"_blank",rel:"noopener"},"Open page");
    ac.append(o);li.append(t,ac);pages.append(li)});

  const m=$("members");m.replaceChildren();
  if(!r.members.length){m.append(emptyMsg("No member names were published in a readable form. Check the SMC page or PDFs."))}
  else{const tb=el("table"),th=el("tr");["Name","Role"].forEach(h=>th.append(el("th",{},h)));tb.append(th);
    r.members.forEach(x=>{const tr=el("tr");tr.append(el("td",{},x.name),el("td",{},x.role||"—"));tb.append(tr)});m.append(tb)}

  const pdfs=$("pdfs");pdfs.replaceChildren();
  if(!r.pdfs.length){const li=el("li");li.append(emptyMsg("No SMC-related PDF files were found."));pdfs.append(li)}
  r.pdfs.forEach(p=>{const li=el("li"),t=el("div",{class:"t"});
    t.append(el("strong",{},p.title),el("small",{},p.url));
    const ac=el("div",{class:"acts"});
    ac.append(el("a",{class:"btn sm",href:"/download?url="+encodeURIComponent(p.url)},"Download"),
              el("a",{class:"btn alt sm",href:p.url,target:"_blank",rel:"noopener"},"View online"));
    li.append(t,ac);pdfs.append(li)});
  $("allf-card").hidden=!r.all_pdfs;$("allp-card").hidden=!r.all_pages;
  if(r.all_pdfs){
    $("allf-n").textContent=r.all_pdfs.length;$("dl-zip-all").href="/zip/"+id+"?all=1";
    const ul=$("allf");ul.replaceChildren();
    if(!r.all_pdfs.length){const li=el("li");li.append(emptyMsg("No PDF files were found on this website."));ul.append(li)}
    r.all_pdfs.forEach(p=>{const li=el("li"),t=el("div",{class:"t"}),st=el("strong",{},p.title||"PDF");
      if(p.smc)st.append(" ",el("span",{class:"badge ok"},"SMC"));
      t.append(st,el("small",{},p.url));
      const ac=el("div",{class:"acts"});
      ac.append(el("a",{class:"btn sm",href:"/download?url="+encodeURIComponent(p.url)},"Download"),el("a",{class:"btn alt sm",href:p.url,target:"_blank",rel:"noopener"},"View online"));
      li.append(t,ac);ul.append(li)});
  }
  if(r.all_pages){
    $("allp-n").textContent=r.all_pages.length;
    const ul=$("allp");ul.replaceChildren();
    r.all_pages.forEach(p=>{const li=el("li"),t=el("div",{class:"t"});
      t.append(el("strong",{},p.title||"(no title)"),el("small",{},p.url));
      const ac=el("div",{class:"acts"});ac.append(el("a",{class:"btn alt sm",href:p.url,target:"_blank",rel:"noopener"},"Open page"));
      li.append(t,ac);ul.append(li)});
  }
  $("out").scrollIntoView({behavior:"smooth"});
}
function renderTable(t,id){
  $("tbl").hidden=false;
  $("t-title").textContent=(t.kind=="list"?"School list: ":"Websites found: ")+t.total+" schools";
  $("t-note").textContent=(t.note?t.note+" ":"")+(t.kind=="list"?"Download this Excel, then use step 2 to find each school's website.":"Then use step 3 with this Excel to find the SMC pages and PDFs.");
  $("t-xlsx").href="/excel/"+id;
  const tb=$("t-table");tb.replaceChildren();
  const hr=el("tr");t.headers.forEach(h=>hr.append(el("th",{},h)));tb.append(hr);
  t.rows.forEach(r=>{const tr=el("tr");r.forEach(c=>{const td=el("td");
    if(/^https?:\/\//.test(c))td.append(el("a",{href:c,target:"_blank",rel:"noopener"},c));else td.textContent=c;tr.append(td)});tb.append(tr)});
  $("t-more").textContent=t.total>t.rows.length?("Showing the first "+t.rows.length+" of "+t.total+". The Excel has all of them."):"";
  $("tbl").scrollIntoView({behavior:"smooth"});
}
function renderBulk(list,id){
  $("bulk").hidden=false;$("b-n").textContent=list.length;
  $("b-site").textContent=list.filter(x=>x.website).length;
  $("b-smc").textContent=list.filter(x=>x.smc_pages.length).length;
  $("b-pdf").textContent=list.reduce((a,x)=>a+x.pdfs.length,0);
  $("b-xlsx").href="/excel/"+id;
  const tb=$("b-table");tb.replaceChildren();
  const hr=el("tr");["#","School","Website","SMC page(s)","PDFs","Status"].forEach(h=>hr.append(el("th",{},h)));tb.append(hr);
  const rows=[];
  list.forEach((x,i)=>{
    const tr=el("tr");tr.dataset.t=(x.name+" "+x.website).toLowerCase();tr.dataset.s=x.smc_pages.length?1:0;tr.dataset.p=x.pdfs.length?1:0;
    const w=el("td");if(x.website){w.append(el("a",{href:x.website,target:"_blank",rel:"noopener"},x.website))}
    const pg=el("td"),pu=el("ul",{class:"plain"});
    x.smc_pages.forEach((p,k)=>{const li=el("li");li.append(el("a",{href:p.url,target:"_blank",rel:"noopener"},"Open page "+(k+1)));pu.append(li)});
    if(!x.smc_pages.length)pu.append(el("li",{},"—"));pg.append(pu);
    const pf=el("td"),fu=el("ul",{class:"plain"});
    x.pdfs.forEach(p=>{const li=el("li");li.append(el("a",{class:"btn sm",href:"/download?url="+encodeURIComponent(p.url)},"Download")," ",el("small",{},(p.title||"PDF").slice(0,50)));fu.append(li)});
    if(!x.pdfs.length)fu.append(el("li",{},"—"));pf.append(fu);
    const st=el("td");st.append(el("span",{class:"badge"+(x.smc_pages.length||x.pdfs.length?" ok":"")},x.status));
    tr.append(el("td",{},i+1),el("td",{},x.name),w,pg,pf,st);tb.append(tr);rows.push(tr)});
  const run=()=>{const t=$("b-q").value.toLowerCase(),m=$("b-f").value;
    rows.forEach(r=>{r.hidden=!(r.dataset.t.includes(t)&&(m=="all"||(m=="smc"&&r.dataset.s=="1")||(m=="pdf"&&r.dataset.p=="1")||(m=="none"&&r.dataset.s=="0"&&r.dataset.p=="0")))})};
  $("b-q").oninput=run;$("b-f").onchange=run;
  $("bulk").scrollIntoView({behavior:"smooth"});
}
["url","uname","name","loc"].forEach(i=>$(i).addEventListener("keydown",e=>{if(e.key=="Enter")$("go").click()}));

const ping=()=>fetch("/api/ping").catch(()=>{});ping();setInterval(ping,5000);
$("quit").onclick=async()=>{
  if(!confirm("Stop the app? Any scan in progress will be lost."))return;
  try{await fetch("/api/quit",{method:"POST"})}catch(e){}
  document.body.replaceChildren(el("main",{},""));
  const m=document.body.firstChild;m.style.padding="40px 20px";
  m.append(el("h1",{},"App stopped"),el("p",{class:"lead"},"You can close this tab. Start the app again whenever you need it."));
};
</script></main></body></html>
"""

def free_port(start=5000):
    for port in range(start, start + 20):
        with socket.socket() as t:
            if t.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


def already_running():
    """If the app is already open, just bring up the page instead of starting a second copy."""
    for port in range(5000, 5020):
        try:
            r = requests.get(f"http://127.0.0.1:{port}/api/ping", timeout=0.7)
            if r.ok and r.json().get("app") == "smc-finder":
                return port
        except Exception:
            pass
    return None


def register_protocol():
    """Windows: lets the 'Start' button in start.html launch this app (done once, no admin needed).
    Works for the packaged .exe and also when running the .py file with Python."""
    if not sys.platform.startswith("win"):
        return
    try:
        import winreg
        if getattr(sys, "frozen", False):
            cmd = f'"{sys.executable}" "%1"'
        else:
            pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            exe = pyw if os.path.exists(pyw) else sys.executable
            cmd = f'"{exe}" "{os.path.abspath(__file__)}" "%1"'
        base = r"Software\Classes\smcfinder"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base) as k:
            winreg.SetValue(k, "", winreg.REG_SZ, "URL:SMC Finder")
            winreg.SetValueEx(k, "URL Protocol", 0, winreg.REG_SZ, "")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base + r"\shell\open\command") as k:
            winreg.SetValue(k, "", winreg.REG_SZ, cmd)
    except Exception:
        pass


def main():
    register_protocol()
    port = already_running()
    if port:
        webbrowser.open(f"http://127.0.0.1:{port}")
        return
    port = free_port()
    url = f"http://127.0.0.1:{port}"
    print(f"School SMC Finder is running at {url}")
    threading.Thread(target=watchdog, daemon=True).start()
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        log = os.path.join(os.path.expanduser("~"), "SMC_Finder_error.log")
        with open(log, "w", encoding="utf-8") as fh:
            fh.write(traceback.format_exc())
        raise
