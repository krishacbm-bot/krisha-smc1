"""
Scrape CBSE schools from schoolsindia.net (all pages) and, optionally, add each school's
website from the official CBSE SARAS affiliation record.

Each school card is read on its own (the biggest container holding only ONE school title),
so Address / Affiliation / School code always belong to the right school.

Setup (once):
    pip install playwright pandas openpyxl
    playwright install chromium

Run:
    python scrape_cbse_schools.py --district Mumbai --find-websites
    python scrape_cbse_schools.py --state Maharashtra --find-websites

Output: cbse_schools_<name>.xlsx / .csv   (with --find-websites it also has a Website column)
Website lookups are cached in website_cache.json, so you can stop and run again; it carries on.
Please check the Terms of the list website before scraping, and keep the pace polite.
"""
import argparse
import json
import os
import re
import sys
import time

import pandas as pd
from playwright.sync_api import sync_playwright

SARAS = "https://saras.cbse.gov.in/SARAS/AffiliatedList/AfflicationDetails/{}"
CACHE = "website_cache.json"
SAFETY_MAX_PAGES = 1000  # only a safety stop against infinite loops

# For each school title (h3): the biggest container that still holds only ONE school title.
# That container is exactly one school card.
JS_CARDS = """
els => els.map(h => {
  let c = h;
  while (c.parentElement && c.parentElement.querySelectorAll('h3').length === 1) {
    c = c.parentElement;
  }
  const a = h.querySelector('a') || c.querySelector('a[href*="/school/"]');
  return { name: h.innerText.trim(), text: c.innerText, href: a ? a.href : '' };
})
"""
LABEL_START = re.compile(r"^(Classes|Board|Affiliation|School code|District|State)\b", re.I)


def slugify(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def field(label, text):
    m = re.search(label + r"\s*:?\s*\n*\s*([^\n]+)", text, re.I)
    return m.group(1).strip() if m else ""


def parse(card):
    t, name = card["text"], card["name"]
    address = field(r"Address", t)
    if not address:
        # Featured schools show the address as the first line under the name, with no "Address" label
        lines = [ln.strip() for ln in t.split("\n") if ln.strip()]
        if name in lines:
            for ln in lines[lines.index(name) + 1:]:
                if not LABEL_START.match(ln):
                    address = ln
                break
    return {
        "School Name": name,
        "Address": address,
        "Board": field(r"Board", t),
        "Affiliation No.": field(r"Affiliation(?: Number)?", t),
        "School Code": field(r"School code", t),
        "District": field(r"District", t),
        "State": field(r"State", t),
        "SchoolsIndia Page": card.get("href", ""),
    }


def scrape_current(page):
    cards = page.eval_on_selector_all("h3", JS_CARDS)
    return [parse(c) for c in cards if "Address" in c["text"] or "District" in c["text"]]


def signature(rows):
    return tuple((r["School Name"], r["Address"], r["Affiliation No."]) for r in rows)


def go_next(page, current):
    """Click the next page. Returns True if a click was made."""
    arrow = page.locator("button.page-arrow[aria-label='Next page']")
    if arrow.count() > 0:
        btn = arrow.last
        if btn.is_disabled():
            return False  # last page reached
        btn.click()
        return True
    nxt = page.get_by_text(str(current + 1), exact=True)
    if nxt.count() > 0:
        nxt.last.click()
        return True
    return False


def scrape_listing(page, url, progress=None):
    rows, seen, n = [], set(), 1
    page.goto(url, wait_until="networkidle")
    while n <= SAFETY_MAX_PAGES:
        data = scrape_current(page)
        sig = signature(data)
        if not data or sig in seen:
            print("No new schools found, stopping.")
            break
        seen.add(sig)
        rows.extend(data)
        print(f"Page {n}: {len(data)} schools (total {len(rows)})")
        if progress:
            progress(n, len(rows), "list")
        if not go_next(page, n):
            print("No next page button, finished.")
            break
        page.wait_for_timeout(1500)
        page.wait_for_load_state("networkidle")
        n += 1
    return rows


# ---- websites from the official CBSE record ---------------------------------
def clean_site(s):
    """Turn what a school typed in CBSE's Website field into one usable URL (or '')."""
    s = re.sub(r"<[^>]+>", " ", s or "").replace("&amp;", "&")
    for tok in re.split(r"[\s,;|]+", s.strip()):
        t = tok.strip("()[]<>.:;,'\"")
        if "." in t and "@" not in t and len(t) > 4 and t.lower() not in ("n.a.", "n/a", "n.a"):
            return t if re.match(r"https?://", t, re.I) else "http://" + t
    return ""


def extract_website(html):
    m = re.search(r"Website\s*</(?:td|th)>\s*<(?:td|th)[^>]*>(.*?)</(?:td|th)>", html, re.S | re.I)
    if m:
        return clean_site(m.group(1))
    text = re.sub(r"<[^>]+>", "\n", html)
    m = re.search(r"Website\s*\|?\s*\n*\s*([^\n]+)", text, re.I)
    if m and not m.group(1).strip().lower().startswith("year of foundation"):
        return clean_site(m.group(1))
    return ""


def lookup_website(ctx, page, affil):
    url = SARAS.format(affil)
    try:
        r = ctx.request.get(url, timeout=30000)
        html = r.text()
        if re.search(r"Website", html, re.I):
            return extract_website(html)
    except Exception:
        pass
    try:  # fallback: let the browser render the page
        page.goto(url, wait_until="networkidle", timeout=45000)
        return extract_website(page.content())
    except Exception:
        return ""


def add_websites(ctx, page, df, progress=None):
    cache = json.load(open(CACHE, encoding="utf-8")) if os.path.exists(CACHE) else {}
    affs = df["Affiliation No."].astype(str).str.replace(r"\.0$", "", regex=True)
    todo = [a for a in affs.unique() if a.isdigit() and a not in cache]
    print(f"Looking up websites for {len(todo)} schools ({len(cache)} already cached)...")
    for i, a in enumerate(todo, 1):
        cache[a] = lookup_website(ctx, page, a)
        if i % 25 == 0:
            json.dump(cache, open(CACHE, "w", encoding="utf-8"))
            print(f"  {i}/{len(todo)} done")
        if progress:
            progress(i, len(todo), "sites")
        time.sleep(0.4)
    json.dump(cache, open(CACHE, "w", encoding="utf-8"))
    df["Website"] = affs.map(cache).fillna("")
    df["Website Check"] = df["Website"].map(
        lambda w: "From CBSE official record" if w else "CBSE record lists no website")
    return df


def save(df, base):
    csv_name = f"{base}.csv"
    df.to_csv(csv_name, index=False, encoding="utf-8-sig")
    print(f"Saved CSV: {csv_name}")
    for name in (f"{base}.xlsx", f"{base}_{int(time.time())}.xlsx"):
        try:
            df.to_excel(name, index=False)
            print(f"Saved Excel: {name}")
            return
        except PermissionError:
            print(f"Could not write {name} (file open in Excel?). Trying another name...")
    print("Excel save failed, but the CSV was saved. Open the CSV in Excel instead.")


def scrape(url, progress=None, websites=False):
    """Returns a DataFrame (empty if nothing was found). progress(n, total, phase) is optional."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(ignore_https_errors=True)
        page = ctx.new_page()
        rows = scrape_listing(page, url, progress)
        if not rows:
            browser.close()
            return pd.DataFrame()
        df = pd.DataFrame(rows)
        if df["SchoolsIndia Page"].ne("").any():
            df = df.drop_duplicates(subset=["SchoolsIndia Page", "School Name"])
        else:
            df = df.drop_duplicates(subset=["School Name", "Affiliation No.", "School Code"])
        df.insert(0, "#", range(1, len(df) + 1))
        dup = int(df["Affiliation No."].duplicated().sum())
        print(f"{len(df)} schools collected. Unique affiliation numbers: {df['Affiliation No.'].nunique()}")
        if dup > 0.3 * len(df):
            print(f"WARNING: {dup} schools share an affiliation number - the page layout may have changed. "
                  "Check the Excel before using it.")
        if websites:
            df = add_websites(ctx, page, df, progress)
        browser.close()
    return df


def main():
    ap = argparse.ArgumentParser(description="Scrape CBSE schools from schoolsindia.net")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--district", help='e.g. "Mumbai", "Pune", "Thane"')
    g.add_argument("--state", help='e.g. "Maharashtra" (all districts)')
    ap.add_argument("--find-websites", action="store_true",
                    help="also add each school's website from the official CBSE record")
    args = ap.parse_args()
    if not (args.district or args.state):
        args.district = input("District / city (for example Mumbai): ").strip()
    if args.district:
        url = f"https://www.schoolsindia.net/cbse-schools-in-{slugify(args.district)}/district"
        base = f"cbse_schools_{slugify(args.district)}"
    else:
        url = f"https://www.schoolsindia.net/cbse-schools-in-{slugify(args.state)}/state"
        base = f"cbse_schools_{slugify(args.state)}"
    if args.find_websites:
        base += "_with_websites"
    print(f"Scraping {url}")
    df = scrape(url, websites=args.find_websites)
    if df.empty:
        sys.exit("No schools found on that page.")
    save(df, base)
    msg = f"Done. {len(df)} schools in total."
    if args.find_websites:
        msg += f" {df['Website'].ne('').sum()} have a website."
    print(msg)


if __name__ == "__main__":
    main()
