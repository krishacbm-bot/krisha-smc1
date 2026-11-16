SCHOOL SMC FINDER
=================
Keep all .py files in this folder together. Replace ALL old copies when you update.

ONE-TIME SETUP (open a command window in this folder: click the address bar, type cmd, Enter)
  pip install flask requests beautifulsoup4 openpyxl lxml
  pip install playwright pandas
  playwright install chromium

START
  python smc_dashboard.py              (your browser opens the page)
  Click "Quit app" (top right) when finished.

ONE-CLICK (button "1. School list from a link")
  Paste the list link, e.g. https://www.schoolsindia.net/cbse-schools-in-mumbai/district
  Keep both boxes ticked:
    - Also get each school's website from the official CBSE record
    - Also find the SMC page and PDFs of every school that has a website
  You get ONE Excel; its first sheet "Final (all data)" has, per school: name, address, affiliation no.,
  website, SMC page links, number of members found and SMC PDF links. The page also shows
  Download buttons for each PDF. Big lists take hours: leave the window open.

STEP BY STEP (if you prefer)
  1. School list (untick the SMC box)  ->  2. Find websites (optional)  ->  3. Find SMC pages & PDFs

Websites come only from CBSE's official record (what each school declared). No guessing.
Schools without a website are listed as "No website found" and skipped.

NOTE ON THE .EXE: steps 2 and 3 work in the .exe. Step 1 needs Python + Playwright.
Please check the school-list website's Terms before scraping it.
