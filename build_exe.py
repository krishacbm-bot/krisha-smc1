"""
Run this ONCE (by you, on a Windows PC that has Python) to make the app for the non-technical person:

    python build_exe.py

It creates:  dist/SMC Finder.exe
Give that single .exe to the other person. They only double-click it:
no Python, no VS Code, no commands, and no black window. The page opens in
their browser and Python runs quietly behind it.

Needs smc_dashboard.py and smc_scraper.py in the same folder as this file.
"""
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pyinstaller",
                       "flask", "requests", "beautifulsoup4", "openpyxl", "lxml"])
subprocess.check_call([
    sys.executable, "-m", "PyInstaller", "--onefile", "--noconsole", "--clean",
    "--name", "SMC Finder",
    "--collect-submodules", "lxml",
    "--hidden-import", "bs4",
    "--hidden-import", "openpyxl",
    "smc_dashboard.py",
])
print("\nDone! Your app is here:  dist\\SMC Finder.exe")
