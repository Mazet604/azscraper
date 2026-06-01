"""
Scrapes the AgencyZoom Monthly Sales Summary report into a CSV.

URL: https://app.agencyzoom.com/report/index?report=monthly

Table structure (tr.parent rows = monthly totals):
  Month | PC Sales | PC Policies | PC Items | PC Premium
        | LH Appts | LH Policies | LH Premium

Close Brave before running — the script launches it with remote debugging.
Output: monthly_report.csv  (same folder as this script)
"""

import csv
import subprocess
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://app.agencyzoom.com/report/index?report=monthly"
BRAVE_EXE = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
BRAVE_PROFILE = r"C:\Users\Mazet\AppData\Local\BraveSoftware\Brave-Browser\User Data"
DEBUG_PORT = 9222
OUTPUT_FILE = Path(__file__).parent / "monthly_report.csv"

FIELDNAMES = [
    "Month",
    "PC_Sales", "PC_Policies", "PC_Items", "PC_Premium",
    "LH_Appts", "LH_Policies", "LH_Premium",
]


def launch_brave():
    cmd = [
        BRAVE_EXE,
        f"--remote-debugging-port={DEBUG_PORT}",
        "--profile-directory=Default",
        f"--user-data-dir={BRAVE_PROFILE}",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    print("Launching Brave with remote debugging...")
    proc = subprocess.Popen(cmd)
    time.sleep(5)
    return proc


def cell_text(row, index):
    """Return stripped inner text of the nth td in a row."""
    tds = row.query_selector_all("td")
    if index >= len(tds):
        return ""
    return tds[index].inner_text().strip()


def scrape():
    proc = launch_brave()

    with sync_playwright() as p:
        print(f"Connecting to Brave via CDP on port {DEBUG_PORT}...")
        try:
            browser = p.chromium.connect_over_cdp(
                f"http://127.0.0.1:{DEBUG_PORT}",
                timeout=15_000,
            )
        except Exception as e:
            print(f"ERROR: Could not connect to Brave — {e}")
            print("Make sure Brave is fully closed before running this script.")
            proc.terminate()
            sys.exit(1)

        context = browser.contexts[0] if browser.contexts else browser.new_context()
        pages = context.pages
        page = pages[0] if pages else context.new_page()

        print(f"Navigating to {URL} ...")
        page.goto(URL, wait_until="domcontentloaded", timeout=60_000)

        print("Waiting for report table to load...")
        try:
            page.wait_for_selector(
                "#detail-container table tbody tr.parent",
                timeout=30_000,
            )
        except PlaywrightTimeoutError:
            print("ERROR: Timed out waiting for the report table.")
            browser.close()
            proc.terminate()
            sys.exit(1)

        rows = page.query_selector_all(
            "#detail-container table tbody tr.parent"
        )

        if not rows:
            print("ERROR: No monthly rows found in the table.")
            browser.close()
            proc.terminate()
            sys.exit(1)

        print(f"Found {len(rows)} month rows. Extracting data...")

        records = []
        for row in rows:
            tds = row.query_selector_all("td")
            # td[0] = "Month" — strip the expand arrow icon text (fa icon has no text)
            month = tds[0].inner_text().strip() if len(tds) > 0 else ""
            records.append({
                "Month":        month,
                "PC_Sales":     tds[1].inner_text().strip() if len(tds) > 1 else "",
                "PC_Policies":  tds[2].inner_text().strip() if len(tds) > 2 else "",
                "PC_Items":     tds[3].inner_text().strip() if len(tds) > 3 else "",
                "PC_Premium":   tds[4].inner_text().strip() if len(tds) > 4 else "",
                "LH_Appts":     tds[5].inner_text().strip() if len(tds) > 5 else "",
                "LH_Policies":  tds[6].inner_text().strip() if len(tds) > 6 else "",
                "LH_Premium":   tds[7].inner_text().strip() if len(tds) > 7 else "",
            })

        browser.close()

    proc.terminate()

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)

    print(f"\nDone! {len(records)} rows written to: {OUTPUT_FILE}")
    for rec in records:
        print(
            f"  {rec['Month']:<15} | Sales={rec['PC_Sales']:>4} "
            f"Policies={rec['PC_Policies']:>4} Items={rec['PC_Items']:>4} "
            f"Premium={rec['PC_Premium']:>14}"
        )


if __name__ == "__main__":
    scrape()
