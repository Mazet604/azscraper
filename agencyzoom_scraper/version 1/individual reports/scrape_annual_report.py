"""
Scrapes the AgencyZoom Annual Sales Summary report into a CSV.

URL: https://app.agencyzoom.com/report/index?report=annual

Table structure — one row per staff member, columns:
  Staff | Jan_Items | Jan_Premium | Feb_Items | Feb_Premium | ...
        | Dec_Items | Dec_Premium | Year_Items | Year_Premium

The last row ("Agency") is the agency-wide total — included in CSV.

Close Brave before running — the script launches it with remote debugging.
Output: annual_report.csv  (same folder as this script)
"""

import csv
import subprocess
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://app.agencyzoom.com/report/index?report=annual"
BRAVE_EXE = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
BRAVE_PROFILE = r"C:\Users\Mazet\AppData\Local\BraveSoftware\Brave-Browser\User Data"
DEBUG_PORT = 9222
OUTPUT_FILE = Path(__file__).parent / "annual_report.csv"

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

FIELDNAMES = ["Staff"] + [
    f"{m}_{t}"
    for m in MONTHS
    for t in ("Items", "Premium")
] + ["Year_Items", "Year_Premium"]


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
                "#detail-container table tbody tr",
                timeout=30_000,
            )
        except PlaywrightTimeoutError:
            print("ERROR: Timed out waiting for the report table.")
            browser.close()
            proc.terminate()
            sys.exit(1)

        rows = page.query_selector_all("#detail-container table tbody tr")

        if not rows:
            print("ERROR: No rows found in the table.")
            browser.close()
            proc.terminate()
            sys.exit(1)

        print(f"Found {len(rows)} staff rows. Extracting data...")

        records = []
        for row in rows:
            # First cell may be <th> (Agency total row) or <td> (staff rows)
            first = row.query_selector("th, td")
            if first is None:
                continue
            staff = first.inner_text().strip()

            # Remaining cells are all <td>
            tds = row.query_selector_all("td")
            # For staff rows: tds[0] is the name td, data starts at tds[1]
            # For Agency row: first cell is <th>, so all tds are data cells
            if first.evaluate("el => el.tagName") == "TH":
                data_cells = tds          # Agency row — all tds are data
            else:
                data_cells = tds[1:]      # Staff row — skip name td

            def cell(i):
                if i >= len(data_cells):
                    return ""
                return data_cells[i].inner_text().strip()

            record = {"Staff": staff}
            col = 0
            for month in MONTHS:
                record[f"{month}_Items"]   = cell(col);   col += 1
                record[f"{month}_Premium"] = cell(col);   col += 1
            record["Year_Items"]   = cell(col);   col += 1
            record["Year_Premium"] = cell(col)
            records.append(record)

        browser.close()

    proc.terminate()

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)

    print(f"\nDone! {len(records)} rows written to: {OUTPUT_FILE}")
    for rec in records:
        print(
            f"  {rec['Staff']:<20} | "
            f"Year: {rec['Year_Items']:>4} items  {rec['Year_Premium']:>14}"
        )


if __name__ == "__main__":
    scrape()
