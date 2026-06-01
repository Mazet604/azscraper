"""
Scrapes the AgencyZoom Monthly Sales Summary report into a CSV.

URL: https://app.agencyzoom.com/report/index?report=monthly

Table structure (tr.parent rows = monthly totals):
  Month | PC Sales | PC Policies | PC Items | PC Premium
        | LH Appts | LH Policies | LH Premium

Brave must be running with --remote-debugging-port=9222.
No need to close Brave — the script opens a new tab and closes only that tab.
Output: monthly_report.csv  (same folder as this script)
"""

import csv
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://app.agencyzoom.com/report/index?report=monthly"
DEBUG_PORT = 9222
OUTPUT_FILE = Path(__file__).parent / "monthly_report.csv"

FIELDNAMES = [
    "Month",
    "PC_Sales", "PC_Policies", "PC_Items", "PC_Premium",
    "LH_Appts", "LH_Policies", "LH_Premium",
]


def cell_text(row, index):
    """Return stripped inner text of the nth td in a row."""
    tds = row.query_selector_all("td")
    if index >= len(tds):
        return ""
    return tds[index].inner_text().strip()


def scrape():
    with sync_playwright() as p:
        print(f"Connecting to Brave via CDP on port {DEBUG_PORT}...")
        try:
            browser = p.chromium.connect_over_cdp(
                f"http://localhost:{DEBUG_PORT}",
                timeout=15_000,
            )
        except Exception as e:
            print(f"ERROR: Could not connect to Brave — {e}")
            print(
                "\nMake sure Brave is running with remote debugging enabled.\n"
                "Right-click your Brave shortcut → Properties → add to Target:\n"
                '  --remote-debugging-port=9222\n'
                "Then restart Brave."
            )
            sys.exit(1)

        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.new_page()

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
            page.close()
            sys.exit(1)

        rows = page.query_selector_all(
            "#detail-container table tbody tr.parent"
        )

        if not rows:
            print("ERROR: No monthly rows found in the table.")
            page.close()
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

        page.close()

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
