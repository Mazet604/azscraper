"""
Scrapes the AgencyZoom P&C Sales by Producer report into a CSV.

Launches Brave with the user's existing Default profile (already logged in)
via subprocess with remote debugging enabled, connects via CDP,
navigates to the sales report, waits for the table, then exports CSV.

If Brave is already open, close it first (the script needs to launch it
with --remote-debugging-port so it can connect).
"""

import csv
import subprocess
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://app.agencyzoom.com/sales-report/index#"
BRAVE_EXE = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
BRAVE_PROFILE = r"C:\Users\Mazet\AppData\Local\BraveSoftware\Brave-Browser\User Data"
DEBUG_PORT = 9222
OUTPUT_FILE = Path(__file__).parent / "sales_report.csv"


def launch_brave():
    """Start Brave with remote debugging on DEBUG_PORT."""
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
    # Give the browser a moment to start and open the debugging socket.
    time.sleep(3)
    return proc


def scrape():
    proc = launch_brave()

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
                "Make sure Brave is not already running before executing this script."
            )
            proc.terminate()
            sys.exit(1)

        # Reuse an existing page or open a new one
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        pages = context.pages
        page = pages[0] if pages else context.new_page()

        print(f"Navigating to {URL} ...")
        page.goto(URL, wait_until="domcontentloaded", timeout=60_000)

        print("Waiting for report table to load...")
        try:
            page.wait_for_selector(
                "#detail-container table tbody tr.line",
                timeout=30_000,
            )
        except PlaywrightTimeoutError:
            print("ERROR: Timed out waiting for the report table.")
            print(
                "The page may still be loading or you may not be logged in. "
                "Check the Brave window and try again."
            )
            browser.close()
            proc.terminate()
            sys.exit(1)

        rows = page.query_selector_all(
            "#detail-container table tbody tr.line.parent"
        )

        if not rows:
            print("ERROR: No data rows found in the table.")
            browser.close()
            proc.terminate()
            sys.exit(1)

        print(f"Found {len(rows)} producer rows. Extracting data...")

        def cell_val(row, selector):
            el = row.query_selector(selector)
            if el is None:
                return ""
            return (el.get_attribute("data-val") or "").strip()

        def fmt_money(val):
            try:
                return f"${int(val):,}"
            except (ValueError, TypeError):
                return val

        records = []
        for row in rows:
            producer_td = row.query_selector("td.lsp")
            producer = (producer_td.get_attribute("data-val") or "").strip()

            records.append({
                "Producer": producer,
                "Premium": fmt_money(cell_val(row, "td.premium")),
                "Policies": cell_val(row, "td.policies"),
                "Items": cell_val(row, "td.items"),
                "Sales": cell_val(row, "td.sales"),
                "Revenue": fmt_money(cell_val(row, "td.commission")),
            })

        browser.close()

    proc.terminate()

    # Write CSV
    fieldnames = ["Producer", "Premium", "Policies", "Items", "Sales", "Revenue"]
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)

    print(f"\nDone! {len(records)} rows written to: {OUTPUT_FILE}")
    for rec in records:
        print(
            f"  {rec['Producer']:<25} | {rec['Premium']:>12} | "
            f"Policies={rec['Policies']} Items={rec['Items']} "
            f"Sales={rec['Sales']} Revenue={rec['Revenue']}"
        )


if __name__ == "__main__":
    scrape()
