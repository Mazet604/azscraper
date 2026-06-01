"""
Scrapes every AgencyZoom Service Center pipeline and writes one sheet per
pipeline, each with two columns: Stage, Ticket Count.

Pipelines are discovered dynamically from the pipeline-selector dropdown on
/pipeline/service-pipeline. All shared logic lives in scraper_core.

Output: agencyzoom_service_pipelines.xlsx
"""

from pathlib import Path

from openpyxl import Workbook

from scraper_core import (
    scraper_session, scrape_all_pipelines, SERVICE_PIPELINES,
    write_sheet, make_unique_sheet_name,
)

SCRAPER_PROFILE = Path(__file__).parent / "scraper_profile"
OUTPUT_FILE     = Path(__file__).parent / "agencyzoom_service_pipelines.xlsx"


def main():
    with scraper_session(SCRAPER_PROFILE) as (page, _):
        print("\nScraping service pipelines...")
        results = scrape_all_pipelines(page, SERVICE_PIPELINES)

    if not results:
        print("No pipelines found. Exiting.")
        return

    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for name, stages in results:
        ws = wb.create_sheet(make_unique_sheet_name(name, used))
        write_sheet(ws, ["Stage", SERVICE_PIPELINES.count_label],
                    [list(s) for s in stages])
    wb.save(OUTPUT_FILE)

    print(f"\nSaved: {OUTPUT_FILE}")
    for name, stages in results:
        total = sum(c for _, c in stages if isinstance(c, int))
        print(f"  {name}: {len(stages)} stage(s), {total} total tickets")


if __name__ == "__main__":
    main()
