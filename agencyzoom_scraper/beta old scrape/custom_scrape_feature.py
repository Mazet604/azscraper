"""
ARCHIVED — Custom (Filtered) Scrape feature.

This module is *not wired into the live scraper*. It preserves the
"Filtered Scrape" feature that let a user pick a From/To date range and
apply it as the Created Date filter on every pipeline before scraping.
The live tool was reverted to a single, unfiltered scrape; this file keeps
the removed code intact so the feature can be restored later.

Two pieces were removed:

  1. scraper_core.py — the date-filter helpers and the `date_filter`
     argument that threaded through the pipeline scrapers
     (`month_range_string`, `date_filter_from_months`,
     `_try_apply_date_filter`, and the `date_filter=` parameters on
     `_scrape_pipeline_stages` / `scrape_all_pipelines`).

  2. lil_scraper.py — the calendar dialog and the "Filtered Scrape"
     button / wiring (`_DateRangePickerDialog`, `_ask_date_range`,
     `_output_path_for_filter`, and the `date_filter` plumbing through
     `run_scrape` / `start_worker`).

To restore: re-add the helpers below to scraper_core.py, re-add the
`date_filter` parameter to `_scrape_pipeline_stages` /
`scrape_all_pipelines` (passing it down to `_try_apply_date_filter`), and
re-add the GUI pieces to lil_scraper.py.
"""

import calendar
import datetime
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError


# ══════════════════════════════════════════════════════════════════════════════
# Removed from scraper_core.py
# ══════════════════════════════════════════════════════════════════════════════

def month_range_string(year, month):
    """Return ('MM/01/YYYY', 'MM/LASTDAY/YYYY') for the given year/month."""
    last_day = calendar.monthrange(year, month)[1]
    return f"{month:02d}/01/{year}", f"{month:02d}/{last_day:02d}/{year}"


def date_filter_from_months(from_year, from_month, to_year, to_month):
    """Build the (startDate, endDate) tuple spanning two month boundaries.

    Both endpoints inclusive; e.g. (2025, 1, 2026, 5) returns
    ('01/01/2025', '05/31/2026').
    """
    start, _ = month_range_string(from_year, from_month)
    _, end = month_range_string(to_year, to_month)
    return start, end


def _try_apply_date_filter(page, start_date, end_date):
    """Apply the Created Date filter on the current pipeline page.

    AgencyZoom's filter form uses an ``ax.DateRangePicker`` instance whose
    in-memory state is the source of truth — writing the hidden inputs
    alone is not reliably picked up when the Filter button is clicked
    (the picker has its own `selectedRange`/start/end state, and the
    button's handler reads from it). So we try the picker's ``.val()``
    API first via every place the instance is plausibly stashed, also
    write the hidden inputs as belt-and-suspenders, then click the
    Filter button (different IDs on lead vs service).

    Returns True if the filter was applied, False if the form isn't there.
    """
    toggle = page.query_selector(
        '.dashboard-header__la-icon[data-toggle="collapse"]'
    )
    created_date_input = "#createddate-control input.timerange-start"
    if toggle and toggle.get_attribute("aria-expanded") != "true":
        toggle.click()
        try:
            # timerange inputs are type="hidden", so use state="attached"
            # — state="visible" would time out on hidden inputs.
            page.wait_for_selector(
                created_date_input, state="attached", timeout=3000)
        except PlaywrightTimeoutError:
            return False
    elif not page.query_selector(created_date_input):
        return False

    # Let the picker library attach after the panel mounts.
    page.wait_for_timeout(800)

    page.evaluate(
        """([startDate, endDate]) => {
            const $ = window.jQuery || null;

            // ─ Strategy 1: invoke the picker's .val() API directly. This
            // is what AZ's own JS does when restoring a saved query, so
            // it's the most reliable path. Hunt the instance in every
            // place it's likely to live.
            const findPicker = () => {
                const direct = [
                    window.createdDateRangePicker,
                    window.leadFilter && window.leadFilter.createdDateRangePicker,
                    window.servicePipelineFilter && window.servicePipelineFilter.createdDateRangePicker,
                    window.servicePipeline && window.servicePipeline.createdDateRangePicker,
                ];
                for (const p of direct) {
                    if (p && typeof p.val === 'function') return p;
                }
                for (const mgrName of ['leadFilter', 'servicePipelineFilter',
                                       'servicePipeline', 'ax']) {
                    const mgr = window[mgrName];
                    if (!mgr) continue;
                    for (const k in mgr) {
                        try {
                            const v = mgr[k];
                            if (v && typeof v.val === 'function'
                                  && /date/i.test(k)) return v;
                        } catch (e) {}
                    }
                }
                if ($) {
                    for (const sel of ['#createddate-control input.timerange-input',
                                       '#createddate-control']) {
                        const $el = $(sel);
                        if (!$el.length) continue;
                        const data = $el.data() || {};
                        for (const k in data) {
                            const v = data[k];
                            if (v && typeof v.val === 'function') return v;
                        }
                    }
                }
                return null;
            };

            const picker = findPicker();
            if (picker) {
                try {
                    picker.val({
                        start: startDate,
                        end: endDate,
                        // 'Customize' matches the data-range-key the picker
                        // uses internally for a custom-date range.
                        selectedRange: 'Customize',
                    });
                } catch (e) { /* fall through to inputs */ }
            }

            // ─ Strategy 2: also write the hidden inputs (and the display
            // input) directly with jQuery change events, in case Strategy
            // 1 silently no-op'd. Scope to #createddate-control so we
            // never touch the page's other .timerange-* control (Next
            // Expiration on lead, Service Due Date on service).
            const root = document.querySelector('#createddate-control');
            if (root) {
                const setField = (sel, val) => {
                    const el = root.querySelector(sel);
                    if (!el) return;
                    el.value = val;
                    if ($) $(el).val(val).trigger('change');
                    else el.dispatchEvent(new Event('change', { bubbles: true }));
                };
                setField('input.timerange-start', startDate);
                setField('input.timerange-end',   endDate);
                setField('input.timerange-range', 'Customize');
                const display = root.querySelector('input.timerange-input');
                if (display) {
                    display.value = startDate + ' - ' + endDate;
                    if ($) $(display).val(display.value).trigger('change');
                }
            }
        }""",
        [start_date, end_date],
    )

    # Filter button: lead -> #confirmFilter, service -> #doFilter.
    for btn_sel in ("#confirmFilter", "#doFilter"):
        btn = page.query_selector(btn_sel)
        if btn:
            btn.click()
            # Per-stage counts refresh via one AJAX call per stage; let the
            # network settle before reading them. Fall back to a fixed wait
            # if networkidle never fires (e.g. background polling).
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except PlaywrightTimeoutError:
                page.wait_for_timeout(2000)
            page.wait_for_timeout(800)
            return True
    return False


# NOTE: `_scrape_pipeline_stages` and `scrape_all_pipelines` in scraper_core.py
# used to accept a `date_filter=None` argument. When set to a
# ('MM/DD/YYYY', 'MM/DD/YYYY') tuple, `_scrape_pipeline_stages` called
# `_try_apply_date_filter(page, date_filter[0], date_filter[1])` after the
# columns mounted but before reading the per-stage counts. That parameter has
# been removed from the live code.


# ══════════════════════════════════════════════════════════════════════════════
# Removed from lil_scraper.py
# ══════════════════════════════════════════════════════════════════════════════

# OUTPUT_FILE = Path(__file__).parent / "agencyzoom_all_data.xlsx"  (live default)

def _output_path_for_filter(output_file, date_filter):
    """Return the workbook path, suffixed with the filter range when present.

    (Originally a closure over the module-level OUTPUT_FILE; parameterized
    here so the archived copy is self-contained.)
    """
    if not date_filter:
        return output_file
    start, end = date_filter           # 'MM/DD/YYYY' strings
    s_mm, s_dd, s_yyyy = start.split("/")
    e_mm, e_dd, e_yyyy = end.split("/")
    suffix = f"_{s_yyyy}-{s_mm}-{s_dd}_to_{e_yyyy}-{e_mm}-{e_dd}"
    return output_file.with_name(f"{output_file.stem}{suffix}{output_file.suffix}")


class _DateRangePickerDialog(tk.Toplevel):
    """Two-month calendar dialog modeled on AgencyZoom's range picker.

    Click any day to set the From date; click another to set the To date.
    A click that lands before the existing From auto-swaps them. A third
    click starts a new selection (resetting From). `result` is set to a
    (start_date, end_date) tuple of datetime.date on Apply, otherwise None.
    """

    DAY_HEADERS = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]

    SELECTED_BG = "#2774f0"
    SELECTED_FG = "#ffffff"
    INRANGE_BG  = "#dde7fa"
    INRANGE_FG  = "#2c3e57"
    NORMAL_BG   = "#ffffff"
    NORMAL_FG   = "#2c3e57"
    OUTMONTH_FG = "#b8c1cd"
    HEADER_FG   = "#7a8597"

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Date Range Filter")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=self.NORMAL_BG)

        today = datetime.date.today()
        self.left_anchor = datetime.date(today.year, today.month, 1)
        self.start_date = None
        self.end_date = None
        self._next_is_start = True
        self.result = None

        self._build_layout()
        self._render()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        dw, dh = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")

    # ── layout ──
    def _build_layout(self):
        outer = tk.Frame(self, padx=14, pady=14, bg=self.NORMAL_BG)
        outer.pack()

        tk.Label(outer,
                 text="Click a day for From, then click another for To.",
                 font=("Segoe UI", 9), fg="#4a5568",
                 bg=self.NORMAL_BG).pack(anchor="w", pady=(0, 8))

        cals = tk.Frame(outer, bg=self.NORMAL_BG)
        cals.pack()
        self.left_frame = tk.Frame(cals, bg=self.NORMAL_BG)
        self.left_frame.pack(side="left", padx=(0, 12))
        self.right_frame = tk.Frame(cals, bg=self.NORMAL_BG)
        self.right_frame.pack(side="left", padx=(12, 0))

        footer = tk.Frame(outer, pady=10, bg=self.NORMAL_BG)
        footer.pack(fill="x")
        self.range_label = tk.Label(footer, text="", font=("Segoe UI", 9),
                                    fg="#4a5568", bg=self.NORMAL_BG)
        self.range_label.pack(side="left")

        tk.Button(footer, text="Apply",  width=8,
                  command=self._on_apply).pack(side="right", padx=(6, 0))
        tk.Button(footer, text="Clear",  width=8,
                  command=self._on_clear).pack(side="right", padx=(6, 0))
        tk.Button(footer, text="Cancel", width=8,
                  command=self.destroy).pack(side="right")

    # ── rendering ──
    def _render(self):
        self._render_month(self.left_frame,  self.left_anchor,
                           show_prev=True,  show_next=False)
        right_anchor = self._add_months(self.left_anchor, 1)
        self._render_month(self.right_frame, right_anchor,
                           show_prev=False, show_next=True)

        s = self.start_date.strftime("%m/%d/%Y") if self.start_date else "—"
        e = self.end_date.strftime("%m/%d/%Y")   if self.end_date   else "—"
        self.range_label.config(text=f"From  {s}    To  {e}")

    def _render_month(self, frame, anchor, show_prev, show_next):
        for w in frame.winfo_children():
            w.destroy()

        hdr = tk.Frame(frame, bg=self.NORMAL_BG)
        hdr.pack()
        self._nav_arrow(hdr, "◀", show_prev, lambda: self._shift(-1))
        tk.Label(hdr, text=anchor.strftime("%B %Y"),
                 font=("Segoe UI", 10, "bold"), width=18, anchor="center",
                 bg=self.NORMAL_BG, fg=self.NORMAL_FG).pack(side="left")
        self._nav_arrow(hdr, "▶", show_next, lambda: self._shift(1))

        dow = tk.Frame(frame, bg=self.NORMAL_BG)
        dow.pack(pady=(2, 1))
        for d in self.DAY_HEADERS:
            tk.Label(dow, text=d, width=3, font=("Segoe UI", 8, "bold"),
                     fg=self.HEADER_FG, bg=self.NORMAL_BG
                     ).pack(side="left", padx=1)

        cal = calendar.Calendar(firstweekday=6)   # Sunday-first
        for week in cal.monthdatescalendar(anchor.year, anchor.month):
            row = tk.Frame(frame, bg=self.NORMAL_BG)
            row.pack()
            for d in week:
                self._make_day(row, d, anchor.month)

    def _nav_arrow(self, parent, text, enabled, on_click):
        lbl = tk.Label(parent, text=text if enabled else " ", width=2,
                       font=("Segoe UI", 10),
                       fg=self.NORMAL_FG if enabled else self.NORMAL_BG,
                       bg=self.NORMAL_BG,
                       cursor="hand2" if enabled else "")
        lbl.pack(side="left")
        if enabled:
            lbl.bind("<Button-1>", lambda _e: on_click())

    def _make_day(self, row, day, current_month):
        in_month = (day.month == current_month)
        bg, fg = self.NORMAL_BG, self.OUTMONTH_FG if not in_month else self.NORMAL_FG

        if in_month and self.start_date and self.end_date \
                and self.start_date <= day <= self.end_date:
            bg, fg = self.INRANGE_BG, self.INRANGE_FG
        if in_month and (day == self.start_date or day == self.end_date):
            bg, fg = self.SELECTED_BG, self.SELECTED_FG

        lbl = tk.Label(row, text=str(day.day), width=3, height=1,
                       bg=bg, fg=fg, font=("Segoe UI", 9),
                       cursor="hand2" if in_month else "")
        lbl.pack(side="left", padx=1, pady=1)
        if in_month:
            lbl.bind("<Button-1>", lambda _e, dt=day: self._on_day_click(dt))

    # ── interaction ──
    def _on_day_click(self, day):
        if self._next_is_start:
            self.start_date = day
            self.end_date = None
            self._next_is_start = False
        else:
            if day < self.start_date:
                self.end_date = self.start_date
                self.start_date = day
            else:
                self.end_date = day
            self._next_is_start = True
        self._render()

    def _shift(self, months):
        self.left_anchor = self._add_months(self.left_anchor, months)
        self._render()

    @staticmethod
    def _add_months(d, months):
        m = d.month - 1 + months
        y = d.year + m // 12
        m = m % 12 + 1
        return datetime.date(y, m, 1)

    def _on_clear(self):
        self.start_date = None
        self.end_date = None
        self._next_is_start = True
        self._render()

    def _on_apply(self):
        if not self.start_date or not self.end_date:
            messagebox.showerror(
                "Incomplete",
                "Please pick both a From and a To date "
                "(click two days on the calendar).",
                parent=self,
            )
            return
        self.result = (self.start_date, self.end_date)
        self.destroy()


def _ask_date_range(parent):
    """Open the calendar picker and return ('MM/DD/YYYY', 'MM/DD/YYYY') or None."""
    dlg = _DateRangePickerDialog(parent)
    dlg.wait_window()
    if dlg.result is None:
        return None
    start, end = dlg.result
    return start.strftime("%m/%d/%Y"), end.strftime("%m/%d/%Y")


# NOTE: the GUI's `main()` wired this in with a second "Filtered Scrape"
# button whose handler did:
#
#     def on_filtered():
#         date_filter = _ask_date_range(root)
#         if date_filter is None:
#             return
#         start_worker(date_filter=date_filter)
#
# and `run_scrape` / `start_worker` carried a `date_filter=None` argument that
# was passed to `scrape_all_pipelines(...)`. The live GUI now has a single
# unfiltered "Scrape" button.
