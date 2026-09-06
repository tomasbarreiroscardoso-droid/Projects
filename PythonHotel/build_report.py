"""
Presentation report builder.

Turns the raw workbooks written by `hotel_rates.py` into one polished Excel
file meant to be opened, read and shown to people - as opposed to
`price_report.py`, which answers single questions on the terminal.

WHAT IT PRODUCES
    Report      built as a comparison FROM Vale Palheiro, not a flat list:
                Vale Palheiro (then Amaria) always sort first and carry a
                neutral highlight in every table (see hotel_order,
                REFERENCE_HOTEL) so the eye has one fixed row to read
                everyone else against.
                  - a KPI summary per hotel, with movement vs the previous run
                  - two charts side by side: occupancy by date on the left,
                    lowest price by date on the right - Vale Palheiro drawn
                    thick, everyone else thin
                  - six tables (hotels x REPORT_DAYS dates): the five in
                    METRICS - occupancy %, availability %, lowest price,
                    highest price, rooms available - plus minimum nights,
                    which is built separately (see write_min_stay_table)
                  - a recent-trend comparison: four more hotels x dates
                    tables (availability and lowest price, each vs the run
                    closest to each TREND_ANCHOR_DAYS value back),
                    cell-for-cell change against an earlier run rather
                    than today's raw value
                  - then, below a heavier visual break, up to two OPTIONAL
                    per-room price sections, each on its own switch in
                    section 0b and each in its own colour: advertised prices
                    (what is on sale now) and estimated sold prices
                    (inferred from stock drops, not published anywhere)
    Data …      three copies of the tidy `data` rows: the latest run, the run
                before it, and the 7th most recent - for drilling into any
                number the summary raises a question about.

LAYOUT RULES
    Nothing is written at a hard-coded row. A cursor walks down the sheet and
    every block records where it landed, so adding hotels (or dates) simply
    pushes later blocks down and the charts follow. This is why the code
    passes `row` around rather than using constants.

COMPARISON BASIS
    "vs previous" (Summary block only) means the run before the latest one,
    NOT literally yesterday: if a day's run is missed, comparing against a
    day that has no data would silently produce an empty report. Both dates
    are printed at the top of the sheet so the basis is always visible.

Usage:
    ./.venv/bin/python build_report.py
    ./.venv/bin/python build_report.py --days 90 --out report.xlsx
    ./.venv/bin/python build_report.py --compare-to 2026-08-02
"""

# ==========================================================================
# QUICK REFERENCE - every command in this project
# --------------------------------------------------------------------------
# Run everything from the project folder, using the project's own
# interpreter rather than the system python. `PY` below is a shorthand for
# it, relative to the project folder, so nothing here depends on where the
# folder lives:
#     cd <the folder holding this file>
#     PY=./.venv/bin/python
# First time on a machine, or after moving the folder:  ./setup.sh
#
# 1. COLLECT   hotel_rates.py   -> output/hotel_rates_<stamp>.xlsx
#     $PY hotel_rates.py                           # uses DAYS_AHEAD, and
#                                                  # each hotel's own nights
#     $PY hotel_rates.py --days 30                 # one-off shorter window
#     $PY hotel_rates.py --days 2 --raw            # dump JSON to inspect
#     $PY hotel_rates.py --hotels other.json       # different property list
#     flags: --days --nights --out --price {nightly_price,total_price}
#            --raw --hotels
#
# 2. PRESENT   build_report.py  -> output/report_<date>.xlsx
#     $PY build_report.py                          # uses REPORT_DAYS
#     $PY build_report.py --days 90                # narrower window, also
#     $PY build_report.py --days 5                 #   resizes the price grids
#     $PY build_report.py --compare-to 2026-08-02  # pick the baseline run
#     flags: --dir --days --out --compare-to
#
# 3. ASK      price_report.py   -> terminal, or a file with --xlsx / --csv
#     $PY price_report.py                          # same as `summary`
#     $PY price_report.py summary
#     $PY price_report.py evolution --checkin 2026-08-14
#     $PY price_report.py by-checkin --room-contains suite
#     $PY price_report.py availability --by checkin
#     $PY price_report.py changes --threshold 10
#     $PY price_report.py changes --flips
#     $PY price_report.py sales                    # estimated sales ledger
#     $PY price_report.py sales --summary
#     $PY price_report.py sales --min-confidence high --hotel Craveiral
#     to a file:  $PY price_report.py --xlsx out.xlsx availability
#     global flags: --dir --price --csv --xlsx
#
# 4. SCHEDULE  run_daily.sh     scrape, then rebuild the report
#     ./run_daily.sh                               # run it by hand
#     ./install_schedule.sh                        # run it every morning
#     ./install_schedule.sh --status | --run-now | --remove
#     logs: output/logs/YYYY-MM-DD.log
# --------------------------------------------------------------------------
# EDITABLE SETTINGS IN THIS FILE
#     REPORT_DAYS               check-in date columns in the TABLES when
#                               --days is omitted (section 0). Override per
#                               run with --days, e.g. `--days 90` - this
#                               constant only changes the default.
#     CHART_DAYS                check-in date columns in the two CHARTS
#                               (section 0). Smaller than REPORT_DAYS on
#                               purpose; the smaller of the two wins.
#
# (section 0b, OPTIONAL SECTIONS)
#     INCLUDE_ADVERTISED_PRICES / INCLUDE_ESTIMATED_SALES
#                               True/False switches for the two optional
#                               per-room price blocks at the bottom
#
# (section 1, LOOK AND FEEL)
#     ACCENT / INK / MUTED      the palette
#     ACCENT_SALES              colour of the estimated-sales block
#     ACCENT_ADVERTISED         colour of the advertised-prices block
#     REFERENCE_HOTEL           which hotel sorts first and gets the neutral
#                               highlight everywhere - see hotel_order
#     REFERENCE_FILL            that highlight's colour
#     CHART_HEIGHT_CM / CHART_WIDTH_CM / CHART_ROWS_PER_CHART
#     REFERENCE_LINE_PT / OTHER_LINE_PT   chart line weights
#     BIG_MOVE 0.10             ratio move that earns a coloured fill (10%)
#     BIG_MOVE_PP 10.0          same for percentage-point tables (10 pp)
#     PP_SCALE 100              writes pp as 20.0, not 0.2
#     DATE_COL_WIDTH / LABEL_COL_WIDTH
#     METRICS                   the five hotels-x-dates tables, their formats,
#                               and whether their change is a ratio or
#                               percentage points
#     SUMMARY_COLS / SALES_SUMMARY_COLS / COMPARISON_COLS   table columns
#
# Availability weights each room type by its physical room count. The true
# capacity per hotel lives in HOTEL_CAPACITY in price_report.py.
# ==========================================================================

from __future__ import annotations

import argparse
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (CharacterProperties, Paragraph,
                                   ParagraphProperties, RichTextProperties)
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import price_report as pr
import hotel_rates as hr

# ==========================================================================
# 0. RUN SIZE
# --------------------------------------------------------------------------
# Four numbers decide how much of the data the report shows.
#
#   REPORT_DAYS   check-in dates in every TABLE - including the advertised-
#                 price and estimated-sales grids - when `--days` is not
#                 given on the command line. Override per run:
#                     build_report.py --days 90
#                     build_report.py --days 5
#
#   CHART_DAYS    check-in dates in the two CHARTS, counted from the first
#                 date. Deliberately smaller: a line plotted across a full
#                 year is an unreadable smear, so the charts show the near
#                 term while the tables still carry everything. The smaller
#                 of the two wins, so a 30-day report draws a 30-day chart
#                 rather than running off the end of the data.
#
#   SUMMARY_DAYS  check-in dates averaged into the SUMMARY block's KPIs and
#                 into the recent-trend comparison - the near window that
#                 actually drives a pricing decision.
#
#   TREND_ANCHOR_DAYS
#                 how far back the recent-trend comparison looks. Add or
#                 remove entries and the comparison grows or shrinks to
#                 match; its headings are generated from this tuple.
#
# EVERY heading and caption in the report is built from these constants, so
# changing one here changes the wording with it. Never write one of these
# numbers into prose - it goes stale the moment the constant changes.
# ==========================================================================
REPORT_DAYS = 365

CHART_DAYS = 90    # date columns drawn in the two charts (see above)

SUMMARY_DAYS = 30            # near-term window behind the Summary KPIs
TREND_ANCHOR_DAYS = (7, 30)  # recent-trend comparison looks back this far

# ==========================================================================
# 0b. OPTIONAL SECTIONS
# --------------------------------------------------------------------------
# Whole report sections, switched on/off with one flag each. Flip to False
# to drop a section from the report entirely (the code stays; it just never
# runs) - flip back to True to bring it back, no other changes needed.
# ==========================================================================
INCLUDE_ADVERTISED_PRICES = True   # per-room grid: cheapest bookable rate, latest run
INCLUDE_ESTIMATED_SALES = False    # per-room grid: inferred sold price, from stock drops

# ==========================================================================
# 1. LOOK AND FEEL
# --------------------------------------------------------------------------
# One place for every colour and font, so the report can be restyled without
# hunting through the layout code.
# ==========================================================================

INK = "1F2933"
MUTED = "7B8794"
ACCENT = "1F4E79"

TITLE_FONT = Font(size=18, bold=True, color=ACCENT)
SUB_FONT = Font(size=10, color=MUTED, italic=True)
SECTION_FONT = Font(size=12, bold=True, color="FFFFFF")
HEAD_FONT = Font(size=9, bold=True, color=INK)
BODY_FONT = Font(size=10, color=INK)
HOTEL_FONT = Font(size=10, bold=True, color=INK)

SECTION_FILL = PatternFill("solid", fgColor=ACCENT)
# The estimated-sales block is inference, not observation, so it gets its own
# colour to keep it visually separate from the measured tables above it.
ACCENT_SALES = "1F7A5A"
SECTION_SALES_FILL = PatternFill("solid", fgColor=ACCENT_SALES)
SALES_HEAD_FILL = PatternFill("solid", fgColor="E4F1EC")
# The advertised-price block is a plain snapshot (not inference like the
# sales block), so it gets its own colour too - a darker green than the
# sales block's, so the two stay distinguishable if both are switched on.
ACCENT_ADVERTISED = "1B4332"
SECTION_ADVERTISED_FILL = PatternFill("solid", fgColor=ACCENT_ADVERTISED)
HEAD_FILL = PatternFill("solid", fgColor="EDF2F7")
WEEKEND_FILL = PatternFill("solid", fgColor="DCE6F1")
NOTE_FILL = PatternFill("solid", fgColor="FFF7E0")
EMPTY_FILL = PatternFill("solid", fgColor="F5F7FA")

# The report is built around ONE reference property: every table puts it
# first (see hotel_order) and tints its row this neutral warm grey - not
# green/red/blue, since none of those "good/bad/section" meanings apply to
# "this is just the row everything else is measured against".
REFERENCE_HOTEL = "Vale Palheiro"
REFERENCE_FILL = PatternFill("solid", fgColor="E8E3D6")

# A plain rule between sections, heavier with more surrounding space where a
# section changes into something conceptually different (inferred sales,
# after the measured tables).
DIVIDER = Side(style="thick", color=MUTED)

# Change formatting: text colour for direction, fill for a move worth noticing.
UP_FONT = Font(size=10, color="1E7B34")
DOWN_FONT = Font(size=10, color="B02418")
BIG_UP_FILL = PatternFill("solid", fgColor="D6F0DC")
BIG_DOWN_FILL = PatternFill("solid", fgColor="FBD9D5")
BIG_MOVE = 0.10  # 10% - threshold at which a ratio cell also gets a fill

# Availability %/Lowest/Highest price tables: how a hotel reads against
# Vale Palheiro on that SAME day (see write_table's compare_mode). Distinct
# from UP_FONT/DOWN_FONT deliberately - those mean "vs an earlier run",
# these mean "vs Vale Palheiro today", and mixing the two shades would blur
# that difference.
BETTER_AVAIL_FONT = Font(size=10, color="1E7B34")   # less available than VP (fuller) = green
BETTER_PRICE_FONT = BETTER_AVAIL_FONT               # cheaper than VP = same green, not a darker shade
REF_BOLD_FONT = Font(size=10, bold=True, color=INK)  # VP's own best/worst day - bold, no colour change

# Availability is held as a fraction (0.30 = 30%), so a 50% -> 30% move
# differences to 0.20. Percentage points are written scaled by 100 so the cell
# reads "-20.0 pp" rather than "-0.2 pp"; the threshold scales with it.
PP_SCALE = 100
BIG_MOVE_PP = 10.0  # 10 percentage points

THIN = Side(style="thin", color="D9E2EC")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

PCT_FMT = "0%"
PRICE_FMT = "#,##0"
INT_FMT = "0"
CHANGE_FMT = "+0.0%;-0.0%;0%"
PP_FMT = "+0.0 pp;-0.0 pp;0 pp"

DATE_COL_WIDTH = 9.5
LABEL_COL_WIDTH = 26

# The two charts sit side by side, bigger than the old ones so more hotels'
# lines stay legible.
CHART_HEIGHT_CM = 13
CHART_WIDTH_CM = 32
CHART_ROWS_PER_CHART = 28  # vertical space reserved for the (single) chart row

# Line weight: the reference hotel gets a heavy, explicitly-coloured line so
# it reads as "the one to compare everyone else against" even with a dozen
# thin, auto-coloured competitor lines crossing it.
REFERENCE_LINE_PT = 3.25
OTHER_LINE_PT = 1.0


# ==========================================================================
# 2. METRICS
# --------------------------------------------------------------------------
# Everything the report shows reduces to four numbers per hotel per check-in
# date, all derived from a single run:
#
#   availability  share of that hotel's known rooms that were on sale
#   min_price     cheapest room on offer
#   max_price     dearest room on offer
#   rooms         how many distinct rooms were on offer
#
# Availability leans on price_report.room_night_grid, which reconstructs the
# full rooms x dates grid from the registry and the coverage sheet - counting
# only rows the API returned would report 100% by construction.
#
# Dates never queried stay blank rather than becoming zero: "we did not look"
# and "nothing was available" must not look alike.
# ==========================================================================

# Each entry: key, heading, number format, note, and how its change is
# expressed. Availability is already a percentage, so a ratio would be
# confusing - 15% -> 30% as "+100%" invites misreading. It is differenced in
# PERCENTAGE POINTS instead; the others divide.
METRICS = [
    ("occupancy", "Occupancy %", PCT_FMT,
     "1 - Availability % (physical rooms sold, not on sale). Green text: "
     "higher occupancy than Vale Palheiro that day (fuller, better). Vale "
     "Palheiro's own cell fills green on its fullest day (or a tie for "
     "fullest), red on its emptiest.",
     "occupancy"),
    ("availability", "Availability %", PCT_FMT,
     "Physical rooms on sale divided by the hotel's total rooms, so a "
     "12-unit category counts for more than a 1-unit one. Green text: "
     "less available than Vale Palheiro that day (fuller, better). Vale "
     "Palheiro's own cell fills green on its best day, red on its worst.",
     "availability"),
    ("min_price", "Lowest price", PRICE_FMT,
     "Cheapest room on offer, per night. Green text: cheaper than "
     "Vale Palheiro that day. Vale Palheiro's own cell is bold on its "
     "cheapest day.", "lowest_price"),
    ("max_price", "Highest price", PRICE_FMT,
     "Dearest room on offer, per night. Vale Palheiro's own cell is bold "
     "on its most expensive day.", "highest_price"),
    ("rooms", "Rooms available", INT_FMT,
     "Number of physical rooms on sale (not room types)", None),
]


def run_dates(df: pd.DataFrame) -> list[pd.Timestamp]:
    """Distinct scrape days, newest first."""
    return sorted(df["scraped_date"].unique(), reverse=True)


def closest_run(runs, target: pd.Timestamp, tolerance_days: int = 3):
    """The run closest to `target`, or None if nothing is within
    `tolerance_days` - an approximate historical anchor is still useful, but
    a run three weeks off calling itself "7 days ago" would mislead.
    """
    candidates = [r for r in runs if abs((pd.Timestamp(r) - target).days) <= tolerance_days]
    if not candidates:
        return None
    return min(candidates, key=lambda r: abs((pd.Timestamp(r) - target).days))


def hotel_order(names) -> list[str]:
    """Vale Palheiro first, Amaria second - the fixed reference and its main
    comparator - then everyone else alphabetical, so the report reads as a
    comparison FROM Vale Palheiro rather than an arbitrary list.
    """
    priority = {REFERENCE_HOTEL: 0, "Amaria": 1}
    return sorted(names, key=lambda n: (priority.get(n, 2), n.casefold()))


def metrics_for_run(df, rooms, coverage, run_date, dates) -> dict[str, pd.DataFrame]:
    """The four METRICS tables for one run, plus a fifth ("min_stay", not in
    METRICS - see write_min_stay_table) - hotels (rows) x `dates` (columns).
    """
    hotels = hotel_order(df["hotel"].unique())
    blank = pd.DataFrame(index=hotels, columns=dates, dtype="float64")

    day = df[df["scraped_date"] == run_date]
    cov_day = coverage[coverage["scraped_date"] == run_date] if not coverage.empty else coverage

    grid = pr.room_night_grid(day, rooms, cov_day) if not day.empty else pd.DataFrame()

    out = {k: blank.copy() for k, *_ in METRICS}
    out["min_stay"] = blank.copy()
    if grid.empty:
        return out

    # Weight each room type by how many physical rooms it holds. Capacity is
    # inferred from the WHOLE history (`df`), not this one run, so the
    # denominator does not shrink on a day the property is nearly full.
    grid = grid.merge(pr.capacity(df), on=["hotel", "room_type_code"], how="left")
    grid["capacity"] = grid["capacity"].fillna(1.0)

    units = grid.pivot_table(index="hotel", columns="checkin",
                             values="units", aggfunc="sum")
    total = grid.pivot_table(index="hotel", columns="checkin",
                             values="capacity", aggfunc="sum")
    avail = units / total
    occ = 1 - avail
    count = units

    # A rate plan can be returned non-bookable (e.g. the stay is shorter than
    # its own min_stay) alongside a bookable one for the same room/date; only
    # a bookable offer is a price anyone could actually pay.
    bookable_day = day[day["is_bookable"].astype(str).str.casefold() == "true"]
    lo = bookable_day.pivot_table(index="hotel", columns="checkin",
                                  values="nightly_price", aggfunc="min")
    hi = bookable_day.pivot_table(index="hotel", columns="checkin",
                                  values="nightly_price", aggfunc="max")

    for key, table in (("occupancy", occ), ("availability", avail), ("rooms", count),
                       ("min_price", lo), ("max_price", hi)):
        out[key] = table.reindex(index=hotels, columns=dates)

    # MINIMUM NIGHTS: per room, the LEAST demanding rate plan seen that day -
    # deliberately NOT bookable-only, since a non-bookable row still reports
    # its own min_stay, which is exactly the signal that room needs more
    # nights than were queried. Per hotel, the STRICTEST such requirement
    # across its rooms - the nights you would need to query to guarantee
    # every room seen today has at least one offer.
    if "min_stay" in day:
        room_ms = day.pivot_table(index=["hotel", "room_type_code"],
                                  columns="checkin", values="min_stay", aggfunc="min")
        if not room_ms.empty:
            hotel_ms = room_ms.groupby(level="hotel").max()
            out["min_stay"] = hotel_ms.reindex(index=hotels, columns=dates)
    return out


# ==========================================================================
# 3. SHEET BUILDING BLOCKS
# --------------------------------------------------------------------------
# Small writers that each return the next free row, so callers can chain them
# without ever naming an absolute position.
# ==========================================================================

def write_section(ws, row: int, title: str, note: str = "", width: int = 12,
                  fill: PatternFill | None = None) -> int:
    """A full-width heading bar. Returns the next free row."""
    ws.cell(row=row, column=1, value=title).font = SECTION_FONT
    for c in range(1, width + 1):
        ws.cell(row=row, column=c).fill = fill or SECTION_FILL
    ws.row_dimensions[row].height = 20
    row += 1
    if note:
        cell = ws.cell(row=row, column=1, value=note)
        cell.font = SUB_FONT
        row += 1
    return row


def write_date_header(ws, row: int, dates, label: str = "Hotel") -> None:
    """Date column headers, with Saturday/Sunday tinted - the calendar
    weekend, and the first thing anyone looks for. (The date in each column
    is the check-in night, so a guest staying through a Saturday/Sunday
    weekend typically checks in on the Saturday - but tinting Friday/Saturday
    to reflect that reads as the wrong weekend at a glance, so this tints
    the actual weekend days instead.)"""
    head = ws.cell(row=row, column=1, value=label)
    head.font = HEAD_FONT
    head.fill = HEAD_FILL
    head.border = BOX
    for i, d in enumerate(dates):
        cell = ws.cell(row=row, column=2 + i, value=d.strftime("%d %b"))
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center")
        cell.fill = WEEKEND_FILL if d.weekday() in (5, 6) else HEAD_FILL
        cell.border = BOX
    ws.row_dimensions[row].height = 18


def _apply_vs_reference_styling(ws, first_row: int, table: pd.DataFrame, dates, mode: str) -> None:
    """Recolour an already-written table relative to REFERENCE_HOTEL's SAME
    DAY value - a whole day's spread across every hotel must be known before
    any one cell in it can be styled, so this is a second pass over columns,
    run after write_table's row-by-row pass has populated every cell.

    mode="availability"  other hotels LOWER than VP that day (fuller, better)
                         get green text; VP's OWN cell fills green when it is
                         the day's minimum (ties count) or red when it is the
                         day's maximum.
    mode="occupancy"     the mirror image of "availability" (1 - it): other
                         hotels HIGHER than VP that day get green text; VP's
                         OWN cell fills green when it is the day's maximum
                         (ties count) or red when it is the day's minimum.
    mode="lowest_price"  other hotels LOWER than VP get green text (same
                         shade as availability's); VP goes bold (no colour
                         change) when it is the day's minimum.
    mode="highest_price" VP goes bold (no colour change) when it is the
                         day's maximum. No styling for other hotels.
    """
    hotels = list(table.index)
    if REFERENCE_HOTEL not in hotels:
        return
    ref_row = first_row + hotels.index(REFERENCE_HOTEL)

    for c, d in enumerate(dates):
        if d not in table.columns:
            continue
        vp_value = table.loc[REFERENCE_HOTEL, d]
        if pd.isna(vp_value):
            continue
        day_values = table[d].dropna()
        day_min = day_values.min() if len(day_values) else None
        day_max = day_values.max() if len(day_values) else None
        col = 2 + c
        ref_cell = ws.cell(row=ref_row, column=col)

        if mode == "availability":
            if day_min is not None and vp_value <= day_min:
                ref_cell.fill = BIG_UP_FILL
            elif day_max is not None and vp_value >= day_max:
                ref_cell.fill = BIG_DOWN_FILL
        elif mode == "occupancy":
            if day_max is not None and vp_value >= day_max:
                ref_cell.fill = BIG_UP_FILL
            elif day_min is not None and vp_value <= day_min:
                ref_cell.fill = BIG_DOWN_FILL
        elif mode == "lowest_price" and day_min is not None and vp_value <= day_min:
            ref_cell.font = REF_BOLD_FONT
        elif mode == "highest_price" and day_max is not None and vp_value >= day_max:
            ref_cell.font = REF_BOLD_FONT

        other_font = None
        if mode in ("availability", "occupancy"):
            other_font = BETTER_AVAIL_FONT
        elif mode == "lowest_price":
            other_font = BETTER_PRICE_FONT
        if other_font is not None:
            for r, hotel in enumerate(hotels):
                if hotel == REFERENCE_HOTEL:
                    continue
                value = table.loc[hotel, d]
                if pd.isna(value):
                    continue
                better = value < vp_value if mode in ("availability", "lowest_price") else value > vp_value
                if better:
                    ws.cell(row=first_row + r, column=col).font = other_font


def write_table(ws, row: int, table: pd.DataFrame, dates, number_format: str,
                compare_mode: str | None = None) -> tuple[int, int]:
    """Write one hotels x dates block.

    `compare_mode` (None, "availability", "occupancy", "lowest_price",
    "highest_price")
    additionally recolours cells relative to REFERENCE_HOTEL's same-day
    value - see _apply_vs_reference_styling. Only meaningful when the
    table's rows really are hotels (not e.g. room names).

    Returns (first_data_row, next_free_row) so later blocks can point at it
    without assuming where it ended up.
    """
    write_date_header(ws, row, dates)
    first = row + 1
    for r, hotel in enumerate(table.index):
        is_ref = hotel == REFERENCE_HOTEL
        label = ws.cell(row=first + r, column=1, value=hotel)
        label.font = HOTEL_FONT
        label.border = BOX
        if is_ref:
            label.fill = REFERENCE_FILL
        for c, d in enumerate(dates):
            value = table.loc[hotel, d] if d in table.columns else None
            cell = ws.cell(row=first + r, column=2 + c)
            if pd.notna(value):
                cell.value = float(value)
                if is_ref:
                    cell.fill = REFERENCE_FILL
            else:
                # Distinguish "nothing on sale / not queried" from a real zero.
                cell.fill = EMPTY_FILL
            cell.number_format = number_format
            cell.font = BODY_FONT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="center")
    if compare_mode:
        _apply_vs_reference_styling(ws, first, table, dates, compare_mode)
    return first, first + len(table.index) + 1


# ==========================================================================
# 3b. MINIMUM-STAY AWARENESS
# --------------------------------------------------------------------------
# hotels.json's per-hotel `nights` (see hotel_rates.py section 1) decides how
# long a stay every date is queried for. A room whose OWN minimum-stay rule
# needs more nights than that never has a bookable offer, and - unlike a
# genuinely sold-out room - looks identical to one price_report.py has never
# heard of, unless something here actually checks for it.
#
# Two different questions, two different tools:
#   MINIMUM NIGHTS table   per hotel PER DAY, from just the one run being
#                          reported on - the strictest requirement any room
#                          revealed that day (see metrics_for_run above).
#   "Minimum stay!" column per hotel, using the WHOLE collected history - a
#                          room can reveal its true requirement on one day
#                          and stay invisible on every other, so a single
#                          day's snapshot alone would miss it.
#
# Both reduce a room's several rate plans to ONE number the same way: the
# LEAST demanding one it has ever shown (min_stay), because a room only
# needs ONE satisfiable rate plan to have a bookable offer at all - a
# stricter sibling rate plan existing alongside it is normal, not a problem
# (confirmed on Craveiral room 4: rate plans requiring 1, 3, 4 and 7 nights
# all sit side by side on the same date; querying 1 night is still fine).
# ==========================================================================

def hotel_nights_config() -> dict[str, int]:
    """Each hotel's currently configured stay length, straight from
    hotels.json - the same source hotel_rates.py itself reads (including
    the null-means-default rule), so this can never drift from what a run
    actually queried for.
    """
    hotels = hr.load_hotels(None)
    return {h["name"]: (h["nights"] if h.get("nights") is not None else hr.NIGHTS)
            for h in hotels}


def rooms_excluded_by_min_stay(df, nights_config: dict) -> set:
    """Hotel names with at least one (room, check-in date) - anywhere in the
    whole collected history - whose least-demanding rate plan on THAT DATE
    still needed more nights than that hotel is currently configured to
    query for.

    Reduced per (room, date) FIRST, not straight to a room's all-time
    cheapest option: a minimum-stay rule is often date-dependent (a season,
    a weekend), so collapsing across dates before comparing would hide a
    real, currently-active exclusion on a stricter date just because the
    same room was more lenient on some unrelated one. Confirmed live on
    Craveiral: every room's all-time-cheapest min_stay is 1, so the old
    all-time-min version never flagged it - but several rooms needed 2
    nights specifically for 2026-10-05 to 10-08, which this version catches.
    """
    if "min_stay" not in df:
        return set()
    per_room_day = (df.dropna(subset=["min_stay"])
                      .groupby(["hotel", "room_type_code", "checkin"])["min_stay"].min())
    flagged = set()
    for (hotel, _code, _checkin), least_demanding in per_room_day.items():
        configured = nights_config.get(hotel)
        if configured is not None and configured < least_demanding:
            flagged.add(hotel)
    return flagged


def write_min_stay_table(ws, row, table, dates, hotels, nights_config: dict) -> tuple[int, int]:
    """MINIMUM NIGHTS: the strictest per-room requirement seen that day, red
    when a hotel's configured nights is below it - at least one room needed
    more nights than this run queried for, so its true rate never appeared.
    """
    write_date_header(ws, row, dates)
    first = row + 1
    for r, hotel in enumerate(hotels):
        is_ref = hotel == REFERENCE_HOTEL
        label = ws.cell(row=first + r, column=1, value=hotel)
        label.font = HOTEL_FONT
        label.border = BOX
        if is_ref:
            label.fill = REFERENCE_FILL
        configured = nights_config.get(hotel)
        for c, d in enumerate(dates):
            value = table.loc[hotel, d] if (hotel in table.index and d in table.columns) else None
            cell = ws.cell(row=first + r, column=2 + c)
            if pd.notna(value):
                cell.value = int(value)
                if configured is not None and configured < value:
                    cell.font = Font(size=10, color="7A150C", bold=True)
                    cell.fill = BIG_DOWN_FILL
                else:
                    cell.font = BODY_FONT
                    if is_ref:
                        cell.fill = REFERENCE_FILL
            else:
                cell.fill = EMPTY_FILL
            cell.number_format = INT_FMT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="center")
    return first, first + len(hotels) + 1


# ==========================================================================
# 3c. RECENT-TREND COMPARISON
# --------------------------------------------------------------------------
# Same near-window (SUMMARY_DAYS) average availability / lowest price the
# Summary block already computes, just measured at a SECOND point in time -
# the run closest to each TREND_ANCHOR_DAYS value before the latest one - so
# a hotel's trajectory
# is visible next to Vale Palheiro's rather than only its own single latest
# number. Blank "then" columns simply mean no run near enough that anchor
# exists yet in the collected history.
# ==========================================================================

def write_divider(ws, row: int, width: int, extra_gap: bool = False) -> int:
    """A plain horizontal rule between report sections: a blank row with a
    heavy bottom border. `extra_gap` adds a blank row either side, for a
    bigger break where the report moves into something conceptually
    different (inferred sales, after the measured tables above).
    """
    if extra_gap:
        row += 1
    for c in range(1, width + 1):
        ws.cell(row=row, column=c).border = Border(bottom=DIVIDER)
    row += 1
    if extra_gap:
        row += 1
    return row


def write_delta_table(ws, row: int, cur_table, anchor_table, dates, hotels,
                      mode: str = "ratio") -> tuple[int, int]:
    """Hotels x dates, same shape as the plain value tables above - but each
    cell is the CHANGE against an earlier run instead of today's raw value.

    `anchor_table` is the matching table (the anchor run's own
    ["availability"]) from a
    metrics_for_run() result for that earlier run, or None if no run close
    enough to it exists yet - every cell is then simply left blank rather
    than guessed. `mode="pp"` differences and scales to percentage points
    (availability, already a fraction); `mode="ratio"` divides (prices).
    """
    big = BIG_MOVE_PP if mode == "pp" else BIG_MOVE
    write_date_header(ws, row, dates)
    first = row + 1
    for r, hotel in enumerate(hotels):
        is_ref = hotel == REFERENCE_HOTEL
        label = ws.cell(row=first + r, column=1, value=hotel)
        label.font = HOTEL_FONT
        label.border = BOX
        if is_ref:
            label.fill = REFERENCE_FILL
        for c, d in enumerate(dates):
            cell = ws.cell(row=first + r, column=2 + c)
            now = cur_table.loc[hotel, d] if (hotel in cur_table.index and d in cur_table.columns) else None
            was = (anchor_table.loc[hotel, d]
                   if (anchor_table is not None and hotel in anchor_table.index
                       and d in anchor_table.columns) else None)
            empty = True
            if pd.notna(now) and pd.notna(was):
                if mode == "pp":
                    cell.value = (float(now) - float(was)) * PP_SCALE
                    empty = False
                elif was != 0:
                    cell.value = float(now) / float(was) - 1
                    empty = False
            if empty:
                cell.fill = EMPTY_FILL
            cell.number_format = PP_FMT if mode == "pp" else CHANGE_FMT
            cell.font = BODY_FONT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="center")

            big_move = False
            if isinstance(cell.value, float):
                v = cell.value
                if v <= -big:
                    cell.font = Font(size=10, color="7A150C", bold=True)
                    cell.fill = BIG_DOWN_FILL
                    big_move = True
                elif v >= big:
                    cell.font = Font(size=10, color="14561F", bold=True)
                    cell.fill = BIG_UP_FILL
                    big_move = True
                elif v < 0:
                    cell.font = DOWN_FONT
                elif v > 0:
                    cell.font = UP_FONT

            if is_ref and not empty and not big_move:
                cell.fill = REFERENCE_FILL
    return first, first + len(hotels) + 1


# ==========================================================================
# 4. SUMMARY BLOCK
# --------------------------------------------------------------------------
# The part most people will actually read. Per hotel, over the next
# SUMMARY_DAYS days:
# how full, how cheap, how dear - each with its movement against the previous
# run so the number has a direction attached rather than sitting alone.
# ==========================================================================

SUMMARY_COLS = [
    ("Hotel", 26), ("Availability %", 15), ("vs prev", 10),
    ("Lowest price", 14), ("vs prev", 10), ("Median lowest", 14),
    ("Highest price", 14), ("Rooms on sale", 14), ("Rooms known", 13),
    ("Dates on sale", 14), ("Nights", 9), ("Min-stay check", 16),
]


def _horizon(table, dates, days=SUMMARY_DAYS):
    """Restrict a metric table to the first `days` columns - the near window
    that actually drives decisions, rather than all REPORT_DAYS of them."""
    if table is None or table.empty:
        return pd.DataFrame()
    return table[[d for d in dates[:days] if d in table.columns]]


def rooms_known(rooms: pd.DataFrame | None, df: pd.DataFrame, hotel: str) -> int | None:
    """How many physical rooms the scraper believes this hotel has.

    This is the denominator behind Availability %, so showing it next to
    "Rooms on sale" makes the percentage self-explanatory: 1 of 51 reads very
    differently from 1 of 7. Counting room *types* here would understate it
    badly - Praia do Canal has 7 types but ~51 rooms - so it uses the
    inferred capacity (price_report section 4b) instead.
    """
    if df is None or df.empty:
        return None
    caps = pr.capacity_by_hotel(df)
    if hotel in caps.index:
        return int(round(float(caps.loc[hotel])))
    return None


def write_summary(ws, row, cur, prev, dates, hotels, rooms=None, df=None,
                  nights_config=None, excluded=None) -> int:
    for i, (name, width) in enumerate(SUMMARY_COLS, start=1):
        cell = ws.cell(row=row, column=i, value=name)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    row += 1

    for hotel in hotels:
        av = _horizon(cur["availability"], dates).loc[hotel] if hotel in cur["availability"].index else pd.Series(dtype=float)
        lo = _horizon(cur["min_price"], dates).loc[hotel] if hotel in cur["min_price"].index else pd.Series(dtype=float)
        hi = _horizon(cur["max_price"], dates).loc[hotel] if hotel in cur["max_price"].index else pd.Series(dtype=float)
        rm = _horizon(cur["rooms"], dates).loc[hotel] if hotel in cur["rooms"].index else pd.Series(dtype=float)

        p_av = p_lo = None
        if prev is not None:
            pa = _horizon(prev["availability"], dates)
            pl = _horizon(prev["min_price"], dates)
            if hotel in pa.index:
                p_av = pa.loc[hotel].mean(skipna=True)
            if hotel in pl.index:
                p_lo = pl.loc[hotel].min(skipna=True)

        av_now = av.mean(skipna=True) if len(av) else None
        lo_now = lo.min(skipna=True) if len(lo) else None
        nights_value = (nights_config or {}).get(hotel)
        warning = "Minimum stay!" if excluded and hotel in excluded else None

        values = [
            hotel,
            av_now,
            # Scaled to real percentage points - see PP_SCALE.
            ((av_now - p_av) * PP_SCALE)
            if (av_now is not None and p_av is not None and pd.notna(p_av)) else None,
            lo_now,
            (lo_now / p_lo - 1) if (lo_now is not None and p_lo not in (None, 0) and pd.notna(p_lo)) else None,
            lo.median(skipna=True) if len(lo) else None,
            hi.max(skipna=True) if len(hi) else None,
            rm.max(skipna=True) if len(rm) else None,
            rooms_known(rooms, df, hotel) if df is not None else None,
            int(av.notna().sum()) if len(av) else 0,
            nights_value,
            warning,
        ]
        formats = [None, PCT_FMT, PP_FMT, PRICE_FMT, CHANGE_FMT,
                   PRICE_FMT, PRICE_FMT, INT_FMT, INT_FMT, INT_FMT,
                   INT_FMT, None]

        for i, (v, fmt) in enumerate(zip(values, formats), start=1):
            cell = ws.cell(row=row, column=i)
            if v is not None and (not isinstance(v, float) or pd.notna(v)):
                cell.value = v
            if fmt:
                cell.number_format = fmt
            cell.font = HOTEL_FONT if i == 1 else BODY_FONT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="left" if i == 1 else "center")
            # Colour the two movement columns by direction.
            if i in (3, 5) and isinstance(v, float) and pd.notna(v):
                cell.font = UP_FONT if v > 0 else DOWN_FONT if v < 0 else BODY_FONT
            # Nights is a configured fact, not a measured value with a
            # direction - bold, but plain black.
            if i == 11:
                cell.font = Font(size=10, bold=True, color=INK)
            if i == 12 and v:
                cell.font = Font(size=10, bold=True, color="B02418")
            if hotel == REFERENCE_HOTEL:
                cell.fill = REFERENCE_FILL
        row += 1
    return row + 1


# ==========================================================================
# 4b. ESTIMATED SALES BLOCK
# --------------------------------------------------------------------------
# Everything above this point is measured: it is what the booking engines
# actually published. This block is inferred - stock that fell between two
# runs, priced at what it was going for beforehand - so it is deliberately a
# different colour, and it carries its caveats on the sheet rather than only
# in the documentation.
#
# One table per hotel, its known rooms down the side and check-in dates across
# the top, showing the unit-weighted average price achieved. The full event
# list, with confidence and stock before/after, goes on its own sheet.
# ==========================================================================

SALES_SUMMARY_COLS = [
    ("Hotel", 26), ("Units sold", 12), ("Rooms", 9), ("Dates", 9),
    ("Avg price", 12), ("Cheapest", 11), ("Dearest", 11),
    ("Est. revenue", 14), ("High conf.", 11),
]

# The "how this works" note sits to the right of the summary table, in its
# own merged block, so it never has to compete for the rows the table needs
# (which vary with the number of hotels).
NOTE_GAP_COLS = 1     # blank columns between the table and the note
NOTE_SPAN_COLS = 9    # width of the merged note block
NOTE_SPAN_ROWS = 20   # height of the merged note block


def _sales_methodology_note(dates, df) -> str:
    """Explain the estimate in terms of the actual run that produced it, so
    the text is always true for whatever --days / history is on hand."""
    n_runs = df["scraped_date"].nunique()
    first_run = pd.Timestamp(df["scraped_date"].min())
    last_run = pd.Timestamp(df["scraped_date"].max())
    return "\n\n".join([
        "HOW THIS IS CALCULATED",
        "Each grid cell is the unit-weighted average of the cheapest "
        "rate-plan price recorded just before a room's stock dropped, "
        "averaged across every drop found for that room and check-in date.",
        "A ‘sale’ is inferred whenever available stock falls between "
        "two consecutive scrapes (stock before minus stock after) - the "
        "room does not need to disappear completely, so multi-unit room "
        "categories are counted correctly rather than only all-or-nothing.",
        "Confidence  HIGH: stock fell but the room stayed on sale (a clean "
        "decrement).  MEDIUM: the room disappeared entirely (sold out or "
        f"withdrawn - can't tell which).  LOW: the two runs were more than "
        f"{pr.MAX_GAP_DAYS} day(s) apart, so more could have happened in "
        "between.",
        "In the summary table above, Cheapest/Dearest are the min/max price "
        "across every detected sale for that hotel; Avg price is estimated "
        "revenue divided by units sold.",
        f"Dataset behind this run: {n_runs} scrape run(s) between "
        f"{first_run:%d %b %Y} and {last_run:%d %b %Y}, covering the "
        f"{len(dates)} check-in date(s) shown in the tables "
        f"({dates[0]:%d %b %Y} – {dates[-1]:%d %b %Y}). This is "
        "recomputed from scratch on every run, never chained onto a "
        "previous report, so widening or narrowing the window (--days, or "
        "just scraping more history) recomputes every figure above.",
        "Known limits: nothing before the first scrape is visible; which "
        "rate plan actually sold isn't known, only the cheapest one on "
        "offer; and a cancellation that raises stock again can look "
        "identical to ‘never sold’, silently masking a real sale. "
        "See the 'Estimated sales' sheet for every individual event and "
        "its confidence.",
    ])


def write_advertised_price_section(ws, row, df, dates, hotels, rooms, latest, width):
    """The advertised-price block: one room x check-in grid per hotel, the
    cheapest BOOKABLE rate on offer as of the latest run. A plain snapshot
    of what a guest could book right now - not an estimate, unlike the
    sales block below it. Returns the next free row."""
    row = write_section(
        ws, row, "  ADVERTISED PRICES — latest run, per room",
        "Cheapest bookable rate on offer for each room and check-in date, "
        "read straight off the latest run only.",
        width=width, fill=SECTION_ADVERTISED_FILL)

    for hotel in hotels:
        grid = pr.advertised_price_grid(df, hotel, dates, as_of=latest)
        row = write_section(ws, row, f"  {hotel.upper()} — advertised price",
                            width=width, fill=SECTION_ADVERTISED_FILL)
        if grid.empty:
            cell = ws.cell(row=row, column=1,
                           value="Nothing bookable for this hotel on the latest run.")
            cell.font = Font(size=10, italic=True, color=MUTED)
            row += 2
            continue
        known = rooms_known_names(rooms, df, hotel)
        grid = grid.reindex(index=[n for n in known if n] or list(grid.index))
        _, row = write_table(ws, row, grid, dates, PRICE_FMT)
    return row


def write_sales_section(ws, row, sales, summary, dates, hotels, rooms, df, width):
    """The estimated-sales block. Returns the next free row."""
    row = write_section(
        ws, row, "  ESTIMATED SOLD PRICES — inferred, not published",
        "Stock that fell between two runs, priced at what it was selling for "
        "just before. An estimate: a room can also leave sale through a "
        "stop-sell or a minimum-stay change. See the 'Estimated sales' sheet "
        "for confidence per event.",
        width=width, fill=SECTION_SALES_FILL)

    if sales is None or sales.empty:
        cell = ws.cell(row=row, column=1,
                       value="No estimated sales yet — this needs at least two "
                             "runs on different days.")
        cell.font = Font(size=10, italic=True, color=INK)
        cell.fill = NOTE_FILL
        return row + 2

    # --- per-hotel summary -------------------------------------------------
    table_header_row = row
    for i, (name, w) in enumerate(SALES_SUMMARY_COLS, start=1):
        cell = ws.cell(row=row, column=i, value=name)
        cell.font = HEAD_FONT
        cell.fill = SALES_HEAD_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    row += 1
    by_hotel = {r["hotel"]: r for _, r in summary.iterrows()} if not summary.empty else {}
    for hotel in hotels:
        s = by_hotel.get(hotel)
        values = [hotel,
                  s["units"] if s is not None else 0,
                  s["rooms"] if s is not None else 0,
                  s["dates"] if s is not None else 0,
                  s["avg_price"] if s is not None else None,
                  s["cheapest"] if s is not None else None,
                  s["dearest"] if s is not None else None,
                  s["revenue"] if s is not None else None,
                  s["high_conf"] if s is not None else 0]
        formats = [None, INT_FMT, INT_FMT, INT_FMT, PRICE_FMT, PRICE_FMT,
                   PRICE_FMT, PRICE_FMT, INT_FMT]
        for i, (v, fmt) in enumerate(zip(values, formats), start=1):
            cell = ws.cell(row=row, column=i)
            if v is not None and (not isinstance(v, float) or pd.notna(v)):
                cell.value = v
            if fmt:
                cell.number_format = fmt
            cell.font = HOTEL_FONT if i == 1 else BODY_FONT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="left" if i == 1 else "center")
            if hotel == REFERENCE_HOTEL:
                cell.fill = REFERENCE_FILL
        row += 1
    row += 1

    # --- methodology note, to the right of the table above -----------------
    note_col = len(SALES_SUMMARY_COLS) + 1 + NOTE_GAP_COLS
    note_end_col = note_col + NOTE_SPAN_COLS - 1
    note_end_row = table_header_row + NOTE_SPAN_ROWS - 1
    ws.merge_cells(start_row=table_header_row, start_column=note_col,
                   end_row=note_end_row, end_column=note_end_col)
    note_cell = ws.cell(row=table_header_row, column=note_col,
                        value=_sales_methodology_note(dates, df))
    note_cell.font = Font(size=9, color=INK)
    note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    note_cell.fill = SALES_HEAD_FILL
    row = max(row, note_end_row + 2)

    # --- one price grid per hotel ------------------------------------------
    for hotel in hotels:
        grid = pr.sales_price_grid(sales, hotel, dates)
        row = write_section(ws, row, f"  {hotel.upper()} — estimated price achieved",
                            width=width, fill=SECTION_SALES_FILL)
        if grid.empty:
            cell = ws.cell(row=row, column=1, value="Nothing estimated sold yet.")
            cell.font = Font(size=10, italic=True, color=MUTED)
            row += 2
            continue
        # Show every room the hotel has, not only those with a sale, so the
        # blanks read as "nothing sold" rather than "room missing".
        known = rooms_known_names(rooms, df, hotel)
        grid = grid.reindex(index=[n for n in known if n] or list(grid.index))
        _, row = write_table(ws, row, grid, dates, PRICE_FMT)
    return row


def write_estimated_sales_block(wb, ws, row, df, rooms, coverage, dates, hotels, width):
    """Everything the estimated-sales section needs: the inference itself,
    the in-sheet block, and the 'Estimated sales' detail sheet. Kept as one
    function so INCLUDE_ESTIMATED_SALES turns the whole thing on or off in
    one place. Returns the next free row."""
    sales = pr.estimated_sales(df, rooms, coverage)
    summary_tbl = pr.sales_summary(sales) if not sales.empty else pd.DataFrame()
    row = write_sales_section(ws, row, sales, summary_tbl, dates, hotels,
                              rooms, df, width=width)

    if not sales.empty:
        detail = wb.create_sheet("Estimated sales")
        detail.sheet_view.showGridLines = False
        detail.append(list(sales.columns))
        for cell in detail[1]:
            cell.font = HEAD_FONT
            cell.fill = SALES_HEAD_FILL
        for rec in sales.itertuples(index=False):
            detail.append([v.to_pydatetime() if isinstance(v, pd.Timestamp) else v
                           for v in rec])
        detail.freeze_panes = "A2"
        detail.auto_filter.ref = detail.dimensions
        for i, name in enumerate(sales.columns, start=1):
            detail.column_dimensions[get_column_letter(i)].width = max(12, len(name) + 3)
    return row


def rooms_known_names(rooms, df, hotel) -> list:
    """Room names for a hotel, registry first so sold-out rooms still appear."""
    if rooms is not None and not rooms.empty and "hotel" in rooms:
        names = rooms.loc[rooms["hotel"] == hotel, "name"].dropna().unique().tolist()
        if names:
            return sorted(names, key=str.casefold)
    return sorted(df.loc[df["hotel"] == hotel, "room_name"].dropna().unique().tolist(),
                  key=str.casefold)


# ==========================================================================
# 5. CHARTS
# --------------------------------------------------------------------------
# Two line charts, one series per hotel, reading straight from the tables
# written below them - so they extend themselves when a hotel is added.
# ==========================================================================

def _label_last_point(series, idx: int) -> None:
    """Print the series name next to one point, so a line is identified where
    it actually runs rather than only in the legend.

    Excel's per-series `showSerName` would label every point - 60 of them per
    line. Attaching a single DataLabel at `idx` with the list-level flags off
    labels that one point and no other.
    """
    props = CharacterProperties(sz=800, b=True)
    text = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)],
                    bodyPr=RichTextProperties())
    label = DataLabel(idx=idx, showSerName=True, showVal=False, showCatName=False,
                      showLegendKey=False, showPercent=False, showBubbleSize=False,
                      dLblPos="r", txPr=text)
    series.dLbls = DataLabelList(dLbl=[label], showSerName=False, showVal=False,
                                 showCatName=False, showLegendKey=False,
                                 showPercent=False, showBubbleSize=False)


def _chart_col_span(width_cm: float, gap_cols: int = 2) -> int:
    """How many columns (starting at A) a chart this wide covers, plus a
    small gap - used to anchor a second chart beside the first without the
    two overlapping. Column pixel widths follow Excel's own approximation
    (chars * 7 + 5 at 96 dpi), since column A (labels) and the date columns
    are set to different widths."""
    target_px = width_cm * 37.795
    col_px = LABEL_COL_WIDTH * 7 + 5
    col = 1
    while col_px < target_px:
        col += 1
        col_px += DATE_COL_WIDTH * 7 + 5
    return col + gap_cols


def add_chart(ws, anchor, title, y_title, first_row, n_hotels, n_dates, header_row,
              pct=False, table=None, dates=None):
    chart = LineChart()
    chart.title = title
    chart.style = 2
    chart.y_axis.title = y_title
    chart.x_axis.title = "Check-in date"
    chart.height = CHART_HEIGHT_CM
    chart.width = CHART_WIDTH_CM
    # openpyxl defaults both axes to axPos "l" and can emit charts whose axes
    # Excel then hides; setting these explicitly avoids an axis-less chart.
    chart.x_axis.axPos = "b"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if pct:
        chart.y_axis.numFmt = "0%"

    # By default Excel expands the plot area to fill the chart, which leaves
    # the axis titles sitting on top of the tick labels. Pinning the inner
    # plot area smaller reserves a margin on the left and along the bottom so
    # each axis title has room of its own.
    # Each line carries its own name at its right-hand end, which does the
    # legend's job in place - so the legend goes, and the space it occupied
    # becomes the margin those labels need.
    chart.legend = None
    chart.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="inner", xMode="edge", yMode="edge",
        x=0.10, y=0.06, w=0.78, h=0.70))
    chart.x_axis.title.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.42, y=0.92))
    chart.y_axis.title.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.005, y=0.30))

    data = Reference(ws, min_col=1, max_col=1 + n_dates,
                     min_row=first_row, max_row=first_row + n_hotels - 1)
    chart.add_data(data, titles_from_data=True, from_rows=True)
    cats = Reference(ws, min_col=2, max_col=1 + n_dates,
                     min_row=header_row, max_row=header_row)
    chart.set_categories(cats)

    for i, s in enumerate(chart.series):
        s.smooth = False

        # A heavy, explicitly-coloured line for the reference hotel; thin
        # and auto-coloured (colour left untouched) for everyone else - a
        # dozen thin competitor lines can cross without burying the one
        # line the report is actually built around.
        is_ref = table is not None and i < len(table.index) and table.index[i] == REFERENCE_HOTEL
        line = LineProperties(w=int((REFERENCE_LINE_PT if is_ref else OTHER_LINE_PT) * 12700))
        if is_ref:
            line.solidFill = ACCENT
        s.graphicalProperties = GraphicalProperties(ln=line)

        # Anchor the name to the last point that actually has a value. Amaria
        # is sold out on most dates, so labelling a blank trailing point would
        # simply show nothing.
        idx = n_dates - 1
        if table is not None and dates is not None and i < len(table.index):
            row_vals = table.loc[table.index[i], dates[:n_dates]]
            valid = [j for j, v in enumerate(row_vals) if pd.notna(v)]
            if not valid:
                continue  # nothing plotted for this hotel; legend still names it
            idx = valid[-1]
        _label_last_point(s, idx)

    ws.add_chart(chart, anchor)


# ==========================================================================
# 6. ASSEMBLY
# --------------------------------------------------------------------------
# Walks down the Report sheet in order, remembering where each block landed
# so the change formulas and charts can point back at them.
# ==========================================================================

def build(df, rooms, coverage, dates, latest, previous, out_path, extra_runs):
    hotels = hotel_order(df["hotel"].unique())
    cur = metrics_for_run(df, rooms, coverage, latest, dates)
    prev = metrics_for_run(df, rooms, coverage, previous, dates) if previous is not None else None
    nights_config = hotel_nights_config()
    excluded = rooms_excluded_by_min_stay(df, nights_config)

    runs = run_dates(df)
    # One entry per TREND_ANCHOR_DAYS value, so adding a third anchor needs
    # no change here or in the headings below.
    trend_anchors = []
    for back in TREND_ANCHOR_DAYS:
        anchor = closest_run(runs, pd.Timestamp(latest) - pd.Timedelta(days=back))
        anchor_cur = (metrics_for_run(df, rooms, coverage, anchor, dates)
                      if anchor is not None else None)
        trend_anchors.append((back, anchor, anchor_cur))

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.sheet_view.showGridLines = False
    # Pin the hotel-name column only. Freezing rows as well would be wrong
    # here: the sheet stacks eight tables, each carrying its own date header,
    # so there is no single header row worth locking to the top.
    ws.freeze_panes = "B1"
    ws.column_dimensions["A"].width = LABEL_COL_WIDTH
    for i in range(len(dates)):
        ws.column_dimensions[get_column_letter(2 + i)].width = DATE_COL_WIDTH

    # --- title ------------------------------------------------------------
    ws["A1"] = "Hotel rate monitor"
    ws["A1"].font = TITLE_FONT
    basis = (f"Latest run {pd.Timestamp(latest):%d %b %Y}"
             + (f"  ·  compared with {pd.Timestamp(previous):%d %b %Y}"
                if previous is not None else "  ·  no earlier run to compare with yet"))
    ws["A2"] = (f"{basis}  ·  {len(hotels)} properties  ·  "
                f"{len(dates)} check-in dates from {dates[0]:%d %b %Y}  ·  "
                f"built {datetime.now():%d %b %Y %H:%M}")
    ws["A2"].font = SUB_FONT
    row = 4

    # --- summary ----------------------------------------------------------
    row = write_section(ws, row, f"  SUMMARY — next {SUMMARY_DAYS} days",
                        width=len(SUMMARY_COLS))
    row = write_summary(ws, row, cur, prev, dates, hotels, rooms, df,
                        nights_config, excluded)

    # Reserve space for the two (side-by-side) charts; they are added once
    # the tables they read from have been written and their positions are
    # known.
    chart_row = row
    row += CHART_ROWS_PER_CHART

    # --- current tables ---------------------------------------------------
    positions = {}
    for key, label, fmt, note, compare_mode in METRICS:
        row = write_section(ws, row, f"  {label.upper()}", note, width=1 + len(dates))
        header = row
        first, row = write_table(ws, row, cur[key], dates, fmt, compare_mode=compare_mode)
        positions[key] = (header, first)

    # --- minimum nights -----------------------------------------------------
    row = write_section(
        ws, row, "  MINIMUM NIGHTS",
        "Per room, the least-demanding rate plan seen that day; per hotel, "
        "the strictest such requirement across its rooms. Red when the "
        "hotel's configured nights (hotels.json) is below this - at least "
        "one room needed more nights than this run queried for.",
        width=1 + len(dates))
    _, row = write_min_stay_table(ws, row, cur["min_stay"], dates, hotels, nights_config)

    # --- separator before the trend comparison ------------------------------
    row = write_divider(ws, row, 1 + len(dates))

    # --- recent-trend comparison vs Vale Palheiro ---------------------------
    row = write_section(
        ws, row, "  COMPARISON — RECENT TREND",
        f"{REFERENCE_HOTEL} shown first and tinted, as everywhere else. Each "
        "table below is the same shape as Availability % / Lowest price "
        "above (hotels x dates), but the cell is the CHANGE against an "
        "earlier run instead of today's raw value - green up, red down, "
        f"filled when the move exceeds {BIG_MOVE:.0%} / {BIG_MOVE_PP:g} pp, "
        "exactly like the measured tables' own logic.",
        width=1 + len(dates))

    for back, anchor_run, anchor_cur in trend_anchors:
        label = f"{back} days"
        basis = (f"vs {pd.Timestamp(anchor_run):%d %b %Y}" if anchor_run is not None
                 else f"no run close enough to {label} ago yet")
        row = write_section(
            ws, row, f"  AVAILABILITY % — {basis}",
            "Today − anchor, in percentage points.  Green up, red down; "
            f"filled when the move exceeds {BIG_MOVE_PP:g} pp.",
            width=1 + len(dates))
        _, row = write_delta_table(
            ws, row, cur["availability"],
            anchor_cur["availability"] if anchor_cur is not None else None,
            dates, hotels, mode="pp")

        row = write_section(
            ws, row, f"  LOWEST PRICE — {basis}",
            "Today ÷ anchor − 1.  Green up, red down; filled when the "
            f"move exceeds {BIG_MOVE:.0%}.", width=1 + len(dates))
        _, row = write_delta_table(
            ws, row, cur["min_price"],
            anchor_cur["min_price"] if anchor_cur is not None else None,
            dates, hotels, mode="ratio")

    # --- big break before the optional, secondary sections ------------------
    if INCLUDE_ADVERTISED_PRICES or INCLUDE_ESTIMATED_SALES:
        row = write_divider(ws, row, 1 + len(dates), extra_gap=True)

    # --- advertised prices ---------------------------------------------------
    if INCLUDE_ADVERTISED_PRICES:
        row = write_advertised_price_section(ws, row, df, dates, hotels, rooms,
                                             latest, width=1 + len(dates))

    # --- estimated sales ---------------------------------------------------
    if INCLUDE_ESTIMATED_SALES:
        row = write_estimated_sales_block(wb, ws, row, df, rooms, coverage,
                                          dates, hotels, width=1 + len(dates))

    # --- charts -------------------------------------------------------------
    # Side by side: occupancy on the left, lowest price on the right.
    # A readable near-term slice rather than every column in the report.
    n_chart_dates = min(len(dates), CHART_DAYS)
    chart2_col = get_column_letter(_chart_col_span(CHART_WIDTH_CM) + 1)
    add_chart(ws, f"A{chart_row}", "Occupancy by check-in date", "Occupied",
              positions["occupancy"][1], len(hotels), n_chart_dates,
              positions["occupancy"][0], pct=True,
              table=cur["occupancy"], dates=dates)
    add_chart(ws, f"{chart2_col}{chart_row}", "Lowest price by check-in date", "EUR",
              positions["min_price"][1], len(hotels), n_chart_dates,
              positions["min_price"][0],
              table=cur["min_price"], dates=dates)

    # --- raw data sheets --------------------------------------------------
    for label, run in extra_runs:
        sheet = wb.create_sheet(label[:31])
        block = df[df["scraped_date"] == run].drop(columns=["scraped_date"], errors="ignore")
        sheet.append(list(block.columns))
        for cell in sheet[1]:
            cell.font = HEAD_FONT
            cell.fill = HEAD_FILL
        for rec in block.itertuples(index=False):
            sheet.append([v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in rec])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for i, col in enumerate(block.columns, start=1):
            sheet.column_dimensions[get_column_letter(i)].width = max(12, min(28, len(str(col)) + 4))

    wb.save(out_path)
    return len(hotels), len(dates)


# ==========================================================================
# 7. CLI
# ==========================================================================

def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Build the presentation Excel report.")
    p.add_argument("--dir", type=Path, default=pr.DEFAULT_DIR, help="workbook directory")
    p.add_argument("--days", type=int, default=REPORT_DAYS,
                   help=f"date columns, incl. the estimated-sales grids (default {REPORT_DAYS})")
    p.add_argument("--out", type=Path, default=None,
                   help="output path (default: <dir>/report_<date>.xlsx)")
    p.add_argument("--compare-to", default=None,
                   help="run date to compare against (default: the previous run)")
    args = p.parse_args(argv)

    try:
        df = pr.load_data(args.dir)
    except (FileNotFoundError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    rooms = pr.load_rooms(args.dir)
    coverage = pr.load_coverage(args.dir)

    runs = run_dates(df)
    latest = runs[0]
    if args.compare_to:
        previous = pd.Timestamp(args.compare_to).normalize()
        if previous not in runs:
            print(f"error: no run on {previous:%Y-%m-%d}. Available: "
                  f"{', '.join(f'{pd.Timestamp(r):%Y-%m-%d}' for r in runs[:8])}",
                  file=sys.stderr)
            return 2
    else:
        previous = runs[1] if len(runs) > 1 else None

    start = date.today() + timedelta(days=1)
    dates = [pd.Timestamp(start + timedelta(days=i)) for i in range(args.days)]

    # The three drill-down sheets: latest run, the one before, and a week back.
    extra = [("Data latest", latest)]
    if len(runs) > 1:
        extra.append(("Data previous", runs[1]))
    if len(runs) > 6:
        extra.append(("Data 7 runs ago", runs[6]))

    out = args.out or args.dir / f"report_{pd.Timestamp(latest):%Y%m%d}.xlsx"
    n_hotels, n_dates = build(df, rooms, coverage, dates, latest, previous, out, extra)

    print(f"Report written to {out}")
    print(f"  {n_hotels} hotel(s) x {n_dates} dates | latest run "
          f"{pd.Timestamp(latest):%Y-%m-%d}"
          + (f" vs {pd.Timestamp(previous):%Y-%m-%d}" if previous is not None
             else " | no comparison run yet"))
    if previous is None:
        print("  Change tables will populate once the scraper runs on a second day.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
