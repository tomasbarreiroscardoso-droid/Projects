"""
Reporting layer over the workbooks produced by `hotel_rates.py`.

Each run of the scraper writes one workbook; this reads them all back and
treats them as a single time series. Nothing here touches the network, so it
is safe to re-run freely.

THE CENTRAL IDEA
    One row = one observation = "on <scraped_at> we saw <hotel> offering
    <room> on <rate plan> for <checkin> at <price>". Every report is a
    different way of collapsing that table:

        fix the check-in date, vary the observation date  -> price_evolution
        fix the observation date, vary the check-in date  -> price_by_checkin
        count what was on sale vs. what exists           -> availability_rate
        compare each observation to the one before       -> daily_changes

WHAT IS AND IS NOT IN THE DATA
    The API only returns what is bookable, so absence carries meaning - and
    two different meanings that must not be confused:

        a room missing from a date  -> that room is sold out
        a date missing entirely     -> either it was sold out completely, or
                                       the run never asked about it

    Nothing in the `data` sheet can tell those apart, which is why the scraper
    also writes a `coverage` sheet listing every date it queried. Reports that
    measure availability build the full room x date grid from `rooms` and
    `coverage` (see room_night_grid) and treat anything not in `data` as sold
    out. Skipping that step overstates availability badly: Amaria reads 100%
    instead of its true ~15%.

USAGE
    Every report is a plain function returning a DataFrame, so this works both
    as a CLI and as a library:

        import price_report as pr
        df    = pr.load_data()
        rooms = pr.load_rooms()
        pr.price_evolution(df, checkin="2026-08-14")

CLI:
    ./.venv/bin/python price_report.py summary
    ./.venv/bin/python price_report.py evolution --checkin 2026-08-14
    ./.venv/bin/python price_report.py by-checkin --room-code 12345
    ./.venv/bin/python price_report.py availability
    ./.venv/bin/python price_report.py changes --threshold 10
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
#     $PY build_report.py --days 90                # narrower window
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
#     DEFAULT_DIR     where the workbooks live (default ./output)
#     KEY             the columns that uniquely identify one observation
#     MAX_GAP_DAYS    runs further apart than this downgrade an estimated
#                     sale to "low" confidence (section 5b)
#     HOTEL_CAPACITY  true room count per hotel, when you know it. Anything
#                     not listed is inferred from the data (section 4b).
# ==========================================================================

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

DEFAULT_DIR = Path(__file__).resolve().parent / "output"

#: Uniquely identifies one observation. A single workbook can legitimately hold
#: several rate plans per room per date, and several runs can share a day, so
#: all five fields are needed.
KEY = ["hotel", "room_type_code", "checkin", "rate_plan_code", "scraped_at"]

NUMERIC = ["nights", "nightly_price", "total_price", "original_price",
           "min_availability", "min_stay", "max_stay"]


# ==========================================================================
# 1. LOADING
# --------------------------------------------------------------------------
# Three loaders, one per sheet the scraper writes:
#
#   load_data      the observations themselves. Adds `days_before` (how far
#                  ahead of check-in we were looking) and `scraped_date` (the
#                  run's calendar day), which most reports group on.
#   load_rooms     the cumulative room registry, from the NEWEST workbook -
#                  it lists every room ever seen, including one that has been
#                  sold out for the entire history and so appears in no
#                  `data` sheet at all.
#   load_coverage  which check-in dates each run asked about.
#
# All three tolerate older workbooks that predate a sheet, so the history
# stays readable as the format grows.
# ==========================================================================

def load_data(directory: Path | str = DEFAULT_DIR, pattern: str = "hotel_rates_*.xlsx") -> pd.DataFrame:
    """Concatenate the `data` sheet of every matching workbook.

    Adds two derived columns used by most reports:
      days_before  how far ahead of check-in the observation was taken
      scraped_date the calendar day of the run (several runs a day collapse here)
    """
    directory = Path(directory)
    files = sorted(directory.glob(pattern))
    if not files:
        raise FileNotFoundError(f"no workbooks matching {pattern!r} in {directory}")

    frames = []
    for f in files:
        try:
            part = pd.read_excel(f, sheet_name="data", engine="openpyxl")
        except ValueError:
            print(f"  ! {f.name}: no 'data' sheet, skipped", file=sys.stderr)
            continue
        if part.empty:
            continue
        part["source_file"] = f.name
        frames.append(part)

    if not frames:
        raise ValueError(f"{len(files)} workbook(s) found in {directory}, but all were empty")

    df = pd.concat(frames, ignore_index=True)

    df["checkin"] = pd.to_datetime(df["checkin"])
    df["scraped_at"] = pd.to_datetime(df["scraped_at"])
    for col in NUMERIC:
        if col in df:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["scraped_date"] = df["scraped_at"].dt.normalize()
    df["days_before"] = (df["checkin"] - df["scraped_date"]).dt.days

    # The same run re-scraped (or a workbook copied) must not double-count.
    df = df.drop_duplicates(subset=KEY + ["occupancy_bucket"], keep="last")
    return df.sort_values(KEY).reset_index(drop=True)


def load_rooms(directory: Path | str = DEFAULT_DIR,
               pattern: str = "hotel_rates_*.xlsx") -> pd.DataFrame:
    """The `rooms` registry sheet from the newest workbook (empty frame if none).

    The registry is cumulative, so the newest workbook lists every room ever
    seen - including one that has been sold out for the whole history.
    """
    files = sorted(Path(directory).glob(pattern))
    for f in reversed(files):
        try:
            rooms = pd.read_excel(f, sheet_name="rooms", engine="openpyxl")
        except ValueError:
            continue
        if not rooms.empty:
            rooms["room_type_code"] = rooms["room_type_code"].astype(str)
            return rooms
    return pd.DataFrame(columns=["hotel", "room_type_code", "name"])


def load_coverage(directory: Path | str = DEFAULT_DIR,
                  pattern: str = "hotel_rates_*.xlsx") -> pd.DataFrame:
    """The `coverage` sheet of every workbook: which dates each run asked about.

    Older workbooks predate this sheet; they simply contribute nothing, and
    availability then falls back to the dates that returned rooms.
    """
    frames = []
    for f in sorted(Path(directory).glob(pattern)):
        try:
            part = pd.read_excel(f, sheet_name="coverage", engine="openpyxl")
        except ValueError:
            continue
        if not part.empty:
            frames.append(part)
    if not frames:
        return pd.DataFrame(columns=["hotel", "checkin", "scraped_at"])

    cov = pd.concat(frames, ignore_index=True)
    cov["checkin"] = pd.to_datetime(cov["checkin"])
    cov["scraped_at"] = pd.to_datetime(cov["scraped_at"])
    cov["scraped_date"] = cov["scraped_at"].dt.normalize()
    return cov.drop_duplicates(["hotel", "checkin", "scraped_at"])


def _price_col(price: str) -> str:
    if price not in ("nightly_price", "total_price"):
        raise ValueError("price must be 'nightly_price' or 'total_price'")
    return price


def _filter(df, hotel=None, room_code=None, rate_plan=None, bookable_only=False):
    out = df
    if hotel:
        out = out[out["hotel"].str.casefold() == str(hotel).casefold()]
    if room_code is not None:
        out = out[out["room_type_code"].astype(str) == str(room_code)]
    if rate_plan is not None:
        out = out[out["rate_plan_code"].astype(str) == str(rate_plan)]
    if bookable_only and "is_bookable" in out:
        out = out[out["is_bookable"].astype(str).str.casefold() == "true"]
    return out


# ==========================================================================
# 2. REPORT - price evolution for one check-in date as it approaches
# --------------------------------------------------------------------------
# Answers "how did each hotel move its rate for 14 August as 14 August got
# closer?" - the yield-management question. Rows are observation dates in
# order, columns are hotels, values are that hotel's cheapest rate for the
# date on that day, plus a delta against the first observation.
#
# Needs at least two runs to say anything, and only sees a check-in date
# while it sits inside the scraper's rolling window.
# ==========================================================================

def price_evolution(
    df: pd.DataFrame,
    checkin: str,
    price: str = "nightly_price",
    hotel: str | None = None,
    room_code: str | None = None,
    rate_plan: str | None = None,
) -> pd.DataFrame:
    """How each hotel moved its rate for one check-in date as inventory filled.

    Rows are observation dates (most recent last), columns are hotels; the
    value is the cheapest rate that hotel offered on that day for the date.
    Also returns the change vs. the first observation.
    """
    col = _price_col(price)
    # A rate plan can be returned non-bookable (e.g. the stay is shorter than
    # its own min_stay) alongside a bookable one for the same room/date; only
    # a bookable offer is a price anyone could actually pay.
    sub = _filter(df, hotel, room_code, rate_plan, bookable_only=True)
    sub = sub[sub["checkin"] == pd.Timestamp(checkin)]
    if sub.empty:
        return pd.DataFrame()

    grid = (
        sub.groupby(["scraped_date", "days_before", "hotel"])[col]
        .min()
        .unstack("hotel")
        .sort_index()
    )
    first = grid.ffill().bfill().iloc[0]
    for h in list(grid.columns):
        grid[f"{h} Δ"] = (grid[h] - first[h]).round(2)
    return grid.reset_index()


# ==========================================================================
# 3. REPORT - price by check-in date across hotels, one room class
# --------------------------------------------------------------------------
# The competitive-set view: a single snapshot comparing what each hotel is
# asking for each arrival date. Reads down the column to see weekend and
# season shape; across to see who is priced above whom.
#
# Comparing like with like is the hard part, since hotels use different room
# codes and names. Fix the class with `room_code` (exact, preferred - codes
# are stable across renames) or `room_contains` (substring of the name, for
# matching "a suite" across properties). Defaults to each hotel's most recent
# observation of every date.
# ==========================================================================

def price_by_checkin(
    df: pd.DataFrame,
    price: str = "nightly_price",
    room_code: str | None = None,
    room_contains: str | None = None,
    rate_plan: str | None = None,
    as_of: str | None = None,
) -> pd.DataFrame:
    """Rate per check-in date, hotels side by side.

    Fix the room class with `room_code` (exact, preferred - codes are stable)
    or `room_contains` (case-insensitive substring of the room name, for
    comparing "a suite" across hotels that use different codes).

    `as_of` picks the observation date; the default uses each hotel's most
    recent observation of every check-in date.
    """
    col = _price_col(price)
    # Same reasoning as price_evolution: a non-bookable rate plan (e.g. too
    # short a stay for its own min_stay) is not a price anyone can get.
    sub = _filter(df, room_code=room_code, rate_plan=rate_plan, bookable_only=True)
    if room_contains:
        sub = sub[sub["room_name"].str.contains(room_contains, case=False, na=False)]
    if as_of:
        sub = sub[sub["scraped_date"] == pd.Timestamp(as_of).normalize()]
    if sub.empty:
        return pd.DataFrame()

    # Keep every row of each hotel/date's newest observation - there may be
    # several rate plans, and `min` below picks the cheapest of them.
    newest = sub.groupby(["hotel", "checkin"])["scraped_at"].transform("max")
    latest = sub[sub["scraped_at"] == newest]

    grid = latest.groupby(["checkin", "hotel"])[col].min().unstack("hotel").sort_index()
    grid.index = grid.index.date
    return grid.reset_index(names="checkin")


# ==========================================================================
# 4. REPORT - availability rate per hotel per week
# --------------------------------------------------------------------------
# How full is each property, by week? Measured as the share of room-nights
# actually on sale out of the room-nights that could have been.
#
# room_night_grid does the real work: it reconstructs the full
# rooms x dates grid that *should* exist (known rooms from the registry,
# queried dates from coverage) and marks each cell available or not. Counting
# only the rows the API returned would define the answer as 100% by
# construction - every returned row is, by definition, available.
#
# `by="checkin"` buckets on the week being sold  -> which weeks are full?
# `by="scraped"` buckets on the week we looked   -> is it filling up lately?
# ==========================================================================

# ==========================================================================
# 4b. CAPACITY - how many physical rooms a property actually has
# --------------------------------------------------------------------------
# A room *type* is not a room. "Nature Park Suite" can be 12 physical suites;
# "Suite 1" is exactly one. Counting types therefore weights a 12-unit
# category the same as a 1-unit one, and a property with a single category
# (Hortas do Rio) can only ever read 0% or 100%.
#
# The engines do publish stock - `min_availability` is a real count for
# Craveiral, Hortas do Rio and Praia do Canal. What none of them publish is
# the DENOMINATOR: total rooms in the category. So it is inferred as the most
# ever seen on sale at once, across the whole history:
#
#     capacity(room type) = max(min_availability) over every date and run
#
# That is a floor, not a certainty - a room never simultaneously bookable is
# invisible. It is self-correcting: the longer the history, the closer it
# gets. Measured against ground truth it lands close - Craveiral infers 37
# against ~36 known, Praia do Canal 51 against 54 published.
#
# Because it is a floor, availability can never exceed 100% by construction:
# today's stock is one of the values the maximum was taken over.
#
# Where the true number is known, put it in HOTEL_CAPACITY and it wins. Give
# either a plain total for the hotel, or a dict of room_type_code -> units
# when the split matters.
# ==========================================================================

#: Known true capacities, overriding what the data infers. Keys are hotel
#: names exactly as they appear in hotels.json.
HOTEL_CAPACITY: dict[str, int | dict[str, int]] = {
    # "Praia do Canal": 54,
}


def capacity(df: pd.DataFrame, overrides: dict | None = None) -> pd.DataFrame:
    """Units per (hotel, room_type_code): the most ever seen on sale at once.

    Returns columns hotel, room_type_code, capacity. A room type that never
    reported a count falls back to 1, which is what "it was bookable" means
    at minimum.
    """
    overrides = HOTEL_CAPACITY if overrides is None else overrides
    sub = df.copy()
    sub["room_type_code"] = sub["room_type_code"].astype(str)

    # Stock is per room, repeated on every rate plan, so take the max across
    # plans before taking the max across dates and runs.
    per = (sub.groupby(["hotel", "room_type_code", "checkin", "scraped_date"])
              ["min_availability"].max().reset_index())
    cap = (per.groupby(["hotel", "room_type_code"])["min_availability"]
              .max().reset_index(name="capacity"))
    cap["capacity"] = cap["capacity"].fillna(1).clip(lower=1)

    for hotel, value in (overrides or {}).items():
        mask = cap["hotel"] == hotel
        if not mask.any():
            continue
        if isinstance(value, dict):
            codes = {str(k): v for k, v in value.items()}
            cap.loc[mask, "capacity"] = (
                cap.loc[mask, "room_type_code"].map(codes)
                   .fillna(cap.loc[mask, "capacity"]))
        else:
            # A plain total: scale the inferred split up to match it, so the
            # relative weight of each category is kept.
            inferred = cap.loc[mask, "capacity"].sum()
            if inferred > 0:
                cap.loc[mask, "capacity"] = (
                    cap.loc[mask, "capacity"] * float(value) / inferred)
    return cap


def capacity_by_hotel(df: pd.DataFrame, overrides: dict | None = None) -> pd.Series:
    """Total inferred rooms per hotel - the denominator of availability."""
    return capacity(df, overrides).groupby("hotel")["capacity"].sum()


def room_night_grid(df: pd.DataFrame, rooms: pd.DataFrame | None = None,
                    coverage: pd.DataFrame | None = None) -> pd.DataFrame:
    """Expand the data to one row per (hotel, run, room, checkin), available or not.

    A sold-out room is simply *absent* from the API payload rather than present
    with a null price, so availability can only be measured against the grid of
    what was asked for:

      rooms     the hotel's known room types, from the `rooms` sheet when
                supplied (authoritative - it survives a room being sold out for
                the whole history), else the rooms seen in the data.
      coverage  the check-in dates each run actually queried, from the
                `coverage` sheet. Without it a date on which *nothing* was
                available leaves no trace at all and availability overstates
                itself - badly, for a property that is often full.
    """
    known = {}
    if rooms is not None and not rooms.empty:
        for hotel, grp in rooms.groupby("hotel"):
            known[hotel] = grp["room_type_code"].astype(str).unique().tolist()

    sub = df.copy()
    sub["room_type_code"] = sub["room_type_code"].astype(str)
    sub["bookable"] = sub["is_bookable"].astype(str).str.casefold() == "true"

    # Every (hotel, run) that queried anything - including runs whose dates were
    # all sold out and so contribute no rows to `sub` at all.
    pairs = set(map(tuple, sub[["hotel", "scraped_date"]].drop_duplicates().values))
    if coverage is not None and not coverage.empty:
        pairs |= set(map(tuple, coverage[["hotel", "scraped_date"]].drop_duplicates().values))

    out = []
    for hotel, scraped in sorted(pairs, key=lambda p: (str(p[0]), p[1])):
        run = sub[(sub["hotel"] == hotel) & (sub["scraped_date"] == scraped)]
        codes = known.get(hotel) or sub.loc[sub["hotel"] == hotel, "room_type_code"].unique().tolist()
        # Only dates this run actually asked about; outside the rolling window
        # an absence means "not requested", not "sold out".
        if coverage is not None and not coverage.empty:
            covered = coverage.loc[
                (coverage["hotel"] == hotel) & (coverage["scraped_date"] == scraped), "checkin"
            ].unique()
        else:
            covered = run["checkin"].unique()
        if len(codes) == 0 or len(covered) == 0:
            continue
        full = pd.MultiIndex.from_product(
            [codes, covered], names=["room_type_code", "checkin"]
        ).to_frame(index=False)

        seen = (
            run.groupby(["room_type_code", "checkin"])
            .agg(available=("bookable", "any"),
                 offers=("bookable", "size"),
                 units=("min_availability", "max"),
                 min_price=("nightly_price", "min"))
            .reset_index()
        )
        merged = full.merge(seen, on=["room_type_code", "checkin"], how="left")
        merged["available"] = merged["available"].fillna(False).astype(bool)
        merged["offers"] = merged["offers"].fillna(0).astype(int)
        # Absent means sold out (0 units). Present but with no count published
        # means at least one, which is all "it was bookable" can tell us.
        merged["units"] = merged["units"].fillna(merged["available"].astype(int))
        merged.loc[~merged["available"], "units"] = 0
        merged["units"] = merged["units"].astype(float)
        merged["hotel"] = hotel
        merged["scraped_date"] = scraped
        out.append(merged)

    grid = pd.concat(out, ignore_index=True)
    names = sub.drop_duplicates(["hotel", "room_type_code"]).set_index(
        ["hotel", "room_type_code"])["room_name"]
    grid["room_name"] = grid.set_index(["hotel", "room_type_code"]).index.map(names)
    return grid


def availability_rate(df: pd.DataFrame, by: str = "checkin",
                      rooms: pd.DataFrame | None = None,
                      coverage: pd.DataFrame | None = None,
                      basis: str = "units") -> pd.DataFrame:
    """Share of the hotel that was on sale, per hotel per week.

    `by="checkin"` buckets on the week being sold (which weeks are full?);
    `by="scraped"` buckets on the week we looked (is it filling up lately?).

    `basis="units"` weighs each room type by how many physical rooms it holds
    (see section 4b) - the honest measure. `basis="types"` is the older
    count-of-categories reading, kept so the two can be compared.
    """
    if by not in ("checkin", "scraped"):
        raise ValueError("by must be 'checkin' or 'scraped'")
    if basis not in ("units", "types"):
        raise ValueError("basis must be 'units' or 'types'")
    bucket = "checkin" if by == "checkin" else "scraped_date"

    grid = room_night_grid(df, rooms, coverage)
    cap = capacity(df)
    grid = grid.merge(cap, on=["hotel", "room_type_code"], how="left")
    grid["capacity"] = grid["capacity"].fillna(1.0)
    grid["week"] = grid[bucket].dt.to_period("W-MON").dt.start_time.dt.date

    out = (
        grid.groupby(["hotel", "week"])
        .agg(
            room_nights=("available", "size"),
            available=("available", "sum"),
            sold_out=("available", lambda s: int((~s).sum())),
            units=("units", "sum"),
            capacity=("capacity", "sum"),
            rooms=("room_type_code", "nunique"),
            dates=("checkin", "nunique"),
            min_price=("min_price", "min"),
        )
        .reset_index()
    )
    if basis == "units":
        out["availability_rate"] = (out["units"] / out["capacity"]).round(3)
    else:
        out["availability_rate"] = (out["available"] / out["room_nights"]).round(3)
    return out.sort_values(["hotel", "week"]).reset_index(drop=True)


# ==========================================================================
# 5. REPORT - day-over-day moves above a threshold
# --------------------------------------------------------------------------
# The alerting view: what changed since the last run, filtered to moves big
# enough to care about (absolute currency by default, percentage with `pct`).
# Compares consecutive observations of the same
# (hotel, room, rate plan, check-in) series.
#
# Price moves and availability moves are reported separately on purpose. A
# room that sells out has no new price to compare - it vanishes from the
# payload - so those events come from availability_changes, which diffs the
# reconstructed grid rather than the price column.
# ==========================================================================

def daily_changes(
    df: pd.DataFrame,
    threshold: float = 5.0,
    pct: bool = False,
    price: str = "nightly_price",
) -> pd.DataFrame:
    """Consecutive-observation price moves per (hotel, room, checkin, rate plan).

    `threshold` is in currency units, or in percent when `pct=True`. Rows where
    the rate appeared or disappeared are reported separately by
    `availability_changes`, not here.
    """
    col = _price_col(price)
    keys = ["hotel", "room_type_code", "rate_plan_code", "checkin"]

    sub = df.dropna(subset=[col]).sort_values(keys + ["scraped_at"]).copy()
    g = sub.groupby(keys, sort=False)
    sub["prev_price"] = g[col].shift(1)
    sub["prev_scraped"] = g["scraped_at"].shift(1)
    sub = sub.dropna(subset=["prev_price"])

    sub["change"] = (sub[col] - sub["prev_price"]).round(2)
    sub["change_pct"] = (sub["change"] / sub["prev_price"] * 100).round(2)

    measure = sub["change_pct"] if pct else sub["change"]
    hits = sub[measure.abs() >= threshold]

    cols = keys + ["room_name", "rate_plan_name", "prev_scraped", "scraped_at",
                   "prev_price", col, "change", "change_pct", "days_before"]
    cols = [c for c in cols if c in hits.columns]
    return (
        hits[cols]
        .sort_values("change", key=lambda s: s.abs(), ascending=False)
        .reset_index(drop=True)
    )


def availability_changes(df: pd.DataFrame, rooms: pd.DataFrame | None = None,
                         coverage: pd.DataFrame | None = None) -> pd.DataFrame:
    """Room/date combinations that sold out or came back between runs.

    Works off `room_night_grid`, because a room that sells out disappears from
    the payload instead of coming back with a null price.
    """
    keys = ["hotel", "room_type_code", "checkin"]
    grid = room_night_grid(df, rooms, coverage).sort_values(keys + ["scraped_date"])
    grid["prev"] = grid.groupby(keys, sort=False)["available"].shift(1)
    flips = grid.dropna(subset=["prev"])
    flips = flips[flips["available"] != flips["prev"].astype(bool)].copy()
    flips["event"] = flips["available"].map({True: "returned", False: "sold_out"})
    flips["days_before"] = (flips["checkin"] - flips["scraped_date"]).dt.days
    cols = keys + ["room_name", "scraped_date", "event", "days_before", "min_price"]
    return flips[cols].sort_values(["scraped_date"] + keys).reset_index(drop=True)


# ==========================================================================
# 5b. ESTIMATED SALES - what appears to have been booked, and at what price
# --------------------------------------------------------------------------
# THE IDEA
#   Nobody publishes their bookings. But if a room had 4 units on sale
#   yesterday at EUR 380 and 3 today, one unit very probably sold, and EUR 380
#   is the best estimate of what it sold for. Comparing consecutive runs turns
#   the scrape history into an estimated sales ledger.
#
# WHY QUANTITY, NOT DISAPPEARANCE
#   Watching only for a room vanishing catches the LAST unit and misses every
#   sale before it - useless for a property like Craveiral with ~36 units
#   across 6 categories. The quantity decrement catches each one.
#   Amaria is the exception: its engine reports 1 no matter the true stock, so
#   only its disappearances are visible. That is a limit of the source, and
#   the confidence column says so rather than hiding it.
#
# WHY THIS IS AN ESTIMATE, NOT A FACT
#   A room can leave sale without being sold: a stop-sell, a rate closing, or
#   - most likely here - the minimum-stay rule moving. Every query asks for a
#   fixed NIGHTS-long stay, so a room drops out when the middle night goes or
#   min-stay rises to 4, while the room itself is still sellable. Hence
#   `confidence`:
#       high    stock fell but the room is still on sale (a clean decrement)
#       medium  the room disappeared entirely (sold out, or withdrawn)
#       low     the two runs are more than MAX_GAP_DAYS apart, so a lot could
#               have happened in between - including sales that cancelled out
#
# WHAT IT CANNOT KNOW
#   - Bookings made before the first scrape. There is no backfill; this ledger
#     only accrues forward.
#   - Which rate plan actually sold. `price` is the cheapest rate on offer
#     before the sale (most bookings take it); price_min/price_max bracket it.
#   - Cancellations, which raise stock again and can mask a real sale.
#   - KNOWN GAP, not yet fixed here: `price`/`price_max` below are taken
#     across every rate plan on offer, including a non-bookable one (e.g. too
#     short a stay for its own min_stay) - the same bug price_evolution and
#     price_by_checkin had before they started filtering to `is_bookable`.
#     Left alone deliberately: changing this shifts every historical
#     estimated-sale price, which deserves its own dedicated check rather
#     than a tag-along fix.
#
# It is computed fresh from the workbooks every time - never appended to a
# previous result - so improving this logic recomputes the entire history.
# ==========================================================================

#: Runs further apart than this make a decrement much weaker evidence.
MAX_GAP_DAYS = 2


def estimated_sales(df: pd.DataFrame, rooms: pd.DataFrame | None = None,
                    coverage: pd.DataFrame | None = None,
                    price: str = "nightly_price",
                    include_low: bool = True) -> pd.DataFrame:
    """One row per estimated sale event: room, stay date, units, price paid.

    Returns columns: hotel, room_type_code, room_name, checkin, sold_on,
    units_sold, price, price_min, price_max, qty_before, qty_after,
    gap_days, days_before, confidence.
    """
    col = _price_col(price)
    keys = ["hotel", "room_type_code", "checkin"]

    sub = df.copy()
    sub["room_type_code"] = sub["room_type_code"].astype(str)

    # One row per room/date/run: stock on hand and the price band on offer.
    per = (sub.groupby(keys + ["scraped_date"])
              .agg(qty=("min_availability", "max"),
                   price=(col, "min"),
                   price_max=(col, "max"),
                   room_name=("room_name", "first"))
              .reset_index())

    # Rooms absent from a run are stock 0 - but only for dates that run
    # actually queried, which is what room_night_grid reconstructs.
    grid = room_night_grid(sub, rooms, coverage)
    if grid.empty:
        return pd.DataFrame()
    grid = grid[["hotel", "room_type_code", "checkin", "scraped_date", "room_name"]]
    full = grid.merge(per.drop(columns=["room_name"]),
                      on=["hotel", "room_type_code", "checkin", "scraped_date"],
                      how="left")
    full["qty"] = full["qty"].fillna(0)

    full = full.sort_values(keys + ["scraped_date"])
    g = full.groupby(keys, sort=False)
    for src, dst in (("qty", "qty_before"), ("price", "price_before"),
                     ("price_max", "price_max_before"), ("scraped_date", "prev_run")):
        full[dst] = g[src].shift(1)
    full = full.dropna(subset=["qty_before"])

    full["units_sold"] = (full["qty_before"] - full["qty"]).clip(lower=0)
    sales = full[full["units_sold"] > 0].copy()
    if sales.empty:
        return pd.DataFrame()

    sales["gap_days"] = (sales["scraped_date"] - sales["prev_run"]).dt.days
    sales["days_before"] = (sales["checkin"] - sales["prev_run"]).dt.days
    sales["confidence"] = "medium"
    sales.loc[sales["qty"] > 0, "confidence"] = "high"
    sales.loc[sales["gap_days"] > MAX_GAP_DAYS, "confidence"] = "low"
    if not include_low:
        sales = sales[sales["confidence"] != "low"]

    # The price attributed is the one seen BEFORE the sale, so drop the
    # current-run price columns first - otherwise the rename collides with
    # them and pandas hands back two columns called "price".
    out = (sales.drop(columns=["price", "price_max"])
                .rename(columns={"scraped_date": "sold_on",
                                 "price_before": "price",
                                 "price_max_before": "price_max",
                                 "qty": "qty_after"}))
    out["price_min"] = out["price"]
    cols = ["hotel", "room_type_code", "room_name", "checkin", "sold_on",
            "units_sold", "price", "price_min", "price_max",
            "qty_before", "qty_after", "gap_days", "days_before", "confidence"]
    out = out[cols].sort_values(["sold_on", "hotel", "room_name", "checkin"])
    return out.reset_index(drop=True)


def sales_price_grid(sales: pd.DataFrame, hotel: str, dates) -> pd.DataFrame:
    """Rooms x check-in dates for one hotel: the estimated price achieved.

    Several units of the same room can sell for the same date on different
    days at different prices, so the cell is the unit-weighted mean - the
    average achieved rate, which is the number a revenue manager would quote.
    """
    if sales.empty:
        return pd.DataFrame()
    h = sales[sales["hotel"] == hotel].dropna(subset=["price"])
    if h.empty:
        return pd.DataFrame()
    h = h.assign(weighted=h["price"] * h["units_sold"])
    agg = (h.groupby(["room_name", "checkin"])
             .agg(weighted=("weighted", "sum"), units=("units_sold", "sum"))
             .reset_index())
    agg["value"] = agg["weighted"] / agg["units"]
    grid = agg.pivot(index="room_name", columns="checkin", values="value")
    return grid.reindex(columns=[d for d in dates], fill_value=None)


def advertised_price_grid(df: pd.DataFrame, hotel: str, dates,
                          as_of: str | None = None,
                          price: str = "nightly_price") -> pd.DataFrame:
    """Rooms x check-in dates for one hotel: the cheapest BOOKABLE rate on
    offer, read straight off one run - the latest, unless `as_of` names an
    earlier one.

    Unlike sales_price_grid (an inferred, historical estimate of what sold),
    this is a plain snapshot of what a guest could book on that run - no
    inference involved.
    """
    col = _price_col(price)
    sub = _filter(df, hotel=hotel, bookable_only=True)
    if sub.empty:
        return pd.DataFrame()
    run = pd.Timestamp(as_of) if as_of is not None else sub["scraped_date"].max()
    sub = sub[sub["scraped_date"] == run]
    if sub.empty:
        return pd.DataFrame()
    grid = sub.groupby(["room_name", "checkin"])[col].min().unstack("checkin")
    return grid.reindex(columns=[d for d in dates], fill_value=None)


def sales_summary(sales: pd.DataFrame) -> pd.DataFrame:
    """Per hotel: how much appears to have sold, and at what average rate."""
    if sales.empty:
        return pd.DataFrame()
    s = sales.assign(weighted=sales["price"] * sales["units_sold"])
    out = (s.groupby("hotel")
             .agg(events=("units_sold", "size"),
                  units=("units_sold", "sum"),
                  rooms=("room_name", "nunique"),
                  dates=("checkin", "nunique"),
                  revenue=("weighted", "sum"),
                  cheapest=("price", "min"),
                  dearest=("price", "max"),
                  high_conf=("confidence", lambda c: int((c == "high").sum())))
             .reset_index())
    out["avg_price"] = (out["revenue"] / out["units"]).round(2)
    return out


# ==========================================================================
# 6. SUMMARY - what history has actually been collected
# --------------------------------------------------------------------------
# Run this first. It shows, per hotel, how many runs exist and what date
# range they cover, which is usually the answer to "why is this report
# empty?" - most reports need at least two runs.
# ==========================================================================

def summary(df: pd.DataFrame) -> pd.DataFrame:
    """One row per hotel: coverage of the collected history."""
    out = (
        df.groupby("hotel")
        .agg(
            rows=("checkin", "size"),
            runs=("scraped_date", "nunique"),
            first_run=("scraped_at", "min"),
            last_run=("scraped_at", "max"),
            rooms=("room_type_code", "nunique"),
            rate_plans=("rate_plan_code", "nunique"),
            checkins=("checkin", "nunique"),
            first_checkin=("checkin", "min"),
            last_checkin=("checkin", "max"),
            min_nightly=("nightly_price", "min"),
            median_nightly=("nightly_price", "median"),
            max_nightly=("nightly_price", "max"),
        )
        .reset_index()
    )
    for c in ("first_run", "last_run", "first_checkin", "last_checkin"):
        out[c] = out[c].dt.strftime("%Y-%m-%d")
    return out


# ==========================================================================
# 7. CLI
# --------------------------------------------------------------------------
# Thin wrapper: parse arguments, load the three sheets once, dispatch to the
# report function, print or write CSV. Deliberately holds no logic of its own
# so the functions above behave identically from a notebook.
#
# Exit codes: 0 report produced, 1 no matching rows, 2 nothing to read.
# ==========================================================================

def _show(df: pd.DataFrame, csv: Path | None, label: str,
          xlsx: Path | None = None) -> int:
    """Print a report, or write it to CSV/Excel. Empty is not an error state
    worth a traceback - it usually just means only one run exists so far."""
    if df.empty:
        print(f"{label}: no matching rows.")
        return 1

    wrote = False
    if csv:
        df.to_csv(csv, index=False)
        print(f"{label}: {len(df)} row(s) -> {csv}")
        wrote = True
    if xlsx:
        # Sheet names are capped at 31 chars by Excel.
        with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=label[:31] or "report", index=False)
        print(f"{label}: {len(df)} row(s) -> {xlsx}")
        wrote = True
    if wrote:
        return 0

    with pd.option_context("display.max_rows", 200, "display.width", 200,
                           "display.max_columns", 50):
        print(f"\n== {label} ==")
        print(df.to_string(index=False))
    return 0


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__.split("\n")[1],
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dir", type=Path, default=DEFAULT_DIR, help="workbook directory")
    p.add_argument("--price", choices=["nightly_price", "total_price"],
                   default="nightly_price")
    p.add_argument("--csv", type=Path, help="write the report to CSV instead of stdout")
    p.add_argument("--xlsx", type=Path, help="write the report to an Excel file instead of stdout")
    # Not required: bare `python price_report.py` runs `summary`, which is
    # the sensible default and avoids an argparse error as a first impression.
    sub = p.add_subparsers(dest="cmd")

    sub.add_parser("summary", help="coverage of the collected history")

    ev = sub.add_parser("evolution", help="one check-in date as it approaches")
    ev.add_argument("--checkin", required=True, help="YYYY-MM-DD")
    ev.add_argument("--hotel")
    ev.add_argument("--room-code")
    ev.add_argument("--rate-plan")

    bc = sub.add_parser("by-checkin", help="price per check-in date across hotels")
    bc.add_argument("--room-code", help="exact room_type_code (preferred)")
    bc.add_argument("--room-contains", help="substring of the room name")
    bc.add_argument("--rate-plan")
    bc.add_argument("--as-of", help="observation date YYYY-MM-DD (default: latest)")

    av = sub.add_parser("availability", help="bookable share per hotel per week")
    av.add_argument("--by", choices=["checkin", "scraped"], default="checkin")

    sl = sub.add_parser("sales", help="estimated sales: what sold, and at what price")
    sl.add_argument("--hotel")
    sl.add_argument("--min-confidence", choices=["low", "medium", "high"], default="low",
                    help="drop events weaker than this (default: keep all)")
    sl.add_argument("--summary", action="store_true",
                    help="one row per hotel instead of the event list")

    ch = sub.add_parser("changes", help="day-over-day moves above a threshold")
    ch.add_argument("--threshold", type=float, default=5.0)
    ch.add_argument("--pct", action="store_true", help="threshold is a percentage")
    ch.add_argument("--flips", action="store_true",
                    help="report appear/disappear events instead of price moves")

    args = p.parse_args(argv)
    if args.cmd is None:
        args.cmd = "summary"
        print("No report given; showing `summary`. "
              "Other reports: evolution, by-checkin, availability, changes.\n")

    try:
        df = load_data(args.dir)
    except (FileNotFoundError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    rooms = load_rooms(args.dir)
    coverage = load_coverage(args.dir)
    print(f"Loaded {len(df)} rows from {df['source_file'].nunique()} workbook(s); "
          f"{len(rooms)} room(s) in the registry; "
          f"{len(coverage)} date(s) queried.")

    if args.cmd == "summary":
        return _show(summary(df), args.csv, "summary", args.xlsx)
    if args.cmd == "evolution":
        out = price_evolution(df, args.checkin, args.price, args.hotel,
                              args.room_code, args.rate_plan)
        return _show(out, args.csv, f"price evolution for {args.checkin}", args.xlsx)
    if args.cmd == "by-checkin":
        out = price_by_checkin(df, args.price, args.room_code, args.room_contains,
                               args.rate_plan, args.as_of)
        return _show(out, args.csv, "price by check-in date", args.xlsx)
    if args.cmd == "availability":
        return _show(availability_rate(df, args.by, rooms, coverage), args.csv,
                     f"availability per hotel per week (by {args.by})", args.xlsx)
    if args.cmd == "sales":
        sales = estimated_sales(df, rooms, coverage, args.price)
        if not sales.empty:
            order = {"low": 0, "medium": 1, "high": 2}
            floor = order[args.min_confidence]
            sales = sales[sales["confidence"].map(order) >= floor]
            if args.hotel:
                sales = sales[sales["hotel"].str.casefold() == args.hotel.casefold()]
        if args.summary:
            return _show(sales_summary(sales), args.csv, "estimated sales by hotel", args.xlsx)
        return _show(sales, args.csv, "estimated sales", args.xlsx)

    if args.cmd == "changes":
        if args.flips:
            return _show(availability_changes(df, rooms, coverage), args.csv,
                         "sold-out / returned events", args.xlsx)
        out = daily_changes(df, args.threshold, args.pct, args.price)
        unit = "%" if args.pct else "EUR"
        return _show(out, args.csv, f"changes >= {args.threshold}{unit}", args.xlsx)
    return 2


if __name__ == "__main__":
    sys.exit(main())
