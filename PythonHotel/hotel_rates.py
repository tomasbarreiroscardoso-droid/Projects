
"""
Hotel rate scraper.

Walks a rolling window of check-in dates for every configured property and
writes the result to a timestamped Excel workbook. One workbook per run; run it
daily and the workbooks become a price history.

THREE BOOKING ENGINES
    Properties do not all share a booking platform, so each hotel declares an
    `engine` and the run routes to the matching fetcher. All three produce
    exactly the same tidy rows, which is why everything downstream - the
    registry, the coverage sheet, price_report.py, build_report.py - is
    engine-agnostic and needed no change when the second or third engine
    arrived.

    guestcentric  (default)  Amaria, Craveiral, Hortas do Rio, Vale Palheiro
        Stateless. One GET per date against crs-api.guestcentric.net's
        `find_offers`, authenticated by a short-lived token (see
        AUTHENTICATION below).

    dedge                    Praia do Canal
        Stateful, and richer. See section 5b.

    cloudbeds                Pensão Agrícola
        Stateless, unauthenticated. See section 5c.

WHAT ONE RUN DOES
    for each hotel:
        open a session for its engine
        for each check-in date in the window:
            ask for a `nights`-long stay for `ADULTS` adults
            flatten the response into tidy rows
            record that the date was queried, even if nothing came back
    write everything to output/hotel_rates_<stamp>.xlsx

AUTHENTICATION (guestcentric)
    The `apikey` alone is never enough: find_offers answers every such call
    with `401 {"error":"Token invalid"}`. The booking engine first calls
    /hola with the apikey (plus channelKey, where the property has one) and
    receives a short-lived PASETO token, which every later call sends as
    `Authorization: Bearer`. Only the token rotates - see TokenCache.

SHEETS PRODUCED
    data       tidy long table, one row per hotel/date/room/rate plan.
               This is the sheet price_report.py reads; everything else is
               a convenience view derived from it.
    rooms      persistent room registry (code, name, first/last seen). Keyed
               on room_type_code so renames update in place.
    coverage   one row per date actually queried, with how many offers came
               back. Without it a sold-out date leaves no trace and any
               availability figure silently reports 100%.
    <rate>     one wide grid per rate plan (rooms x dates), for eyeballing.
    BestPrice  wide grid of the lowest rate per room/date.

DESIGN NOTES
    - Politeness is deliberate: one request at a time, REQUEST_DELAY apart,
      never parallel. A 60-date window over 2 hotels is ~120 requests / ~2 min.
    - Nothing about a hotel is hardcoded. Room names, rate plan names and
      occupancy buckets are all discovered from the payload, so a property can
      rename or add rooms without a code change.
    - A run that fails outright writes no workbook, so a broken job cannot
      quietly fill output/ with empty files and poison the history.

Usage:
    ./.venv/bin/python hotel_rates.py
    ./.venv/bin/python hotel_rates.py --days 30 --out ./output
    ./.venv/bin/python hotel_rates.py --days 2 --raw   # dump JSON to inspect
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
#     run_daily.sh passes no --days/--nights on purpose, so the settings
#     below (and hotels.json) are the only place those are decided.
# --------------------------------------------------------------------------
# EDITABLE SETTINGS IN THIS FILE (section 1, CONFIG)
#     HOTELS          properties to scrape - but prefer hotels.json, which
#                     overrides it and needs no code change
#     DAYS_AHEAD      how many check-in dates to walk forward. THE setting to
#                     change for a longer or shorter scrape; --days overrides
#                     it for one run only.
#     NIGHTS          DEFAULT stay length, used only for a hotel whose
#                     hotels.json entry has no "nights" of its own. Every
#                     hotel currently sets its own, so this is a fallback.
#     ADULTS 2 / CHILDREN 0     occupancy asked for
#     CURRENCY EUR / LANGUAGE en
#     REQUEST_DELAY 0.5         seconds between requests - politeness
#     TIMEOUT 20 / MAX_RETRIES 3
#     DEDGE_CULTURE en-US       language for D-EDGE properties (section 5b)
# ==========================================================================

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ==========================================================================
# 1. CONFIG
# --------------------------------------------------------------------------
# Everything you would normally want to change lives in this block.
#
# Hotels can be listed here OR - preferably - in `hotels.json` next to this
# file, which overrides this list entirely (see load_hotels). Keeping them in
# JSON means adding a property or fixing a credential never touches code that
# the scheduled job depends on.
#
# Per hotel:
#   name         label used in sheets and filenames; yours to choose
#   hotel_id     the property's permanent id, also part of the URL path
#   apikey       long-lived; published in the hotel site's "Book Now" link
#   channel_key  OPTIONAL. Amaria's booking engine sends one, Craveiral's does
#                not. Omit the field entirely rather than passing "".
#   nights       OPTIONAL. Stay length for JUST this hotel, overriding
#                --nights / NIGHTS for everyone else in the same run - use it
#                for a property whose minimum-stay rule needs more nights
#                than the global default. Check the min_stay/max_stay columns
#                in the `data` sheet to see what a hotel is actually
#                enforcing before setting this. `"nights": null` (JSON has no
#                bare "key with no value") or omitting the field entirely are
#                equivalent - both fall back to the global default.
#   extra_params OPTIONAL dict of extra query parameters, merged into every
#                request. Escape hatch for the day GuestCentric adds a new
#                required parameter - no code change needed.
#
# dedge and cloudbeds hotels need a different set of fields - see sections 5b
# and 5c below.
#
# hotels.json has no native way to add a comment, so an array entry with no
# "name" key (e.g. {"#": "some note"}) is treated as one and skipped by
# load_hotels - it never reaches validation or the scrape loop.
# ==========================================================================

HOTELS = [
    {
        "name": "Amaria",
        "hotel_id": "6444eeaa74e1a0a2",
        "apikey": "fc26985da12de073e9bdc9b211d4f4e1",
        "channel_key": "974d82150b02b308a6a79ba2f87acad6",
    },
    {
        "name": "Craveiral",
        "hotel_id": "6a90df52e4240f81",
        "apikey": "1979030c88f09c6c6626521262c5e7ff",
        # no channel_key - this booking engine does not send one
    },
    {
        "name": "Hortas do Rio",
        "hotel_id": "ac2a571772d8a7cd",
        "apikey": "b948ee3b792cca2a7111abfad6153bf7",
        "channel_key": "58e5de4e971fc00be29aa10492813ad4",
    },
]

# How many check-in dates to walk forward from tomorrow. This is THE setting
# for the size of a run - the scheduled job in run_daily.sh passes no --days,
# so this value is what it uses. `--days N` overrides it for one run only.
# Cost: roughly DAYS_AHEAD x number of hotels requests, REQUEST_DELAY apart,
# so 365 days x 6 hotels at 0.5s is about 20 minutes.
DAYS_AHEAD = 365

# Default stay length, used ONLY for a hotel whose hotels.json entry has no
# "nights" of its own (or has "nights": null). A hotel's own value always
# wins, so today, with every hotel setting one, this changes nothing.
NIGHTS = 3
ADULTS = 2
CHILDREN = 0
CURRENCY = "EUR"
LANGUAGE = "en"

REQUEST_DELAY = 0.5      # seconds between requests - be polite
TIMEOUT = 20
MAX_RETRIES = 3

API_BASE = "https://crs-api.guestcentric.net/v1.0/bookingengine"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "origin": "https://hypercommerce.guestcentric.net",
    "referer": "https://hypercommerce.guestcentric.net/",
    "user-agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36"
    ),
}


# ==========================================================================
# 2. HOTEL LIST
# --------------------------------------------------------------------------
# Loads the properties to scrape, preferring hotels.json over the HOTELS
# constant above. Validates each entry up front so a typo fails immediately
# with a clear message, rather than 40 requests into an overnight run.
# ==========================================================================

def load_hotels(path: Path | None):
    """Hotels from `hotels.json` if present, else the HOTELS list above.

    GuestCentric apikeys expire, so keeping them in a data file means
    refreshing a key is an edit to JSON, not to code (and not to whatever the
    cron job is pointed at).
    """
    if path is None:
        path = Path(__file__).resolve().parent / "hotels.json"
    if not path.exists():
        return HOTELS
    hotels = json.loads(path.read_text(encoding="utf-8"))
    # JSON has no native comment syntax. An entry with no "name" key - e.g.
    # {"#": "some note"} - is a comment, not a hotel, and is dropped here so
    # it never reaches validation or the scrape loop.
    hotels = [h for h in hotels if isinstance(h, dict) and "name" in h]
    # Each engine needs different credentials, so validate against the right
    # set rather than a lowest common denominator that would let a broken
    # entry through to fail mid-run.
    required_by_engine = {
        "guestcentric": {"name", "hotel_id", "apikey"},  # channel_key optional
        "dedge": {"name", "hotel_id", "engine_id"},
        "cloudbeds": {"name", "widget_property"},  # origin, language optional
    }
    for h in hotels:
        engine = h.get("engine", "guestcentric")
        if engine not in required_by_engine:
            raise SystemExit(f"{path}: {h.get('name')!r} has unknown engine "
                             f"{engine!r}; expected one of "
                             f"{sorted(required_by_engine)}")
        missing = required_by_engine[engine] - h.keys()
        if missing:
            raise SystemExit(f"{path}: {h.get('name')!r} ({engine}) is missing "
                             f"{sorted(missing)}")
        # `nights` may be a positive int, absent, or explicitly null - JSON
        # has no way to write "the key exists but has no value" other than
        # null, and null means "use the global --nights default" (see main).
        if h.get("nights") is not None and (
            not isinstance(h["nights"], int) or h["nights"] < 1
        ):
            raise SystemExit(f"{path}: {h.get('name')!r} has invalid 'nights' "
                             f"{h['nights']!r}; expected a positive integer, "
                             f"null, or the field omitted")
    print(f"Loaded {len(hotels)} hotel(s) from {path.name}")
    return hotels


# ==========================================================================
# 3. ROOM REGISTRY
# --------------------------------------------------------------------------
# Problem this solves: the API only returns rooms that are *available* for the
# dates asked about. A room that is sold out simply is not in the payload, so
# a run alone can never tell you the full set of rooms a hotel has - and a
# wide rooms-x-dates grid built only from today's payload would grow and
# shrink rows from run to run.
#
# So every room ever seen is remembered in output/room_registry.json, shared
# across runs. The grids are built from the registry, which keeps their shape
# stable, and price_report.py uses the same list to tell "sold out" apart from
# "this room does not exist".
#
# The registry is keyed on (hotel, room_type_code), never on the name: hotels
# rename rooms, and keying on names would create phantom duplicate rows. A
# rename updates the entry in place and keeps the old name in previous_names.
# ==========================================================================

class RoomRegistry:
    """Tracks every room ever seen, keyed by (hotel, room_type_code).

    Keying on the numeric code rather than the name means a hotel renaming
    "Suite 1" to "Ocean Suite" updates the existing entry instead of creating
    a phantom duplicate row.
    """

    def __init__(self, path: Path):
        self.path = path
        self.rooms: dict[str, dict] = {}
        if path.exists():
            try:
                self.rooms = json.loads(path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError) as exc:
                print(f"  ! could not read registry ({exc}); starting fresh")

    @staticmethod
    def key(hotel: str, code) -> str:
        return f"{hotel}||{code}"

    def observe(self, hotel: str, code, name: str, today: str) -> None:
        k = self.key(hotel, code)
        entry = self.rooms.get(k)
        if entry is None:
            self.rooms[k] = {
                "hotel": hotel,
                "room_type_code": code,
                "name": name,
                "first_seen": today,
                "last_seen": today,
                "previous_names": [],
            }
            print(f"  + new room registered: {name} (code {code})")
        else:
            if entry["name"] != name:
                print(f"  ~ room {code} renamed: {entry['name']} -> {name}")
                prev = entry.setdefault("previous_names", [])
                if entry["name"] not in prev:
                    prev.append(entry["name"])
                if name in prev:
                    prev.remove(name)
                entry["name"] = name
            entry["last_seen"] = today

    def for_hotel(self, hotel: str) -> list[dict]:
        rows = [r for r in self.rooms.values() if r["hotel"] == hotel]
        return sorted(rows, key=lambda r: r["name"].casefold())

    def all_sorted(self) -> list[dict]:
        return sorted(
            self.rooms.values(),
            key=lambda r: (r["hotel"].casefold(), r["name"].casefold()),
        )

    def save(self) -> None:
        self.path.write_text(
            json.dumps(self.rooms, indent=2, ensure_ascii=False), encoding="utf-8"
        )


# ==========================================================================
# 4. FETCHING
# --------------------------------------------------------------------------
# The network layer, and the only part that talks to GuestCentric.
#
# Two ideas do most of the work here:
#
# TokenCache  holds the short-lived session token and re-mints it on demand,
#             so a long run survives its own token expiring mid-flight.
#
# FetchError  makes failure explicit. The original version of this script
#             returned None on failure, which the caller could not tell apart
#             from "this date is sold out" - so a run where every request was
#             rejected looked like a hotel with no availability, wrote a
#             workbook full of nothing, and exited 0. Now a failed request
#             raises, and `fatal` marks the case where continuing is pointless
#             (bad credentials) rather than worth a retry (a blip).
# ==========================================================================

class FetchError(Exception):
    """A request failed. Distinct from a successful response with no rooms."""

    def __init__(self, message: str, fatal: bool = False):
        super().__init__(message)
        self.fatal = fatal  # fatal -> retrying/continuing this hotel is pointless


def auth_params(hotel: dict) -> dict:
    """The credential query parameters sent on every call for a property.

    `channel_key` is optional: Amaria's booking engine sends `channelKey` on
    every call, Craveiral's sends none at all. Passing an empty one is not the
    same as omitting it, so only include it when the hotel actually has one.

    `extra_params` is a deliberate escape hatch. If GuestCentric starts
    requiring a parameter this script knows nothing about, it can be added to
    hotels.json and it will ride along on both /hola and /find_offers, with no
    change to this file.
    """
    params = {"apikey": hotel["apikey"]}
    if hotel.get("channel_key"):
        params["channelKey"] = hotel["channel_key"]
    params.update(hotel.get("extra_params") or {})
    return params


class TokenCache:
    """Mints and holds the booking engine's session token.

    `apikey` alone is never accepted by find_offers: the booking engine first
    calls /hola with the apikey + channelKey and gets back a short-lived PASETO
    token, which every later call sends as `Authorization: Bearer`. The token
    expires well within a long run, so it is re-minted on demand.
    """

    def __init__(self, session, hotel: dict):
        self.session = session
        self.hotel = hotel
        self._token: str | None = None

    def value(self) -> str:
        if self._token is None:
            self._token = self._mint()
        return self._token

    def invalidate(self) -> None:
        self._token = None

    def _mint(self) -> str:
        params = auth_params(self.hotel)
        try:
            resp = self.session.get(
                f"{API_BASE}/hola", params=params, headers=HEADERS, timeout=TIMEOUT
            )
        except requests.RequestException as exc:
            raise FetchError(f"could not reach /hola: {exc}", fatal=True) from exc
        if resp.status_code != 200:
            raise FetchError(
                f"/hola rejected the credentials (HTTP {resp.status_code}: "
                f"{resp.text[:120]}) - check apikey and channel_key",
                fatal=True,
            )
        token = resp.json()
        if not isinstance(token, str) or not token:
            raise FetchError(f"/hola returned an unexpected token: {token!r}", fatal=True)
        return token


def fetch_offers(session, hotel: dict, checkin: date, checkout: date, auth: TokenCache):
    """Return the parsed JSON payload.

    Raises FetchError if the request never succeeded, so callers can tell a
    genuine "no availability" (empty payload) from a failed call.
    """
    occupancies = json.dumps(
        [{"adults": ADULTS, "children": CHILDREN, "babies": 0, "children_ages": []}],
        separators=(",", ":"),
    )
    params = {
        **auth_params(hotel),
        "hotel_id": hotel["hotel_id"],
        "language": LANGUAGE,
        "currency": CURRENCY,
        "checkin": checkin.isoformat(),
        "checkout": checkout.isoformat(),
        "promotion_code": "[]",
        "campaign_code": "",
        "is_mobile": "false",
        "login": "",
        "occupancies": occupancies,
    }
    url = f"{API_BASE}/{hotel['hotel_id']}/find_offers"

    last = "no attempt made"
    reminted = False
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            headers = {**HEADERS, "Authorization": f"Bearer {auth.value()}"}
            resp = session.get(url, params=params, headers=headers, timeout=TIMEOUT)
            if resp.status_code == 200:
                return resp.json()
            body = resp.text[:200].replace("\n", " ")
            last = f"HTTP {resp.status_code}: {body}"
            if resp.status_code in (401, 403):
                # Almost always just an expired token. Re-mint once and retry;
                # a second rejection means the credentials themselves are bad.
                if reminted:
                    raise FetchError(f"{last} (token refresh did not help)", fatal=True)
                print(f"  ~ token expired on {checkin}; re-minting")
                auth.invalidate()
                reminted = True
                continue
            print(f"  ! {last} on {checkin} (attempt {attempt})")
        except (requests.RequestException, json.JSONDecodeError) as exc:
            last = f"{type(exc).__name__}: {exc}"
            print(f"  ! {last} on {checkin} (attempt {attempt})")
        time.sleep(2 * attempt)
    raise FetchError(last)


# ==========================================================================
# 5. PARSING
# --------------------------------------------------------------------------
# Flattens one nested find_offers payload into flat rows.
#
# Shape of the payload:
#   rooms[]       room_type_code -> name, description, images...
#   rates[]       rate_plan_code -> name (the commercial rate plan)
#   room_rates{}  keyed by OCCUPANCY BUCKET ("200" = 2 adults, 0 children,
#                 0 babies), each holding the priced combinations of
#                 room_type_code x rate_plan_code. This is where prices live.
#
# The three lists are joined on those codes to produce one readable row per
# offer. Quirks encountered against the live API, all handled here:
#
#   - on a sold-out date `room_rates` is an empty LIST, not a dict; `or {}`
#     normalises it so .items() is always safe.
#   - a rate_plan_code appearing in room_rates can be MISSING from rates[]
#     (Craveiral's plan 45), hence the "Rate {code}" fallback name.
#   - prices are NOT always whole euros. Amaria quotes 1450, Craveiral quotes
#     350.92. Values are stored exactly as returned - never rounded.
#   - average_price_per_night_per_room is normally present; it is recomputed
#     from the total only when the API omits it.
#   - each offer carries a `restrictions` object with `minimum_stay` and
#     `maximum_stay` (confirmed on Vale Palheiro, 2026-09-04). This is the
#     minimum-stay signal price_report.py could never see before: a room
#     dropping out of the payload for a short window may be this, not a
#     genuine sold-out. `maximum_stay: 0` means "no maximum", not zero
#     nights, so it is stored exactly as returned rather than translated to
#     None - a 0 legitimately means something here.
# ==========================================================================

def parse_offers(payload, hotel_name, checkin, checkout, registry, scraped_at):
    """Flatten one find_offers payload into tidy records."""
    if not payload:
        return []

    today = scraped_at[:10]

    rooms = {r["room_type_code"]: r for r in payload.get("rooms", []) or []}
    rates = {r["rate_plan_code"]: r for r in payload.get("rates", []) or []}

    for code, room in rooms.items():
        registry.observe(hotel_name, code, room.get("name", f"Room {code}"), today)

    records = []
    room_rates = payload.get("room_rates") or {}
    # room_rates is keyed by occupancy bucket ("200" = 2 adults, 0 children)
    for bucket, entries in room_rates.items():
        for rr in entries or []:
            rcode = rr.get("room_type_code")
            pcode = rr.get("rate_plan_code")
            room = rooms.get(rcode, {})
            rate = rates.get(pcode, {})
            restrictions = rr.get("restrictions") or {}

            nights = (checkout - checkin).days
            total = rr.get("total_price")
            nightly = rr.get("average_price_per_night_per_room")
            if nightly is None and total is not None and nights:
                nightly = round(total / nights, 2)

            records.append(
                {
                    "hotel": hotel_name,
                    "checkin": checkin.isoformat(),
                    "checkout": checkout.isoformat(),
                    "nights": nights,
                    "occupancy_bucket": bucket,
                    "room_type_code": rcode,
                    "room_name": room.get("name", f"Room {rcode}"),
                    "rate_plan_code": pcode,
                    "rate_plan_name": rate.get("name", f"Rate {pcode}"),
                    "meal_plan": (rr.get("meal_plan") or {}).get("name", ""),
                    "nightly_price": nightly,
                    "total_price": total,
                    "original_price": rr.get("original_price"),
                    "currency": rr.get("currency", CURRENCY),
                    "is_bookable": bool(rr.get("is_bookable")),
                    "best_offer": bool(rr.get("best_offer")),
                    "min_availability": min(
                        (d.get("availability", 0) for d in rr.get("daily_rates") or []),
                        default=None,
                    ),
                    "min_stay": restrictions.get("minimum_stay"),
                    "max_stay": restrictions.get("maximum_stay"),
                    "scraped_at": scraped_at,
                }
            )
    return records


# ==========================================================================
# 5b. D-EDGE ENGINE
# --------------------------------------------------------------------------
# Praia do Canal runs D-EDGE (the FastBooking/Availpro lineage), not
# GuestCentric. Same job, different protocol, so it lives behind the same
# interface: given a hotel and a stay, hand back tidy records.
#
# The conversation is stateful, unlike GuestCentric's stateless query string:
#
#   POST  /session {"engineId": ...}        once per run -> a session id
#   PATCH /{sid}/session/display            once per run -> culture/currency
#   PATCH /{sid}/session/context            per date     -> the stay dates
#   GET   /{sid}/search/rooms               per date     -> what is on sale
#   GET   /{sid}/description/rooms          once per run -> the full catalogue
#
# Two useful consequences:
#
#   - `description/rooms` returns EVERY room the hotel has, so a sold-out room
#     is a fact (catalogue minus search) rather than something inferred from
#     its absence. That is strictly better than the GuestCentric side, where
#     the registry exists precisely because absence is all we get.
#   - `availableQuantity` is a real count of rooms left, which GuestCentric
#     never exposes.
#
# Culture matters: a fresh session defaults to the property's pt-BR and returns
# Portuguese room names. Setting display culture to en-US keeps this hotel's
# names consistent with the others in the report.
# ==========================================================================

DEDGE_API = "https://booking-api-hub.decms.eu"
DEDGE_CULTURE = "en-US"


def _amount(node, *path):
    """Dig a numeric amount out of D-EDGE's deeply wrapped money objects.

    Prices arrive as {converted|original: {text, value: {amount, currency}}},
    several layers down. Returns (amount, currency), or (None, None).
    """
    cur = node
    for key in path:
        if not isinstance(cur, dict):
            return None, None
        cur = cur.get(key)
    if not isinstance(cur, dict):
        return None, None
    value = (cur.get("converted") or cur.get("original") or {}).get("value") or {}
    amount = value.get("amount")
    return (amount if amount else None), value.get("currency")


class DEdgeClient:
    """One booking session for one property, reused across the date window."""

    def __init__(self, session, hotel: dict):
        self.http = session
        self.hotel = hotel
        self.sid = None

    def _headers(self):
        return {**HEADERS, "content-type": "application/json",
                "origin": self.hotel.get("origin", "https://www.praiadocanal.pt"),
                "referer": self.hotel.get("origin", "https://www.praiadocanal.pt") + "/"}

    def start(self) -> None:
        body = {"engineId": self.hotel["engine_id"]}
        try:
            r = self.http.post(f"{DEDGE_API}/session", json=body,
                               headers=self._headers(), timeout=TIMEOUT)
        except requests.RequestException as exc:
            raise FetchError(f"could not reach D-EDGE: {exc}", fatal=True) from exc
        if r.status_code >= 300:
            raise FetchError(f"session refused (HTTP {r.status_code}: {r.text[:120]}) "
                             f"- check engine_id", fatal=True)
        self.sid = r.json()["sessionId"]
        # Without this the API answers in the property's default language.
        self.http.patch(f"{DEDGE_API}/{self.sid}/session/display",
                        json={"culture": DEDGE_CULTURE, "currency": CURRENCY,
                              "countryCode": "PT"},
                        headers=self._headers(), timeout=TIMEOUT)

    def catalogue(self):
        """Every room the property has, available or not."""
        r = self.http.get(f"{DEDGE_API}/{self.sid}/description/rooms",
                          headers=self._headers(), timeout=TIMEOUT)
        if r.status_code >= 300:
            return []
        body = r.json()
        return body.get("rooms", []) if isinstance(body, dict) else (body or [])

    def rate_names(self) -> dict:
        """rateId -> {code, name} for the property's rate plans.

        search/rooms returns only ids, so without this every rate plan would be
        labelled "Rate R361572/O0" instead of "Best Available Rate".
        """
        r = self.http.get(f"{DEDGE_API}/{self.sid}/description/rates",
                          headers=self._headers(), timeout=TIMEOUT)
        if r.status_code >= 300:
            return {}
        body = r.json()
        rates = body.get("rates", []) if isinstance(body, dict) else (body or [])
        return {str(x.get("rateId")): {"code": x.get("code"), "name": x.get("name")}
                for x in rates}

    def search(self, checkin: date, checkout: date):
        """Rooms on sale for one stay. Raises FetchError like the other engine."""
        ctx = {"stayPeriod": {"arrivalDate": checkin.isoformat(),
                              "departureDate": checkout.isoformat()},
               "guests": {"adults": ADULTS, "children": CHILDREN, "infants": 0},
               "selectedHotelId": self.hotel["hotel_id"]}
        last = ""
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                p = self.http.patch(f"{DEDGE_API}/{self.sid}/session/context",
                                    json=ctx, headers=self._headers(), timeout=TIMEOUT)
                if p.status_code >= 300:
                    last = f"context HTTP {p.status_code}: {p.text[:120]}"
                else:
                    r = self.http.get(f"{DEDGE_API}/{self.sid}/search/rooms",
                                      headers=self._headers(), timeout=TIMEOUT)
                    if r.status_code == 200:
                        return r.json()
                    last = f"search HTTP {r.status_code}: {r.text[:120]}"
            except (requests.RequestException, json.JSONDecodeError) as exc:
                last = f"{type(exc).__name__}: {exc}"
            print(f"  ! {last} on {checkin} (attempt {attempt})")
            time.sleep(2 * attempt)
            # A dropped connection usually takes the server-side session with
            # it, so retrying the same request just fails again. Re-establish
            # the session before the next attempt.
            if attempt < MAX_RETRIES:
                try:
                    self.start()
                except FetchError:
                    pass
        raise FetchError(last)


def parse_dedge(payload, hotel_name, checkin, checkout, registry, scraped_at,
                rate_names=None):
    """Flatten one D-EDGE search/rooms payload into the same tidy records.

    Deliberate mapping choices:
      room_type_code  roomId - numeric and stable, like GuestCentric's code
      rate_plan_code  offerId ("R390789/O94109"). NOT rateId, which embeds the
                      stay dates ("20260909;20260912;390789") and would make
                      every date look like a brand-new rate plan.
      nightly_price   total / nights, matching how the GuestCentric side
                      reports an average nightly rate. Per-night detail is in
                      dailyPrice if it is ever needed.
    """
    if not payload:
        return []

    nights = (checkout - checkin).days
    records = []
    for room in payload.get("rooms") or []:
        rcode = room.get("roomId")
        rname = room.get("name", f"Room {rcode}")
        bookable = (room.get("availabilityType") or "").casefold() == "available"

        rates = room.get("rates") or []
        prices = [_amount(rt, "price", "totalPrice")[0] for rt in rates]
        cheapest = min((p for p in prices if p is not None), default=None)

        for rt in rates:
            total, currency = _amount(rt, "price", "totalPrice")
            crossed, _ = _amount(rt, "price", "crossedOutPrice")
            offer = rt.get("offerId") or rt.get("rateId")
            # search/rooms sometimes inlines a description and sometimes does
            # not, so fall back to the rate catalogue keyed on the numeric id
            # buried in rateId ("20260909;20260912;390789" -> "390789").
            desc = rt.get("description") or {}
            if not isinstance(desc, dict):
                desc = {}
            plan, code = desc.get("name"), desc.get("code")
            if not plan and rate_names:
                rid = str(rt.get("rateId") or "").split(";")[-1]
                looked = rate_names.get(rid) or {}
                plan, code = looked.get("name"), looked.get("code")
            # Several offers can share a rate plan (a promo variant of the same
            # BAR); keep the plan name but qualify it so the grids stay distinct.
            suffix = offer.split("/")[-1] if isinstance(offer, str) and "/" in offer else ""
            label = plan or f"Rate {code or offer}"
            if suffix and suffix != "O0":
                label = f"{label} ({suffix})"

            records.append({
                "hotel": hotel_name,
                "checkin": checkin.isoformat(),
                "checkout": checkout.isoformat(),
                "nights": nights,
                "occupancy_bucket": f"{ADULTS}{CHILDREN}0",
                "room_type_code": rcode,
                "room_name": rname,
                "rate_plan_code": offer,
                "rate_plan_name": label,
                "meal_plan": "Breakfast" if str(rt.get("breakfastIncluded")).casefold()
                             == "included" else "",
                "nightly_price": round(total / nights, 2) if total and nights else None,
                "total_price": total,
                "original_price": crossed,
                "currency": currency or CURRENCY,
                "is_bookable": bookable,
                "best_offer": total is not None and total == cheapest,
                "min_availability": rt.get("availableQuantity"),
                # No min/max-stay field has been observed in D-EDGE's payload
                # yet (unlike GuestCentric's restrictions and Cloudbeds'
                # los_min/los_max) - left blank rather than guessed.
                "min_stay": None,
                "max_stay": None,
                "scraped_at": scraped_at,
            })
    return records


# ==========================================================================
# 5c. CLOUDBEDS ENGINE
# --------------------------------------------------------------------------
# Pensão Agrícola runs Cloudbeds, not GuestCentric or D-EDGE. Simplest of the
# three: one stateless, unauthenticated POST per date - no session, no token.
#
#   POST /booking/rooms
#     checkin, checkout, currency_code, lang, widget_property
#
# `widget_property` is the property's numeric id (NOT the short slug in its
# public URL, e.g. ".../reservation/7gNOci/" - that slug never appears in
# this request). Read it from the captured request's form data, or from the
# reservation page's own JS. There is no key or token; anyone can call this.
#
# Confirmed against the live API with a 3-night stay (2026-10-14 -> 17):
#   - `rate_basic` is the TOTAL price for the whole queried stay, not a
#     nightly rate - it equalled the sum of `detailed_rates[].rate` exactly
#     (258 + 277 + 324 = 859). nightly_price is derived by dividing by
#     nights, same as the other two engines.
#   - Every response lists BOTH bookable rooms (`room_types`) and sold-out
#     ones (`unavailable_room_types`), each with its own `room_type_id` and
#     name. Unlike GuestCentric, the room registry is therefore complete
#     from the very first query rather than needing history to accumulate -
#     the same benefit D-EDGE's `description/rooms` gives.
#   - `los_min`/`los_max` (length-of-stay) are published directly per rate -
#     real minimum-stay visibility, cleaner than either other engine gives.
#     A room in `unavailable_room_types` still carries its `los_min`, but
#     that is the room's *general* policy, not necessarily *why* it is
#     unavailable on this date - it can equally just be sold out. Confirmed
#     directly: a 1-night query left one room unavailable (its own los_min
#     was 2), but a 3-night query - which satisfies every room's los_min -
#     still left four *different* rooms unavailable, genuinely sold out for
#     those dates. The two causes are not distinguishable from this payload
#     alone.
#   - The rate plan name is `package_external_name` when the offer is a
#     package (e.g. "Base Rate", or a season-specific "3 Noites Julho a
#     Setembro" confirmed live - a package can itself force an exact
#     los_min == los_max, independent of the room's own policy). The
#     unpackaged "direct" rate has package_id "0" and no name field at all;
#     `room_type_title`'s "<room> (<plan>)" suffix is tried next, then
#     "Rate {code}" as a last resort.
# ==========================================================================

CLOUDBEDS_API = "https://us2.cloudbeds.com/booking/rooms"

RATE_PLAN_SUFFIX = re.compile(r"\(([^)]+)\)\s*$")


def fetch_cloudbeds(session, hotel: dict, checkin: date, checkout: date):
    """One stateless POST for one stay. Raises FetchError like the other engines.

    A bad `widget_property` fails identically on every date in the window, so
    400/404 are treated as fatal rather than retried down the whole run.
    """
    data = {
        "checkin": checkin.isoformat(),
        "checkout": checkout.isoformat(),
        "currency_code": CURRENCY,
        "lang": hotel.get("language", "en-us"),
        "widget_property": hotel["widget_property"],
    }
    origin = hotel.get("origin", "https://us2.cloudbeds.com")
    headers = {**HEADERS, "content-type": "application/x-www-form-urlencoded",
               "origin": origin, "referer": origin + "/"}

    last = "no attempt made"
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.post(CLOUDBEDS_API, data=data, headers=headers, timeout=TIMEOUT)
            if resp.status_code == 200:
                return resp.json()
            body = resp.text[:200].replace("\n", " ")
            last = f"HTTP {resp.status_code}: {body}"
            if resp.status_code in (400, 404):
                raise FetchError(f"{last} - check widget_property", fatal=True)
            print(f"  ! {last} on {checkin} (attempt {attempt})")
        except (requests.RequestException, json.JSONDecodeError) as exc:
            last = f"{type(exc).__name__}: {exc}"
            print(f"  ! {last} on {checkin} (attempt {attempt})")
        time.sleep(2 * attempt)
    raise FetchError(last)


def parse_cloudbeds(payload, hotel_name, checkin, checkout, registry, scraped_at):
    """Flatten one /booking/rooms payload into the same tidy records.

    Deliberate mapping choices:
      room_type_code  room_type_id - numeric and stable, like the other engines.
      rate_plan_code  package_id (falls back to rate_id) - stable across the
                      window, unlike rate_id which can vary per request.
      rate_plan_name  parsed out of room_type_title's "(...)" suffix; see
                      section 5c above.
      total_price     rate_basic, falling back to summing detailed_rates if
                      a future response ever omits it.
    """
    if not payload:
        return []

    today = scraped_at[:10]
    nights = (checkout - checkin).days

    room_types = payload.get("room_types") or []
    unavailable = payload.get("unavailable_room_types") or []

    # Cloudbeds returns every room in every response - available and sold
    # out alike - so the registry is complete from the very first query,
    # unlike GuestCentric where only bookable rooms ever appear.
    for room in room_types:
        registry.observe(hotel_name, room.get("room_type_id"),
                         room.get("room_type_name", "Room"), today)
    for room in unavailable:
        registry.observe(hotel_name, room.get("room_type_id"),
                         room.get("room_type_name", "Room"), today)

    by_room = {}
    for rt in room_types:
        by_room.setdefault(rt.get("room_type_id"), []).append(rt)

    records = []
    for rcode, entries in by_room.items():
        totals = [e.get("rate_basic") for e in entries if e.get("rate_basic") is not None]
        cheapest = min(totals, default=None)

        for rt in entries:
            rname = rt.get("room_type_name", f"Room {rcode}")
            title = rt.get("room_type_title") or rname
            pcode = rt.get("package_id") or rt.get("rate_id")

            plan_name = rt.get("package_external_name")
            if not plan_name:
                m = RATE_PLAN_SUFFIX.search(title)
                plan_name = m.group(1) if m else f"Rate {pcode}"

            total = rt.get("rate_basic")
            if total is None:
                detailed = rt.get("detailed_rates") or []
                total = round(sum(d.get("rate", 0) for d in detailed), 2) if detailed else None
            nightly = round(total / nights, 2) if total is not None and nights else None

            records.append({
                "hotel": hotel_name,
                "checkin": checkin.isoformat(),
                "checkout": checkout.isoformat(),
                "nights": nights,
                "occupancy_bucket": f"{ADULTS}{CHILDREN}0",
                "room_type_code": rcode,
                "room_name": rname,
                "rate_plan_code": pcode,
                "rate_plan_name": plan_name,
                "meal_plan": "",
                "nightly_price": nightly,
                "total_price": total,
                "original_price": None,
                "currency": CURRENCY,
                "is_bookable": True,
                "best_offer": total is not None and total == cheapest,
                "min_availability": rt.get("remaining"),
                "min_stay": rt.get("los_min"),
                "max_stay": rt.get("los_max"),
                "scraped_at": scraped_at,
            })
    return records


# ==========================================================================
# 6. EXCEL OUTPUT
# --------------------------------------------------------------------------
# `data` is the source of truth - one row per observation, the sheet every
# other script reads. Everything below it is a derived view for reading by
# eye, and can be regenerated from `data` at any time.
#
# The wide grids are built rooms(registry) x dates(window) and filled from a
# lookup keyed on (room_type_code, checkin) rather than on the room name, so a
# hotel renaming a room mid-window does not blank out its row.
#
# Excel constrains sheet names: 31 characters, and none of : \ / ? * [ ].
# Rate plan names come from the API and respect none of that, so
# safe_sheet_name sanitises and de-duplicates them.
# ==========================================================================

TIDY_COLUMNS = [
    "hotel", "checkin", "checkout", "nights", "occupancy_bucket",
    "room_type_code", "room_name", "rate_plan_code", "rate_plan_name",
    "meal_plan", "nightly_price", "total_price", "original_price",
    "currency", "is_bookable", "best_offer", "min_availability",
    "min_stay", "max_stay", "scraped_at",
]

BOLD = Font(bold=True)


def safe_sheet_name(name: str, used: set) -> str:
    """Excel sheet names: <=31 chars, no : \\ / ? * [ ]"""
    clean = re.sub(r"[:\\/?*\[\]]", "-", name).strip() or "Sheet"
    clean = clean[:31]
    base, n = clean, 2
    while clean.casefold() in used:
        suffix = f"_{n}"
        clean = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(clean.casefold())
    return clean


def _freeze_and_size(ws, widths):
    ws.freeze_panes = "B2"
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for cell in ws[1]:
        cell.font = BOLD
        cell.alignment = Alignment(horizontal="center")


def write_tidy_sheet(wb, records):
    ws = wb.create_sheet("data")
    ws.append(TIDY_COLUMNS)
    for rec in records:
        ws.append([rec.get(c) for c in TIDY_COLUMNS])
    _freeze_and_size(ws, [16] * len(TIDY_COLUMNS))
    ws.auto_filter.ref = ws.dimensions


def write_coverage_sheet(wb, coverage):
    """One row per date actually queried, whether or not anything came back.

    Without this a sold-out date is indistinguishable from a date the run never
    asked about, and any availability metric silently reports 100%.
    """
    ws = wb.create_sheet("coverage")
    cols = ["hotel", "checkin", "checkout", "nights", "offers", "rooms_offered", "scraped_at"]
    ws.append(cols)
    for c in coverage:
        ws.append([c.get(k) for k in cols])
    _freeze_and_size(ws, [20, 14, 14, 8, 8, 14, 20])


def write_rooms_sheet(wb, registry):
    ws = wb.create_sheet("rooms")
    cols = ["hotel", "room_type_code", "name", "first_seen", "last_seen", "previous_names"]
    ws.append(cols)
    for r in registry.all_sorted():
        ws.append([
            r["hotel"], r["room_type_code"], r["name"],
            r["first_seen"], r["last_seen"],
            "; ".join(r.get("previous_names", [])),
        ])
    _freeze_and_size(ws, [20, 14, 32, 14, 14, 30])


def write_grid(wb, used_names, sheet_label, dates, rooms, lookup, price_field):
    """rooms x dates grid.

    `rooms` is the registry list (already alphabetical); `lookup` is keyed by
    (room_type_code, checkin) so a rename mid-window doesn't blank the row.
    """
    ws = wb.create_sheet(safe_sheet_name(sheet_label, used_names))
    ws.append(["room"] + [d.isoformat() for d in dates])
    for room in rooms:
        row = [room["name"]]
        for d in dates:
            row.append(lookup.get((room["room_type_code"], d.isoformat())))
        ws.append(row)
    _freeze_and_size(ws, [32] + [12] * len(dates))
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.number_format = "#,##0"
    return ws


def write_workbook(path, records, registry, dates, price_field="nightly_price",
                   coverage=()):
    wb = Workbook()
    default_sheet = wb.active
    if default_sheet is not None:
        wb.remove(default_sheet)

    write_tidy_sheet(wb, records)
    write_rooms_sheet(wb, registry)
    write_coverage_sheet(wb, coverage)

    used_names = {"data", "rooms", "coverage"}
    hotels = sorted({r["hotel"] for r in records} | {r["hotel"] for r in registry.all_sorted()})
    multi_hotel = len(hotels) > 1

    for hotel in hotels:
        hotel_recs = [r for r in records if r["hotel"] == hotel]
        rooms = registry.for_hotel(hotel)
        rate_names = sorted({r["rate_plan_name"] for r in hotel_recs})

        for rate_name in rate_names:
            lookup = {}
            for r in hotel_recs:
                if r["rate_plan_name"] == rate_name and r[price_field] is not None:
                    lookup[(r["room_type_code"], r["checkin"])] = r[price_field]
            label = f"{hotel} - {rate_name}" if multi_hotel else rate_name
            write_grid(wb, used_names, label, dates, rooms, lookup, price_field)

        # lowest BOOKABLE price per room/date across all rate plans - a
        # non-bookable rate (e.g. too short a stay for its own min_stay) can
        # be cheaper on paper but isn't a price anyone could actually get.
        best = {}
        for r in hotel_recs:
            if r[price_field] is None or not r["is_bookable"]:
                continue
            k = (r["room_type_code"], r["checkin"])
            if k not in best or r[price_field] < best[k]:
                best[k] = r[price_field]
        label = f"{hotel} - BestPrice" if multi_hotel else "BestPrice"
        write_grid(wb, used_names, label, dates, rooms, best, price_field)

    wb.save(path)


# ==========================================================================
# 7. MAIN
# --------------------------------------------------------------------------
# Walks hotels x dates, sleeping REQUEST_DELAY between calls, and decides what
# a run is worth at the end:
#
#   nothing queried successfully   -> write nothing, exit 1. A broken job must
#                                     not leave empty workbooks behind, since
#                                     price_report would read them as real
#                                     history.
#   queried, but all sold out      -> write the workbook anyway. "Everything
#                                     is full" is a genuine and interesting
#                                     observation; the coverage sheet records
#                                     it.
#   some requests failed           -> write, but warn and exit 1 so the log
#                                     shows the run is incomplete.
#
# The window starts tomorrow: today is a same-day booking and prices behave
# differently, so it is excluded deliberately.
# ==========================================================================

def main(argv=None):
    p = argparse.ArgumentParser(description="Scrape hotel room rates to Excel.")
    p.add_argument("--days", type=int, default=DAYS_AHEAD, help="check-in dates ahead")
    p.add_argument("--nights", type=int, default=NIGHTS,
                   help="default stay length; a hotel's 'nights' in hotels.json overrides this")
    p.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).resolve().parent / "output",
        help="output directory",
    )
    p.add_argument("--price", choices=["nightly_price", "total_price"],
                   default="nightly_price", help="value shown in the wide grids")
    p.add_argument("--raw", action="store_true", help="also dump raw JSON responses")
    p.add_argument("--hotels", type=Path, default=None,
                   help="JSON file of hotels (default: ./hotels.json if it exists)")
    args = p.parse_args(argv)

    args.out.mkdir(parents=True, exist_ok=True)
    registry = RoomRegistry(args.out / "room_registry.json")

    start = date.today() + timedelta(days=1)
    dates = [start + timedelta(days=i) for i in range(args.days)]
    scraped_at = datetime.now().isoformat(timespec="seconds")

    session = requests.Session()
    records = []
    coverage = []
    failures = 0

    for hotel in load_hotels(args.hotels):
        engine = hotel.get("engine", "guestcentric")
        # .get(..., default) only falls back when the KEY is absent - an
        # explicit "nights": null still returns None, so that is checked for
        # separately and treated the same as omitting the field entirely.
        nights = hotel.get("nights")
        if nights is None:
            nights = args.nights
        print(f"\n=== {hotel['name']} [{engine}] === ({nights} night(s))")

        auth = client = rate_names = None
        if engine == "dedge":
            client = DEdgeClient(session, hotel)
            try:
                client.start()
                rate_names = client.rate_names()
                # The catalogue lists rooms that are sold out for every date in
                # the window, which no search result would ever reveal.
                for room in client.catalogue():
                    registry.observe(hotel["name"], room.get("roomId"),
                                     room.get("name", "Room"), scraped_at[:10])
            except FetchError as exc:
                failures += 1
                print(f"  ! {hotel['name']} unavailable: {exc}")
                continue
        elif engine == "guestcentric":
            auth = TokenCache(session, hotel)
        # cloudbeds needs neither a session nor a token: each request is a
        # self-contained, unauthenticated POST (see section 5c).

        for i, checkin in enumerate(dates, start=1):
            checkout = checkin + timedelta(days=nights)
            try:
                if engine == "dedge":
                    payload = client.search(checkin, checkout)
                elif engine == "cloudbeds":
                    payload = fetch_cloudbeds(session, hotel, checkin, checkout)
                else:
                    payload = fetch_offers(session, hotel, checkin, checkout, auth)
            except FetchError as exc:
                failures += 1
                print(f"  [{i:>3}/{len(dates)}] {checkin} -> FAILED ({exc})")
                if exc.fatal:
                    print(f"  ! aborting {hotel['name']}: credentials rejected. "
                          f"Re-capture credentials for this engine (see README).")
                    break
                time.sleep(REQUEST_DELAY)
                continue

            if args.raw:
                raw_dir = args.out / "raw" / hotel["name"]
                raw_dir.mkdir(parents=True, exist_ok=True)
                (raw_dir / f"{checkin}.json").write_text(
                    json.dumps(payload, ensure_ascii=False), encoding="utf-8"
                )

            if engine == "dedge":
                new = parse_dedge(payload, hotel["name"], checkin, checkout,
                                  registry, scraped_at, rate_names)
            elif engine == "cloudbeds":
                new = parse_cloudbeds(payload, hotel["name"], checkin, checkout,
                                      registry, scraped_at)
            else:
                new = parse_offers(payload, hotel["name"], checkin, checkout,
                                   registry, scraped_at)
            records.extend(new)
            coverage.append({
                "hotel": hotel["name"],
                "checkin": checkin.isoformat(),
                "checkout": checkout.isoformat(),
                "nights": nights,
                "offers": len(new),
                "rooms_offered": len({r["room_type_code"] for r in new}),
                "scraped_at": scraped_at,
            })
            status = f"{len(new):>3} offers" if new else "  no availability"
            print(f"  [{i:>3}/{len(dates)}] {checkin} -> {status}")
            time.sleep(REQUEST_DELAY)

    if not coverage:
        # Nothing was successfully queried at all. Writing an empty timestamped
        # workbook every night would quietly poison the price history, so
        # refuse and fail loudly instead.
        print(f"\nNo dates queried successfully ({failures} failed request(s)). "
              f"Nothing written; run with --raw to inspect responses.")
        return 1

    registry.save()

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_file = args.out / f"hotel_rates_{stamp}.xlsx"
    write_workbook(out_file, records, registry, dates, args.price, coverage)

    if not records:
        # A genuine "everything is sold out" run is still worth keeping: the
        # coverage sheet is what makes that visible later.
        print(f"\nNo offers available on any of the {len(coverage)} dates queried.")
    print(f"\nDone. {len(records)} rows across {len(registry.rooms)} known rooms.")
    if failures:
        print(f"WARNING: {failures} request(s) failed; this run is incomplete.")
    print(f"Written to {out_file}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
