# Hotel rate scraper

Scrapes room rates from GuestCentric, D-EDGE and Cloudbeds booking engines
into Excel workbooks, one run per day, and reports on how prices move over
time.

**New here? Read [GUIDE.md](GUIDE.md) first** — it explains what every file
does in plain language and walks through installing this on a fresh Mac.
This README is the reference: how each engine works, what was verified
against the live APIs, and why the report is built the way it is.

| File | Purpose |
| --- | --- |
| `hotel_rates.py` | Scraper. Writes a timestamped `output/hotel_rates_*.xlsx` per run. |
| `build_report.py` | Builds the polished `output/report_<date>.xlsx` for reading and presenting. |
| `hotels.json` | Which properties to scrape, their credentials, and each one's stay length. |
| `price_report.py` | Reporting library + CLI over every workbook in `output/`. |
| `run_daily.sh` | Wrapper for the scheduled run: scrape, then rebuild the report. Logs to `output/logs/`. |
| `setup.sh` | One-time install: creates `.venv/` and installs `requirements.txt`. |
| `install_schedule.sh` | Installs / removes the daily launchd job. Fills the project path into the plist. |
| `com.tomas.hotelrates.plist` | Template for that job (daily 06:30). Not used directly — `install_schedule.sh` renders it. |
| `requirements.txt` | The three packages this needs: `requests`, `pandas`, `openpyxl`. |
| `output/*.xlsx` | Per run: `data` (tidy), `rooms` (registry), `coverage` (dates queried), a grid per rate plan, `BestPrice`. |
| `output/room_registry.json` | Rooms remembered across runs, so one that stops being offered is never lost. |
| `output/logs/` | One log per day from the scheduled run; pruned after 90 days. |
| `GUIDE.md` | Plain-language tour of the project + full installation instructions. |

### Portability

Everything resolves paths relative to its own location, so **the folder can
be moved or copied anywhere and still works** — no path is written into any
script. Two things do not travel:

- **`.venv/`** hardcodes absolute paths to this folder and to the exact
  Python it was built from, so it is kept out of git and must be recreated on
  each machine with `./setup.sh`. That is what `requirements.txt` is for.
- **the installed launchd job**, which needs absolute paths by design. After
  moving the folder, re-run `./install_schedule.sh` to point it at the new
  location.

`hotels.json` is tracked, not ignored — the project cannot run without it.
The `apikey` values in it are **public by design**: each one is published in
that hotel's own "Book Now" link and is readable by anyone who visits the
site, which is exactly where they were captured from. They grant read-only
access to the same availability search a browser already performs, so this
repository is fine to make public. What must never be committed is anything
that is genuinely private — `.env`, `*.pem`, `secrets.json` and
`credentials.json` are ignored for that reason.

**Where the ignore rules live.** There is no `.gitignore`. The rules
(`output/`, `.venv/`, `__pycache__/`, the private files above, …) are in
`.git/info/exclude` at the repository root (`Projects/`), which git reads but
never pushes. So they apply on this Mac only: a fresh clone has **no** ignore
rules, and `git add -A` there would pick up `.venv/` and `output/` — copy
those lines into the clone's own `.git/info/exclude` first. To share one
specific report while `output/` stays ignored, add it by name with
`git add -f PythonHotel/output/<file>.xlsx`; the rule itself is untouched.
Once added, that file is tracked, so rebuilding it later shows up as a change
to commit.

## Authentication

`find_offers` cannot be called with an `apikey` alone — that always returns
`401 {"error":"Token invalid"}`, no matter which key or headers you send. The
booking engine performs a two-step handshake, and the scraper now does the same:

1. `GET /v1.0/bookingengine/hola?apikey=<apikey>&channelKey=<channel_key>`
   returns a short-lived PASETO session token (`v4.local.…`).
2. Every later call sends it as `Authorization: Bearer <token>`, alongside
   `apikey` and `channelKey` as query parameters.

The token expires well inside a 60-date run, so `TokenCache` re-mints it
automatically on the first 401 and only gives up if a freshly minted token is
also rejected.

**The token is the only thing that rotates.** Verified 2026-07-31: the apikey
published in amaria.pt's "Book Now" link is stable, an apikey captured from
DevTools hours earlier still authenticates, and both work with or without
`channelKey`. So `hotels.json` needs no maintenance — a `401 Token invalid` in
the logs means the handshake broke, *not* that a credential needs re-capturing.
Expect to revisit these values only if a property rebuilds its website.

## Credentials for a property

The two values are stable, so this is a one-time capture per hotel:

- **`apikey`** — on the hotel's own site, the "Book Now" link points at
  `secure.guestcentric.net/api/bg/book.php?apikey=…`. That value is the apikey.
- **`channel_key`** and **`hotel_id`** — follow that link, then read the
  booking engine's own URL and its `find_offers` request in DevTools
  (Network -> Fetch/XHR).

```json
[
  { "name": "Amaria",    "hotel_id": "6444eeaa74e1a0a2",
    "apikey": "fc26985da12de073e9bdc9b211d4f4e1",
    "channel_key": "974d82150b02b308a6a79ba2f87acad6" },
  { "name": "Craveiral", "hotel_id": "6a90df52e4240f81",
    "apikey": "1979030c88f09c6c6626521262c5e7ff" }
]
```

**`channel_key` is optional** — Amaria's booking engine sends `channelKey` on
every call, Craveiral's sends none. Omit the field entirely when the property
has none; sending an empty one is not the same as omitting it. `name`,
`hotel_id` and `apikey` are required and a malformed entry fails at startup.

**`channel_key` can be shared across sister properties.** Vale Palheiro and
Hortas do Rio are on the same GuestCentric channel and share one
`channel_key` — each hotel still has its own `hotel_id` and `apikey`. Confirmed
2026-09-04 against a captured `find_offers` request from Vale Palheiro's own
booking domain (`book.valepalheiro.com`).

**`nights` is optional, per hotel.** Some properties enforce a longer minimum
stay than the global default (`NIGHTS` in `hotel_rates.py`, or `--nights` for
one run). Set `"nights": N` on just that hotel's entry to override it for
that property only. A hotel's own value always wins over the global default,
and today **every** hotel in `hotels.json` sets its own (1 or 2 nights), so
`NIGHTS` is currently a fallback that nothing reaches:

```json
{ "name": "Some Hotel", "hotel_id": "...", "apikey": "...", "nights": 4 }
```

To go back to the global default for one hotel, either delete its `"nights"`
key or set it to `"nights": null` — JSON has no way to write "the key exists
with no value", so `null` is the explicit form of "use the default", and is
treated identically to omitting the field.

`hotels.json` also has no native way to add a comment (`#` is not valid JSON
syntax), so a plain string value on a `"#"` key is used as a convention
instead — any array entry with no `"name"` key is treated as a comment and
skipped by `load_hotels()` before validation, e.g. the first entry in the
current file explaining how the per-hotel `nights` field works.

Check the `min_stay` / `max_stay` columns in the `data` sheet before guessing
at this — see "Minimum-stay visibility" below.

### Properties on D-EDGE

Not every hotel runs GuestCentric. Praia do Canal is on **D-EDGE** (the
FastBooking/Availpro lineage), declared with `"engine": "dedge"`:

```json
{ "name": "Praia do Canal", "engine": "dedge",
  "engine_id": "JRAL", "hotel_id": 24933,
  "origin": "https://www.praiadocanal.pt" }
```

Its credentials come from the hotel site itself: open the booking page and read
`window.QW_HOTEL_ENV.ENGINE_ID` in the console for `engine_id`, and
`GET {api}/{session}/session` reports `selectedHotelId` for `hotel_id`. There
is no key or token — the session is anonymous.

The conversation is stateful rather than a query string:

```
POST  /session {"engineId": …}     once per run  → a session id
PATCH /{sid}/session/display       once per run  → culture/currency
PATCH /{sid}/session/context       per date      → the stay dates
GET   /{sid}/search/rooms          per date      → what is on sale
GET   /{sid}/description/rooms     once per run  → the full room catalogue
GET   /{sid}/description/rates     once per run  → rate plan names
```

Two things this engine does better:

- `description/rooms` lists **every** room the property has, so "sold out" is
  catalogue-minus-search — a fact rather than something inferred from absence.
  Praia do Canal's registry is complete from the first run.
- `availableQuantity` is a real count of rooms left, which GuestCentric never
  exposes.

Three quirks handled in `parse_dedge`:

- `rateId` embeds the stay dates (`20260909;20260912;390789`), so
  `rate_plan_code` uses `offerId` instead — otherwise every date would look
  like a brand-new rate plan and the grids would explode.
- `search/rooms` often omits rate names, so they are looked up from
  `description/rates` (`BARHIGH` → "Best Available Rate").
- A fresh session answers in the property's default language. The run sets
  culture to `en-US` so room names match the rest of the report.

Costs two requests per date instead of one, and the connection is occasionally
dropped mid-run — a retry re-opens the session rather than just repeating the
request, since a dropped connection takes the server-side session with it.

### Properties on Cloudbeds

Pensão Agrícola is on **Cloudbeds**, declared with `"engine": "cloudbeds"`:

```json
{ "name": "Pensão Agrícola", "engine": "cloudbeds",
  "widget_property": "161682624164059",
  "origin": "https://us2.cloudbeds.com" }
```

Simplest of the three engines: one stateless, **unauthenticated** POST per
date, no key or token at all —

```
POST /booking/rooms
  checkin, checkout, currency_code, lang, widget_property
```

`widget_property` is the property's numeric id — **not** the short slug in
its public URL (e.g. `.../reservation/7gNOci/`; that slug never appears in
the request). Capture it from DevTools -> Network -> Fetch/XHR -> the POST to
`.../booking/rooms` -> its form data. `origin` is just the hotel's Cloudbeds
subdomain (`https://us2`, `us3`, … `.cloudbeds.com`, per property).

Confirmed live 2026-09-04, with a 3-night stay:

- `rate_basic` is the **total** price for the whole stay, not nightly — it
  matched the sum of `detailed_rates[].rate` exactly.
- every response lists both bookable rooms (`room_types`) **and** sold-out
  ones (`unavailable_room_types`), each with its own id and name — so, unlike
  GuestCentric, the room registry is complete from the very first run.
- `los_min` / `los_max` (length-of-stay) are published directly per rate —
  real minimum-stay data, more explicit than either other engine gives (see
  "Minimum-stay visibility" below). A packaged rate can force an exact
  `los_min == los_max` independent of the room's own policy — confirmed on a
  season-specific 3-night-only package.

`hotels.json` overrides the `HOTELS` list in the script, so adding a property
or refreshing a credential never means editing code the scheduled job depends
on.

## Minimum-stay visibility

Every offer now carries `min_stay` and `max_stay` in the `data` sheet:

- **GuestCentric** — from `restrictions.minimum_stay` / `maximum_stay` on
  each offer. `maximum_stay: 0` means "no maximum", not zero nights, and is
  stored exactly as returned rather than translated to `None`.
- **Cloudbeds** — from `los_min` / `los_max` directly, per rate. A room
  missing from the response for a short stay can genuinely be sold out, or
  can be excluded purely because the stay is shorter than its `los_min` —
  Cloudbeds does not say which, though a room reappearing at a longer stay
  length with the same `min_availability` is a strong hint it was the latter.
- **D-EDGE** — not yet observed in the payload; both columns are blank for
  Praia do Canal until a raw response is inspected and the fields (if any)
  are identified.

This is what makes a hotel's minimum-stay rule visible without guessing: run
`--raw` once for the hotel in question, or just check these two columns in
the next `data` sheet.

### Retrying past a minimum stay (`--probe-min-stay`, off by default)

Some dates come back with real listings where **not one offer is bookable**,
because every rate plan's `min_stay` is longer than the hotel's `nights`.
Craveiral on 2026-10-08 is the example: 3 rooms, 13 offers, every one
non-bookable, `min_stay` 2–7 against a 1-night query. Left alone, that date
reads 0% available even though rooms are on sale.

With `--probe-min-stay`, such a date is queried **once more**, at the
shortest `min_stay` any of those offers asked for (2 nights, in that
example):

- **Only that date changes.** Every other date keeps the hotel's own
  `nights`, and `hotels.json` is never touched.
- **One retry, never more.** Whatever the retry returns replaces the first
  attempt's rows for that date, so two rows for the same date and rate plan
  never collide in `price_report.py` (whose key does not include `nights`).
  If the retry returns nothing at all, the first attempt's rows are kept. It
  never tries an even longer stay.
- **An empty date is never retried.** A date with no listings at all (Amaria
  on 2026-12-23) may be sold out or closed, and carries no `min_stay` to
  retry with, so it costs no extra request.
- **Every affected row is flagged.** `nights_overridden` is `True` in both
  the `data` and `coverage` sheets, and `nights` holds the stay length
  actually queried. The reports do not treat these rows differently yet: a
  2-night price sits in the same grids as 1-night prices, so filter on the
  flag when comparing prices.

Tested on a 30-day run: 5 Craveiral dates (2026-10-04 to 10-08) went from 0
bookable offers to 2–6 each, for 5 extra requests. None of the 52 empty
dates across the other hotels was retried.

`PROBE_MIN_STAY` in `hotel_rates.py` sets the default (`False`); the flag
turns it on for one run. `run_daily.sh` does not pass it.

## Running

First time on a machine, or after moving the folder, run `./setup.sh` once —
see [GUIDE.md](GUIDE.md) for the full installation walkthrough. Then:

```bash
.venv/bin/python hotel_rates.py            # a full run, DAYS_AHEAD wide
.venv/bin/python hotel_rates.py --days 2 --raw   # a quick look at the JSON
```

Useful flags: `--days` (check-in dates ahead — **defaults to `DAYS_AHEAD` in
`hotel_rates.py`**), `--nights` (fallback stay length, only
used for a hotel whose `hotels.json` entry sets no `"nights"` of its own),
`--price total_price` (value shown in the wide grids), `--raw` (dump JSON
responses to `output/raw/<hotel>/<date>.json`; a retried date also writes
`<date>_retry<N>n.json`), `--hotels path.json`, `--probe-min-stay` (retry a
date blocked only by minimum stay — see "Retrying past a minimum stay").

`--days` and `--nights` are for one-off runs. To change the size of *every*
run, including the scheduled one, edit `DAYS_AHEAD` in `hotel_rates.py` and
each hotel's `"nights"` in `hotels.json` — `run_daily.sh` deliberately passes
neither flag, so those are the single source of truth.

Budget roughly `DAYS_AHEAD × properties × REQUEST_DELAY` for a run — at the
settings in the repo today (240 × 6 × 0.5 s) that is about 12 minutes of
pauses, plus the time each response takes. D-EDGE costs a second request per
date, and so does each date retried with `--probe-min-stay`.

Exit code is `0` only when every request succeeded; `1` if any failed. A run in
which nothing could be queried at all writes nothing, so a broken job cannot
quietly fill `output/` with empty workbooks and poison the price history. A run
where every date is genuinely sold out *is* written — the `coverage` sheet is
what makes that visible later.

### Verified against the live API

Confirmed on 2026-07-31 against `output/raw/Amaria/2026-08-09.json` and the
response captured from the booking engine in a browser:

- prices are whole euros — `total_price` 4350 is EUR 4,350, exactly 3 × the
  `average_price_per_night_per_room` of 1450
- the only occupancy bucket returned is `"200"` (2 adults, 0 children)
- rooms and rate plans are discovered correctly: codes 7 / 10, plans 12 / 14

Three payload quirks worth knowing, all handled:

- on a date with no availability the API returns `"room_rates": []` — a
  **list**, where an available date returns a **dict** keyed by occupancy bucket
- **whole euros are an Amaria habit, not an API rule.** Craveiral returns
  `350.92` per night. Prices are stored exactly as returned; don't round.
- a `rate_plan_code` present in `room_rates` can be **missing from the `rates`
  array** (Craveiral's plan 45). The name then falls back to `Rate 45`, which
  is why a sheet by that name appears — it is real data, not a bug.

## Reports

```bash
.venv/bin/python price_report.py                       # same as `summary`
.venv/bin/python price_report.py summary
.venv/bin/python price_report.py evolution --checkin 2026-08-14 --room-code 102
.venv/bin/python price_report.py by-checkin --room-contains suite
.venv/bin/python price_report.py availability --by checkin
.venv/bin/python price_report.py changes --threshold 10
.venv/bin/python price_report.py changes --flips
.venv/bin/python price_report.py sales --summary
```

Reports print to the terminal by default. To get a file instead:

```bash
.venv/bin/python price_report.py --xlsx availability.xlsx availability
.venv/bin/python price_report.py --csv changes.csv changes --threshold 10
```

Global flags: `--dir` (workbook directory), `--price`, `--csv`, `--xlsx`.

Note the two scripts differ here: `hotel_rates.py` always writes a workbook (it
is collecting data), while `price_report.py` prints unless asked for a file (it
is answering a question). `evolution` and `changes` need at least two runs and
report "no matching rows" until then — that is not an error.

Every report is a function returning a DataFrame, so the same pieces work from
a notebook:

```python
import price_report as pr
df    = pr.load_data()
rooms = pr.load_rooms()
pr.price_evolution(df, "2026-08-14")
pr.availability_rate(df, by="checkin", rooms=rooms)
```

A sold-out room **disappears from the payload** rather than returning a null
price, and a fully sold-out date returns nothing at all. Availability is
therefore measured against the grid of what was *asked for* (`room_night_grid`):
the hotel's known rooms from the cumulative `rooms` sheet, crossed with the
dates each run queried from the `coverage` sheet. This matters a great deal
here — Amaria is sold out on most dates, and measuring only the rows that came
back reports 100% availability when the true figure is 15%.

### Availability counts rooms, not room types

A room *type* is not a room. "Nature Park Suite" is 12 physical suites;
Amaria's "Suite 1" is exactly one. Counting types weights those equally, and a
property with a single category can only ever read 0% or 100% — which is
exactly what Hortas do Rio did: 38 dates reporting "100% available" when the
truth ranged from 29% to 100%.

So availability is **units on sale ÷ total rooms**, and occupancy is
**1 − availability**. Both are per hotel and per check-in date: units summed
across the hotel's room types, divided by room counts summed the same way.

The engines publish stock (`min_availability`) but never the total, so each
room type's count is inferred as **the most units it was seen with at once**
— bookable or not, since a min-stay-blocked offer still reports its stock.
The question is *over which runs*:

```python
pr.capacity(df)                           # over every run in the history
pr.capacity_by_hotel(df)                  #   ...summed per hotel
pr.run_capacity(df, run_date, 20, rooms)  # over ONE run, with fallback
pr.availability_rate(df, basis="units")   # CLI report - uses capacity()
pr.availability_rate(df, basis="types")   # the old reading, for comparison
```

**All history (`capacity`) never forgets.** Vale Palheiro's Villa Terracotta
was on sale for one week (seen 2026-09-04 to 09-06 only), yet kept Vale
Palheiro at 13 rooms after it disappeared: a date with 9 villas on sale read
69% available (9/13) when the hotel was really offering 12 (75%). A
renumbering is worse. If Vale Palheiro merged 9 one-unit villas into 3
three-unit room types, the 9 dead codes would keep counting next to the 3 new
ones — 22 rooms instead of 13, permanently. The same goes for any room a hotel
retires or downsizes.

**One run (`run_capacity`) is what the presentation report uses.** Per room
type, the most units it showed on any check-in date of the latest run; a
known room type that run never listed counts **0**. Retired, downsized and
renumbered rooms drop out on the next run by themselves — the renumbering
above reads the true count, not 22. This relies on the run being wide enough
to catch each room at full stock somewhere in its window, which far-off dates
usually provide. So a hotel whose latest run had listings on fewer than
`ROOM_COUNT_MIN_DAYS` (20, in `build_report.py`) dates — a short `--days` run,
or a property closed for most of the window — **falls back to all history**
for that report. It counts dates *with listings*, not dates queried: a closed
date says nothing about how many rooms a hotel has, and Pensão Agrícola was
queried on 30 dates in one test but listed on only 4.

Either way the count is an estimate, not a fact. It misses a room never on
sale in the window measured, and it can also come out a little *above* the
real number, because it trusts each engine's own stock count — Craveiral and
Praia do Canal both read a room or two high. Counts as of the 2026-09-13 run:

| Hotel | All history | Latest run | Actual | Room types |
| --- | --- | --- | --- | --- |
| Praia do Canal | 55 | 55 | ~54 | 7 |
| Craveiral | 38 | 38 | ~36 | 6 |
| Vale Palheiro | 13 | 12 | 12 on sale | 13 |
| Amaria | 10 | 10 | 10 | 10 |
| Pensão Agrícola | 7 | 6 | ? | 7 |
| Hortas do Rio | 7 | 7 | ? | 1 |

Vale Palheiro's gap is Villa Terracotta. Pensão Agrícola's is a room Cloudbeds
lists only as unavailable, so it is registered but never had an offer.

Where the true number is known, set it in `HOTEL_CAPACITY` in
`price_report.py` and it wins over both — either a plain total (the inferred
split is scaled to match) or a `{room_type_code: units}` dict, where `0`
retires a room type outright.

**Amaria is unaffected**, and correctly so: each of its "room types" is one
named physical room (`Suite 1`, `Junior Suite 2`), so units and types agree
exactly. The correction is large elsewhere — Craveiral reads 3% where types
said 19%.

Rows are de-duplicated on `(hotel, room_type_code, checkin, rate_plan_code,
scraped_at, occupancy_bucket)`, so re-reading a copied workbook is harmless.

## The presentation report

```bash
.venv/bin/python build_report.py
```

Writes `output/report_<latest run>.xlsx`, built as a comparison **from Vale
Palheiro** rather than a flat hotel list:

- **Vale Palheiro sorts first everywhere** (Amaria second), and its row
  carries a neutral warm-grey tint in every table on the sheet - one fixed
  row to read everyone else against. Adding or removing a hotel still just
  reflows everything (`hotel_order` / `REFERENCE_HOTEL` in `build_report.py`
  are the only places that know this).
- **Report**:
  - a KPI summary per hotel over the next `SUMMARY_DAYS` days, with its own
    conditional formatting (see below) - and, beside it rather than below
    it, **SUMMARY OCCUPANCY**: the same Occupancy % figure re-averaged over
    each of the next 12 `SUMMARY_DAYS`-long periods back to back, so a
    hotel that looks fine near-term but fills up badly a few months out is
    visible without scanning the full OCCUPANCY % table by eye. Both
    tables' Vale Palheiro rows are kept aligned to the same sheet row.
  - two charts side by side, bigger than before: occupancy on the left,
    lowest price on the right, both by check-in date. Vale Palheiro draws
    thick and in the report's own accent colour; everyone else draws thin,
    auto-coloured by Excel
  - six tables of hotels x `REPORT_DAYS` dates (the charts show only the
    first `CHART_DAYS` of them): **occupancy %** (1 - availability,
    placed first), availability %, lowest price, highest price, rooms
    available, **MINIMUM NIGHTS** (see below)
  - a divider, then **COMPARISON — RECENT TREND**: the same near-term KPIs
    measured again against the run closest to each `TREND_ANCHOR_DAYS`
    value back, so a
    hotel's trajectory sits next to Vale Palheiro's rather than only its own
    latest number. Blank "then"/"Δ" cells just mean no run close enough to
    that anchor exists yet - expected until there's enough history behind it
  - a bigger break, then two optional, independently-switchable sections
    (see below): **ADVERTISED PRICES** and **ESTIMATED SOLD PRICES**
- **Number of rooms** — the room count behind every Availability % and
  Occupancy % figure. One block per hotel: its rooms down the side, the count
  each room is given in column B (the hotel's total — the Summary's "Rooms
  known" — on the block's header row), then the units each room showed on
  each check-in date of the latest run. The block title says which basis the
  hotel is on. A retired room keeps its row (the registry never deletes) and
  reads 0.
- **Data latest / previous / 7 runs ago** — the tidy rows for drilling in.

Flags: `--days` (date columns — defaults to `REPORT_DAYS` in
`build_report.py`), `--compare-to YYYY-MM-DD`, `--out`, `--dir`.

**One room count per report.** The count is measured once, from the latest
run (see "Availability counts rooms, not room types"), and every run the
report compares — the previous run and each trend anchor — is measured
against that same count. So a move in the trend tables is a change in rooms
sold, never partly a change in how many rooms a hotel is counted as having.
An older run that had more units of a room than today's count allows (Villa
Terracotta on 2026-09-04) is capped at today's count, so no run reads above
100% available. Tomorrow's report re-measures from tomorrow's run.
`ROOM_COUNT_MIN_DAYS` near the top of `build_report.py` sets the fallback
threshold, and the terminal output names any hotel that fell back.

`--days` only changes one run. To change the default width of every report,
including the scheduled one, edit `REPORT_DAYS` near the top of
`build_report.py`. It does not need to match `DAYS_AHEAD` in the scraper, but
the report can only show dates that were actually scraped: asking for more
days than were collected simply leaves the extra columns empty.

**Removed**: the old "CHANGE IN X — vs previous run" tables and the
`Previous` sheet backing them, along with `--formulas` (which only ever
controlled those tables). The new recent-trend comparison above supersedes
them with a steadier, more meaningful anchor (7/30 days) instead of
whatever the previous run happened to be.

**ADVERTISED PRICES / ESTIMATED SOLD PRICES** — two optional, per-room
price grids (rooms x check-in dates, one grid per hotel), each switched on
or off with one flag near the top of `build_report.py`:

```python
INCLUDE_ADVERTISED_PRICES = True   # per-room grid: cheapest bookable rate, latest run
INCLUDE_ESTIMATED_SALES = False    # per-room grid: inferred sold price, from stock drops
```

Flip either to `False`/`True` and rebuild - no other change needed; the
code stays either way, it just doesn't run when its flag is off. They're
independent of each other and each carries its own colour in the section
header, neither a light shade, so it's obvious which block is which: a
darker green for Advertised, the original teal-green for Estimated.

- **Advertised prices** (`pr.advertised_price_grid`) is a plain snapshot,
  not an estimate: for each room and check-in date, the cheapest *bookable*
  rate on offer on the latest run, read straight off the data. Cheap to
  compute (one run) and always available as soon as there's at least one
  run.
- **Estimated sold prices** (unchanged) infers a sale from a stock drop
  between two runs and prices it at what the room was selling for just
  before - needs at least two runs on different days, and is an estimate,
  not an observation (see the methodology note next to its summary table).

**Occupancy %, Availability %, Lowest price and Highest price are coloured
relative to Vale Palheiro, per day** (`write_table`'s `compare_mode`,
applied only to these four - not Rooms available):

- Occupancy (1 - Availability): a hotel *more* occupied than Vale Palheiro
  that day gets green text. Vale Palheiro's own cell fills green on its
  fullest day (the day's maximum across every hotel, ties included) and red
  on its emptiest (the day's minimum, ties included).
- Availability: the mirror image - a hotel showing *less* availability than
  Vale Palheiro that day (fuller, the better reading) gets green text; the
  cell's own fill is untouched. Vale Palheiro's own cell fills green on its
  best day (the day's minimum, ties included) and red on its worst (the
  day's maximum, ties included).
- Lowest price: a hotel cheaper than Vale Palheiro that day gets the same
  green text as Occupancy/Availability's "better" case. Vale Palheiro's own
  cell goes bold (no colour change) on its cheapest day.
- Highest price: no styling for other hotels. Vale Palheiro's own cell goes
  bold (no colour change) on its most expensive day.

**The SUMMARY block uses its OWN rule** for the same four ideas, applied to
30-day averages/extremes instead of one date at a time -
`_apply_summary_conditional_formatting` in `build_report.py`, not
`compare_mode` above:

- **Occupancy %**: Vale Palheiro's cell never changes fill - it always keeps
  its plain reference tint, best or worst. Only its text moves: black
  unless it is the single fullest hotel of the run (ties count), in which
  case it turns green **and bold**. Every other hotel keeps the usual green
  text when it beats Vale Palheiro, and turns bold too if it happens to be
  the outright fullest hotel - not merely fuller than Vale Palheiro, the
  actual maximum. Nothing in this column is ever red.
- **Lowest price / Median lowest / Highest price**: not reference-relative
  at all - whichever hotel actually has the lowest figure (or, for Highest
  price, the highest) gets bold text, regardless of whether that happens to
  be Vale Palheiro.
- **SUMMARY OCCUPANCY** (the 12-period side table) is the exception: it
  reuses the per-day OCCUPANCY % rule above unchanged (Vale Palheiro's cell
  fills green/red on its best/worst period), since it is that table's own
  figures, just bucketed into periods instead of shown one day at a time.

**Lowest/highest price only ever count a *bookable* offer.** A rate plan can
come back non-bookable (its own `min_stay` is longer than the stay queried)
alongside a bookable one for the same room/date - Vale Palheiro's Casita
Jasmim on 2026-10-01 is a real example: a non-bookable offer at 256.5
sitting next to a bookable one at 285. Before this was fixed, "Lowest price"
would have reported 256.5, a price nobody could actually get; it now
correctly reports 285. The same fix applies to `hotel_rates.py`'s
`BestPrice` sheet and to `price_report.py`'s `evolution` / `by-checkin`
reports. **Not yet applied** to the "Estimated sales" ledger - see the
comment above `estimated_sales` in `price_report.py`.

**MINIMUM NIGHTS** — how the report surfaces a hotel's minimum-stay rule
(GuestCentric's `restrictions.minimum_stay`, Cloudbeds' `los_min` - see
"Minimum-stay visibility" above), rather than leaving a room's absence
ambiguous between "sold out" and "stay too short":

- The grid: per room, the *least*-demanding rate plan seen that day (a room
  needs only one satisfiable rate plan to have a bookable offer at all - a
  stricter sibling rate plan alongside it is normal, not a problem). Per
  hotel, the strictest such requirement across its rooms that day. Red when
  the hotel's configured `nights` (`hotels.json`) is below it.
- The Summary table's **Nights** column (bold, black - a configured fact, not
  a measured value) shows what each hotel is currently set to; **Min-stay
  check**, right beside it, reads "Minimum stay!" in red when the MINIMUM
  NIGHTS table above shows at least one date still needing more nights than
  the current setting. It is literally that table's own red cells rolled up
  to one flag per hotel (`hotels_needing_more_nights` in `build_report.py`),
  so the two can never disagree, and it inherits the same per-(room, date)
  reduction: a room needs only one satisfiable rate plan on a given day, so
  what matters is its *least*-demanding rate plan on that specific date,
  then the *strictest* such requirement across a hotel's rooms that day - a
  minimum-stay rule is often date-dependent (a season, a weekend), so
  collapsing across dates before comparing would hide a real exclusion on a
  stricter date just because the same room was lenient elsewhere.
- **Reads the LATEST run only, deliberately** - not the whole collected
  history. A requirement an older run once observed but that has since gone
  quiet in the data (the room retired, the season passed, the rule relaxed)
  is not flagged; only what the current scrape itself still proves, on a
  date still ahead of it. An earlier version instead scanned every run ever
  collected, which could flag a hotel indefinitely over something years out
  of date with nothing on the sheet to say whether it was still true. Fixed
  after a live case on Amaria: three runs in early September 2026 once
  showed several dates needing 3 nights against its 2-night setting, but by
  the next run that room/date pair had simply dropped out of the scrape
  entirely - under the old rule Amaria stayed flagged with no way to tell
  why from the sheet alone.

One thing worth knowing: **"vs previous" in the Summary block means the
previous *run*, not literally yesterday.** If a day is missed, comparing
against a date with no data would silently produce a misleading figure.
Both dates are printed at the top of the sheet.

Nothing is written at a hard-coded row: blocks are placed by a cursor that
records where each landed, so adding hotels pushes later sections down and the
charts extend themselves. Verified with a 5-hotel fixture.

## Scheduling

**Not currently installed** — the files are ready but no job is loaded.

```bash
./install_schedule.sh              # install and load it
./install_schedule.sh --status     # is it loaded? when did it last run?
./install_schedule.sh --run-now    # run once now to prove it works
./install_schedule.sh --remove     # unload and delete it
```

`com.tomas.hotelrates.plist` is a **template**: launchd cannot expand
variables, so every path in it must be absolute. `install_schedule.sh`
substitutes wherever the folder actually lives, validates the result with
`plutil`, writes it to `~/Library/LaunchAgents/` and loads it. Re-run it
after moving the folder; edit `StartCalendarInterval` in the template to
change the time of day.

The job runs `run_daily.sh`, which scrapes and then rebuilds the report. It
passes **no** `--days` or `--nights`, so the run is exactly the size that
`DAYS_AHEAD` and `hotels.json` say it is — there is no second place where
those get decided. A failure to build the report is logged but does not mark
the scrape as failed: the data is what matters and the report can always be
regenerated.

launchd rather than `crontab`: cron on macOS needs Full Disk Access granted to
`/usr/sbin/cron`, and cron simply skips a run if the Mac is asleep at 06:30,
whereas launchd runs it on the next wake — though not if the Mac was shut
down. Logs land in `output/logs/YYYY-MM-DD.log` (kept `LOG_KEEP_DAYS`, 90);
`room_registry.json` is shared across runs so rooms are never lost when a
hotel stops offering them.

## Constraints honoured

- One request at a time, `REQUEST_DELAY` (0.5s) apart, retries with backoff —
  never parallel. A 401
  re-mints the token once rather than retrying blindly; if the fresh token is
  also rejected the hotel is abandoned immediately instead of making 3 doomed
  requests per date.
- The room registry is keyed on `room_type_code`, never on the name; a rename
  updates the entry in place and keeps the old name in `previous_names`.
- No room names are hardcoded anywhere; everything is discovered from the API.
