# Brief: hotel rate scraper

I have a working Python script, `hotel_rates.py`, that scrapes room rates from
GuestCentric booking engines into a structured Excel workbook. It has been
tested offline against a captured API response, but **never run against the
live API**. Help me get it running and then extend it.

## What it does

Calls `https://crs-api.guestcentric.net/v1.0/bookingengine/{hotel_id}/find_offers`
once per check-in date, walking forward 60 days with a constant 3-night stay.
Writes a tidy `data` sheet (one row per hotel/date/room/rate-plan), a persistent
`rooms` registry, one wide rooms-x-dates grid per rate plan, and a `BestPrice`
grid.

## Task 1 — verify against the live API

```bash
pip install requests openpyxl
python hotel_rates.py --days 2 --raw
```

If it fails, inspect the response and fix. Likely causes:
- a required header I omitted (the browser sends `origin` and `referer`; I kept
  both, but there may be more)
- rate limiting -> increase `REQUEST_DELAY`
- an unexpected response shape on dates with no availability

If it succeeds, check `output/raw/*.json` against what the booking engine shows
in a browser for the same date. Confirm specifically:
- prices are whole euros (a 3-night total of 4350 means EUR 4,350)
- the occupancy bucket is `"200"` (2 adults, 0 children)

## Task 2 — add the other hotels

For each additional property: open its booking engine in Chrome, DevTools ->
Network -> filter Fetch/XHR -> find the `find_offers` request -> copy its
`apikey` and `hotel_id`. Append to the `HOTELS` list at the top of the script.

## Task 3 — schedule it

Set up a daily cron job on macOS so I build a price history over time. Each run
writes a timestamped xlsx; `room_registry.json` is shared across runs.

## Task 4 — the reporting script (separate file)

Write `price_report.py` that reads the `data` sheets from every xlsx in
`./output/`, concatenates them into one DataFrame keyed on
`(hotel, room_type_code, checkin, rate_plan_code, scraped_at)`, and reports:

- price evolution for a given date as the date approaches (how each hotel moves
  its rate as inventory fills)
- price by check-in date across hotels, for a fixed room class
- availability rate per hotel per week
- day-over-day changes above a threshold

Keep it as a library of functions plus a CLI, so I can call pieces from a
notebook.

## Constraints

- Keep requests polite: 1/second, retries with backoff, never parallel.
- The room registry must stay keyed on `room_type_code`, not name -- hotels
  rename rooms and I don't want duplicate rows.
- Don't hardcode room names anywhere; everything is discovered from the API.
