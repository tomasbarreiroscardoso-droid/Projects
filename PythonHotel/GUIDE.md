# Guide — what everything is, and how to install it

A plain-language tour of this project. If you want the deep technical
detail — how each booking engine works, what was verified against the live
APIs, why availability is calculated the way it is — that is in
[README.md](README.md).

**What this project does, in one sentence:** every morning it visits each
hotel's own booking website, writes down what every room costs for the
coming months (`DAYS_AHEAD` days ahead), and turns the accumulated history
into one Excel report.

---

## Part 1 — The files

### The four that matter

Almost everything you will ever touch is one of these.

#### `hotel_rates.py` — **collects the data**

Goes out to the internet and asks each hotel's booking engine "what do you
have available, and for how much?", for every check-in date from tomorrow
onwards. Writes one timestamped workbook per run into `output/`. Run it
daily; the workbooks stack up into a price history.

It is the only file that touches the network. Nothing overwrites anything —
each run is a new file — so a bad run can never destroy history.

| Setting | Line | What it does |
| --- | --- | --- |
| `DAYS_AHEAD` | **221** | **The big one.** How many check-in dates ahead to collect. |
| `NIGHTS` | 226 | Fallback stay length, used only for a hotel that sets no `nights` of its own in `hotels.json`. Currently nothing reaches it. |
| `PROBE_MIN_STAY` | 259 | `False`. When `True` (or with `--probe-min-stay`), a date that has listings but nothing bookable only because of minimum stay is asked for once more at that minimum. Affected rows are marked `nights_overridden`. |
| `ADULTS` / `CHILDREN` | 260–261 | The occupancy to ask prices for. `2` adults, `0` children. |
| `CURRENCY` / `LANGUAGE` | 262–263 | `EUR`, `en`. |
| `REQUEST_DELAY` | 265 | Seconds to wait between requests. `0.5` — politeness. Raising it makes runs slower but gentler. |
| `TIMEOUT` / `MAX_RETRIES` | 266–267 | How long to wait on a slow reply, and how many times to retry. |
| `HOTELS` | 195 | A built-in property list, **ignored whenever `hotels.json` exists**. Edit the JSON, not this. |

```bash
./.venv/bin/python hotel_rates.py               # a normal full run
./.venv/bin/python hotel_rates.py --days 3      # quick test, 3 dates only
./.venv/bin/python hotel_rates.py --days 2 --raw    # also save the raw JSON
./.venv/bin/python hotel_rates.py --hotels other.json   # a different list
./.venv/bin/python hotel_rates.py --price total_price   # grids show stay total
./.venv/bin/python hotel_rates.py --probe-min-stay      # retry min-stay-blocked dates
./.venv/bin/python hotel_rates.py --help
```

`--days`, `--nights` and `--probe-min-stay` change **one run only**. To
change every run, edit `DAYS_AHEAD` (line 221), `PROBE_MIN_STAY` (line 259)
and each hotel's `nights` in `hotels.json`.

> **What `--probe-min-stay` does.** Some dates return rooms where every offer
> needs a longer stay than the hotel's `nights` — so nothing is bookable and
> the date looks full. With the flag, that one date is asked for again, once,
> at the shortest stay those offers need. A date that returns nothing at all
> is never retried. Details: README.md → "Retrying past a minimum stay".

> **How long does a run take?** Roughly `DAYS_AHEAD × hotels ×
> REQUEST_DELAY`. Work it out from whatever those are set to — with
> `DAYS_AHEAD` at 240 that is about 12 minutes of pauses, plus the time each
> response takes. Use `--days 3` when you are just testing something.

#### `build_report.py` — **builds the Excel report**

Reads every workbook `hotel_rates.py` has ever written and produces one
polished `output/report_<date>.xlsx`: a KPI summary (plus **SUMMARY
OCCUPANCY** beside it - the same Occupancy % figure re-averaged over 12
consecutive `SUMMARY_DAYS`-long periods, so a hotel filling up months out is
visible without scanning the full table), two charts, six hotels × dates
tables, a trend comparison, optional per-room price grids, and a **Number of
rooms** sheet showing how many rooms each hotel is counted as having. This is
the file that produces the thing you actually open and show to people.

It never touches the network — it only reads what is already in `output/`,
so it is safe to run as many times as you like.

| Setting | Line | What it does |
| --- | --- | --- |
| `REPORT_DAYS` | **200** | **The big one.** How many check-in date columns wide every TABLE is. |
| `CHART_DAYS` | 202 | How many date columns the two CHARTS draw. Kept below `REPORT_DAYS` on purpose — a line across a whole year is unreadable. |
| `SUMMARY_DAYS` | 204 | The near-term window averaged into the Summary KPIs and the trend comparison - also the length of each of the 12 periods in SUMMARY OCCUPANCY. |
| `TREND_ANCHOR_DAYS` | 205 | How far back the trend comparison looks. Add an entry and it grows a section; headings follow automatically. |
| `ROOM_COUNT_MIN_DAYS` | 225 | `20`. How many dates with listings a hotel needs in the latest run before its room count is read from that run. Below it, the count uses all history instead. See "How does the report know how many rooms a hotel has?" below. |
| `INCLUDE_ADVERTISED_PRICES` | 234 | `True`/`False`. Show the per-room "what it costs right now" grids. |
| `INCLUDE_ESTIMATED_SALES` | 235 | `True`/`False`. Show the per-room inferred-sold-price grids. Off by default. |
| `REFERENCE_HOTEL` | 275 | Which hotel sorts first and is highlighted everywhere as the baseline. `"Vale Palheiro"`. |
| `ACCENT` | 246 | The report's main colour. |
| `ACCENT_ADVERTISED` / `ACCENT_SALES` | 264 / 258 | The two optional sections' header colours. |
| `CHART_HEIGHT_CM` / `CHART_WIDTH_CM` | 323–324 | Size of the two charts, in centimetres. |
| `BIG_MOVE` / `BIG_MOVE_PP` | 288 / 307 | How large a change has to be before a cell is coloured. The captions in the report quote whatever you set here. |
| `DATE_COL_WIDTH` / `LABEL_COL_WIDTH` | 318–319 | Column widths. |
| `METRICS` | 357 | The five hotels × dates tables, their number formats, and their explanatory captions. |
| `SUMMARY_PERIODS` | 1037 | `12`. How many `SUMMARY_DAYS`-long periods SUMMARY OCCUPANCY covers (360 days by default). |
| `SUMMARY_SIDE_GAP_COLS` | 1036 | `2`. Blank columns between the Summary table and SUMMARY OCCUPANCY beside it. |

```bash
./.venv/bin/python build_report.py                       # normal
./.venv/bin/python build_report.py --days 90             # narrower report
./.venv/bin/python build_report.py --days 30             # one month
./.venv/bin/python build_report.py --compare-to 2026-08-02   # pick baseline
./.venv/bin/python build_report.py --out ~/Desktop/report.xlsx
./.venv/bin/python build_report.py --help
```

`REPORT_DAYS` does not have to match `DAYS_AHEAD`, but the report can only
show dates that were actually collected — ask for more and the extra columns
are simply empty.

#### `hotels.json` — **which hotels, and their settings**

The one file you edit to add, remove or reconfigure a property. No code
change needed; `hotel_rates.py` reads it at startup and a malformed entry
fails immediately with a clear message rather than halfway through a run.

```json
{
  "name": "Vale Palheiro",
  "hotel_id": "d9d3aed39b92e11b",
  "apikey": "c591968b891e4d7414bf61c860e796ae",
  "channel_key": "58e5de4e971fc00be29aa10492813ad4",
  "nights": 1
}
```

- **`name`** — the label used everywhere in the reports. Yours to choose.
- **`nights`** — **the stay length for this hotel specifically.** This is
  the setting to change when a property's minimum stay changes. Delete it,
  or set it to `null`, to fall back to `NIGHTS` in `hotel_rates.py`.
- **`engine`** — omit it for GuestCentric (the default); `"dedge"` or
  `"cloudbeds"` for the other two. Each needs a different set of fields —
  see README.md.
- **A `{"#": "..."}` entry is a comment.** JSON has no comment syntax, so
  any entry without a `"name"` key is skipped. That is what the first entry
  in the file is.

Where the credentials come from is documented in README.md → "Credentials
for a property".

#### `run_daily.sh` — **the daily job**

Runs the scraper, then rebuilds the report, and writes everything it did to
`output/logs/YYYY-MM-DD.log`. This is what the scheduler calls each morning;
you can also run it by hand with `./run_daily.sh`.

| Setting | Line | What it does |
| --- | --- | --- |
| `LOG_KEEP_DAYS` | 40 | How many days of logs to keep. `90`. |

**It deliberately passes no `--days` and no `--nights`.** A setting belongs
in one place: if this script also named a number, editing `DAYS_AHEAD` would
appear to do nothing on the scheduled run. If you need a one-off different
window, run the script by hand with the flag instead of adding it here.

---

### The supporting cast

| File / folder | What it is |
| --- | --- |
| `price_report.py` | The engine room, plus a terminal tool. `build_report.py` imports it for all the actual maths (availability, capacity, price history, estimated sales). You can also run it directly to ask one quick question without building a whole spreadsheet — see below. |
| `setup.sh` | Run once per machine: creates `.venv/` and installs the packages. Also the fix if the environment ever breaks. |
| `install_schedule.sh` | Installs, checks, triggers or removes the daily scheduled job. |
| `com.tomas.hotelrates.plist` | A **template** for that job, holding the schedule (06:30 daily, lines 42–45). Do not copy it by hand — `install_schedule.sh` fills in the folder path and installs it. |
| `requirements.txt` | The three packages needed: `requests`, `pandas`, `openpyxl`. |
| `.git/info/exclude` | In the repository root (`Projects/`), not this folder. The ignore rules: keeps `.venv/`, `output/` and `__pycache__/` out of version control. Local to this Mac and never pushed — there is no `.gitignore`. |
| `output/` | Everything produced. `hotel_rates_<stamp>.xlsx` = one scrape; `report_<date>.xlsx` = the readable report. |
| `output/room_registry.json` | Every room ever seen, remembered across runs — so a room that sells out for months is still counted, not silently forgotten. |
| `output/logs/` | One log per day from the scheduled run. First place to look when something did not happen. |
| `.venv/` | The private Python installation for this project. Machine-specific — never copy it between computers, recreate it with `./setup.sh`. |
| `__pycache__/` | Python's own speed-up cache. Ignore it; safe to delete at any time. |
| `.vscode/` | Tells VS Code to use `.venv` automatically. |
| `claude_code_brief.md` | Historical notes from when the project was first built. |

`price_report.py`, if you want it:

```bash
./.venv/bin/python price_report.py summary
./.venv/bin/python price_report.py availability --by checkin
./.venv/bin/python price_report.py evolution --checkin 2026-10-14
./.venv/bin/python price_report.py by-checkin --room-contains suite
./.venv/bin/python price_report.py changes --threshold 10
./.venv/bin/python price_report.py sales --summary
./.venv/bin/python price_report.py --xlsx out.xlsx availability   # to a file
```

Its own settings worth knowing: `HOTEL_CAPACITY` (line 407) — the true room
count for a hotel, when you know it, overriding the figure inferred from the
data (a room type set to `0` stops counting altogether); and `MAX_GAP_DAYS`
(line 756), how far apart two runs can be before an estimated sale is
downgraded to low confidence.

---

## Part 2 — Installing on a Mac

Start to finish, on a Mac that has never run this before. Everything happens
in **Terminal** (press `Cmd + Space`, type `Terminal`, press Enter).

> You do **not** need VS Code, PyCharm or any other editor to run this. A
> plain text editor is enough to change a setting. VS Code is only a
> convenience — see the optional step at the end.

### Step 1 — Put the folder somewhere

Anywhere is fine: Desktop, Documents, Downloads. Nothing in the project
records where it lives, so it works from any location and can be moved
later. Two mild suggestions:

- **Avoid iCloud-synced folders** (Desktop and Documents, if you have
  "Desktop & Documents Folders" turned on in iCloud). The daily run writes a
  file every morning and iCloud will keep re-uploading them.
- **Avoid folder names with unusual characters.** Spaces are handled fine.

Then, in Terminal, move into the folder. The easy way: type `cd` and a
space, then drag the folder from Finder into the Terminal window and press
Enter.

```bash
cd ~/Desktop/PythonHotel     # or wherever you put it
ls                           # you should see hotel_rates.py, setup.sh, ...
```

### Step 2 — Check whether you have Python

```bash
python3 --version
```

You need **3.10 or newer** (3.13 is what this was built on).

- Prints `Python 3.13.7`, or anything `3.10`+ → **skip to step 3.**
- Prints `Python 3.9.6` or lower, or "command not found", or opens a pop-up
  offering to install developer tools → **install Python first**, below.

<details>
<summary><b>Installing Python (only if you need to)</b></summary>

macOS ships an old Python that cannot run this. Two options:

**Option A — Homebrew** (recommended; makes future updates easy)

```bash
# 1. install Homebrew, if you don't have it. It will ask for your password.
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

At the end Homebrew prints two or three "Next steps" commands to run —
**run them.** On Apple Silicon Macs they add Homebrew to your PATH, and
skipping them is the most common reason `brew` is "not found" afterwards.
Then:

```bash
brew install python@3.13
python3 --version        # should now say 3.13.x
```

If `python3 --version` still shows the old version, close Terminal, open a
new window and try again.

**Option B — the official installer**

Download from <https://www.python.org/downloads/>, open the `.pkg`, click
through it. Simpler, but you update it by hand each time.

</details>

### Step 3 — Run the setup script

```bash
./setup.sh
```

This creates `.venv/` — a private Python installation just for this project,
so its packages can never clash with anything else on your Mac — and
installs `requests`, `pandas` and `openpyxl` into it. Takes a minute or two;
it needs an internet connection.

<details>
<summary><b>If it says "permission denied"</b></summary>

The scripts lost their executable flag (this happens with some ways of
copying or unzipping a folder). Fix it once:

```bash
chmod +x *.sh
./setup.sh
```

Or just run it through bash, which does not care about the flag:

```bash
bash setup.sh
```

</details>

<details>
<summary><b>Other things that can go wrong here</b></summary>

- **"no Python 3.10 or newer found"** — go back to step 2.
- **The download hangs or fails** — no internet, or a corporate network
  blocking `pypi.org`. On a company Mac, try a home network or a phone
  hotspot.
- **"cannot be opened because it is from an unidentified developer"** —
  macOS Gatekeeper. It does not apply to scripts you run from Terminal, only
  to double-clicking them, so run them from Terminal as shown here.
- **You moved the folder after setting up** — just run `./setup.sh` again.
  It rebuilds the environment from scratch, which is exactly the fix.

</details>

### Step 4 — Collect some data

Start small, so you find out in 30 seconds rather than a quarter of an hour if
something is wrong:

```bash
./.venv/bin/python hotel_rates.py --days 3
```

You should see each hotel being visited, date by date, and finally
`Written to .../output/hotel_rates_<stamp>.xlsx`. Then do a real run:

```bash
./.venv/bin/python hotel_rates.py
```

<details>
<summary><b>If a hotel fails</b></summary>

- **`401 Token invalid`** — the authentication handshake with that hotel's
  booking engine broke. See README.md → "Authentication".
- **One hotel fails, the others work** — that property most likely rebuilt
  its website and its ids changed. README.md → "Credentials for a property".
- **Everything fails** — check your internet connection first.

A run where nothing at all could be collected writes no workbook on purpose,
so a broken job can never fill `output/` with empty files and poison the
history.

</details>

### Step 5 — Build the report

```bash
./.venv/bin/python build_report.py
```

Prints where it wrote the file. Open it:

```bash
open output/report_*.xlsx
```

**This is the finish line** — if a spreadsheet opens with hotels down the
side and dates across the top, everything works.

> Some things need more than one run to appear: the "vs previous" movement
> column, the recent-trend comparison, and estimated sales all compare two
> runs. On day one those are blank. That is expected, not a fault.

### Step 6 — Optional: run it automatically every morning

```bash
./install_schedule.sh
```

That is the whole thing. It fills your actual folder path into the schedule
template, installs it into `~/Library/LaunchAgents/`, and loads it. From
then on, every morning at **06:30**, your Mac scrapes and rebuilds the
report by itself.

```bash
./install_schedule.sh --status     # is it loaded? how did the last run go?
./install_schedule.sh --run-now    # trigger a run right now, to prove it works
./install_schedule.sh --remove     # stop it (your data is untouched)
```

**To change the time**, edit `com.tomas.hotelrates.plist` — `Hour` on line
43 and `Minute` on line 45, both 24-hour — then run `./install_schedule.sh`
again.

Things worth knowing:

- **The Mac has to be awake.** If it was asleep at 06:30, the run happens on
  the next wake. If it was shut down, that day is skipped. This is why the
  scraper is scheduled with launchd rather than `cron` — `cron` would simply
  drop the run entirely.
- **If you move the folder, re-run `./install_schedule.sh`.** The scheduler
  needs absolute paths, so it has the old location written into it until you
  reinstall. This is the one part of the project that does not follow the
  folder on its own.
- **Check on it** with `--status`, or read `output/logs/<today>.log`.
- **macOS may ask for permission** the first time the job runs on its own —
  typically to let it access a folder. Approve it. If you never see a
  prompt and nothing runs, look in System Settings → Privacy & Security.

### Step 7 — Optional: VS Code

Only if you want a nicer place to edit the files than TextEdit.

1. Download from <https://code.visualstudio.com/>, drag it to Applications.
2. Open it, then **File → Open Folder** and choose this project folder.
3. When it offers to install the **Python** extension, accept.
4. It will find `.venv` automatically — `.vscode/settings.json` already
   tells it to.

The built-in terminal (**Terminal → New Terminal**) opens already in the
project folder, so every command in this guide works there unchanged.

---

## Part 3 — Common questions

**Do I have to type `./.venv/bin/python` every time?**
Yes, unless you activate the environment first:

```bash
source .venv/bin/activate      # now plain `python` means the project's one
python build_report.py
deactivate                     # when you're done
```

Both are equivalent. The long form is used throughout this guide because it
works whether or not you remembered to activate.

**Will the folder still work if I move it?**
Yes. Every script finds its own location at runtime — nothing has a path
written into it. The single exception is the **scheduled job**: re-run
`./install_schedule.sh` after moving, and it will point at the new place.

**Can I copy the whole folder to another Mac and just run it?**
Copy it, then run `./setup.sh` on the new Mac. You cannot reuse `.venv/`:
it hardcodes absolute paths to the folder it was built in and to the exact
Python it was built from, so a copied one looks fine but fails in confusing
ways — and an Intel-built one cannot work on Apple Silicon at all. That is
precisely what `requirements.txt` and `setup.sh` exist for, and why
`.venv/` is kept out of git.

**Does this work on Windows?**
The three Python files do — they use no macOS-specific code. The four
`.sh` scripts do not: they are bash. On Windows you would run the Python
commands directly (`.venv\Scripts\python.exe hotel_rates.py`), create the
environment with `py -m venv .venv` and `.venv\Scripts\pip install -r
requirements.txt`, and use **Task Scheduler** instead of
`install_schedule.sh`. It has not been tested there. On Linux everything
works except `install_schedule.sh` (use `cron`, calling `run_daily.sh`).

**What do I put on GitHub?**
Everything except what the ignore rules exclude — `hotels.json` included.
Its `apikey` values are public by design: each is published in that hotel's
own "Book Now" link, readable by any visitor, and that is where they were
captured from. They allow the same read-only availability search a browser
already makes, so a public repository is fine. Genuinely private files
(`.env`, `*.pem`, `secrets.json`, `credentials.json`) are ignored and should
stay that way.

The ignore rules are **not** in a `.gitignore`. They live in
`.git/info/exclude` at the repository root (`Projects/`), which git never
pushes — so they only exist on this Mac. On a fresh clone, copy those lines
into the new `.git/info/exclude` before running `git add`, or `.venv/` and
`output/` will be picked up.

`output/` is excluded, which means a fresh clone starts with no price
history. That is usually what you want — it is large, regenerated daily, and
nobody should be merging spreadsheets. To share one particular report
anyway, add it by name — the ignore rule stays as it is:

```bash
git add -f PythonHotel/output/report_20260913.xlsx
```

After that git tracks that one file, so rebuilding it later shows up as a
change to commit.

**How do I change how far ahead it looks?**
`DAYS_AHEAD` on line 221 of `hotel_rates.py` (what gets collected) and
`REPORT_DAYS` on line 200 of `build_report.py` (how wide the tables are).
Change them and everything follows — the report's own headings and
captions are generated from the constants, so nothing will still be
claiming the old number. `run_daily.sh` passes no day count of its own.

**How do I change a hotel's stay length?**
Its `"nights"` in `hotels.json`. Nowhere else.

**How does the report know how many rooms a hotel has?**
No booking engine says, so the report works it out from the latest run. For
each room type it takes the most units that room showed on any date of the
run, then adds those up per hotel. That total is what Availability % and
Occupancy % divide by.

- A room the latest run never listed counts **0**. So a room a hotel stops
  selling, or re-lists under a new code, stops counting on the very next run.
  It still has a row in the report, because `room_registry.json` never
  deletes anything.
- If a hotel had listings on fewer than `ROOM_COUNT_MIN_DAYS` dates (line 225
  of `build_report.py`, default 20), one run is too little to go on. That
  hotel uses the most units ever recorded across all runs instead. A short
  test run like `--days 3` will do this for every hotel.
- The same count is used for every run inside one report, so "vs 7 days ago"
  compares rooms sold, not a changed room count.

The **Number of rooms** sheet in the report shows the count for every room,
and which method each hotel used. If you know a hotel's real number, set
`HOTEL_CAPACITY` in `price_report.py` and it wins. Details: README.md →
"Availability counts rooms, not room types".

**Something broke — what do I try first?**

```bash
./setup.sh                                  # rebuild the environment
./.venv/bin/python hotel_rates.py --days 3  # a fast end-to-end test
./install_schedule.sh --status              # if it's the scheduled run
cat output/logs/$(date +%Y-%m-%d).log       # what the job actually did
```

Rebuilding the environment fixes a surprising share of problems and costs a
minute. It cannot lose data — `output/` is untouched.
