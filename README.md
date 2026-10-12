# Daily-Market-Review

Scripts for Phil's daily indicator monitor and monthly momentum ETF schedule.

Only the scripts live here. Everything that changes from day to day stays in the engine
artifact (https://claude.ai/artifact/LzQbtVpQmPcGyrJ62Wc5Jt, file `engine_bundle.txt`):
history CSVs, `monitor_state.json`, `context_notes.json`, `ai_demand.json` and `momentum/`
(prices, unemployment, `ETF schedule.xlsx`).

`ai_demand.json` holds AI demand figures published only as web pages (a16z consumer data,
Ramp token price and spend). The daily run adds each new release as a new observation and
never edits old ones. Ramp business adoption comes automatically from FRED instead.

## How the daily run uses this repository

Code and data are kept apart, so nothing unpacked from the artifact is ever run.

1. Clone this repository (attached to the session) and unpack `engine_bundle.txt` into a
   separate data folder. The bundle holds data only; any `.py` file found in it is deleted.
2. `python3 <repo>/run_all.py <data>`: pull, monthly schedule, build and checks. The scripts run
   from the repository and read and write the data folder.
3. Research and write the reading, then rebuild with `python3 <repo>/build_monitor4.py <data>`.
4. `python3 <repo>/pack.py <data>`, save the bundle back to the engine artifact, publish the page.

Change the scripts here, not in the bundle.

| File | Job |
|---|---|
| run_all.py | Whole daily job: pull, monthly schedule, build, checks; writes run_summary.txt |
| monitor_pull4.py | Collects the figures (needs FRED_API_KEY) |
| build_monitor4.py | Builds dashboard.html and runs the checks, including the research end point |
| dashboard_template4.html | Page template |
| momentum_monthly.py | Produces the monthly ETF schedule when one is due |
| momentum/momentum_schedule.py | Momentum rules |
| pack.py | Packs history and state (never code) into engine_bundle.txt |
