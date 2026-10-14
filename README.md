# Daily-Market-Review

Scripts for Phil's daily indicator monitor and monthly momentum ETF schedule.

Only the scripts live here. Everything that changes from day to day stays in the engine
artifact (https://claude.ai/artifact/LzQbtVpQmPcGyrJ62Wc5Jt) as plain, readable files, each
saved separately and listed in `engine_files.json`: history CSVs, `monitor_state.json`,
`context_notes.json`, `ai_demand.json` and `momentum/prices.csv` and `momentum/unemployment.csv`.
There is no archive to unpack. `ETF schedule.xlsx` is not stored (the artifact cannot serve
.xlsx files); `run_all.py` rebuilds it from the two momentum CSVs at the start of each run.

`ai_demand.json` holds AI demand figures published only as web pages (a16z consumer data,
Ramp token price and spend). The daily run adds each new release as a new observation and
never edits old ones. Ramp business adoption comes automatically from FRED instead.

## How the daily run uses this repository

Code and data are kept apart, so nothing unpacked from the artifact is ever run.

1. Clone this repository (attached to the session). Read `engine_files.json` from the engine
   artifact, then read each file it lists and copy them into a separate data folder.
2. `python3 <repo>/run_all.py <data>`: pull, monthly schedule, build and checks. The scripts run
   from the repository and read and write the data folder.
3. Research and write the reading, then rebuild with `python3 <repo>/build_monitor4.py <data>`.
4. `python3 <repo>/pack.py <data>` (writes `engine_files.json`), copy the listed files into the
   scratchpad, publish them to the engine artifact as separate files, publish the page.

Change the scripts here, not in the bundle.

| File | Job |
|---|---|
| run_all.py | Whole daily job: pull, monthly schedule, build, checks; writes run_summary.txt |
| monitor_pull4.py | Collects the figures (needs FRED_API_KEY) |
| build_monitor4.py | Builds dashboard.html and runs the checks, including the research end point |
| dashboard_template4.html | Page template |
| momentum_monthly.py | Produces the monthly ETF schedule when one is due |
| momentum/momentum_schedule.py | Momentum rules |
| pack.py | Lists the history and state files (never code) to save, in engine_files.json |
