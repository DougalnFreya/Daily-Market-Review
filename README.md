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

1. Unpack `engine_bundle.txt` from the engine artifact into a working folder (state).
2. Copy this repository's scripts over it, so the code always comes from here.
3. Run `run_all.py`, research and write the reading, rebuild with `build_monitor4.py`.
4. `pack.py`, then save the bundle back to the engine artifact and publish the page.

Change the scripts here, not in the bundle: the run overwrites the bundle's copies with these.

| File | Job |
|---|---|
| run_all.py | Whole daily job: pull, monthly schedule, build, checks; writes run_summary.txt |
| monitor_pull4.py | Collects the figures (needs FRED_API_KEY) |
| build_monitor4.py | Builds dashboard.html and runs the checks, including the research end point |
| dashboard_template4.html | Page template |
| momentum_monthly.py | Produces the monthly ETF schedule when one is due |
| momentum/momentum_schedule.py | Momentum rules |
| pack.py | Packs state and scripts into engine_bundle.txt |
