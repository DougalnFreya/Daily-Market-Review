"""List the engine's history and state (not code, not the day's outputs) in engine_files.json,
so each data file can be saved to the engine artifact as a plain, readable file of its own.

There is no archive to unpack: the daily run reads engine_files.json from the artifact, then
reads each file it lists straight into the data folder. Scripts are never listed: they come
from the repository on every run.

Writes engine_files.json: {"saved": "YYYY-MM-DD", "files": {"<published path>": "<path in DATA>"}}.
The two paths are the same; the map is the shape the artifact publish takes as "files".

Usage: python3 /path/to/repo/pack.py /path/to/data
"""
import datetime, json, os, sys

DATA = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
SKIP_NAMES = {"monitor_data.json", "monitor_latest.csv", "dashboard.html", "run_summary.txt",
              "monitor_checks.txt", "engine_bundle.txt", "engine_files.json", "runall.out",
              "digest.json", "readings.json", "research.json", "review.json", "dashboard_template4.html",
              "engine_index.html"}
# Only types the artifact will serve. ETF schedule.xlsx is not one of them: run_all.py rebuilds it
# each run from momentum/prices.csv and unemployment.csv, which are saved.
DATA_TYPES = (".csv", ".json")

files = {}
for root, dirs, names in os.walk(DATA):
    dirs[:] = [d for d in dirs if d not in ("__pycache__", ".git")]
    for fn in sorted(names):
        if fn in SKIP_NAMES or not fn.endswith(DATA_TYPES):
            continue
        rel = os.path.relpath(os.path.join(root, fn), DATA).replace(os.sep, "/")
        files[rel] = rel
files["engine_files.json"] = "engine_files.json"
manifest = {"saved": datetime.date.today().isoformat(), "files": dict(sorted(files.items()))}
json.dump(manifest, open(os.path.join(DATA, "engine_files.json"), "w"), indent=1)
print("listed", len(files), "files:", ", ".join(sorted(files)))
