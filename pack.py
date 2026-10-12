"""Pack the engine's history and state (not code, not the day's outputs) into engine_bundle.txt:
a gzipped tar, base64-encoded, so it can be stored as one text file in the engine artifact.
Scripts are never packed: they come from the repository on every run.
Unpack with:  base64 -d engine_bundle.txt | tar xz

Usage: python3 /path/to/repo/pack.py /path/to/data
"""
import base64, io, os, sys, tarfile

DATA = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
SKIP_NAMES = {"monitor_data.json", "monitor_latest.csv", "dashboard.html", "run_summary.txt",
              "monitor_checks.txt", "engine_bundle.txt", "runall.out",
              "digest.json", "readings.json", "research.json", "dashboard_template4.html"}
buf = io.BytesIO()
with tarfile.open(fileobj=buf, mode="w:gz") as tar:
    for root, dirs, files in os.walk(DATA):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", ".git")]
        for fn in sorted(files):
            if fn in SKIP_NAMES or fn.endswith((".log", ".before", ".pyc", ".py")):
                continue
            full = os.path.join(root, fn)
            tar.add(full, arcname=os.path.relpath(full, DATA))
open(os.path.join(DATA, "engine_bundle.txt"), "w").write(base64.b64encode(buf.getvalue()).decode())
print("packed", len(buf.getvalue()), "bytes")
