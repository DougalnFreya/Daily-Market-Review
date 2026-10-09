"""Pack the engine (scripts, history and state, not the day's outputs) into engine_bundle.txt:
a gzipped tar, base64-encoded, so it can be stored as one text file in the engine artifact.
Unpack with:  base64 -d engine_bundle.txt | tar xz
"""
import base64, io, os, tarfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKIP_NAMES = {"monitor_data.json", "monitor_latest.csv", "dashboard.html", "run_summary.txt",
              "monitor_checks.txt", "engine_bundle.txt", "runall.out",
              "digest.json", "readings.json", "research.json"}
buf = io.BytesIO()
with tarfile.open(fileobj=buf, mode="w:gz") as tar:
    for root, dirs, files in os.walk(HERE):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        for fn in sorted(files):
            if fn in SKIP_NAMES or fn.endswith((".log", ".before", ".pyc")):
                continue
            full = os.path.join(root, fn)
            tar.add(full, arcname=os.path.relpath(full, HERE))
open(os.path.join(HERE, "engine_bundle.txt"), "w").write(base64.b64encode(buf.getvalue()).decode())
print("packed", len(buf.getvalue()), "bytes")
