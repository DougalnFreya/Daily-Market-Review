"""The whole daily job in one go: collect the figures, produce the monthly ETF schedule if one
is due, build the dashboard page and run its checks. Writes run_summary.txt for the scheduled
task to read. Needs FRED_API_KEY in the environment.

Code and data are kept apart. The scripts run from this repository; the history and state
(the plain data files read from the engine artifact) live in a separate data folder, which holds no code.

Usage: python3 /path/to/repo/run_all.py /path/to/data
"""
import os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else HERE
os.chdir(DATA)
env = dict(os.environ)
env["MOMENTUM_DIR"] = os.path.join(DATA, "momentum")
if os.path.exists("/root/.ccr/ca-bundle.crt"):          # the cloud's proxy certificate
    env.setdefault("SSL_CERT_FILE", "/root/.ccr/ca-bundle.crt")


def run(script, args, timeout):
    try:
        r = subprocess.run([sys.executable, os.path.join(HERE, script)] + args, capture_output=True,
                           text=True, env=env, cwd=DATA, timeout=timeout)
        return r.returncode, r.stdout + r.stderr
    except subprocess.TimeoutExpired:
        return 99, f"timed out after {timeout}s"


out = []

# The engine artifact cannot store the .xlsx workbook, so it is not saved between runs. It holds
# nothing that is not in prices.csv and unemployment.csv (its "Rule you followed" column has
# never been filled in the cloud copy), so rebuild it from those with the unchanged calculation.
# Nothing is fetched: this reproduces the last schedule exactly (checked 12 October 2026: every
# sheet identical). momentum_monthly.py then decides from it whether a new month is due.
mdir = env["MOMENTUM_DIR"]
wb = os.path.join(mdir, "ETF schedule.xlsx")
if not os.path.exists(wb) and os.path.exists(os.path.join(mdir, "prices.csv")):
    try:
        r = subprocess.run([sys.executable, os.path.join(HERE, "momentum", "momentum_schedule.py"),
                            "prices.csv", "unemployment.csv", "ETF schedule.xlsx"], cwd=mdir,
                           capture_output=True, text=True, env=env, timeout=300)
        ok = r.returncode == 0 and os.path.exists(wb)
        out.append("WORKBOOK rebuilt from stored prices" if ok else
                   "WORKBOOK rebuild failed: " + (r.stdout + r.stderr)[-500:])
    except subprocess.TimeoutExpired:
        out.append("WORKBOOK rebuild failed: timed out")

code, log = run("monitor_pull4.py", [DATA], 1500)
open("pull.log", "w").write(log)
tail = [l for l in log.splitlines() if "series," in l or l.strip().startswith("FAILED")]
out.append(f"PULL exit={code}: " + ("; ".join(tail[-8:]) or log[-300:]))

code, log = run("momentum_monthly.py", [], 600)
open("monthly.log", "w").write(log)
out.append(f"MONTHLY exit={code}:\n{log.strip()[:2500]}")

code, log = run("build_monitor4.py", [DATA], 600)
open("build.log", "w").write(log)
out.append(f"BUILD exit={code}")
checks = open("monitor_checks.txt").read() if os.path.exists("monitor_checks.txt") else log[-1500:]
out.append(checks)

open("run_summary.txt", "w").write("\n".join(out) + "\n")
print("\n".join(out))
