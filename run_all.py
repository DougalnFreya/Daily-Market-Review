"""The whole daily job in one go: collect the figures, produce the monthly ETF schedule if one
is due, build the dashboard page and run its checks. Writes run_summary.txt for the scheduled
task to read. Needs FRED_API_KEY in the environment.

Code and data are kept apart. The scripts run from this repository; the history and state
(the unpacked engine bundle) live in a separate data folder, which holds no code.

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
