"""Leading indicator monitor - page builder (version 4).

Version 4 puts the measures from the original brief in a first panel: Baltic Clean and Dirty
Tanker Indices, the diesel crack against Brent, NBP front-month gas, UK wind curtailment and
negative-price hours, the Drewry Shanghai-Rotterdam rate, and UK and US 10-year yields.
Everything else follows as supporting detail.

Adds to version 2: the direction each indicator has moved (six-month trace, one- and
three-month change with the value and date each is measured from), six-month charts of the
funds you hold, the monthly momentum schedule read directly from ETF schedule.xlsx, and a set
of checks run on every build and shown at the top of the page.

Reads monitor_data.json (written by monitor_pull4.py) and ETF schedule.xlsx, and writes
dashboard.html from dashboard_template4.html. Standard library only.
Run it after the pull; Claude publishes the result.

Everything the pull covers is derived here. The handful of figures with no free
machine-readable source are in MANUAL below - that block is the only thing that
needs editing by hand, and only when a company reports.

Usage:  python build_monitor4.py [data_dir]
"""
import sys, os, json, re
from datetime import datetime, timezone

# --------------------------------------------------------------- hand-entered
# Update when the underlying source reports. Everything else comes from the data file.
MANUAL = {
    # value, and the date the source last reported it. The page shows the date and marks
    # anything older than 120 days as stale, so a figure nobody has refreshed cannot sit
    # there looking current.
    "oracle_backlog_bn":       (664,  "2026-09-09"),   # contracted revenue not yet billing
    "oracle_revenue_bn":       (71.8, "2026-09-09"),   # trailing twelve months
    "oracle_backlog_prior_bn": (455,  "2026-03-10"),   # earlier in 2026, for direction
    "oracle_one_customer_bn":  (300,  "2026-09-09"),   # backlog from a single counterparty
    "hyper_capex_bn":          (434,  "2026-05-01"),   # 4 hyperscalers, trailing 4 quarters
    "hyper_depreciation_bn":   (149,  "2026-05-01"),
    "amazon_capex_pct_ocf":    (102,  "2026-05-01"),
    "uk_auction_cover":        (3.24, "2026-09-10"),   # DMO - not verified from source
    "uk_auction_tail_bp":      (0.3,  "2026-09-10"),
    "ai_disclosed_arr_bn":     (2,    "2026-09-01"),   # largest disclosed franchises, combined
    "ai_token_price_fall_pct": (41,   "2026-09-01"),   # effective price per million tokens
}
STALE_AFTER_DAYS = 120


def M(key):
    """Hand-entered value."""
    return MANUAL[key][0]


def M_date(key):
    return MANUAL[key][1]


def stale_manual(today=None):
    """Which hand-entered figures have gone past their shelf life."""
    today = today or datetime.now(timezone.utc).date()
    out = []
    for k, (_, d) in MANUAL.items():
        age = (today - datetime.fromisoformat(d).date()).days
        if age > STALE_AFTER_DAYS:
            out.append((k, d, age))
    return out


THRESHOLD_US10Y = 5.25   # where shares and bonds have historically stopped cushioning each other


# --------------------------------------------------------------- helpers

def last(s):
    return s[-1][1] if s else None


def back(s, n):
    """Value n observations ago."""
    return s[-1 - n][1] if s and len(s) > n else None


def rng(s):
    vals = [v for _, v in s] if s else []
    return (min(vals), max(vals)) if vals else (0, 1)


def pct_off_high(daily, window=252):
    """How far below its own recent high a price sits, as a negative percentage."""
    if not daily:
        return None
    vals = [v for _, v in daily[-window:]]
    return round((vals[-1] / max(vals) - 1) * 100, 1)


def align(a, b):
    """Pair two series on matching dates.

    Never zip these by position: the monthly histories start on different dates (SPY from
    1993, RSP from 2003), so index i is a different month in each and every derived ratio
    comes out nonsense.
    """
    bd = dict(b)
    return [[d, v, bd[d]] for d, v in a if d in bd]


def ratio_series(a, b):
    return [[d, x / y] for d, x, y in align(a, b) if y]


def diff_series(a, b, scale_a=1.0):
    return [[d, x * scale_a - y] for d, x, y in align(a, b)]


def chg_pct(s, n):
    a, b = last(s), back(s, n)
    return round((a / b - 1) * 100, 1) if a and b else None


def row(name, sub, v, m12, lo, hi, unit, state, window, thr=None, thrlbl=None):
    d = {"n": name, "s": sub, "v": v, "m12": m12, "lo": lo, "hi": hi,
         "u": unit, "st": state, "w": window}
    if thr is not None:
        d["thr"], d["thrlbl"] = thr, thrlbl
    return d


def band(v, lo, hi, warn=0.75, bad=0.9, invert=False):
    """Position in range -> calm / watch / extreme. invert for measures where low is bad."""
    if hi == lo:
        return "calm"
    p = (v - lo) / (hi - lo)
    if invert:
        p = 1 - p
    return "extreme" if p >= bad else "watch" if p >= warn else "calm"


# --------------------------------------------------------------- workbook reader
# Reads the ETF schedule workbook with the standard library only, so the build needs
# nothing installed on the laptop.
import zipfile
import xml.etree.ElementTree as ET

_M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _col(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def read_xlsx(path):
    """{sheet name: [row, ...]} with each row a list of cell values (None for blanks)."""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    ss = []
    if "xl/sharedStrings.xml" in names:
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(f"{{{_M}}}si"):
            ss.append("".join(t.text or "" for t in si.iter(f"{{{_M}}}t")))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rmap = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for s in wb.find(f"{{{_M}}}sheets"):
        tgt = rmap[s.get(f"{{{_R}}}id")].lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        rows = []
        for row in ET.fromstring(z.read(tgt)).iter(f"{{{_M}}}row"):
            cells = {}
            for c in row.findall(f"{{{_M}}}c"):
                t, v = c.get("t"), c.find(f"{{{_M}}}v")
                if t == "s":
                    val = ss[int(v.text)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(f"{{{_M}}}t"))
                elif v is None:
                    val = None
                elif t == "b":
                    val = v.text == "1"
                elif t in ("str", "e"):
                    val = v.text
                else:
                    val = float(v.text)
                cells[_col(c.get("r"))] = val
            if any(x is not None for x in cells.values()):
                rows.append([cells.get(i) for i in range(max(cells) + 1)])
        out[s.get("name")] = rows
    return out


# --------------------------------------------------------------- direction of travel
# Each row carries a six-month trace and its change over one and three months, each with the
# value and date it is measured from, so every figure can be checked against the source.
# The arrow is simply the sign of the three-month change. It is coloured by the same
# convention as the state chip: red when the move is towards the stressed end of the range.

from datetime import date, timedelta

SPARK_DAYS = 183     # six months of trace


def _d(x):
    """'YYYY-MM-DD' or 'YYYY-MM' -> date."""
    return date.fromisoformat(x if len(x) == 10 else x + "-01")


def at_or_before(s, target):
    """Last value on or before a date. Series must be in date order."""
    v = None
    for d, x in s:
        if _d(d) <= target:
            v = x
        else:
            break
    return v


def trend(s, bad_up, mode, span=None, scale=1.0, obs=False):
    """Six-month trace, one- and three-month change, and direction.

    bad_up: True if a rise points towards stress, False if a fall does, None if neither.
    mode:   how the change is expressed - 'bp' basis points, 'pp' percentage points,
            'pct' percent change, 'abs' in the row's own unit.
    span:   not used; kept so existing calls read the same.
    obs:    measure 1 and 3 "months" as 1 and 3 observations back (for monthly events such
            as auctions, whose dates do not fall a fixed number of days apart).
    """
    s = sorted([[d, v * scale] for d, v in (s or []) if v is not None], key=lambda p: _d(p[0]))
    if len(s) < 2:
        return None
    end = _d(s[-1][0])
    win = [p for p in s if _d(p[0]) >= end - timedelta(days=SPARK_DAYS)] or s[-2:]
    if len(win) < 2:
        win = s[-2:]
    step = max(1, len(win) // 40)                    # thin daily data to about weekly
    sp = win[::-1][::step][::-1]
    v = s[-1][1]

    def pick(days):
        if obs:
            k = {30: 1, 91: 3}[days]
            return s[-1 - k] if len(s) > k else None
        tgt = end - timedelta(days=days)
        prev = None
        for d, x in s:
            if _d(d) <= tgt:
                prev = [d, x]
            else:
                break
        return prev

    def chg(days):
        p = pick(days)
        b = p[1] if p else None
        if b is None:
            return None
        if mode == "pct":
            return round((v / b - 1) * 100, 1) if b else None
        if mode == "bp":
            return round((v - b) * 100)
        return round(v - b, 2)

    def base(days):
        p = pick(days)
        return [p[0], round(p[1], 4)] if p else None

    b3p = pick(91)
    b3 = b3p[1] if b3p else None
    move = (v - b3) if b3 is not None else 0.0
    direction = "up" if move > 0 else "down" if move < 0 else "flat"
    toward_stress = None if (bad_up is None or move == 0) else ((move > 0) == bad_up)
    out_roc = rate_of_change(s, mode, bad_up)
    return {"roc": out_roc,
            "sp": [[d, round(x, 4)] for d, x in sp], "c1": chg(30), "c3": chg(91),
            "b1": base(30), "b3": base(91), "now": [s[-1][0], round(v, 4)],
            "m": mode, "dir": direction, "stress": toward_stress, "from": sp[0][0], "to": s[-1][0]}


def _cadence_pts(s):
    """Reduce a series to one point per natural period: per ISO week for daily or weekly
    data, or the observations themselves when they already arrive monthly or less often."""
    s = sorted(s)
    if len(s) < 3:
        return s, None
    gaps = sorted((_d(s[i][0]) - _d(s[i - 1][0])).days for i in range(1, len(s)))
    typical = gaps[len(gaps) // 2]
    if typical > 10:
        return s, "1m"                       # monthly or slower: a week means nothing
    buck = {}
    for dd, v in s:
        buck[_d(dd).isocalendar()[:2]] = [dd, v]
    return [buck[k] for k in sorted(buck)], "1w"


def rate_of_change(s, mode, bad_up=None):
    """How fast a measure is moving right now, judged against its own recent history.

    Three plain facts: the change over the latest period, whether a move that size is
    unusual for this measure, and whether the pace over the last month is rising or easing.
    Level tells you where you are; this tells you how quickly it is changing.
    """
    pts, lab = _cadence_pts(s)
    if not lab or len(pts) < 6:
        return None
    def delta(a, b):
        if mode == "pct":
            return (b / a - 1) * 100 if a else None
        if mode == "bp":
            return (b - a) * 100
        return b - a
    chg = []
    for i in range(1, len(pts)):
        c = delta(pts[i - 1][1], pts[i][1])
        if c is not None:
            chg.append([pts[i][0], c, pts[i - 1]])
    if len(chg) < 6:
        return None
    latest = chg[-1]
    window = [abs(c) for _, c, _ in chg[-52:]]
    rank = round(sum(1 for c in window if c <= abs(latest[1])) / len(window) * 100)
    # whether THIS period's move is towards the stressed end - not the three-month one, which
    # can point the other way and would otherwise group a measure with its opposite.
    stress = None if (bad_up is None or latest[1] == 0) else ((latest[1] > 0) == bad_up)
    out = {"lab": lab, "chg": round(latest[1], 2), "from": [latest[2][0], round(latest[2][1], 4)],
           "rank": rank, "n": len(window), "stress": stress}
    if len(chg) >= 12:
        recent = sum(abs(c) for _, c, _ in chg[-4:]) / 4.0
        earlier = sum(abs(c) for _, c, _ in chg[-12:-4]) / 8.0
        if earlier:
            out["ratio"] = round(recent / earlier, 2)
            out["pace"] = ("faster" if recent > earlier * 1.25 else
                           "slower" if recent < earlier * 0.8 else "steady")
    return out


def yield_dollar_count(y10, usd, window=20, start="2006-01-01"):
    """Days in the last `window` trading days when the US 10-year yield rose and the broad
    dollar fell on the same day, for every date since `start`.

    Both series are matched on the dates they share, and each day's change is measured from
    the previous shared date. Rising yields with a falling dollar is what investors selling
    US bonds and the currency together looks like; rising yields with a rising dollar is the
    US paying more and still attracting money.
    """
    y = {d: v for d, v in (y10 or []) if v is not None}
    pts = [(d, v, y[d]) for d, v in sorted(usd or []) if v is not None and d in y and d >= start]
    flags = [(pts[i][0], 1 if (pts[i][2] > pts[i - 1][2] and pts[i][1] < pts[i - 1][1]) else 0)
             for i in range(1, len(pts))]
    return [[flags[i][0], sum(f for _, f in flags[i - window + 1:i + 1])]
            for i in range(window - 1, len(flags))]


def with_trend(r, tr):
    if tr:
        r["tr"] = tr
    return r


# --------------------------------------------------------------- derive rows

def span_of(lo, hi):
    return (hi - lo) or None


def build_rows(S):
    g = S.get
    R = {}
    T = with_trend

    # ---- auction demand -------------------------------------------------
    auctions = g("us_auctions", [])
    def auc(term):
        a = [x for x in auctions if x["term"] == term]
        if not a:
            return None, None, None, None
        prior = a[1:13]
        avg = lambda k: round(sum(x[k] for x in prior) / len(prior), 2)  # noqa: E731
        covers = [x["cover"] for x in a]
        dealers = [x["dealer_pct"] for x in a]
        return a[0], {"cover": avg("cover"), "dealer_pct": avg("dealer_pct")}, \
               {"cover": (min(covers), max(covers)), "dealer": (0, max(dealers))}, a

    R["p-auction"] = []
    for term, label in (("10-Year", "10-year"), ("30-Year", "30-year")):
        l, av, rr, a = auc(term)
        if not l:
            continue
        cov = [[x["date"], x["cover"]] for x in a]
        dea = [[x["date"], x["dealer_pct"]] for x in a]
        R["p-auction"].append(T(row(
            f"US {label} auction cover", f"{l['date']} &middot; bids per dollar sold &middot; "
            f"last twelve average {av['cover']} &middot; overseas-type buyers took "
            f"{l.get('indirect_pct', '?')}%", l["cover"],
            at_or_before(sorted(cov), _d(l["date"]) - timedelta(days=365)),
            rr["cover"][0], rr["cover"][1], "x",
            band(l["cover"], *rr["cover"], invert=True), "since 2018"),
            trend(cov, False, "abs", obs=True)))
        R["p-auction"].append(T(row(
            f"US {label} dealer share", f"{l['date']} &middot; the obliged buyer of last resort "
            f"&middot; lower is stronger demand &middot; average {av['dealer_pct']}%",
            l["dealer_pct"], at_or_before(sorted(dea), _d(l["date"]) - timedelta(days=365)),
            0, rr["dealer"][1], "%",
            band(l["dealer_pct"], 0, rr["dealer"][1]), "since 2018"),
            trend(dea, True, "pp", obs=True)))
    # Every 2- to 30-year auction in the 30 days to each auction date, weighted by the amount
    # sold. One reading per measure instead of one per maturity, so the weak 2-, 5- and
    # 7-year sales of late September show up without adding a dozen rows.
    def rolling(key, days=30):
        a = sorted((x for x in auctions if x.get("size_bn")), key=lambda x: x["date"])
        out, counts = [], {}
        for d in sorted({x["date"] for x in a}):
            end = _d(d)
            w = [x for x in a if end - timedelta(days=days) < _d(x["date"]) <= end]
            tot = sum(x["size_bn"] for x in w)
            if tot:
                out.append([d, round(sum(x[key] * x["size_bn"] for x in w) / tot, 2)])
                counts[d] = len(w)
        return out, counts

    for key, nm, invert, what in (
            ("indirect_pct", "US auctions, last month: overseas-type share", True,
             "the &lsquo;indirect&rsquo; bidders: foreign central banks and other investors "
             "bidding through a dealer, which includes some US funds &middot; lower is weaker"),
            ("dealer_pct", "US auctions, last month: dealer share", False,
             "what the obliged buyers of last resort were left holding &middot; higher is weaker")):
        ser, counts = rolling(key)
        if not ser:
            continue
        v, dl = ser[-1][1], ser[-1][0]
        yr = [x for d, x in ser[:-1] if _d(d) >= _d(dl) - timedelta(days=365)]
        avg = round(sum(yr) / len(yr), 1) if yr else None
        vals = [x for _, x in ser]
        lo, hi = min(vals), max(vals)
        inwin = [x for x in auctions if x.get("size_bn")
                 and _d(dl) - timedelta(days=30) < _d(x["date"]) <= _d(dl)]
        worst = (min if invert else max)(inwin, key=lambda x: x[key])
        R["p-auction"].append(T(row(
            nm, f"every 2- to 30-year note and bond sold in the 30 days to {dl} "
            f"({counts[dl]} auctions), weighted by size &middot; {what} &middot; "
            f"weakest: {worst['term'].replace('-Y', '-y')} on {worst['date']} at {worst[key]}% "
            f"&middot; average over the past year {avg}%",
            v, at_or_before(ser, _d(dl) - timedelta(days=365)), lo, hi, "%",
            band(v, lo, hi, invert=invert), "since 2018"),
            trend(ser, not invert, "pp")))

    # Yields up, dollar down. Computed from the raw series when the pull has them; a series
    # computed the same way elsewhere is accepted only if the raw dollar series is absent.
    ydd = yield_dollar_count(g("us_treasury_10y", []), g("us_dollar_broad", [])) \
        or g("us_yield_up_dollar_down_precomputed", [])
    if ydd:
        v, dl = ydd[-1][1], ydd[-1][0]
        vals = [x for _, x in ydd]
        lo, hi = min(vals), max(vals)
        avg = round(sum(vals) / len(vals), 1)
        rare = round(100 * sum(1 for x in vals if x >= v) / len(vals))
        R["p-auction"].append(T(row(
            "US yields up, dollar down",
            f"days in the 20 trading days to {dl} when the US 10-year yield rose and the broad "
            f"dollar fell &middot; investors selling US bonds and the currency together &middot; "
            f"average since 2006 {avg}; a reading this high or higher on {rare}% of days "
            f"&middot; the dollar index is published weekly, so this runs up to a week behind",
            v, at_or_before(ydd, _d(dl) - timedelta(days=365)), lo, hi, "of20",
            band(v, lo, hi), "since 2006"),
            trend(ydd, True, "abs")))

    R["p-auction"].append(row(
        "UK gilt auction cover", f"DMO &middot; reported at a {M('uk_auction_tail_bp')}bp tail "
        "&middot; not verified from source", M("uk_auction_cover"), 3.0, 1.5, 4.0, "x",
        "calm", "illustrative"))

    # ---- UK rates -------------------------------------------------------
    uk = {}
    for k in ("uk_gilt_5y", "uk_gilt_10y", "uk_gilt_20y", "uk_bank_rate", "uk_sonia",
              "uk_real_10y", "uk_infl_implied_5y", "uk_infl_implied_10y"):
        uk[k] = g(k, [])
    spread_s0 = diff_series(uk["uk_gilt_10y"], uk["uk_bank_rate"])
    spread = (last(spread_s0),
              back(uk["uk_gilt_10y"], 252) - back(uk["uk_bank_rate"], 252))
    spread_s = diff_series(uk["uk_gilt_10y"], uk["uk_bank_rate"])

    def ukrow(name, sub, k, bad_up=True, state=None):
        s = uk[k]
        lo, hi = rng(s)
        return T(row(name, sub, last(s), back(s, 252), lo, hi, "%",
                     state or band(last(s), lo, hi), "since 2005"),
                 trend(s, bad_up, "bp", span_of(lo, hi)))

    R["p-uk"] = [
        ukrow("20-year gilt yield", "Nominal par yield", "uk_gilt_20y"),
        ukrow("10-year gilt yield", "Nominal par yield", "uk_gilt_10y"),
        ukrow("5-year gilt yield", "Nominal par yield", "uk_gilt_5y"),
        T(row("10-year gilt over Bank Rate",
              f"The market's premium over policy &middot; {spread[1]:.2f}pp a year ago",
              round(spread[0], 3), round(spread[1], 3), -1.5, 3.0, "pp",
              band(spread[0], -1.5, 3.0), "illustrative"),
          trend(spread_s, True, "bp", 4.5)),
        ukrow("UK 10-year expected inflation", "Implied by index-linked gilts, RPI basis",
              "uk_infl_implied_10y"),
        ukrow("UK 5-year expected inflation",
              "Near-term &middot; the front of the curve is where an energy shock shows",
              "uk_infl_implied_5y"),
        ukrow("UK 10-year real yield", "Index-linked gilts", "uk_real_10y"),
        ukrow("Bank Rate", "Bank of England policy rate", "uk_bank_rate", None, "calm"),
        ukrow("SONIA", "Overnight sterling rate", "uk_sonia", None, "calm"),
    ]

    # ---- labour ---------------------------------------------------------
    ukun, ukvac, ukemp = g("uk_unemployment", []), g("uk_vacancies", []), g("uk_employment", [])
    usun, claims, sahm = g("us_unemployment", []), g("us_claims", []), g("us_sahm", [])
    r10 = lambda s: rng(s[-120:])  # noqa: E731  - last ten years for labour series
    c_lo = min(v for _, v in claims[-520:]) / 1000
    c_hi = max(v for _, v in claims[-520:]) / 1000
    sahm_lo = min(v for _, v in sahm[-120:])

    R["p-jobs"] = [
        T(row("UK unemployment rate", f"ONS, aged 16+ &middot; latest {ukun[-1][0]}",
              last(ukun), back(ukun, 12), *r10(ukun), "%",
              band(last(ukun), *r10(ukun)), "10 years"),
          trend(ukun, True, "pp", span_of(*r10(ukun)))),
        T(row("UK vacancies", f"ONS &middot; latest {ukvac[-1][0]}", last(ukvac), back(ukvac, 12),
              *r10(ukvac), "k", band(last(ukvac), *r10(ukvac), invert=True), "10 years"),
          trend(ukvac, False, "pct", span_of(*r10(ukvac)))),
        T(row("UK employment rate", f"ONS, aged 16 to 64 &middot; latest {ukemp[-1][0]}",
              last(ukemp), back(ukemp, 12), *r10(ukemp), "%",
              band(last(ukemp), *r10(ukemp), invert=True), "10 years"),
          trend(ukemp, False, "pp", span_of(*r10(ukemp)))),
        T(row("US unemployment rate", f"Latest {usun[-1][0]}", last(usun), back(usun, 12),
              *r10(usun), "%", band(last(usun), *r10(usun)), "10 years"),
          trend(usun, True, "pp", span_of(*r10(usun)))),
        T(row("US initial jobless claims", f"Week to {claims[-1][0]}",
              round(last(claims) / 1000), round(back(claims, 52) / 1000),
              round(c_lo), round(c_hi), "k", band(last(claims) / 1000, c_lo, c_hi), "10 years"),
          trend(claims, True, "pct", c_hi - c_lo, scale=0.001)),
        T(row("US Sahm recession gauge", "Triggers at +0.50", last(sahm), back(sahm, 12),
              sahm_lo, 1.2, "pp", band(last(sahm), sahm_lo, 1.2), "10 years"),
          trend(sahm, True, "pp", 1.2 - sahm_lo)),
    ]

    # ---- UK power -------------------------------------------------------
    power = g("uk_power", {})
    months = sorted(power)
    if months:
        cur, prior = power[months[-1]], power[months[-13]] if len(months) > 12 else power[months[0]]
        means = [power[m]["mean_price"] for m in months if power[m]["mean_price"]]
        negs = [power[m]["negative_hours"] for m in months]
        R["p-power"] = [
            T(row("UK day-ahead power", f"Monthly mean &middot; {months[-1]} (month to date)",
                  cur["mean_price"], prior["mean_price"], min(means), max(means), "£/MWh",
                  band(cur["mean_price"], min(means), max(means)), f"{len(months)} months"),
              trend([[m, power[m]["mean_price"]] for m in months], True, "pct",
                    max(means) - min(means))),
            T(row("Negative price hours", f"Hours at a negative day-ahead price &middot; {months[-1]} (month to date)",
                  cur["negative_hours"], prior["negative_hours"], min(negs), max(negs), "h",
                  band(cur["negative_hours"], min(negs), max(negs)), f"{len(months)} months"),
              trend([[m, power[m]["negative_hours"]] for m in months], True, "abs",
                    (max(negs) - min(negs)) or None)),
        ]

    # ---- oil ------------------------------------------------------------
    ho, wti, brent = g("heating_oil", []), g("wti", []), g("brent", [])
    hom, wtim, brentm = g("heating_oil_monthly", []), g("wti_monthly", []), g("brent_monthly", [])
    crack_d = diff_series(ho, wti, 42)
    crack = last(crack_d)
    crack12 = back(ho, 252) * 42 - back(wti, 252)
    crack_hist = [v for _, v in diff_series(hom, wtim, 42)]
    bw_d = diff_series(brent, wti)
    bw = last(bw_d)
    bw_hist = [v for _, v in diff_series(brentm, wtim)]
    fro, dht = g("frontline", []), g("dht", [])

    R["p-energy"] = [
        T(row("Diesel crack against US crude", "Heating oil converted to barrels, less WTI",
              round(crack, 2), round(crack12, 2), round(min(crack_hist), 2),
              round(max(crack_hist), 2), "$/bbl",
              band(crack, min(crack_hist), max(crack_hist)), "since 2005"),
          trend(crack_d, True, "abs", max(crack_hist) - min(crack_hist))),
        T(row("Crude tanker owners (Frontline)", "Frontline, 12-month share price change &middot; trace is the share price",
              chg_pct(fro, 252), 0, -90, 150, "pct",
              band(chg_pct(fro, 252), -90, 150), "since 2005"),
          trend(fro, True, "pct")),
        T(row("Crude tanker owners (DHT)", "DHT Holdings, 12-month share price change &middot; trace is the share price",
              chg_pct(dht, 252), 0, -90, 120, "pct",
              band(chg_pct(dht, 252), -90, 120), "since 2005"),
          trend(dht, True, "pct")),
        T(row("Brent crude", f"Front month &middot; {brent[-1][0]}", last(brent),
              back(brent, 252), *rng(brentm), "$/bbl", band(last(brent), *rng(brentm)), "since 2007"),
          trend(brent, True, "pct", span_of(*rng(brentm)))),
        T(row("Brent less WTI", "Waterborne against landlocked crude", round(bw, 2),
              round(back(brent, 252) - back(wti, 252), 2), round(min(bw_hist), 2),
              round(max(bw_hist), 2), "$/bbl",
              band(bw, min(bw_hist), max(bw_hist)), "since 2007"),
          trend(bw_d, True, "abs", max(bw_hist) - min(bw_hist))),
    ]

    # ---- AI, the bet ----------------------------------------------------
    orcl, crwv, nvda = g("oracle", []), g("coreweave", []), g("nvidia", [])
    smh, spy, rsp = g("smh", []), g("spy", []), g("rsp", [])
    smhm, spym, rspm = g("smh_monthly", []), g("spy_monthly", []), g("rsp_monthly", [])
    rsp_spy, smh_spy = ratio_series(rspm, spym), ratio_series(smhm, spym)
    smh_spy_d, rsp_spy_d = ratio_series(smh, spy), ratio_series(rsp, spy)
    bl = M("oracle_backlog_bn") / M("oracle_revenue_bn")
    ss_lo, ss_hi = min(v for _, v in smh_spy), max(v for _, v in smh_spy)
    rs_lo, rs_hi = min(v for _, v in rsp_spy), max(v for _, v in rsp_spy)

    R["p-bet"] = [
        row("Oracle backlog to revenue",
            "Contracted but not yet billing &middot; a normal software company runs at one to two times",
            round(bl, 1), round(M("oracle_backlog_prior_bn") / M("oracle_revenue_bn"), 1),
            0, 10, "x", "extreme", "illustrative"),
        T(row("Oracle", "Share price against its own 12-month high &middot; spent the capital, not yet paid",
              pct_off_high(orcl), 0, -60, 0, "pctHigh", "extreme", "1 year"),
          trend(orcl, False, "pct")),
        T(row("CoreWeave", "Share price against its own 12-month high &middot; pure-play compute landlord",
              pct_off_high(crwv), 0, -60, 0, "pctHigh", "extreme", "1 year"),
          trend(crwv, False, "pct")),
        T(row("Nvidia", "Share price against its own 12-month high &middot; its sales <em>are</em> the "
                        "buildout spending, so this belongs here, not on the return side",
              pct_off_high(nvda), 0, -60, 0, "pctHigh", "calm", "1 year"),
          trend(nvda, False, "pct")),
        T(row("Semiconductors vs market", "Chip shares relative to the S&amp;P 500",
              round(last(smh_spy_d), 4), round(smh_spy[-13][1], 4),
              round(ss_lo, 4), round(ss_hi, 4), "ratio",
              band(last(smh_spy_d), ss_lo, ss_hi), "since 2003"),
          trend(smh_spy_d, True, "pct", ss_hi - ss_lo)),
        T(row("Market breadth", "Equal-weighted against cap-weighted S&amp;P 500",
              round(last(rsp_spy_d), 4), round(rsp_spy[-13][1], 4),
              round(rs_lo, 4), round(rs_hi, 4), "ratio",
              band(last(rsp_spy_d), rs_lo, rs_hi, invert=True), "since 2003"),
          trend(rsp_spy_d, False, "pct", rs_hi - rs_lo)),
    ]

    # ---- financial stress -----------------------------------------------
    hyg, lqd = g("hyg", []), g("lqd", [])
    hygm, lqdm = g("hyg_monthly", []), g("lqd_monthly", [])
    hl = ratio_series(hygm, lqdm)
    hl_d = ratio_series(hyg, lqd)
    hl_lo, hl_hi = min(v for _, v in hl), max(v for _, v in hl)
    vix, vixm = g("vix", []), g("vix_monthly", [])
    us10, us30, us3m = g("us_treasury_10y", []), g("us_treasury_30y", []), g("us_treasury_3m", [])
    real, be, tp = g("us_real_10y", []), g("us_breakeven_10y", []), g("us_term_premium", [])
    gold, goldm = g("gold", []), g("gold_monthly", [])
    curve_d = diff_series(us10, us3m)
    curve = last(curve_d)

    R["p-fin"] = [
        T(row("High yield vs investment grade", "Credit stress proxy",
              round(last(hl_d), 4), round(hl[-13][1], 4),
              round(hl_lo, 4), round(hl_hi, 4), "ratio",
              band(last(hl_d), hl_lo, hl_hi, invert=True), "since 2007"),
          trend(hl_d, False, "pct", hl_hi - hl_lo)),
        T(row("Equity volatility", "VIX, month end", last(vix), back(vix, 252), *rng(vixm), "",
              band(last(vix), *rng(vixm)), "since 1990"),
          trend(vix, True, "abs", span_of(*rng(vixm)))),
        T(row("US 10-year yield",
              f"{abs(round((THRESHOLD_US10Y - last(us10)) * 100))}bp "
              f"{'below' if last(us10) < THRESHOLD_US10Y else 'above'} the level where bonds "
              "stop hedging shares", last(us10), back(us10, 252), 0.5, 5.5, "%",
              band(last(us10), 0.5, 5.5), "since 1962",
              thr=THRESHOLD_US10Y, thrlbl=f"{THRESHOLD_US10Y}% correlation threshold"),
          trend(us10, True, "bp", 5.0)),
        T(row("US 10-year real yield", "Index-linked &middot; the discount rate on everything",
              last(real), back(real, 252), *rng(real), "%",
              band(last(real), *rng(real)), "since 2003"),
          trend(real, True, "bp", span_of(*rng(real)))),
        T(row("US 10-year expected inflation", "Breakeven", last(be), back(be, 252), *rng(be), "%",
              band(last(be), *rng(be)), "since 2003"),
          trend(be, True, "bp", span_of(*rng(be)))),
        T(row("US term premium", "Extra return demanded for lending long to the government",
              last(tp), back(tp, 252), *rng(tp), "%", band(last(tp), *rng(tp)), "since 1990"),
          trend(tp, True, "bp", span_of(*rng(tp)))),
        T(row("US 30-year yield", "Long-end government borrowing cost", last(us30),
              back(us30, 252), 1.2, 5.5, "%", band(last(us30), 1.2, 5.5), "since 1977"),
          trend(us30, True, "bp", 4.3)),
        T(row("US yield curve", "10-year less 3-month", round(curve, 3),
              round(back(us10, 252) - back(us3m, 252), 3), -1.6, 3.8, "pp",
              band(curve, -1.6, 3.8, invert=True), "since 1962"),
          trend(curve_d, False, "bp", 5.4)),
        T(row("Gold", "Spot, month end", last(gold), back(gold, 252), *rng(goldm), "$/oz",
              band(last(gold), *rng(goldm)), "since 2000"),
          trend(gold, True, "pct", span_of(*rng(goldm)))),
    ]
    return R




# --------------------------------------------------------------- before the Gulf closure
# Hormuz closed in early March 2026. For measures of the Gulf shock the page shows the
# January-February 2026 average beside today's level, so a figure can be judged now.
PRE_FROM, PRE_TO = "2026-01-01", "2026-02-28"


def pre_crisis(series):
    vals = [v for d, v in (series or []) if PRE_FROM <= d[:10] <= PRE_TO and v is not None]
    if not vals:
        return None
    avg = sum(vals) / len(vals)
    return {"avg": round(avg, 4), "n": len(vals)}


def with_pre(r, series):
    p = pre_crisis(series)
    if p and r.get("v") is not None:
        p["chg_pct"] = round((r["v"] / p["avg"] - 1) * 100, 1) if p["avg"] else None
        p["chg_abs"] = round(r["v"] - p["avg"], 4)
        r["pre"] = p
    return r


def year_ago(series, d=None):
    s = sorted(series or [])
    if not s:
        return None
    end = _d(s[-1][0]) if d is None else d
    return at_or_before(s, end - timedelta(days=365))


def hist_row(name, sub, series, unit, bad_up, note):
    """A measure with a rebuilt history (weekly reports, then the daily ticker)."""
    s = sorted(series or [])
    if len(s) < 2:
        return na_row(name, sub, "no readings")
    vals = [v for _, v in s]
    lo, hi = min(vals), max(vals)
    ya = year_ago(s)
    r = row(name, sub, s[-1][1], ya, lo, hi, unit, band(s[-1][1], lo, hi), f"since {s[0][0]}")
    r["prov"] = note
    r["short"] = {"first": s[0], "last": s[-1], "n": len(s), "chg": None, "note": note}
    tr = trend(s, bad_up, "pct")
    if tr:
        r["tr"] = tr
    return with_pre(r, s)


CLEANED = {}


def drop_glitches(series, name):
    """Remove one-day readings that break from both neighbours while the neighbours agree.

    A stock level cannot jump and reverse in a day, so such a reading is a data error in the
    source. Zeros in a stock series are treated the same way. Removals are counted and shown
    in the checks.
    """
    s = sorted(series or [])
    keep, dropped = [], []
    for i, (d, v) in enumerate(s):
        if 0 < i < len(s) - 1:
            a, b = s[i - 1][1], s[i + 1][1]
            if a and b and abs(a / b - 1) < 0.10 and (v == 0 or abs(v / ((a + b) / 2) - 1) > 0.25):
                dropped.append(d)
                continue
        if v == 0:
            dropped.append(d)
            continue
        keep.append([d, v])
    CLEANED[name] = dropped
    return keep


def rolling(series, n=7):
    s = sorted(series or [])
    out = []
    for i in range(n - 1, len(s)):
        w = [v for _, v in s[i - n + 1:i + 1]]
        out.append([s[i][0], round(sum(w) / n, 2)])
    return out


def add_series(*ss):
    maps = [dict(x) for x in ss]
    common = sorted(set.intersection(*[set(m) for m in maps]))
    return [[d, round(sum(m[d] for m in maps), 2)] for d in common]


# --------------------------------------------------------------- what is moving yields

def build_split(S):
    """Change in each 10-year (and UK 5-year) yield, split into real yield and expected
    inflation. UK uses the Bank of England zero-coupon curves throughout (nominal, real,
    implied inflation on the RPI basis), so the parts add up to the total bar a small
    compounding difference, shown as 'other'."""
    g = S.get
    specs = [("UK 10-year", g("uk_nom_zc_10y"), g("uk_real_10y"), g("uk_infl_implied_10y")),
             ("UK 5-year", g("uk_nom_zc_5y"), g("uk_real_5y"), g("uk_infl_implied_5y")),
             ("US 10-year", g("us_treasury_10y"), g("us_real_10y"), g("us_breakeven_10y"))]
    rows = []
    for name, nom, real, inf in specs:
        if not (nom and real and inf):
            continue
        nd, rd, idd = dict(nom), dict(real), dict(inf)
        common = sorted(set(nd) & set(rd) & set(idd))
        if not common:
            continue
        end = common[-1]
        out = {"name": name, "date": end, "now": {"nom": nd[end], "real": rd[end], "inf": idd[end]}}
        for label, days in (("3m", 91), ("12m", 365)):
            tgt = (_d(end) - timedelta(days=days)).isoformat()
            base = [c for c in common if c <= tgt]
            if not base:
                continue
            b = base[-1]
            n_ = round((nd[end] - nd[b]) * 100)
            r_ = round((rd[end] - rd[b]) * 100)
            i_ = round((idd[end] - idd[b]) * 100)
            out[label] = {"from": b, "nom": n_, "real": r_, "inf": i_, "other": n_ - r_ - i_}
        rows.append(out)

    def sentence(r):
        x = r.get("12m")
        if not x:
            return ""
        parts = sorted((("the real yield", x["real"]), ("expected inflation", x["inf"])),
                       key=lambda p: -abs(p[1]))
        verb = "rose" if x["nom"] > 0 else "fell" if x["nom"] < 0 else "was unchanged"
        return (f"{r['name']} yield {verb} {abs(x['nom'])}bp over 12 months: {parts[0][0]} "
                f"{'+' if parts[0][1] >= 0 else '−'}{abs(parts[0][1])}bp, {parts[1][0]} "
                f"{'+' if parts[1][1] >= 0 else '−'}{abs(parts[1][1])}bp.")
    return {"rows": rows, "sentences": [sentence(r) for r in rows if sentence(r)]}


# --------------------------------------------------------------- warning signs, measured now

WARN_RULES = [
    # key, label, how to compute, threshold, direction, how often it warned before past peaks
    ("breadth", "Market breadth (equal vs cap weight)", -2.7, "down", "All three"),
    ("vix", "Equity volatility (VIX)", 8.8, "up", "All three"),
    ("hyig", "High yield vs investment grade", -4.5, "down", "One of three"),
]


def build_warn(S):
    """The three measures tested before the 2007, 2020 and 2022 peaks, measured now the
    same way: change over the last six month-ends (the latest month-end bar is the month
    to date). The warning level is the smallest change seen before any of those peaks."""
    g = S.get
    ser = {"breadth": ratio_series(g("rsp_monthly", []), g("spy_monthly", [])),
           "vix": g("vix_monthly", []),
           "hyig": ratio_series(g("hyg_monthly", []), g("lqd_monthly", []))}
    out = []
    for key, label, thr, direction, record in WARN_RULES:
        s = sorted(ser[key] or [])
        if len(s) < 7:
            out.append({"key": key, "n": label, "na": True})
            continue
        now, then = s[-1], s[-7]
        chg = round((now[1] / then[1] - 1) * 100, 1)
        on = chg <= thr if direction == "down" else chg >= thr
        out.append({"key": key, "n": label, "chg": chg, "from": then[0], "to": now[0],
                    "thr": thr, "dir": direction, "on": on, "record": record})
    return out


# --------------------------------------------------------------- the brief

def na_row(name, sub, reason):
    return {"n": name, "s": sub, "na": reason, "st": "none"}


def short_row(name, sub, series, unit, bad_up, source_note):
    """A measure whose history only builds from the first run (tanker indices, Drewry)."""
    s = sorted(series or [])
    if not s:
        return na_row(name, sub, "no readings yet")
    vals = [v for _, v in s]
    lo, hi = min(vals), max(vals)
    # no padding: with a short history the page shows no range bar at all
    first, lastp = s[0], s[-1]
    r = row(name, sub, lastp[1], None, lo, hi, unit, "none", f"since {first[0]}")
    r["short"] = {"first": first, "last": lastp, "n": len(s),
                  "chg": round((lastp[1] / first[1] - 1) * 100, 1) if first[1] else None,
                  "note": source_note}
    tr = trend(s, bad_up, "pct")
    if tr:
        r["tr"] = tr
    return r


def build_brief(S, R):
    """The measures named in the original brief, in the order given."""
    g = S.get
    T = with_trend
    out = []

    def take(panel, name):
        for i, r in enumerate(R.get(panel, [])):
            if r["n"] == name:
                return R[panel].pop(i)
        return None

    bal = g("baltic") or {}
    prov = ("Weekly figures to 25 Sep 2026 rebuilt from published shipbroker reports (each "
            "source is listed in shipping_history_sources.csv); daily from the Baltic Exchange "
            "public ticker after that.")
    out.append(hist_row("Baltic Clean Tanker Index",
                        "Cost of shipping refined fuel", bal.get("BCTI"), "idx", True, prov))
    out.append(hist_row("Baltic Dirty Tanker Index",
                        "Cost of shipping crude, VLCCs included. The VLCC route rate itself is "
                        "licensed only; this is the closest measure with a published history",
                        bal.get("BDTI"), "idx", True, prov))
    ho, brent = g("heating_oil", []), g("brent", [])
    hom, brentm = g("heating_oil_monthly", []), g("brent_monthly", [])
    crack = diff_series(ho, brent, 42)
    hist = [v for _, v in diff_series(hom, brentm, 42)]
    if crack and hist:
        out.append(T(row("Diesel crack against Brent",
                         "NYMEX ultra-low sulphur diesel converted to barrels, less Brent. "
                         "ICE gasoil is licensed only",
                         round(last(crack), 2), round(back(crack, 252), 2) if len(crack) > 252 else None,
                         round(min(hist), 2), round(max(hist), 2), "$/bbl",
                         band(last(crack), min(hist), max(hist)), "since 2007"),
                     trend(crack, True, "abs")))
        with_pre(out[-1], crack)
    else:
        out.append(na_row("Diesel crack against Brent", "", "diesel or Brent prices missing"))

    nbp, nbpm = g("nbp", []), g("nbp_monthly", [])
    if nbp and nbpm:
        lo, hi = rng(nbpm)
        lo, hi = min(lo, min(v for _, v in nbp)), max(hi, max(v for _, v in nbp))
        out.append(T(row("NBP gas, front month",
                         "NYMEX UK NBP contract via Yahoo, which does not state the unit; the "
                         "prices match Dutch TTF converted to dollars per MMBtu (see checks)",
                         last(nbp), back(nbp, 252), lo, hi, "$/MMBtu",
                         band(last(nbp), lo, hi), "since 2017"),
                     trend(nbp, True, "pct")))
        with_pre(out[-1], nbp)

    else:
        out.append(na_row("NBP gas, front month", "", "no NBP prices pulled"))
    gas = g("uk_gas") or {}
    if gas.get("storage"):
        st = drop_glitches(gas["storage"], "UK gas in storage")
        vals = [v for _, v in st]
        out.append((T(row("UK gas in storage",
                                  "Stock in all UK storage sites, National Gas. Storage falls every "
                                  "winter, so compare with the same day a year ago rather than the range",
                                  st[-1][1], year_ago(st), min(vals), max(vals), "GWh", "none",
                                  f"since {st[0][0][:4]}"),
                              trend(st, False, "pct"))))  # seasonal: no Jan-Feb comparison
    else:
        out.append(na_row("UK gas in storage", "", "National Gas data missing"))
    if gas.get("lng") and gas.get("beach") and gas.get("interconnector"):
        tot = rolling(add_series(gas["lng"], gas["beach"], gas["interconnector"]))
        lng = rolling(gas["lng"])
        for nm, sub, ser in (
                ("Gas arriving in the UK",
                 "All supply into the national grid: North Sea and Norway, LNG and "
                 "interconnectors. Seven-day average, GWh a day, National Gas", tot),
                ("LNG arriving in the UK",
                 "Liquefied gas by ship, the supply route most exposed to Qatar. Seven-day "
                 "average, GWh a day, National Gas", lng)):
            vals = [v for _, v in ser]
            out.append(T(row(nm, sub, ser[-1][1], year_ago(ser), min(vals), max(vals),
                             "GWh/d", "none", f"since {ser[0][0][:4]}"),
                         trend(ser, False, "pct")))  # seasonal: compare with a year ago
    else:
        out.append(na_row("Gas arriving in the UK", "", "National Gas data missing"))

    con = g("uk_constraint") or {}
    months = sorted(con)
    full = [m for m in months if m != months[-1] or con[m]["days"] >= 28] if months else []
    if len(full) >= 13:
        cur = full[-1]
        vals = [con[m]["gwh_down"] for m in full]
        ser = [[m, con[m]["gwh_down"]] for m in full]
        yago = f"{int(cur[:4]) - 1}{cur[4:]}"
        mtd = months[-1] if months[-1] not in full else None
        sub = (f"Energy turned down to relieve grid limits, mostly Scottish wind, NESO. "
               f"{datetime.fromisoformat(cur + '-01'):%B %Y}: cost £{con[cur]['cost_m']:.0f}m")
        if mtd:
            sub += (f" &middot; {datetime.fromisoformat(mtd + '-01'):%B} so far: "
                    f"{con[mtd]['gwh_down']:,.0f} GWh in {con[mtd]['days']} days")
        r = row("UK wind curtailment", sub, con[cur]["gwh_down"],
                con[yago]["gwh_down"] if yago in con else None,
                min(vals), max(vals), "GWh", "none", f"since {full[0]}")
        out.append(T(r, trend(ser, None, "pct")))
    else:
        out.append(na_row("UK wind curtailment", "", "NESO constraint data missing"))

    neg = take("p-power", "Negative price hours")
    out.append(neg or na_row("Negative price hours", "", "Elexon data missing"))

    dr = g("drewry") or {}
    out.append(hist_row("Drewry container rate, Shanghai to Rotterdam",
                        "US dollars per 40ft container, weekly",
                        dr.get("SHA_RTM"), "$/40ft", True,
                        "Weekly figures to 24 Sep 2026 rebuilt from Drewry's published releases "
                        "and the news sites that carry them (sources in "
                        "shipping_history_sources.csv); read from Drewry's page each week after that."))

    # A figure with no history cannot be judged good or bad, so it is recorded but not shown
    # until half a year of readings exists.
    held_back = [r for r in out if r.get("short") and
                 (_d(r["short"]["last"][0]) - _d(r["short"]["first"][0])).days < 180]
    out = [r for r in out if r not in held_back]
    R["_held_back"] = [{"n": r["n"], "first": r["short"]["first"][0], "count": r["short"]["n"]}
                       for r in held_back]

    for panel, name in (("p-uk", "10-year gilt yield"), ("p-fin", "US 10-year yield")):
        r = take(panel, name)
        if r:
            if name == "10-year gilt yield":
                r["n"] = "UK 10-year gilt yield"
            out.append(r)
    return out


# --------------------------------------------------------------- funds

FUND_WINDOW_DAYS = 183


def build_funds(F):
    """Six months of each fund, rebased to 100 at the start of the window."""
    out = {}
    for key, f in (F or {}).items():
        s = sorted(f.get("series") or [], key=lambda p: p[0])
        if len(s) < 2:
            continue
        end = _d(s[-1][0])
        win = [p for p in s if _d(p[0]) >= end - timedelta(days=FUND_WINDOW_DAYS)]
        if len(win) < 2:
            continue
        base = win[0][1]
        v = s[-1][1]
        def chg(days):
            b = at_or_before(s, end - timedelta(days=days))
            return round((v / b - 1) * 100, 1) if b else None
        if f.get("group") not in ("held", "held_ref"):
            continue
        out[key] = {"name": f["name"], "group": f["group"],
                    "p0": [win[0][0], win[0][1]], "p1": [s[-1][0], v],
                    "pts": [[d, round(x / base * 100, 2)] for d, x in win],
                    "c1": chg(30), "c3": chg(91), "c6": round((v / base - 1) * 100, 1),
                    "to": s[-1][0]}
    return out


# --------------------------------------------------------------- momentum schedule
# The basket section shows the monthly schedule's own output, read from ETF schedule.xlsx.
# Nothing here ranks funds. The build re-does the workbook's arithmetic from its own
# "Prices used" sheet purely as a check, and reports whether every figure agrees.

SCHEDULE_DIR = os.environ.get("MOMENTUM_DIR", r"C:\Users\phil_\Documents\Momentum ETF schedule")
FACTOR_T = ["IUMF", "IUQF", "IUVF", "CUS1", "IEFM", "IEFQ", "IEFV", "XXSC"]
SECTOR_T = ["ICDU", "ICSU", "IESU", "IHCU", "IISU", "IITU", "IMSU"]
MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August",
               "September", "October", "November", "December"]


def ym_add(ym, k):
    y, m = int(ym[:4]), int(ym[5:7])
    t = y * 12 + (m - 1) + k
    return f"{t // 12:04d}-{t % 12 + 1:02d}"


def build_schedule(data_dir, today=None):
    """Read the workbook, check its arithmetic, and return what the page shows."""
    import re
    today = today or datetime.now(timezone.utc).date()
    checks = []
    path = None
    for d in (SCHEDULE_DIR, data_dir):
        p = os.path.join(d, "ETF schedule.xlsx")
        if os.path.exists(p):
            path = p
            break
    if not path:
        checks.append(("Schedule workbook found", False, "ETF schedule.xlsx not found"))
        return None, checks
    wb = read_xlsx(path)
    mtime = datetime.fromtimestamp(os.path.getmtime(path), timezone.utc).strftime("%Y-%m-%d")
    checks.append(("Schedule workbook found", True, f"{path}, saved {mtime}"))
    if os.path.dirname(path) != SCHEDULE_DIR:
        checks.append(("Schedule read from the workbook the monthly task writes", False,
                       f"{SCHEDULE_DIR} was not reachable, so a copy in {data_dir} was read "
                       f"instead; it may be an older month's schedule"))

    this = wb.get("This month", [])
    title = str(this[0][0]) if this else ""
    m = re.search(r"for (\w+) (\d{4})", title)
    sched_month = f"{m.group(2)}-{MONTHS_LONG.index(m.group(1)) + 1:02d}" if m else None
    cur = f"{today.year:04d}-{today.month:02d}"
    ok = sched_month == cur or today.day <= 4          # the monthly task runs on days 1-4
    checks.append(("Schedule is for the current month", ok,
                   f"workbook: {title or 'no title'}; today: {MONTHS_LONG[today.month - 1]} {today.year}"
                   + ("" if sched_month == cur else " (the monthly task runs on the first trading day)")))

    # "Prices used" -> month-end prices
    pu = wb.get("Prices used", [])
    head = [str(x) for x in pu[1]]
    prices = {}
    for r in pu[2:]:
        if r and r[0]:
            prices[str(r[0])] = {head[i]: r[i] for i in range(1, min(len(head), len(r))) if r[i] is not None}
    months = sorted(prices)
    M = months[-1]

    # "Rankings"
    rk = wb.get("Rankings", [])
    rank_title = str(rk[0][0]) if rk else ""
    rows = []
    for r in rk[2:]:
        if r and r[1]:
            rows.append({"basket": r[0], "t": r[1], "name": r[2], "r1": r[3], "r3": r[4], "r6": r[5],
                         "r12": r[6], "blend": r[7], "rb": int(r[8]), "r6rank": int(r[9])})

    # check: every return in Rankings from the Prices used sheet
    n, worst, bad = 0, 0.0, []
    for row in rows:
        for k, fld in ((1, "r1"), (3, "r3"), (6, "r6"), (12, "r12")):
            b = prices.get(ym_add(M, -k), {}).get(row["t"])
            e = prices[M].get(row["t"])
            if b is None or e is None:
                bad.append(f"{row['t']} {k}m: price missing")
                continue
            d = abs((e / b - 1) - row[fld]) * 100
            worst = max(worst, d)
            n += 1
            if d > 0.0001:
                bad.append(f"{row['t']} {k}m: prices give {(e / b - 1) * 100:.3f}%, Rankings shows {row[fld] * 100:.3f}%")
        bl = (row["r1"] + row["r3"] + row["r6"] + row["r12"]) / 4
        if abs(bl - row["blend"]) > 1e-9:
            bad.append(f"{row['t']} blend is not the average of the four")
    checks.append(("Rankings agree with the workbook's own prices", not bad,
                   f"{n} returns recalculated from 'Prices used' (end of {M}); largest difference "
                   f"{worst:.6f} percentage points" + ("; " + "; ".join(bad[:6]) if bad else "")))

    # check: ranks, and the funds to hold are the top-ranked
    rbad = []
    for basket in ("Factor", "Sector"):
        rs = [r for r in rows if r["basket"] == basket]
        for fld, col in (("blend", "rb"), ("r6", "r6rank")):
            order = sorted(rs, key=lambda r: -r[fld])
            for i, r in enumerate(order):
                if r[col] != i + 1:
                    rbad.append(f"{r['t']} {col} is {r[col]}, should be {i + 1}")
    picks = {}
    for r in this:
        if r and r[0] in ("Blend 1-3-6-12", "6-month") and len(r) >= 3:
            picks[r[0]] = {"factor": str(r[1]).split(" - ")[0], "sector": str(r[2]).split(" - ")[0],
                           "factor_name": r[1], "sector_name": r[2], "change": r[3] if len(r) > 3 else ""}
    for rule, col in (("Blend 1-3-6-12", "rb"), ("6-month", "r6rank")):
        for basket, key in (("Factor", "factor"), ("Sector", "sector")):
            top = next((r["t"] for r in rows if r["basket"] == basket and r[col] == 1), None)
            got = picks.get(rule, {}).get(key)
            if top != got:
                rbad.append(f"{rule} {basket}: This month shows {got}, Rankings rank 1 is {top}")
    checks.append(("Rankings are in order and the funds to hold are the top-ranked", not rbad,
                   "; ".join(rbad[:6]) if rbad else "both rules, both baskets"))

    # check: the protection line against the workbook's SPY prices and US unemployment
    lines = [str(r[0]) for r in this if r and r[0]]
    pbad, pdet = [], []
    spy_line = next((l for l in lines if l.startswith("US market (SPY)")), "")
    nums = [float(x) for x in re.findall(r"\d+\.\d+", spy_line)]
    spy = [prices[ym_add(M, -k)].get("SPY") for k in range(10)]
    if len(nums) >= 2 and None not in spy:
        avg10 = sum(spy) / 10
        if abs(nums[0] - spy[0]) > 0.006 or abs(nums[1] - avg10) > 0.006:
            pbad.append(f"workbook says {nums[0]} vs {nums[1]}; prices give {spy[0]:.2f} vs {avg10:.2f}")
        else:
            pdet.append(f"SPY {spy[0]:.2f} against 10-month average {avg10:.2f} recalculated")
    else:
        pbad.append("could not read the SPY line or the SPY prices")
    un_path = os.path.join(os.path.dirname(path), "unemployment.csv")
    un_line = next((l for l in lines if l.startswith("US unemployment")), "")
    if os.path.exists(un_path) and un_line:
        with open(un_path, encoding="utf-8") as f:
            un = [(r[0], float(r[1])) for r in (l.strip().split(",") for l in f.readlines()[1:])
                  if len(r) > 1 and r[1] not in ("", ".")]
        un = [(mth, v) for mth, v in un if mth <= M]
        un = [(mth, v) for mth, v in un if mth <= ym_add(M, -1)] if un else un
        u3 = sum(v for _, v in un[-3:]) / 3
        low12 = min(v for _, v in un[-12:])
        wn = [float(x) for x in re.findall(r"\d+\.\d+", un_line)]
        if len(wn) >= 2 and abs(wn[0] - u3) < 0.006 and abs(wn[1] - low12) < 0.06:
            pdet.append(f"unemployment {u3:.2f}% against low {low12:.1f}% recalculated")
        else:
            pbad.append(f"unemployment line says {wn[:2]}; unemployment.csv gives {u3:.2f} and {low12:.1f}")
    checks.append(("Protection line agrees with the prices and unemployment data", not pbad,
                   "; ".join(pbad or pdet)))

    # charts: month-end prices from the workbook, six months, rebased to 100 at the start
    base = ym_add(M, -6)
    chart_months = [mm for mm in months if base <= mm <= M]
    series = {}
    for t in FACTOR_T + SECTOR_T:
        b = prices.get(base, {}).get(t)
        if b:
            series[t] = [[mm, round(prices[mm][t] / b * 100, 2)] for mm in chart_months
                         if prices[mm].get(t) is not None]
    history = []
    for r in wb.get("History", [])[2:]:
        if r and r[0]:
            history.append({"m": str(r[0]), "bf": r[1], "bs": r[2], "sf": r[3], "ss": r[4],
                            "prot": bool(r[5]), "br": r[6], "sr": r[7]})
    return {"title": title, "lines": lines, "picks": picks, "rank_title": rank_title,
            "rows": rows, "M": M, "base": base, "series": series,
            "prices": {mm: {t: prices[mm].get(t) for t in FACTOR_T + SECTOR_T} for mm in chart_months},
            "history": history[-6:], "path": path, "saved": mtime}, checks


# --------------------------------------------------------------- the three risks
# The page is organised by the three named risks, not by where the data comes from, so each
# row can be judged on whether it would actually tell you that risk is crystallising.
# build_rows and build_brief still derive every row exactly as before; this layer only decides
# which block each one belongs in. Anything that earns no place is still built and still
# checked, and sits in the fold.

GULF = ["Baltic Dirty Tanker Index", "Baltic Clean Tanker Index", "Crude tanker owners (Frontline)",
        "Diesel crack against Brent", "Brent less WTI", "NBP gas, front month",
        "UK gas in storage", "Gas arriving in the UK"]
DEBT = ["UK gilt auction cover", "US auctions, last month: overseas-type share",
        "US auctions, last month: dealer share", "US yields up, dollar down",
        "US 10-year auction cover", "US 10-year dealer share",
        "UK 10-year gilt yield", "20-year gilt yield", "UK 10-year real yield",
        "UK 10-year expected inflation", "US 10-year yield", "US 30-year yield"]
TRANS = ["High yield vs investment grade", "Equity volatility", "Market breadth", "Gold",
         "US unemployment rate", "US Sahm recession gauge", "US initial jobless claims"]
AI_BLOCK = ["Oracle backlog to revenue"]
MARKET_FOLD = ["Oracle", "CoreWeave", "Nvidia", "Semiconductors vs market"]
RENAME = {"20-year gilt yield": "UK 20-year gilt yield"}

# A chip and a range bar say "this reading can be judged good or bad". Where the lowest and
# highest were never measured, they cannot, so the row shows its value and direction only.
NO_SOURCED_RANGE = {
    "Oracle backlog to revenue":
        "No measured range: the comparison is that a normal software company runs at one to "
        "two times, not a recorded high and low.",
    "UK gilt auction cover":
        "Typed in from the DMO report and not verified from source, so there is no range to "
        "judge it against.",
}


def off_high_series(s, win=252):
    """Percent below the trailing twelve-month high, for every day in the series.

    Gives these rows a range that is their own recorded history rather than a round number.
    """
    s = sorted(s or [])
    out = []
    for i in range(len(s)):
        wnd = [v for _, v in s[max(0, i - win + 1):i + 1]]
        hi = max(wnd)
        if hi:
            out.append([s[i][0], round((s[i][1] / hi - 1) * 100, 2)])
    return out


def chg_pct_series(s, n=252):
    """Twelve-month percent change, for every day - again so the row has a real range."""
    s = sorted(s or [])
    return [[s[i][0], round((s[i][1] / s[i - n][1] - 1) * 100, 2)]
            for i in range(n, len(s)) if s[i - n][1]]


def real_range(r, hist):
    """Replace a typed-in range with the measure's own recorded distribution."""
    if not hist:
        return r
    vals = [v for _, v in hist]
    lo, hi = min(vals), max(vals)
    r["lo"], r["hi"] = round(lo, 2), round(hi, 2)
    r["m12"] = None
    r["w"] = "its own range since %s" % hist[0][0][:4]
    r["st"] = band(r["v"], lo, hi, invert=True)   # nearer the bottom of the range is worse
    return r


def build_blocks(S, R):
    """Redistribute every derived row into the four risk blocks, plus the fold."""
    g = S.get
    T = with_trend
    pool = [r for rows in R.values() for r in rows]
    taken = set()

    def grab(name):
        for r in pool:
            if r["n"] == name and id(r) not in taken:
                taken.add(id(r))
                r["n"] = RENAME.get(r["n"], r["n"])
                return r
        return None

    def block(names):
        return [r for r in (grab(n) for n in names) if r]

    gulf, debt, trans = block(GULF), block(DEBT), block(TRANS)
    ai = block(AI_BLOCK)

    # ---- ranges that were typed in rather than measured ----------------
    for r in pool:
        if r["n"] in NO_SOURCED_RANGE:
            r["norange"] = NO_SOURCED_RANGE[r["n"]]
            r["st"] = "none"
    for nm, key in (("Oracle", "oracle"), ("CoreWeave", "coreweave"), ("Nvidia", "nvidia")):
        for r in pool:
            if r["n"] == nm:
                real_range(r, off_high_series(g(key, [])))
    for nm, key in (("Crude tanker owners (Frontline)", "frontline"),
                    ("Crude tanker owners (DHT)", "dht")):
        for r in pool:
            if r["n"] == nm:
                real_range(r, chg_pct_series(g(key, [])))

    # ---- ranges labelled with a history the row does not actually use ---
    # These carried round numbers under a "since 1962" label. Recomputed from the pulled
    # series over the same twenty-year window the gilt rows use, so the label is true and
    # the chip is earned. The threshold marker still does the work on the US 10-year.
    def since(key, start="2005-01-01"):
        return [[d, v] for d, v in g(key, []) if d >= start]

    def reband(name, ser, invert=False, label="since 2005"):
        if not ser:
            return
        vals = [v for _, v in ser]
        for r in pool:
            if r["n"] == name:
                r["lo"], r["hi"] = round(min(vals), 2), round(max(vals), 2)
                r["w"] = label
                r["st"] = band(r["v"], min(vals), max(vals), invert=invert)

    us10s, us3ms = since("us_treasury_10y"), since("us_treasury_3m")
    reband("US 10-year yield", us10s)
    reband("US 30-year yield", since("us_treasury_30y"))
    reband("US yield curve", diff_series(us10s, us3ms), invert=True)
    reband("10-year gilt over Bank Rate",
           diff_series(since("uk_gilt_10y"), since("uk_bank_rate")))
    # The Sahm gauge has no useful historical range - the 2020 spike swamps it - but it does
    # have a published rule. Judge it against that, and say so rather than implying a range.
    for r in pool:
        if r["n"] == "US Sahm recession gauge":
            r["lo"], r["hi"] = 0.0, 0.5
            r["w"] = "against the +0.50 trigger"
            r["st"] = band(r["v"], 0.0, 0.5)
            r["thr"], r["thrlbl"] = 0.5, "+0.50 recession trigger"
            r["ruleband"] = True

    # ---- crude freight against product freight ------------------------
    # A spread is only legible once you can see which leg moved, and the gap between moving
    # crude and moving refined product is the clearest read on where the dislocation sits.
    bal = g("baltic") or {}
    bc, bd = dict(bal.get("BCTI") or []), dict(bal.get("BDTI") or [])
    both = sorted([d0, bd[d0] / bc[d0]] for d0 in bd if d0 in bc and bc[d0])
    if len(both) >= 8:
        vals = [v for _, v in both]
        r = T(row("Crude against product freight",
                  "Baltic Dirty divided by Baltic Clean &middot; how much dearer it is to move "
                  "crude than refined product &middot; rising means the crude leg is "
                  "deteriorating faster",
                  round(both[-1][1], 2), None, round(min(vals), 2), round(max(vals), 2), "x",
                  band(both[-1][1], min(vals), max(vals)), "since %s" % both[0][0]),
              trend(both, True, "abs"))
        with_pre(r, both)
        # sits with the two indices it is derived from, not at the end of the block
        after = max((i for i, x in enumerate(gulf) if x["n"].startswith("Baltic")), default=-1)
        gulf.insert(after + 1, r)

    # ---- which leg moved a spread ------------------------------------
    # A crack can fall because product got cheaper or because crude got dearer. Those mean
    # opposite things, so the row shows both legs rather than only the difference.
    def split_legs(name, a_key, b_key, a_lab, b_lab, a_scale=1.0):
        for r in pool:
            if r["n"] != name:
                continue
            roc = (r.get("tr") or {}).get("roc")
            if not roc:
                return
            A, B = dict(g(a_key, [])), dict(g(b_key, []))
            d0, d1 = roc["from"][0], r["tr"]["now"][0]
            if d0 in A and d1 in A and d0 in B and d1 in B:
                r["legs"] = {"from": d0, "to": d1,
                             "a": round(A[d1] * a_scale - A[d0] * a_scale, 2), "alab": a_lab,
                             "b": round(B[d1] - B[d0], 2), "blab": b_lab}

    split_legs("Diesel crack against Brent", "heating_oil", "brent", "diesel", "Brent", 42.0)
    split_legs("Diesel crack against US crude", "heating_oil", "wti", "diesel", "WTI", 42.0)
    split_legs("Brent less WTI", "brent", "wti", "Brent", "WTI")

    # ---- sterling: named in the transmission block, never shown before -
    gu = g("gbpusd", [])
    if gu:
        # The feed quotes pounds per dollar; sterling is read the other way round.
        cable = [[d, 1.0 / v] for d, v in gu if v]
        vals = [v for _, v in cable]
        trans.append(T(row("Sterling against the dollar",
                           "Dollars per pound &middot; a fall alongside rising long yields is "
                           "the adverse combination named in the debt risk",
                           round(cable[-1][1], 4),
                           round(cable[-253][1], 4) if len(cable) > 252 else None,
                           round(min(vals), 4), round(max(vals), 4), "ratio",
                           band(cable[-1][1], min(vals), max(vals), invert=True),
                           "since %s" % cable[0][0][:4]),
                       trend(cable, False, "pct")))

    # ---- the AI return side, and the gaps named rather than papered over
    ai.append(row("Disclosed enterprise AI revenue",
                  "Largest auditable franchises in company filings, combined &middot; against "
                  "the capital spending above &middot; typed in %s, not verified from source"
                  % MANUAL["ai_disclosed_arr_bn"][1],
                  MANUAL["ai_disclosed_arr_bn"][0], None, None, None, "$bn", "none",
                  "no range"))
    ai[-1]["norange"] = ("One reported figure, no history: shown for direction only. The return "
                         "side is barely measured anywhere.")
    prod = g("us_productivity", [])
    if prod:
        lo, hi = rng(prod[-80:])
        ai.append(T(row("US productivity growth",
                        "Nonfarm business output per hour, BLS &middot; quarterly, so the two "
                        "changes are one and three quarters back &middot; whether the wider economy "
                        "is getting enough value from AI to fund the revenue the labs need",
                        last(prod), back(prod, 4), lo, hi, "idx",
                        band(last(prod), lo, hi, invert=True), "20 years"),
                    trend(prod, False, "pct", obs=True)))
    # ---- AI demand (added 10 October 2026): how many pay, and how much. Penetration and
    # monetisation are what the build-out has to grow into; the rate of change matters more
    # than the level, so these carry no state chip.
    for key, nm, sub in (
            ("ramp_ai_all", "Businesses paying for AI",
             "Share of US businesses on Ramp paying for any AI tool, from card and bill payments "
             "&middot; Ramp AI Index via FRED, monthly &middot; penetration of business demand"),
            ("ramp_ai_anthropic", "Businesses paying Anthropic",
             "Share of US businesses on Ramp paying Anthropic (Claude) &middot; Ramp AI Index via "
             "FRED, monthly"),
            ("ramp_ai_openai", "Businesses paying OpenAI",
             "Share of US businesses on Ramp paying OpenAI &middot; Ramp AI Index via FRED, "
             "monthly")):
        ser = g(key, [])
        if ser:
            lo, hi = rng(ser)
            ai.append(T(row(nm, sub, round(last(ser), 2), round(back(ser, 12), 2) if len(ser) > 12 else None,
                            round(lo, 2), round(hi, 2), "%", "none", "since %s" % ser[0][0][:4]),
                        trend(ser, False, "pp", obs=True)))
            ai[-1]["norange"] = ("Adoption rising is the build-out working. The warning is the "
                                 "monthly gain shrinking or turning negative: watch the change, "
                                 "not the level.")
            ai[-1]["demand"] = True
    for key, nm, sub in (
            ("us_bbb_yield", "BBB borrowing cost",
             "ICE BofA BBB effective yield &middot; Oracle&rsquo;s credit tier, not Oracle itself "
             "&middot; what the investment-grade end of the financing chain pays"),
            ("us_hy_yield", "High-yield borrowing cost",
             "ICE BofA high-yield effective yield &middot; CoreWeave&rsquo;s credit tier &middot; "
             "the measure that bites first in a refinancing squeeze")):
        ser = g(key, [])
        if ser:
            lo, hi = rng(ser)
            ai.append(T(row(nm, sub, last(ser), back(ser, 252), lo, hi, "%",
                            band(last(ser), lo, hi), "since 1997"),
                        trend(ser, True, "bp", span_of(lo, hi))))
    for nm, sub, why in (
            ("Backlog conversion at the five providers",
             "Microsoft, Alphabet, Amazon, Oracle and CoreWeave &middot; is contracted backlog "
             "converting to revenue as guided, and are there downward revisions",
             "needs the quarterly filings typed in - only Oracle is entered so far"),
            ("Lab compute spend as a share of revenue",
             "OpenAI and Anthropic against the 50-60% working ceiling &middot; both unaudited",
             "no filed source exists; needs the reported figures typed in each round"),
            ("Borrowing costs and productivity",
             "Credit-tier borrowing costs and US output per hour &middot; added to the pull on "
             "1 October 2026",
             "added to the pull on 1 October 2026 - first reading on the next run")
            if not (g("us_bbb_yield") or g("us_productivity")) else (None, None, None),
            ("Stake gains at Amazon and Alphabet",
             "Gains booked on their holdings in AI labs, as a share of reported profit &middot; "
             "profit that is a mark-up on their own customer",
             "needs the quarterly filings typed in")):
        if nm:
            ai.append(na_row(nm, sub, why))

    detail = [r for r in pool if id(r) not in taken and r["n"] not in MARKET_FOLD]
    market = [r for r in pool if r["n"] in MARKET_FOLD and id(r) not in taken]
    for r in market:
        taken.add(id(r))

    out = {"p-gulf": gulf, "p-debt": debt, "p-ai": ai, "p-trans": trans,
           "p-fold-detail": detail, "p-fold-market": market}
    # Every row build_rows produced must land in exactly one block.
    placed = [r for rows in out.values() for r in rows]
    assert len(placed) == len(set(id(r) for r in placed)), "a row was placed twice"
    missing = [r["n"] for r in pool if not any(r is p for p in placed)]
    assert not missing, "rows lost in assembly: %s" % missing
    return out


# --------------------------------------------------------------- AI demand, typed in
# Figures published only as web pages (a16z consumer data, Ramp prices and spend). Each new
# release is added to ai_demand.json by the daily run as a new observation; the page shows the
# latest, the change on the previous release and the change on a year earlier.

AI_DEMAND_FILE = "ai_demand.json"
DEMAND_NOTES = []


def _ym(p):
    try:
        y, m = p.split("-")[:2]
        return int(y), int(m)
    except Exception:
        return None


def demand_rows(d, pulled):
    p = os.path.join(d, AI_DEMAND_FILE)
    if not os.path.exists(p):
        DEMAND_NOTES.append(("AI demand figures on file", "warn", "no ai_demand.json"))
        return []
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        DEMAND_NOTES.append(("AI demand figures on file", "warn", f"ai_demand.json unreadable: {e}"))
        return []
    out, stale = [], []
    for sr in data.get("series") or []:
        obs = sorted([o for o in sr.get("obs") or [] if o.get("value") is not None],
                     key=lambda o: o.get("period", ""))
        if not obs:
            out.append(na_row(sr["name"], sr.get("sub", ""), "no release recorded yet"))
            continue
        lt, prev = obs[-1], (obs[-2] if len(obs) > 1 else None)
        mode = sr.get("change", "pct")

        def chg(a, b):
            if a is None or b is None:
                return None
            return round(a - b, 2) if mode == "pp" else (round((a / b - 1) * 100, 1) if b else None)
        ym = _ym(lt["period"])
        yago = None
        if ym:
            want = f"{ym[0] - 1:04d}-{ym[1]:02d}"
            yago = next((o for o in obs if o["period"] == want), None)
        src = f"typed in from {lt.get('source', 'source not named')}"
        r = row(sr["name"], sr.get("sub", "") + " &middot; " + src, lt["value"],
                yago["value"] if yago else None, None, None, sr.get("unit", ""), "none", "no range")
        r["norange"] = ("Typed in from each release; too few readings for a range. The change "
                        "between releases is the signal.")
        r["demand"] = True
        r["dem"] = {"latest": lt, "prev": prev, "yago": yago, "mode": mode,
                    "chg_prev": chg(lt["value"], prev["value"]) if prev else None,
                    "chg_12m": chg(lt["value"], yago["value"]) if yago else None,
                    "n": len(obs), "first": obs[0]["period"]}
        if ym:
            age = (pulled - datetime(ym[0], ym[1], 1).date()).days
            if age > sr.get("max_age_days", 120):
                stale.append(f"{sr['name']}: latest release covers {lt['period']}")
        out.append(r)
    DEMAND_NOTES.append(("AI demand figures are recent", "warn" if stale else True,
                         "; ".join(stale) + " - check for a newer release" if stale else
                         f"{len(out)} typed-in series, each within its release cycle"))
    return out


# --------------------------------------------------------------- what changed since last run
# The page says what moved rather than leaving it to be spotted. State is written each run and
# read back on the next one, so "changed" means changed, not "looks different to me".

STATE_FILE = "monitor_state.json"


def build_standing(d, R, payload):
    now = {r["n"]: r.get("st") for rows in R.values() for r in rows if not r.get("na")}
    prev = {}
    p = os.path.join(d, STATE_FILE)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                saved = json.load(f)
            # A second build of the same pull (after the reading is written) must compare
            # against the run before, not against itself.
            if saved.get("pulled") == payload["generated"] and "prev" in saved:
                prev = saved.get("prev") or {}
            else:
                prev = saved.get("states") or {}
        except Exception:
            prev = {}
    moved = [{"n": k, "from": prev[k], "to": v}
             for k, v in now.items() if k in prev and prev[k] != v and v != "none"]
    pulled = datetime.fromisoformat(payload["generated"]).date()
    age = (datetime.now(timezone.utc).date() - pulled).days
    blocks = []
    for pid, label in (("p-gulf", "The Gulf"), ("p-debt", "Government debt"),
                       ("p-ai", "AI overreach"), ("p-trans", "Transmission")):
        allrows = R.get(pid, [])
        rows = [r for r in allrows if not r.get("na") and r.get("st") in ("watch", "extreme")]
        graded = [r for r in allrows if not r.get("na") and not r.get("norange")]
        blocks.append({"n": label,
                       "extreme": [r["n"] for r in rows if r["st"] == "extreme"],
                       "watch": [r["n"] for r in rows if r["st"] == "watch"],
                       "pending": len([r for r in allrows if r.get("na")]),
                       "nojudge": len([r for r in allrows if r.get("norange")]),
                       "total": len(allrows), "graded": len(graded)})
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"pulled": payload["generated"], "states": now, "prev": prev}, f)
    fast = []
    for pid, label in (("p-gulf", "The Gulf"), ("p-debt", "Government debt"),
                       ("p-ai", "AI overreach"), ("p-trans", "Transmission")):
        for r in R.get(pid, []):
            roc = (r.get("tr") or {}).get("roc")
            if not roc or roc.get("rank") is None:
                continue
            fast.append({"n": r["n"], "block": label, "lab": roc["lab"], "chg": roc["chg"],
                         "rank": roc["rank"], "pace": roc.get("pace"), "u": r.get("u"),
                         "m": r["tr"]["m"], "stress": roc.get("stress")})
    fast.sort(key=lambda x: -x["rank"])
    return {"age": age, "pulled": pulled.isoformat(), "moved": moved, "blocks": blocks,
            "fast": [f for f in fast if f["rank"] >= 75][:6], "first_run": not prev}


# --------------------------------------------------------------- what was happening
# The build can say what moved. It cannot say why: that needs news, which no series carries.
# context_notes.json holds dated, sourced notes written by a person (or by the weekly publish
# run, which has web access). They are shown beside the measures they bear on, and labelled as
# reporting rather than measurement so the two can never be confused.

CONTEXT_FILE = "context_notes.json"


def build_context(d):
    p = os.path.join(d, CONTEXT_FILE)
    if not os.path.exists(p):
        return []
    try:
        with open(p, encoding="utf-8") as f:
            notes = (json.load(f) or {}).get("notes") or []
    except Exception:
        return []
    for n in notes:
        n.setdefault("measures", [])
        n.setdefault("source", "")
        n.setdefault("url", "")
    notes.sort(key=lambda n: n.get("date", ""), reverse=True)
    return notes[:12]


# --------------------------------------------------------------- digest and reading
# The build says what moved. The reading says what it means. digest.json is a compact summary
# of every measure, written for whoever writes the reading (the daily task). readings.json is
# that reading. Every figure quoted in the reading's own words must be one the digest holds, so
# a number cannot be invented; figures from news reports are allowed only inside the reported
# items, which must name a source, a date and a link.

DIGEST_FILE = "digest.json"
READINGS_FILE = "readings.json"
BLOCK_KEYS = {"gulf": ("p-gulf", "The Gulf"), "debt": ("p-debt", "Government debt"),
              "ai": ("p-ai", "AI overreach"), "trans": ("p-trans", "Transmission")}
_MONTHS = ("January|February|March|April|May|June|July|August|September|October|November|"
           "December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec")


def _r(x, k=4):
    return None if x is None else round(float(x), k)


def digest_row(r, block):
    tr = r.get("tr") or {}
    roc = tr.get("roc") or {}
    sp = [p for p in (tr.get("sp") or []) if p and p[1] is not None]
    out = {"n": r["n"], "block": block, "about": r.get("s"), "unit": r.get("u"),
           "state": r.get("st"), "now": _r(r.get("v")), "at": tr.get("to"),
           "range_window": r.get("w"), "range_low": _r(r.get("lo")), "range_high": _r(r.get("hi")),
           "year_ago": _r(r.get("m12"))}
    if r.get("na"):
        out["missing"] = r["na"]
        return out
    if tr:
        out["change_mode"] = "percent" if tr.get("m") == "pct" else "absolute"
        out["direction_3m"] = tr.get("dir")
        out["change_1m"] = _r(tr.get("c1")); out["change_1m_from"] = tr.get("b1")
        out["change_3m"] = _r(tr.get("c3")); out["change_3m_from"] = tr.get("b3")
    if roc:
        out["change_latest"] = {"period": roc.get("lab"), "change": _r(roc.get("chg")),
                                "from": roc.get("from"), "bigger_than_pct_of_periods": roc.get("rank"),
                                "pace": roc.get("pace"), "stress_direction": roc.get("stress")}
    if sp:
        hi = max(sp, key=lambda p: p[1]); lo = min(sp, key=lambda p: p[1]); now = sp[-1][1]
        out["six_month_high"] = [hi[0], _r(hi[1])]
        out["six_month_low"] = [lo[0], _r(lo[1])]
        out["from_six_month_high"] = {"absolute": _r(now - hi[1]),
                                      "percent": _r((now / hi[1] - 1) * 100, 1) if hi[1] else None}
        out["from_six_month_low"] = {"absolute": _r(now - lo[1]),
                                     "percent": _r((now / lo[1] - 1) * 100, 1) if lo[1] else None}
        out["recent_points"] = [[p[0], _r(p[1])] for p in sp[-10:]]
    if r.get("pre"):
        out["before_hormuz_closed"] = {"avg_jan_feb_2026": _r(r["pre"].get("avg")),
                                       "change_percent": _r(r["pre"].get("chg_pct"), 1),
                                       "change_absolute": _r(r["pre"].get("chg_abs"))}
    if r.get("legs"):
        L = r["legs"]
        out["legs"] = {"from": L.get("from"), "to": L.get("to"),
                       L.get("alab", "a"): _r(L.get("a")), L.get("blab", "b"): _r(L.get("b"))}
    if r.get("short"):
        out["short_history"] = r["short"]
    if r.get("dem"):
        dm = r["dem"]
        out["releases"] = [{"period": o.get("period"), "value": o.get("value"),
                            "source": o.get("source"), "stated": o.get("stated")}
                           for o in ([dm["prev"]] if dm["prev"] else []) + [dm["latest"]]]
        out["change_on_previous_release"] = dm["chg_prev"]
        out["change_on_year_earlier"] = dm["chg_12m"]
        out["change_mode"] = "percentage points" if dm["mode"] == "pp" else "percent"
        if dm["yago"]:
            out["year_earlier_release"] = {"period": dm["yago"]["period"], "value": dm["yago"]["value"]}
    return out


def build_digest(d, R, STANDING, CONTEXT, payload):
    rows = []
    for key, (pid, label) in BLOCK_KEYS.items():
        rows += [dict(digest_row(r, label), block_key=key) for r in R.get(pid, [])]
    for pid in ("p-fold-detail", "p-fold-market"):
        rows += [dict(digest_row(r, "Detail (in the fold)"), block_key="detail") for r in R.get(pid, [])]
    dg = {"how_to_read": [
              "One entry per measure on the page. block_key says which risk it sits under "
              "(gulf, debt, ai, trans) or 'detail' for the fold.",
              "now/at: the latest reading and its date. state: calm, watch or extreme against "
              "range_low/range_high over range_window. For several measures that range is built "
              "from month-end values, so a daily reading can sit outside it; six_month_high/low "
              "come from the daily trace.",
              "change_mode says whether change_1m/change_3m are percent or absolute. For "
              "measures in % (yields, shares), absolute changes are in basis points for yields "
              "and percentage points for shares.",
              "change_latest is the most recent weekly (or monthly) move, with how unusual it is "
              "(bigger_than_pct_of_periods) and whether it is speeding up.",
              "legs split a spread's latest move into its two parts. before_hormuz_closed "
              "compares with the January-February 2026 average.",
              "Quote figures as the digest holds them (rounding is fine). The build rejects a "
              "reading containing a figure it cannot find among the block's measures or those "
              "listed under 'uses'."],
          "pulled": datetime.fromisoformat(payload["generated"]).date().isoformat(),
          "changed_state_since_last_run": STANDING.get("moved"),
          "moving_fastest": STANDING.get("fast"),
          "notes_on_file": CONTEXT, "measures": rows}
    with open(os.path.join(d, DIGEST_FILE), "w", encoding="utf-8") as f:
        json.dump(dg, f, indent=1)
    return dg


def _numbers_in(obj, acc):
    if isinstance(obj, bool) or obj is None:
        return
    if isinstance(obj, (int, float)):
        acc.add(abs(float(obj)))
    elif isinstance(obj, str):
        if re.fullmatch(r"\d{4}-\d{2}(-\d{2})?", obj):     # dates are not figures
            return
        for t in re.findall(r"\d[\d,]*\.?\d*", obj):
            try:
                acc.add(abs(float(t.replace(",", ""))))
            except ValueError:
                pass
    elif isinstance(obj, dict):
        for v in obj.values():
            _numbers_in(v, acc)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            _numbers_in(v, acc)


def _allowed(dg, key, uses=()):
    """Figures a block's reading may quote: those of its own measures, plus any measure from
    elsewhere on the page that the reading lists by name under "uses"."""
    acc = set()
    uses = {u.strip().lower() for u in uses or ()}
    for m in dg["measures"]:
        if m.get("block_key") != key and m["n"].lower() not in uses:
            continue
        tmp = set()
        _numbers_in({k: v for k, v in m.items() if k not in ("about", "n", "missing")}, tmp)
        acc |= tmp                    # yield changes are already held in basis points
    acc |= {20.0, 52.0, 100.0}
    return acc


def _figures(text):
    t = re.sub(r"\b\d{1,2}(st|nd|rd|th)?\s+(%s)\b" % _MONTHS, " ", text)      # 6 October
    t = re.sub(r"\b(%s)\s+\d{1,2}\b" % _MONTHS, " ", t)                        # October 6
    t = re.sub(r"\b(19|20)\d{2}\b", " ", t)                                     # years
    t = re.sub(r"\b\d+(-|‑)(year|month|week|day|hour|quarter)s?\b", " ", t)  # 10-year
    t = re.sub(r"\b\d+(-\d+)+\b", " ", t)                                       # 1-3-6-12
    out = []
    for m in re.finditer(r"\d[\d,]*\.?\d*", t):
        tok = m.group(0).rstrip(".,")
        try:
            out.append((tok, float(tok.replace(",", ""))))
        except ValueError:
            pass
    return out


def _known(tok, x, allowed):
    if x <= 12 and float(x).is_integer():
        return True
    dec = len(tok.split(".")[1]) if "." in tok else 0
    tol = 0.5 * 10 ** -dec + 1e-9                 # the figure as written must round from the data
    return any(abs(x - v) <= tol for v in allowed)


# ---------------------------------------------------------- research end point (added 9 Oct 2026)
# The reading is only shown if the research behind it reached a defined end point, checked here
# rather than judged by whoever wrote it. research.json must hold, for this pull:
#   questions:  every standing question below, each answered with at least one source
#               (url + date), status "answered" or "nothing new" (a source was still checked);
#   leads:      every follow-up raised by a story; none may be "open". "closed" needs a url,
#               "unanswerable" needs a reason;
#   saturation: per risk, the last two searches run and found_new false for both;
#   feeds:      rolling news feeds read, with every relevant item mapped to a question or lead
#               ("covered_by"), or a reason the feeds could not be read (shown as a note);
# and every reported item in readings.json must come from a source recorded in research.json.
# If any of this fails, the whole reading is held back: no reading is better than a thin one.
RESEARCH_FILE = "research.json"
STANDING_QUESTIONS = {
    "gulf-talks":   ("gulf", "Where do the US-Iran talks and the fighting stand, and what changed since the last run?"),
    "gulf-ships":   ("gulf", "What happened to shipping through the Strait of Hormuz and Gulf oil flows?"),
    "gulf-prices":  ("gulf", "What moved oil, diesel and gas prices, and why?"),
    "debt-uk":      ("debt", "What moved UK gilt yields, and what is expected from the Bank of England and the Budget?"),
    "debt-us":      ("debt", "What moved US Treasury yields, and what is expected from the Federal Reserve?"),
    "debt-demand":  ("debt", "Is demand at government bond auctions holding up?"),
    "debt-infl":    ("debt", "What did the latest inflation figures show, and what is due next?"),
    "ai-openai":    ("ai", "What is new on OpenAI's revenue, funding and compute commitments?"),
    "ai-anthropic": ("ai", "What is new on Anthropic's revenue, funding and compute commitments?"),
    "ai-builders":  ("ai", "What is new on the hyperscalers', Oracle's, CoreWeave's and Nvidia's spending, borrowing and results?"),
    "ai-shares":    ("ai", "What moved AI and technology shares and valuations, and why?"),
    "trans-credit": ("trans", "What is new in company credit: spreads, defaults, refinancing, private credit?"),
    "trans-market": ("trans", "What moved shares, volatility, gold and sterling, and why?"),
    "trans-dates":  ("trans", "Which dates in the next six weeks matter to these risks?"),
}


def check_research(d, dg, rd):
    """Return a list of reasons the research is not finished (empty list = finished)."""
    p = os.path.join(d, RESEARCH_FILE)
    if not os.path.exists(p):
        return ["no research.json - research not recorded"], None
    try:
        with open(p, encoding="utf-8") as f:
            rs = json.load(f)
    except Exception as e:
        return [f"research.json unreadable: {e}"], None
    if rs.get("pulled") != dg["pulled"]:
        return [f"research is for {rs.get('pulled')}, data pulled {dg['pulled']}"], None
    gaps, srcs = [], set()
    qs = {q.get("id"): q for q in rs.get("questions") or []}
    for qid, (blk, text) in STANDING_QUESTIONS.items():
        q = qs.get(qid)
        if not q:
            gaps.append(f"question not answered: {text}"); continue
        good = [s for s in q.get("sources") or [] if s.get("url") and s.get("date")]
        if q.get("status") not in ("answered", "nothing new") or not (q.get("answer") or "").strip() or not good:
            gaps.append(f"question not answered with a source: {text}")
        srcs |= {s["url"] for s in good}
    for ld in rs.get("leads") or []:
        st = ld.get("status")
        if st == "closed" and ld.get("url"):
            srcs.add(ld["url"])
        elif st == "unanswerable" and (ld.get("reason") or "").strip():
            pass
        else:
            gaps.append(f"lead still open: {ld.get('lead', '?')}")
    sat = rs.get("saturation") or {}
    for blk in ("gulf", "debt", "ai", "trans"):
        last = (sat.get(blk) or {}).get("last_searches") or []
        if len(last) < 2 or any(s.get("found_new") is not False for s in last[-2:]):
            gaps.append(f"{BLOCK_KEYS[blk][1]}: searching stopped before two searches in a row found nothing new")
    feeds = rs.get("feeds") or {}
    feed_note = None
    if feeds.get("items") is not None:
        miss = [it.get("title", "?") for it in feeds["items"]
                if it.get("relevant") and not it.get("covered_by")]
        if miss:
            gaps.append(f"feed items not followed up: {'; '.join(miss[:3])}")
        srcs |= {it["url"] for it in feeds["items"] if it.get("url") and it.get("covered_by")}
    else:
        feed_note = f"news feeds not read: {feeds.get('unavailable') or 'no reason given'}"
        if not feeds.get("unavailable"):
            gaps.append(feed_note)
    for key, b in (rd.get("blocks") or {}).items():
        for it in b.get("reported") or []:
            if it.get("url") and it["url"] not in srcs:
                gaps.append(f"{BLOCK_KEYS.get(key, (None, key))[1]}: reported item not in the research record ({it['url'][:60]})")
    return gaps, feed_note


def build_readings(d, dg):
    """Read readings.json, check it, and return (readings for the page, check results)."""
    p = os.path.join(d, READINGS_FILE)
    checks = []
    if not os.path.exists(p):
        checks.append(("Reading written for this pull", "warn",
                       "no readings.json - the page shows the measures without a reading"))
        return None, checks
    try:
        with open(p, encoding="utf-8") as f:
            rd = json.load(f)
    except Exception as e:
        checks.append(("Reading written for this pull", "warn", f"readings.json unreadable: {e}"))
        return None, checks
    if rd.get("pulled") != dg["pulled"]:
        checks.append(("Reading written for this pull", "warn",
                       f"reading is for {rd.get('pulled')}, data pulled {dg['pulled']} - not shown"))
        return None, checks
    gaps, feed_note = check_research(d, dg, rd)
    if gaps:
        checks.append(("Research reached its end point", "warn",
                       f"{len(gaps)} gap(s): " + "; ".join(gaps[:6]) +
                       " - whole reading held back"))
        return None, checks
    checks.append(("Research reached its end point", True,
                   f"all {len(STANDING_QUESTIONS)} standing questions sourced, no open leads, "
                   "every risk searched to saturation" + (f"; note: {feed_note}" if feed_note else "")))
    allowed_all = set().union(*(_allowed(dg, k, [m["n"] for m in dg["measures"]
                                                 if m.get("block_key") == "detail"])
                                for k in BLOCK_KEYS))
    held, out_blocks = [], {}
    for key, b in (rd.get("blocks") or {}).items():
        if key not in BLOCK_KEYS:
            held.append(f"{key}: not a block on the page"); continue
        rep_nums = set()
        bad_rep = [i for i, it in enumerate(b.get("reported") or [])
                   if not (it.get("source") and it.get("date") and it.get("url"))]
        _numbers_in([it.get("text", "") for it in b.get("reported") or []], rep_nums)
        unknown = []
        names = {m["n"].lower() for m in dg["measures"]}
        badnames = [u for u in b.get("uses") or [] if u.strip().lower() not in names]
        allowed = _allowed(dg, key, b.get("uses"))
        for fld in ("summary", "measured", "reading", "watch"):
            for tok, x in _figures(b.get(fld) or ""):
                if not _known(tok, x, allowed | rep_nums):
                    unknown.append(tok)
        if badnames:
            held.append(f"{BLOCK_KEYS[key][1]}: 'uses' names no measure on the page: {', '.join(badnames[:3])}")
        elif bad_rep:
            held.append(f"{BLOCK_KEYS[key][1]}: a reported item lacks a source, date or link")
        elif unknown:
            held.append(f"{BLOCK_KEYS[key][1]}: {', '.join(sorted(set(unknown))[:5])} not found in the data")
        elif b.get("verdict") not in ("building", "stable", "easing", "mixed"):
            held.append(f"{BLOCK_KEYS[key][1]}: verdict must be building, stable, easing or mixed")
        else:
            out_blocks[key] = b
    unknown = [tok for tok, x in _figures(rd.get("headline") or "") if not _known(tok, x, allowed_all)]
    headline = rd.get("headline") if not unknown else None
    if unknown:
        held.append(f"headline: {', '.join(sorted(set(unknown))[:5])} not found in the data")
    checks.append(("Every figure in the reading is in the data", "warn" if held else True,
                   "; ".join(held) + " - held back" if held else
                   f"{len(out_blocks)} risk readings checked against the measures"))
    return ({"pulled": rd["pulled"], "headline": headline, "blocks": out_blocks,
             "written": rd.get("written", "")}, checks)


# --------------------------------------------------------------- checks
# Run on every build. The page shows the result at the top, and the Saturday publish refuses
# to publish if anything fails. Maximum ages allow for each source's normal publication lag.

MAX_AGE = {  # days between the pull and the latest reading
    "Baltic Clean Tanker Index": 5, "Baltic Dirty Tanker Index": 5,
    "Drewry container rate, Shanghai to Rotterdam": 10, "UK wind curtailment": 65,
    "NBP gas, front month": 5, "Diesel crack against Brent": 5, "p-brief": 7,
    "Negative price hours": 35, "UK gas in storage": 5, "Gas arriving in the UK": 5,
    "LNG arriving in the UK": 5,
    "US 10-year auction cover": 45, "US 10-year dealer share": 45,
    "US 30-year auction cover": 45, "US 30-year dealer share": 45,
    "US auctions, last month: overseas-type share": 45, "US auctions, last month: dealer share": 45,
    "US yields up, dollar down": 14,   # dollar index published weekly, a few days in arrears
    # quarterly, dated to the start of the quarter and published about five weeks after it
    # ends, so the newest reading is normally four to seven months old
    "US productivity growth": 230,
    "UK day-ahead power": 35, "Diesel crack against US crude": 5, "Brent crude": 5,
    "Crude against product freight": 5,
    "Brent less WTI": 5, "Crude tanker owners (Frontline)": 5, "Crude tanker owners (DHT)": 5,
    "Oracle": 5, "CoreWeave": 5, "Nvidia": 5, "Semiconductors vs market": 5,
    "Market breadth": 5, "Sterling against the dollar": 5,
    "p-gulf": 7, "p-debt": 7, "p-ai": 7, "p-trans": 7,
    "p-fold-detail": 7, "p-fold-market": 7,
    # ONS dates each three-month window by its middle month (May-July shows as June) and
    # publishes roughly monthly; the gap between releases is sometimes five weeks (15 September
    # to 20 October 2026), when the newest reading reaches about 140 days. Raised 10 October
    # 2026 from 130 and 100, which failed in the week before a release.
    "US term premium": 21, "UK unemployment rate": 145, "UK employment rate": 145,
    "UK vacancies": 115, "US unemployment rate": 60, "US Sahm recession gauge": 60,
    "US initial jobless claims": 14,
    # Ramp AI Index: monthly, dated the first of the month, published about five to six weeks
    # after the month ends, so the newest reading is normally 40 to 75 days old
    "Businesses paying for AI": 80, "Businesses paying Anthropic": 80, "Businesses paying OpenAI": 80,
}


HELD_BACK = []
SPLIT = {}


def run_checks(payload, R, FUNDS, SCHED, sched_checks):
    out = []

    def add(name, ok, detail=""):
        out.append({"name": name, "ok": ok, "detail": detail})

    pulled = datetime.fromisoformat(payload["generated"]).date()
    errs = payload.get("errors") or {}
    add("Every data source answered", not errs,
        "; ".join(f"{k}: {v}" for k, v in list(errs.items())[:6]) if errs else
        f"{len(payload['series'])} series pulled on {pulled:%d %B %Y}")

    stale, gaps, mism, n = [], [], [], 0
    for pid, rows in R.items():
        for r in rows:
            tr = r.get("tr")
            if not tr or r.get("na"):
                continue
            n += 1
            limit = MAX_AGE.get(r["n"], MAX_AGE.get(pid, 7))
            age = (pulled - _d(tr["to"])).days
            if age > limit:
                stale.append(f"{r['n']}: latest {tr['to']}, {age} days old (limit {limit})")
            monthly_event = ("auction cover" in r["n"] or "dealer share" in r["n"]
                             or r["n"].startswith("US auctions, last month"))
            wins = (("b1", 14, 50), ("b3", 60, 120)) if monthly_event else (("b1", 20, 45), ("b3", 80, 110))
            if r["n"] == "US productivity growth":     # quarterly: one and three quarters back
                wins = (("b1", 85, 95), ("b3", 268, 278))
            if r["n"] == "UK wind curtailment" or r["n"] == "Negative price hours":
                wins = (("b1", 25, 35), ("b3", 85, 95))
            if r.get("short") and (_d(tr["to"]) - _d(r["short"]["first"][0])).days < 100:
                wins = ()                         # history still building: no 3-month base yet
            for key, lo, hi in wins:
                b = tr.get(key)
                gap = (_d(tr["to"]) - _d(b[0])).days if b else None
                if gap is None or not lo <= gap <= hi:
                    gaps.append(f"{r['n']}: {'1' if key == 'b1' else '3'}-month change measured "
                                f"from {b[0] if b else 'nothing'}")
            # the number in the Now column must be the last point of its own trace
            if r["u"] not in ("pct", "pctHigh") and r["n"] != "UK wind curtailment":
                lastv = tr["now"][1]
                if abs(r["v"] - lastv) > max(0.006, abs(lastv) * 0.006):
                    mism.append(f"{r['n']}: shows {r['v']}, series ends at {lastv}")
            if not (r["v"] == r["v"]) or r["v"] is None:
                mism.append(f"{r['n']}: no value")
    na = [f"{r['n']} ({r['na']})" for pid in ("p-gulf", "p-debt", "p-trans")
          for r in R.get(pid, []) if r.get("na")]
    berr = payload.get("brief_errors") or {}
    add("Every measure shown in your panel has this week's data", "warn" if (na or berr) else True,
        "; ".join(na + [f"{k}: {v}" for k, v in berr.items()]) if (na or berr) else
        f"{sum(len(R.get(p, [])) for p in ('p-gulf', 'p-debt', 'p-ai', 'p-trans'))} measures "
        f"in the four risk blocks")
    held = HELD_BACK
    if held:
        add("Recorded but not shown", "warn", "; ".join(
            f"{h['n']}: recording since {h['first']} ({h['count']} readings)" for h in held) +
            ". Each appears on the page once it has six months of history to be judged against.")
    for nm, dropped in CLEANED.items():
        add(f"{nm}: source errors removed", "warn" if dropped else True,
            (f"{len(dropped)} one-day readings that broke from both neighbours were left out, "
             f"latest {dropped[-1]}" if dropped else "none found"))
    for r in SPLIT.get("rows", []):
        if r["name"].startswith("UK"):
            bad = [f"{k}: other {r[k]['other']:+}bp" for k in ("3m", "12m") if k in r and abs(r[k]["other"]) > 10]
            add(f"{r['name']} yield split adds up", not bad,
                "; ".join(bad) if bad else "nominal change equals real plus expected inflation "
                "within 10bp (" + ", ".join(f"{k} other {r[k]['other']:+}bp" for k in ("3m", "12m") if k in r) + ")")
    nbp = payload["series"].get("nbp") or []
    ttf = dict(payload["series"].get("ttf") or [])
    eur = dict(payload["series"].get("eurusd") or [])
    both = [(dd, v, ttf[dd], eur[dd]) for dd, v in nbp if dd in ttf and dd in eur]
    if both:
        dd, v, t, e = both[-1]
        conv = t * e / 3.412
        ok = abs(v / conv - 1) < 0.25
        add("NBP gas is in dollars per MMBtu", ok,
            f"{dd}: NBP {v:.2f} against Dutch TTF {t:.2f} EUR/MWh = {conv:.2f} $/MMBtu "
            f"({(v / conv - 1) * 100:+.0f}%). A unit mismatch would be several times out.")
    else:
        add("NBP gas is in dollars per MMBtu", "warn", "not checked: TTF or EUR/USD prices missing")
    add("Every indicator is up to date", not stale,
        "; ".join(stale[:8]) if stale else f"{n} automated measures, each within its source's normal lag")
    add("Every 1- and 3-month change is measured from the right date", not gaps,
        "; ".join(gaps[:8]) if gaps else "all within 20-45 and 80-110 days")
    add("Every value shown matches the latest reading in its series", not mism,
        "; ".join(mism[:8]) if mism else f"{n} measures")

    stale_m = stale_manual()
    add("Typed-in company figures are recent", "warn" if stale_m else True,
        "; ".join(f"{k.replace('_', ' ')} last reported {d} ({a} days)" for k, d, a in stale_m)
        if stale_m else "all reported within 120 days")

    fbad = []
    for k, f in (FUNDS or {}).items():
        age = (pulled - _d(f["to"])).days
        if age > 5:
            fbad.append(f"{f['name']}: latest price {f['to']}, {age} days old")
        pts = f["pts"]
        jumps = [(pts[i][0], pts[i][1] / pts[i - 1][1] - 1) for i in range(1, len(pts))]
        big = [(d, j) for d, j in jumps if abs(j) > 0.12]
        if big:
            fbad.append(f"{f['name']}: {big[0][1]:+.1%} in one day on {big[0][0]}, check the source")
    add("Fund prices are current and free of bad readings", bool(FUNDS) and not fbad,
        "; ".join(fbad) if fbad else (f"{len(FUNDS)} series" if FUNDS else "no fund prices pulled"))

    # --- structure and honesty of the page itself -----------------------
    live = {pid: [r for r in rows if not r.get("na")] for pid, rows in R.items()}
    empty = [pid for pid in ("p-gulf", "p-debt", "p-ai", "p-trans") if not live.get(pid)]
    add("Every risk block has something in it", not empty,
        "; ".join(f"{p} is empty behind a live heading" for p in empty) if empty else
        ", ".join(f"{p.replace('p-', '')} {len(live.get(p, []))}" for p in
                  ("p-gulf", "p-debt", "p-ai", "p-trans")))

    ungrounded = []
    for pid, rows in R.items():
        for r in rows:
            if r.get("na"):
                continue
            chipped = r.get("st") in ("calm", "watch", "extreme")
            if chipped and r.get("norange"):
                ungrounded.append(f"{r['n']}: chip shown with no sourced range")
            if chipped and (r.get("lo") is None or r.get("hi") is None or r["lo"] == r["hi"]):
                ungrounded.append(f"{r['n']}: chip computed against {r.get('lo')}-{r.get('hi')}")
    chips = [r for p in R.values() for r in p if r.get("st") in ("calm", "watch", "extreme")]
    ruled = [r["n"] for r in chips if r.get("ruleband")]
    add("No state chip is drawn from a range that was not measured", not ungrounded,
        "; ".join(ungrounded[:6]) if ungrounded else
        f"{len(chips) - len(ruled)} chips against a recorded high and low"
        + (f"; {len(ruled)} against a published rule instead ({', '.join(ruled)})" if ruled else ""))

    aimiss = [r["n"] for r in R.get("p-ai", [])
              if not r.get("na") and not r.get("tr") and "typed in" not in (r.get("s") or "")
              and not r.get("norange")]
    add("Every AI figure says where it came from", not aimiss,
        "; ".join(aimiss) if aimiss else
        f"{len(R.get('p-ai', []))} rows: {sum(1 for r in R.get('p-ai', []) if r.get('na'))} "
        "named as not yet measured, the rest carry a source and a date")

    seen, dup = set(), []
    for pid, rows in R.items():
        for r in rows:
            if id(r) in seen:
                dup.append(r["n"])
            seen.add(id(r))
    add("No measure appears in two blocks", not dup,
        "; ".join(dup) if dup else f"{len(seen)} rows, each in one block")

    bad_legs = []
    for pid, rows in R.items():
        for r in rows:
            L, roc = r.get("legs"), (r.get("tr") or {}).get("roc")
            if not L or not roc:
                continue
            if abs((L["a"] - L["b"]) - roc["chg"]) > 0.05:
                bad_legs.append(f"{r['n']}: {L['alab']} {L['a']} less {L['blab']} {L['b']} "
                                f"does not equal the {roc['lab']} change {roc['chg']}")
    legged = [r["n"] for p in R.values() for r in p if r.get("legs")]
    add("Each spread's two legs add back to the spread", not bad_legs,
        "; ".join(bad_legs) if bad_legs else
        f"{len(legged)} spreads checked: " + ", ".join(legged))

    ctx = build_context(os.path.dirname(os.path.abspath(__file__)))
    newest = max((c.get("date", "") for c in ctx), default="")
    cage = (pulled - _d(newest)).days if newest else None
    add("Context notes are present and recent", bool(ctx) and cage is not None and cage <= 14,
        (f"{len(ctx)} notes, newest {newest} ({cage} days before the pull)" if ctx else
         "none written - the page can say what moved but nothing about why")
        + ("" if (cage is not None and cage <= 14) else
           "; add a note or the 'what was happening' section goes stale"))

    age = (datetime.now(timezone.utc).date() - pulled).days
    add("The page is current", age <= 2,
        f"data pulled {pulled:%d %B %Y}, {age} day{'' if age == 1 else 's'} ago"
        + ("" if age <= 2 else " - the local run has not refreshed, so nothing here is this week's"))

    for name, ok, det in sched_checks:
        add(name, ok, det)
    return out


# --------------------------------------------------------------- page

TEMPLATE = "dashboard_template4.html"


def fill_manual(html, generated):
    g = lambda k: MANUAL[k][0]  # noqa: E731
    dt = lambda k: datetime.fromisoformat(MANUAL[k][1]).strftime("%d %B %Y").lstrip("0")  # noqa: E731
    vals = {
        "ORACLE_X": f"{g('oracle_backlog_bn') / g('oracle_revenue_bn'):.1f}",
        "ORACLE_BL": f"{g('oracle_backlog_bn'):,.0f}", "ORACLE_REV": f"{g('oracle_revenue_bn'):,.1f}",
        "ORACLE_ONE": f"{g('oracle_one_customer_bn'):,.0f}", "ORACLE_PRIOR": f"{g('oracle_backlog_prior_bn'):,.0f}",
        "CAPEX": f"{g('hyper_capex_bn'):,.0f}", "DEPR": f"{g('hyper_depreciation_bn'):,.0f}",
        "CAPEX_GAP": f"{g('hyper_capex_bn') - g('hyper_depreciation_bn'):,.0f}",
        "AMZN": f"{g('amazon_capex_pct_ocf'):.0f}", "AI_ARR": f"{g('ai_disclosed_arr_bn'):,.0f}",
        "TOKEN_FALL": f"{g('ai_token_price_fall_pct'):.0f}",
        "D_ORACLE": dt("oracle_backlog_bn"), "D_ORACLE_PRIOR": dt("oracle_backlog_prior_bn"),
        "D_CAPEX": dt("hyper_capex_bn"), "D_AMZN": dt("amazon_capex_pct_ocf"),
        "D_TOKEN": dt("ai_token_price_fall_pct"),
    }
    for k, v in vals.items():
        html = html.replace("{" + k + "}", v)
    return html


def render(R, FUNDS, SCHED, CHECKS, payload, template_path):
    import re
    with open(template_path, encoding="utf-8") as f:
        html = f.read()
    j = lambda x: json.dumps(x, separators=(",", ":"))  # noqa: E731
    for name, val, empty in (("DATA", R, "{}"), ("FUNDS", FUNDS, "{}"), ("SCHED", SCHED, "null"),
                             ("CHECKS", CHECKS, "[]")):
        html = html.replace(f"const {name} = {empty};", f"const {name} = {j(val)};", 1)
    generated = payload["generated"]
    stamp = datetime.fromisoformat(generated).strftime("%d %B %Y").lstrip("0")
    html = fill_manual(html, generated)
    ys = payload["series"].get("spy") or []
    html = html.replace("{YAHOO_TO}", datetime.fromisoformat(ys[-1][0]).strftime("%d %B %Y").lstrip("0") if ys else "unknown")
    html = html.replace("{BUILT}", datetime.now(timezone.utc).strftime("%d %B %Y").lstrip("0"))
    html = re.sub(r'(<p class="stamp">)[^<]*(</p>)', rf'\1Data pulled {stamp}\2', html, count=1)
    return html


def main():
    d = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(d, "monitor_data.json"), encoding="utf-8") as f:
        payload = json.load(f)
    R = build_rows(payload["series"])
    brief = build_brief(payload["series"], R)
    global HELD_BACK, SPLIT
    HELD_BACK = R.pop("_held_back", [])
    SPLIT = build_split(payload["series"])
    WARN = build_warn(payload["series"])
    R = {"p-brief": brief, **R}
    R = build_blocks(payload["series"], R)
    _pulled = datetime.fromisoformat(payload["generated"]).date()
    _drows = demand_rows(d, _pulled)
    _ai = R.get("p-ai", [])
    _at = max([i for i, r in enumerate(_ai) if r.get("demand")], default=len(_ai) - 1) + 1
    R["p-ai"] = _ai[:_at] + _drows + _ai[_at:]
    FUNDS = build_funds(payload["series"].get("funds"))
    SCHED, sched_checks = build_schedule(d)
    STANDING = build_standing(d, R, payload)
    CONTEXT = build_context(d)
    DIGEST = build_digest(d, R, STANDING, CONTEXT, payload)
    READINGS, rd_checks = build_readings(d, DIGEST)
    CHECKS = run_checks(payload, R, FUNDS, SCHED, sched_checks)
    CHECKS += [{"name": n, "ok": ok, "detail": det} for n, ok, det in rd_checks]
    CHECKS += [{"name": n, "ok": ok, "detail": det} for n, ok, det in DEMAND_NOTES]
    html = render(R, FUNDS, SCHED, CHECKS, payload, os.path.join(d, TEMPLATE))
    html = html.replace("const SPLIT = null;", "const SPLIT = " + json.dumps(SPLIT, separators=(",", ":")) + ";", 1)
    html = html.replace("const WARN = null;", "const WARN = " + json.dumps(WARN, separators=(",", ":")) + ";", 1)
    html = html.replace("const STANDING = null;", "const STANDING = " + json.dumps(STANDING, separators=(",", ":")) + ";", 1)
    html = html.replace("const CONTEXT = null;", "const CONTEXT = " + json.dumps(CONTEXT, separators=(",", ":")) + ";", 1)
    html = html.replace("const READINGS = null;", "const READINGS = " + json.dumps(READINGS, separators=(",", ":")) + ";", 1)
    out = os.path.join(d, "dashboard.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    fails = [c for c in CHECKS if c["ok"] is False]
    warns = [c for c in CHECKS if c["ok"] == "warn"]
    lines = [f"Built {out}", f"{sum(len(v) for v in R.values())} rows across {len(R)} panels; "
             f"{len(FUNDS)} fund series; schedule {'read' if SCHED else 'NOT READ'}",
             f"CHECKS: {len(CHECKS) - len(fails) - len(warns)} passed, {len(warns)} to note, {len(fails)} failed"]
    for c in CHECKS:
        mark = "PASS" if c["ok"] is True else "NOTE" if c["ok"] == "warn" else "FAIL"
        lines.append(f"  {mark}  {c['name']} - {c['detail']}")
    with open(os.path.join(d, "monitor_checks.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    sys.exit(main())
