"""Leading indicator monitor - data pull (version 4).

Version 4 adds the measures from the original brief that version 3 lacked: the Baltic Clean
and Dirty Tanker Indices, NBP front-month gas, Brent (for the diesel crack), UK grid
constraint volumes (wind curtailment) and the Drewry Shanghai-Rotterdam container rate.

Fetches every series on the monitor that has a free, machine-readable source and writes
them to one file. Run it on a schedule; Claude reads the output and rebuilds the page.

Sources, all keyless and free:
  Bank of England IADB   gilt curve, real yields, implied inflation, Bank Rate, SONIA
  FRED                   US rates, breakevens, term premium, claims, unemployment, Sahm
  US Treasury Fiscal Data auction cover and bidder split, all nominal notes and bonds since 2018 (commercial reuse permitted)
  ONS                    UK unemployment, employment, vacancies
  Elexon BMRS            UK day-ahead power price and negative price hours
  Yahoo Finance          prices and equities (undocumented endpoint - see NOTE below)

NOTE on Yahoo: it carries every price series here and has no published terms or stability
guarantee. If it stops working the rest of the file still populates, and the failure is
recorded in the "errors" block rather than crashing the run.

Not fetched - no free machine-readable source. These stay hand-entered:
  UK gilt auction cover and tail (DMO report resists export)
  Drewry container index, Baltic tanker indices, NBP gas futures (licensed)
  Oracle backlog, hyperscaler capex (quarterly filings)
  Ramp AI index (published as charts)

Usage:  python monitor_pull4.py [output_dir]
Writes: monitor_data.json   full series, for rebuilding the page
        monitor_latest.csv  one row per indicator, for eyeballing

Standard library only - no pip install needed.
"""
import sys, os, json, csv, re, ssl, time
import urllib.request, urllib.error, urllib.parse
from datetime import datetime, timedelta, timezone

UA = {"User-Agent": "Mozilla/5.0 (monitor-pull)"}
TIMEOUT = 30
CTX = ssl.create_default_context()

errors = {}


def get(url, as_json=False, tries=3):
    """Fetch a URL, returning text or parsed JSON. Retries on transient failures."""
    last, attempt = None, 0
    while attempt < tries:
        attempt += 1
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=TIMEOUT, context=CTX) as r:
                raw = r.read().decode("utf-8", "replace")
            return json.loads(raw) if as_json else raw
        except Exception as e:                      # noqa: BLE001 - want every failure mode
            last = e
            if getattr(e, "code", None) == 429:     # rate-limited: back off properly, try longer
                tries = max(tries, 5)
                time.sleep(8 * attempt)
            else:
                time.sleep(1.5 * attempt)
    raise last


def record_error(key, exc):
    errors[key] = f"{type(exc).__name__}: {exc}"
    print(f"  FAILED {key}: {exc}", file=sys.stderr)


# ---------------------------------------------------------------- Bank of England

BOE = {
    "uk_bank_rate":        "IUDBEDR",
    "uk_sonia":            "IUDSOIA",
    "uk_gilt_5y":          "IUDSNPY",
    "uk_gilt_10y":         "IUDMNPY",
    "uk_gilt_20y":         "IUDLNPY",
    "uk_real_5y":          "IUDSRZC",
    "uk_real_10y":         "IUDMRZC",
    "uk_real_20y":         "IUDLRZC",
    "uk_infl_implied_5y":  "IUDSIZC",
    "uk_infl_implied_10y": "IUDMIZC",
    "uk_infl_implied_20y": "IUDLIZC",
    "uk_nom_zc_5y":        "IUDSNZC",   # nominal zero-coupon: same basis as the real and
    "uk_nom_zc_10y":       "IUDMNZC",   # inflation curves, so the yield split adds up
}
MONTHS = {m: i + 1 for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}


def pull_boe(series):
    """Bank of England Interactive Database. Daily. Returns [[iso_date, value], ...]."""
    url = ("https://www.bankofengland.co.uk/boeapps/database/_iadb-fromshowcolumns.asp"
           "?csv.x=yes&Datefrom=01/Jan/2005&Dateto=31/Dec/2035&CSVF=TN&UsingCodes=Y"
           "&VPD=Y&VFD=N&SeriesCodes=" + series)
    out = []
    for line in get(url).strip().splitlines()[1:]:
        parts = line.strip().split(",")
        if len(parts) < 2 or not parts[1]:
            continue
        try:
            d, m, y = parts[0].split()
            out.append([f"{y}-{MONTHS[m]:02d}-{int(d):02d}", float(parts[1])])
        except (ValueError, KeyError):
            continue
    return out


# ---------------------------------------------------------------- FRED

FRED = {
    "us_unemployment":   "UNRATE",
    "us_claims":         "ICSA",
    "us_sahm":           "SAHMREALTIME",
    "us_breakeven_10y":  "T10YIE",
    "us_real_10y":       "DFII10",
    "us_term_premium":   "THREEFYTP10",
    "us_treasury_10y":   "DGS10",
    "us_treasury_30y":   "DGS30",
    "us_treasury_3m":    "DTB3",
    # added 1 October 2026 for the AI financing chain: what the two credit tiers behind the
    # buildout actually pay, and whether the economy is producing the value that funds it.
    "us_bbb_yield":      "BAMLC0A4CBBBEY",
    "us_hy_yield":       "BAMLH0A0HYM2EY",
    "us_productivity":   "OPHNFB",
    # added 5 October 2026: broad trade-weighted dollar, for the count of days when US yields
    # rose and the dollar fell together (investors leaving US bonds and the currency at once).
    # Published weekly on Mondays, so the latest reading is a few days old.
    "us_dollar_broad":   "DTWEXBGS",
    # added 10 October 2026: AI demand. Share of US businesses on Ramp paying for AI, from
    # their card and bill spending (Ramp AI Index), monthly since 2023, carried by FRED.
    "ramp_ai_all":       "RAMPAIALL",
    "ramp_ai_anthropic": "RAMPAIANTHROPIC",
    "ramp_ai_openai":    "RAMPAIOPENAI",
    # added 14 October 2026: actual inflation, not only what markets expect. Pulled as the
    # change on a year earlier (FRED's pc1 transformation), as the headline figure is reported.
    # The Fed's upper target lets the build see a rate decision the day it happens.
    "us_cpi_yoy":        ("CPIAUCSL", "pc1"),
    "us_core_cpi_yoy":   ("CPILFESL", "pc1"),
    "us_fed_upper":      "DFEDTARU",
}


def pull_fred(series):
    """FRED. With FRED_API_KEY set (the cloud run), uses the FRED API, because FRED refuses
    cloud servers on its graph download. Without it (the laptop), uses the graph CSV, which
    needs no key. Missing values are '.' and are dropped either way.
    series is a FRED id, or (id, units) for a transformation such as 'pc1' (change on a
    year earlier, in percent)."""
    series, units = series if isinstance(series, tuple) else (series, None)
    key = os.environ.get("FRED_API_KEY")
    if key:
        url = (f"https://api.stlouisfed.org/fred/series/observations?series_id={series}"
               f"&api_key={key}&file_type=json&limit=100000" + (f"&units={units}" if units else ""))
        out = []
        for o in get(url, as_json=True).get("observations", []):
            if o.get("value") in (None, "", "."):
                continue
            try:
                out.append([o["date"], float(o["value"])])
            except ValueError:
                continue
        return out
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series}" + (
        f"&transformation={units}" if units else "")
    out = []
    for row in csv.reader(get(url).strip().splitlines()[1:]):
        if len(row) < 2 or row[1].strip() in ("", "."):
            continue
        try:
            out.append([row[0], float(row[1])])
        except ValueError:
            continue
    return out


# ---------------------------------------------------------------- US Treasury auctions

def pull_us_auctions(since="2018-01-01"):
    """Every nominal note and bond auction (2- to 30-year): cover ratio and bidder split.

    Inflation-linked and floating-rate issues are excluded - their yields are not comparable.
    Announced-but-unsettled auctions carry null results, so rows without a cover ratio are
    dropped rather than overwriting good data.

    Shares are of the competitive amount sold to the public (dealers + indirect + direct).
    The Treasury's published total also includes the Federal Reserve's own rollover
    (SOMA add-ons) and non-competitive bids; dividing by that total understates every share
    at auctions where the Fed rolls over maturing holdings. The published press-release
    percentages use the competitive denominator, so these now match them.
    Since 2018 so the range includes the weak auctions of December 2018 and March 2020.
    """
    base = ("https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v1"
            "/accounting/od/auctions_query")
    fields = ("auction_date,original_security_term,bid_to_cover_ratio,high_yield,"
              "primary_dealer_accepted,indirect_bidder_accepted,direct_bidder_accepted")
    url = (f"{base}?fields={fields}"
           f"&filter=security_type:in:(Note,Bond),inflation_index_security:eq:No,"
           f"floating_rate:eq:No,auction_date:gte:{since}"
           f"&sort=-auction_date&page[size]=5000")
    rows = []
    for x in get(url, as_json=True).get("data", []):
        if not x.get("bid_to_cover_ratio") or x["bid_to_cover_ratio"] == "null":
            continue
        try:
            dealer = float(x["primary_dealer_accepted"])
            ind = float(x["indirect_bidder_accepted"])
            dire = float(x["direct_bidder_accepted"])
            comp = dealer + ind + dire
            if comp < 1e9:              # odd small operations, not regular auctions
                continue
            rows.append({
                "date":   x["auction_date"],
                "term":   x["original_security_term"],
                "cover":  float(x["bid_to_cover_ratio"]),
                "yield":  float(x["high_yield"]),
                "size_bn":      round(comp / 1e9, 3),
                "dealer_pct":   round(100 * dealer / comp, 1),
                "indirect_pct": round(100 * ind / comp, 1),
                "direct_pct":   round(100 * dire / comp, 1),
            })
        except (ValueError, KeyError, ZeroDivisionError, TypeError):
            continue
    return rows


# ---------------------------------------------------------------- ONS

ONS = {
    "uk_unemployment": ("employmentandlabourmarket/peoplenotinwork/unemployment", "mgsx"),
    "uk_employment":   ("employmentandlabourmarket/peopleinwork/employmentandemployeetypes", "lf24"),
    "uk_vacancies":    ("employmentandlabourmarket/peopleinwork/employmentandemployeetypes", "ap2y"),
    # added 14 October 2026: UK consumer price inflation, annual rates, from the MM23 dataset
    "uk_cpi":          ("economy/inflationandpriceindices", "d7g7", "mm23"),
    "uk_services_cpi": ("economy/inflationandpriceindices", "d7nn", "mm23"),
}


def pull_ons(path, code, dataset="lms"):
    """ONS time series JSON. Monthly values, returned as [['YYYY-MM', value], ...]."""
    url = f"https://www.ons.gov.uk/{path}/timeseries/{code}/{dataset}/data"
    out = []
    for m in get(url, as_json=True).get("months", []):
        try:
            y, mon = m["date"].split()
            out.append([f"{y}-{MONTHS[mon.title()]:02d}", float(m["value"])])
        except (ValueError, KeyError):
            continue
    return out


# ---------------------------------------------------------------- Elexon

def pull_elexon(days=400):
    """UK day-ahead power price, aggregated by month.

    The market index endpoint caps a request at seven days, so this walks the window
    backwards a week at a time. Returns mean price and count of negative half-hours.
    """
    base = "https://data.elexon.co.uk/bmrs/api/v1/balancing/pricing/market-index"
    agg = {}
    end = datetime.now(timezone.utc).date()
    start = end - timedelta(days=days)
    d = start
    while d < end:
        a, b = d, min(d + timedelta(days=6), end)
        url = f"{base}?from={a}T00:00Z&to={b}T23:30Z"
        try:
            for x in get(url, as_json=True).get("data", []):
                if x.get("dataProvider") != "APXMIDP":
                    continue
                month = (x.get("settlementDate") or "")[:7]
                price = x.get("price")
                if not month or price is None:
                    continue
                s = agg.setdefault(month, {"n": 0, "neg": 0, "sum": 0.0, "min": None})
                s["n"] += 1
                s["sum"] += price
                if price < 0:
                    s["neg"] += 1
                if s["min"] is None or price < s["min"]:
                    s["min"] = price
        except Exception as e:                      # noqa: BLE001
            record_error(f"elexon_{a}", e)
        d = b + timedelta(days=1)
    return {m: {"periods": s["n"], "negative_hours": round(s["neg"] / 2, 1),
                "mean_price": round(s["sum"] / s["n"], 2) if s["n"] else None,
                "min_price": s["min"]}
            for m, s in sorted(agg.items())}


# ---------------------------------------------------------------- Yahoo

YAHOO = {
    "ftse100": "^FTSE", "sp500": "^GSPC", "vix": "^VIX",
    "brent": "BZ=F", "wti": "CL=F", "heating_oil": "HO=F", "gold": "GC=F",
    "hyg": "HYG", "lqd": "LQD",
    "frontline": "FRO", "dht": "DHT",
    "smh": "SMH", "spy": "SPY", "rsp": "RSP",
    "oracle": "ORCL", "coreweave": "CRWV", "nvidia": "NVDA", "microsoft": "MSFT",
    "gbpusd": "GBP=X",
    "nbp": "NBP=F",   # NYMEX UK NBP natural gas, front month (Yahoo does not state the unit)
    "ttf": "TTF=F",   # Dutch TTF gas, EUR/MWh - kept as a cross-check on NBP
    "eurusd": "EURUSD=X",  # converts TTF for that cross-check
}


# added 14 October 2026: same-day US Treasury yields (CBOE indices), used only to carry the
# FRED daily series forward past FRED's one- to two-day publication lag. Daily only.
YAHOO_SAME_DAY = {"us_treasury_10y_live": "^TNX", "us_treasury_30y_live": "^TYX"}


def pull_yahoo(ticker, rng="5y", interval="1d", period=None):
    """Closes at the given interval. Undocumented endpoint - a failure here is expected, not fatal.

    Pass period=(start_epoch, end_epoch) rather than rng for long histories. With range=max
    Yahoo silently decimates a monthly request down to roughly quarterly, which drops the
    peaks and troughs and makes any min/max over the series wrong. Explicit period bounds
    return the true month-end closes.
    """
    q = (f"period1={period[0]}&period2={period[1]}" if period else f"range={rng}")
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{urllib.parse.quote(ticker)}?{q}&interval={interval}")
    r = get(url, as_json=True)["chart"]["result"][0]
    ts = r["timestamp"]
    close = r["indicators"]["quote"][0]["close"]
    out = []
    for i, t in enumerate(ts):
        c = close[i]
        if c is None:
            continue
        out.append([datetime.fromtimestamp(t, timezone.utc).strftime("%Y-%m-%d"), round(float(c), 4)])
    return out



# ---------------------------------------------------------------- brief: shipping and grid

BROWSER_UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                            "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def get_browser(url, as_json=False):
    req = urllib.request.Request(url, headers=BROWSER_UA)
    with urllib.request.urlopen(req, timeout=TIMEOUT, context=CTX) as r:
        raw = r.read().decode("utf-8", "replace")
    return json.loads(raw) if as_json else raw


def _iso_date(s):
    """Accept an ISO date or a UK day/month/year one and return ISO."""
    s = (s or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        return "%s-%02d-%02d" % (m.group(3), int(m.group(2)), int(m.group(1)))
    return s[:10]


def append_history(path, rows):
    """Keep a dated history for sources that only publish the latest figure.

    rows: [(date, key, value)]. Existing rows for the same date and key are replaced.
    """
    hist = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            rd = csv.DictReader(f)
            cols = [c for c in (rd.fieldnames or []) if c]
            if "key" in cols and "value" in cols:
                for r in rd:
                    hist[(_iso_date(r["date"]), r["key"])] = float(r["value"])
            else:
                # A history rebuilt by hand arrives wide: date,KEY1,KEY2,... Read it
                # anyway and rewrite it long below, so a hand-edited file cannot
                # silently kill the run.
                for r in rd:
                    when = _iso_date(r[cols[0]])
                    for k in cols[1:]:
                        if r.get(k) not in (None, ""):
                            hist[(when, k)] = float(r[k])
    for d, k, v in rows:
        hist[(d, k)] = float(v)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["date", "key", "value"])
        for (d, k), v in sorted(hist.items()):
            w.writerow([d, k, v])
    out = {}
    for (d, k), v in sorted(hist.items()):
        out.setdefault(k, []).append([d, v])
    return out


def pull_baltic(out_dir):
    """Baltic Clean and Dirty Tanker Indices from the Baltic Exchange's public ticker.

    The ticker carries only today's and the previous day's figure, so each run adds them to
    baltic_history.csv and the history grows from the first run. Baltic data is for your own
    use: the Exchange's data policy does not allow republishing it.
    """
    rows = []
    for x in get_browser("https://blacksun-api.balticexchange.com/api/ticker", as_json=True):
        name = x.get("indexDataSetName")
        if name not in ("BCTI", "BDTI"):
            continue
        for when in ("previous", "current"):
            p = x.get(when) or {}
            if p.get("value") is not None and p.get("indexDate"):
                rows.append((p["indexDate"][:10], name, p["value"]))
    if not rows:
        raise ValueError("ticker returned no BCTI or BDTI figures")
    return append_history(os.path.join(out_dir, "baltic_history.csv"), rows)


def pull_drewry(out_dir):
    """Drewry World Container Index: the Shanghai to Rotterdam rate from the weekly text.

    Drewry publishes the route rate in prose each Thursday. Each run records it in
    drewry_history.csv. The full history is a paid product.
    """
    import re
    url = ("https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/"
           "world-container-index-assessed-by-drewry")
    html = get_browser(url)
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text)
    m = (re.search(r"\$([\d,]+)\s*per 40ft container from Shanghai to Rotterdam", text) or
         re.search(r"Shanghai to Rotterdam[^.$]{0,80}?\$([\d,]+)", text))
    if not m:
        raise ValueError("Shanghai to Rotterdam rate not found in the page text")
    rate = float(m.group(1).replace(",", ""))
    t = re.search(r"World Container Index\s*-\s*(\d{1,2}) (\w{3})", html)
    if not t:
        raise ValueError("assessment date not found in the page title")
    mon = MONTHS[t.group(2).title()]
    today = datetime.now(timezone.utc).date()
    year = today.year if (mon, int(t.group(1))) <= (today.month, today.day) else today.year - 1
    d = f"{year}-{mon:02d}-{int(t.group(1)):02d}"
    return append_history(os.path.join(out_dir, "drewry_history.csv"), [(d, "SHA_RTM", rate)])


def pull_neso_constraints(years=4):
    """UK grid constraint volumes from NESO: energy turned down, mostly Scottish wind.

    'Thermal constraints volume' is the balancing volume taken to relieve network limits;
    negative values are generation turned down. Summed by month, in GWh, with cost in GBP m.
    """
    base = "https://api.neso.energy/api/3/action/"
    pkg = get(base + "package_show?id=constraint-breakdown", as_json=True)["result"]
    res = sorted(pkg["resources"], key=lambda r: r["name"])[-years:]
    daily = {}
    for r in res:
        recs = get(f"{base}datastore_search?resource_id={r['id']}&limit=500",
                   as_json=True)["result"]["records"]
        for x in recs:
            try:
                daily[x["Date"][:10]] = (float(x["Thermal constraints volume"]),
                                         float(x["Thermal constraints cost"]))
            except (KeyError, TypeError, ValueError):
                continue
    months = {}
    for d, (vol, cost) in sorted(daily.items()):
        m = months.setdefault(d[:7], {"days": 0, "gwh_down": 0.0, "cost_m": 0.0, "last": d})
        m["days"] += 1
        m["gwh_down"] += max(0.0, -vol) / 1000
        m["cost_m"] += cost / 1e6
        m["last"] = d
    return {k: {kk: (round(vv, 2) if isinstance(vv, float) else vv) for kk, vv in v.items()}
            for k, v in months.items()}

# ---------------------------------------------------------------- UK gas supply and storage

NATIONAL_GAS = {
    # key: (data item, what it is). Daily energy in kWh, converted to GWh.
    "storage":        ("PUBOBJ330", "Storage, daily aggregated stock level"),
    "lng":            ("PUBOBJ633", "System entry, LNG importation"),
    "beach":          ("PUBOBJ631", "System entry, UK continental shelf and Norway (beach)"),
    "interconnector": ("PUBOBJ632", "System entry, interconnectors"),
}


def pull_national_gas(start="2019-01-01"):
    """National Gas Transmission data portal, daily, no key. Returns {key: [[date, GWh], ...]}."""
    import re
    end = datetime.now(timezone.utc).date().isoformat()
    out = {}
    for key, (item, _) in NATIONAL_GAS.items():
        url = ("https://data.nationalgas.com/api/find-gas-data-download?applicableFor=Y"
               f"&dateFrom={start}&dateTo={end}&dateType=GASDAY&latestFlag=Y&ids={item}&type=CSV")
        vals = {}
        for line in get(url).strip().splitlines()[1:]:
            m = re.match(r'^[^,]*,(\d\d)/(\d\d)/(\d{4}),".*?",([-\d.]+)', line)
            if m:
                vals[f"{m.group(3)}-{m.group(2)}-{m.group(1)}"] = round(float(m.group(4)) / 1e6, 2)
        if not vals:
            raise ValueError(f"{item} returned no rows")
        out[key] = [[d, vals[d]] for d in sorted(vals)]
    return out


# ---------------------------------------------------------------- funds

# Funds charted over six months on the page. Yahoo symbols were matched to the ISINs on the
# ii statements (the 0P... codes are Yahoo's identifiers for unit trusts). All are priced in
# sterling. One year of daily closes is pulled so the page can show six months.
FUNDS = [
    # key,        Yahoo symbol,    short name,                     group
    ("man_income", "0P000029CC.L", "Man Income",                   "held"),
    ("fid_special", "0P0000WUU1.L", "Fidelity Special Situations", "held"),
    ("artemis_inc", "0P0000KKC3.L", "Artemis Income",              "held"),
    ("ftse_all",   "^FTAS",        "FTSE All-Share",               "held_ref"),
]
# The factor and sector funds are not pulled here: the page shows the monthly schedule's own
# figures from ETF schedule.xlsx, so there is one set of basket prices, not two.


# ---------------------------------------------------------------- run

def summarise(name, series):
    """Latest value plus the comparison points the page shows."""
    if not series or not isinstance(series, list):
        return None
    vals = [v for _, v in series]
    n = len(series)
    back = lambda k: series[n - 1 - k][1] if n - 1 - k >= 0 else None  # noqa: E731
    return {"indicator": name, "date": series[-1][0], "latest": series[-1][1],
            "prev": back(1), "back_21": back(21), "back_252": back(252),
            "min": min(vals), "max": max(vals), "points": n}


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    data, summary = {}, []

    jobs = []
    jobs += [(f"boe:{k}", k, lambda s=s: pull_boe(s)) for k, s in BOE.items()]
    jobs += [(f"fred:{k}", k, lambda s=s: pull_fred(s)) for k, s in FRED.items()]
    jobs += [(f"ons:{k}", k, lambda a=a: pull_ons(*a)) for k, a in ONS.items()]
    jobs += [(f"yahoo:{k}", k, lambda t=t: pull_yahoo(t)) for k, t in YAHOO.items()]
    jobs += [(f"yahoo:{k}", k, lambda t=t: pull_yahoo(t, rng="3mo")) for k, t in YAHOO_SAME_DAY.items()]
    # Month-end closes over the full history, so the page can show where today sits in a
    # twenty-year range. The daily pull above only goes back five years.
    jobs += [(f"yahoo:{k}:monthly", f"{k}_monthly",
              lambda t=t: pull_yahoo(t, interval="1mo", period=(315532800, 2000000000)))
             for k, t in YAHOO.items()]

    for label, key, fn in jobs:
        print(f"  {label}")
        try:
            series = fn()
            data[key] = series
            s = summarise(key, series)
            if s:
                summary.append(s)
        except Exception as e:                      # noqa: BLE001
            record_error(label, e)

    funds = {}
    for key, sym, name, group in FUNDS:
        print(f"  fund:{key}")
        try:
            funds[key] = {"name": name, "group": group, "symbol": sym,
                          "series": pull_yahoo(sym, rng="1y")}
        except Exception as e:                      # noqa: BLE001
            record_error(f"fund:{key}", e)
    data["funds"] = funds

    brief_errors = {}
    # Baltic and Drewry history before the first run was rebuilt from published weekly
    # reports (sources in shipping_history_sources.csv); each run adds the latest figures.
    for label, key, fn in [("baltic:tankers", "baltic", lambda: pull_baltic(out_dir)),
                           ("drewry:wci", "drewry", lambda: pull_drewry(out_dir)),
                           ("neso:constraints", "uk_constraint", pull_neso_constraints),
                           ("nationalgas:supply", "uk_gas", pull_national_gas)]:
        print(f"  {label}")
        try:
            data[key] = fn()
        except Exception as e:                      # noqa: BLE001
            brief_errors[label] = f"{type(e).__name__}: {e}"
            print(f"  FAILED {label}: {e}", file=sys.stderr)

    for label, key, fn in [("treasury:auctions", "us_auctions", pull_us_auctions),
                           ("elexon:power", "uk_power", pull_elexon)]:
        print(f"  {label}")
        try:
            data[key] = fn()
        except Exception as e:                      # noqa: BLE001
            record_error(label, e)

    payload = {"generated": datetime.now(timezone.utc).isoformat(timespec="seconds"),
               "series": data, "errors": errors, "brief_errors": brief_errors}

    json_path = os.path.join(out_dir, "monitor_data.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, separators=(",", ":"))

    csv_path = os.path.join(out_dir, "monitor_latest.csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["indicator", "date", "latest", "prev",
                                          "back_21", "back_252", "min", "max", "points"])
        w.writeheader()
        w.writerows(sorted(summary, key=lambda r: r["indicator"]))

    print(f"\nWrote {json_path}")
    print(f"Wrote {csv_path}")
    print(f"{len(data)} series, {len(errors)} failures")
    if errors:
        print("Failures:")
        for k, v in errors.items():
            print(f"  {k}: {v}")
    return 1 if len(errors) > (len(jobs) + len(FUNDS)) // 2 else 0


if __name__ == "__main__":
    sys.exit(main())
