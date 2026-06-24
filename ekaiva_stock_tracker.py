#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EKAIVA — Mid / Small / Micro-Cap 100-EMA Trend Dashboard
========================================================
Screens the full Indian mid/small/micro-cap universe daily and flags stocks
trading ABOVE their 100-EMA on BOTH the daily and weekly timeframes.

UNIVERSE (AMFI, Option B):
  Mid-cap   = AMFI "Mid Cap"   stocks                       (overall rank 101-250)
  Small-cap = top 250 AMFI "Small Cap" stocks by avg mcap   (overall rank 251-500)
  Micro-cap = all remaining AMFI "Small Cap" stocks         (overall rank 501+)

CONDITION (the headline box):
  latest daily close  > daily 100-EMA   AND
  latest weekly close > weekly 100-EMA

Watch lists shown alongside: "Daily-100 only" and "Weekly-100 only".
Also keeps the 0-6 daily EMA score (5/10/20/50/100/200) for context.

INSTALL:  pip install yfinance nselib pandas openpyxl pyarrow
RUN:      python ekaiva_stock_tracker.py     (daily, after ~16:15 IST)
"""

import os, json, time, datetime as dt
import pandas as pd

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
EMAS              = [5, 10, 20, 50, 100, 200]
YEARS_BACK        = 5                 # 100-week EMA needs depth
WEEKS_SHOWN       = 104               # weeks in the history modal
MIN_DAILY_BARS    = 120               # skip stocks with less (can't form 100-EMA)
THIN_WEEKLY_BARS  = 150               # flag (don't drop) if fewer weekly bars (~3 yrs)
SMALLCAP_TOP_N    = 250               # smallcaps 1..250 = Small, 251+ = Micro
MICROCAP_MIN_MCAP_CR = 500            # min market cap (Rs cr) for microcaps; 0 = keep ALL, raise to drop illiquid tail
BATCH             = 50                # yfinance tickers per request (smaller = gentler on Yahoo)
SLEEP_BETWEEN     = 1.5               # seconds between batches (rate-limit safety)

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_OUT      = os.path.join(OUT_DIR, "dashboard_stocks.html")
HIST_CSV      = os.path.join(OUT_DIR, "history_weekly_stocks.csv")
QUALIFIED_CSV = os.path.join(OUT_DIR, "qualified_today.csv")
PRICE_CACHE   = os.path.join(OUT_DIR, "prices_cache.parquet")

# Universe source files. Auto-fetch is attempted; if it fails, drop these in the
# folder manually (AMFI list changes only in Jan & Jul).
AMFI_LOCAL     = os.path.join(OUT_DIR, "amfi_categorization.xlsx")
EQUITY_L_LOCAL = os.path.join(OUT_DIR, "EQUITY_L.csv")
AMFI_PAGE      = "https://www.amfiindia.com/research-information/other-data/categorization-of-stocks"
EQUITY_L_URL   = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"

# Score -> colour (same mapping as the index tracker)
SIGNAL_COLOURS = {0:"#7a0d0d",1:"#c0392b",2:"#e8604c",3:"#f39c12",4:"#f1c40f",5:"#4caf50",6:"#1f9d3a"}

# ============================================================================
# 1. UNIVERSE
# ============================================================================
def _find_amfi_file():
    if os.path.exists(AMFI_LOCAL):
        return AMFI_LOCAL
    xl = [f for f in os.listdir(OUT_DIR) if f.lower().endswith((".xlsx", ".xls"))]
    pref = [f for f in xl if "categ" in f.lower() or "amfi" in f.lower()]
    cand = pref or xl
    return os.path.join(OUT_DIR, cand[0]) if cand else None

def _col(cols, *needles, exact=None):
    if exact:
        for c in cols:
            if str(c).strip().lower() == exact:
                return c
    for c in cols:
        cl = str(c).strip().lower()
        if all(n in cl for n in needles):
            return c
    return None

def build_universe():
    """Read the AMFI Excel directly — it already carries the NSE Symbol and the
    official Large/Mid/Small label. Mid = "Mid Cap"; Small split by market cap into
    top 250 (Small) and the rest (Micro). BSE-only rows (no NSE symbol) are skipped."""
    path = _find_amfi_file()
    if path is None:
        raise FileNotFoundError(
            f"Could not find the AMFI Excel in this folder. Download it from {AMFI_PAGE}")
    print(f"  AMFI file: {os.path.basename(path)}")

    raw = pd.read_excel(path, header=None)
    hdr = None
    for i in range(min(15, len(raw))):
        row = " ".join(str(x).lower() for x in raw.iloc[i].tolist())
        if "company" in row and "isin" in row:
            hdr = i; break
    df = pd.read_excel(path, header=hdr if hdr is not None else 1)
    cols = list(df.columns)

    c_name = _col(cols, "company") or _col(cols, "name")
    c_sym  = _col(cols, exact="nse symbol") or _col(cols, "nse", "symbol")
    c_cat  = _col(cols, "categoriz") or _col(cols, "sebi") or _col(cols, "classification")
    c_mc   = (_col(cols, "average", "all") or _col(cols, "nse", "market", "cap")
              or _col(cols, "market", "cap"))
    if not (c_name and c_sym and c_cat):
        raise ValueError(f"Unexpected AMFI columns: {cols}")

    d = pd.DataFrame({
        "name":   df[c_name].astype(str).str.strip(),
        "symbol": df[c_sym].astype(str).str.strip(),
        "mcap":   pd.to_numeric(df[c_mc], errors="coerce") if c_mc else 0.0,
        "cat":    df[c_cat].astype(str).str.strip().str.lower(),
    })
    d["bucket"] = d["cat"].apply(
        lambda c: "Mid" if "mid" in c else ("Small" if "small" in c else "Large"))
    d = d[d["bucket"].isin(["Mid", "Small"])]
    d = d[~d["symbol"].str.lower().isin(["-", "", "nan", "none"])].dropna(subset=["symbol"])

    small = d[d["bucket"] == "Small"].sort_values("mcap", ascending=False)
    small_top = set(small.head(SMALLCAP_TOP_N).index)
    d["cap"] = d.apply(
        lambda r: "Midcap" if r["bucket"] == "Mid"
        else ("Smallcap" if r.name in small_top else "Microcap"), axis=1)

    if MICROCAP_MIN_MCAP_CR > 0:                       # optional dead-tail guard (0 = keep all)
        before = len(d)
        d = d[~((d["cap"] == "Microcap") & (d["mcap"] < MICROCAP_MIN_MCAP_CR))]
        print(f"  microcap min-mcap filter (>= Rs {MICROCAP_MIN_MCAP_CR} cr): dropped {before-len(d)}")

    uni = d[["symbol", "name", "cap"]].drop_duplicates("symbol").reset_index(drop=True)
    print(f"  universe: {len(uni)} NSE stocks "
          f"(Mid {sum(uni.cap=='Midcap')}, Small {sum(uni.cap=='Smallcap')}, Micro {sum(uni.cap=='Microcap')})")
    return uni


# ============================================================================
# 2. PRICES  (yfinance bulk + parquet cache; nselib fallback)
# ============================================================================
def fetch_prices(symbols, start, end):
    import yfinance as yf, logging
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)   # silence per-ticker noise
    logging.getLogger("urllib3").setLevel(logging.CRITICAL)

    cache = pd.DataFrame()
    if os.path.exists(PRICE_CACHE):
        try: cache = pd.read_parquet(PRICE_CACHE)
        except Exception: cache = pd.DataFrame()
    have = set(cache.columns)
    last = cache.index.max() if len(cache) else None
    dl_start = (last - pd.Timedelta(days=7)).date() if last is not None and have.issuperset(symbols) else start
    yf_end = (pd.Timestamp(end) + pd.Timedelta(days=1)).date()   # yfinance end is EXCLUSIVE -> +1 to include today

    def _grab(tickers, sleep):
        got = {}
        for i in range(0, len(tickers), BATCH):
            chunk = tickers[i:i+BATCH]
            try:
                data = yf.download(chunk, start=dl_start, end=yf_end, progress=False,
                                   group_by="ticker", auto_adjust=False, threads=True)
                for t in chunk:
                    try:
                        s = (data[t]["Close"].dropna() if isinstance(data.columns, pd.MultiIndex)
                             else data["Close"].dropna())
                        if len(s): got[t[:-3]] = s.rename(t[:-3])
                    except Exception: pass
            except Exception: pass
            print(f"    fetched {min(i+BATCH, len(tickers))}/{len(tickers)}  (got data for {len(got)} so far)")
            time.sleep(sleep)
        return got

    tickers = [f"{s}.NS" for s in symbols]
    got = _grab(tickers, SLEEP_BETWEEN)

    # one retry pass for whatever failed first time (usually Yahoo throttling, not real)
    missing = [f"{s}.NS" for s in symbols if s not in got]
    if missing:
        print(f"  retrying {len(missing)} stocks Yahoo skipped (pausing to avoid throttling)...")
        time.sleep(15)
        got.update(_grab(missing, SLEEP_BETWEEN * 2))

    frames = list(got.values())
    fresh = pd.concat(frames, axis=1) if frames else pd.DataFrame()
    if len(fresh): fresh.index = pd.to_datetime(fresh.index)
    merged = fresh if not len(cache) else cache.combine_first(fresh)
    if len(fresh):                                   # let fresh overwrite overlap
        merged.loc[fresh.index, fresh.columns] = fresh
    merged = merged.sort_index()
    try: merged.to_parquet(PRICE_CACHE)
    except Exception as e: print(f"    cache save skipped: {e}")
    print(f"  price data obtained for {merged.shape[1]} of {len(symbols)} stocks")
    return merged

# ============================================================================
# 3. ANALYSIS
# ============================================================================
def analyze(symbol, name, cap, closes):
    df = closes.dropna().to_frame("close")
    if len(df) < MIN_DAILY_BARS:
        return None
    for p in EMAS:
        df[f"ema{p}"] = df["close"].ewm(span=p, adjust=False).mean()
    df["score"] = sum((df["close"] > df[f"ema{p}"]).astype(int) for p in EMAS)

    # weekly view (resample includes the current partial week as last row)
    wk = df.resample("W-FRI").last().dropna(subset=["close"]).copy()
    wk["wema100"] = wk["close"].ewm(span=100, adjust=False).mean()
    wk["d100"] = wk["close"] > wk["ema100"]      # daily 100-EMA, as of week end
    wk["w100"] = wk["close"] > wk["wema100"]     # weekly 100-EMA
    wk["qual"] = wk["d100"] & wk["w100"]

    last  = df.iloc[-1]
    prev  = df.iloc[-2]
    wlast = wk.iloc[-1]
    d100  = bool(last["close"] > last["ema100"])
    w100  = bool(wlast["close"] > wlast["wema100"])
    qual_today = d100 and w100

    # --- did it become qualified TODAY (was NOT qualified as of yesterday)? ---
    d100_y = bool(prev["close"] > prev["ema100"])           # daily leg, as of yesterday
    wk_y = df["close"].iloc[:-1].resample("W-FRI").last().dropna()   # weekly, through yesterday only
    if len(wk_y) >= 1:
        wema_y = wk_y.ewm(span=100, adjust=False).mean()
        w100_y = bool(wk_y.iloc[-1] > wema_y.iloc[-1])
    else:
        w100_y = False
    qual_yesterday = d100_y and w100_y
    crossed_today = qual_today and not qual_yesterday

    hist = [{"date": idx.strftime("%d/%m/%Y"),
             "close": round(float(r["close"]), 2),
             "score": int(r["score"]),
             "qual":  bool(r["qual"])}
            for idx, r in wk.tail(WEEKS_SHOWN).iterrows()]

    return {
        "sym": symbol, "name": name, "cap": cap,
        "close": round(float(last["close"]), 2),
        "chg":   round(float((last["close"]/prev["close"]-1)*100), 2),
        "score": int(last["score"]),
        "d100":  d100, "w100": w100, "qual": qual_today,
        "crossed": crossed_today,
        "ema":   {str(p): round(float(last[f"ema{p}"]), 2) for p in EMAS},
        "thin":  len(wk) < THIN_WEEKLY_BARS,
        "date":  df.index[-1].strftime("%d/%m/%Y"),
        "hist":  hist,
    }

def update_weekly_history(records):
    """Accumulate weekly rows (retain old, refresh/extend recent) per stock."""
    store = {}                                       # (sym, iso) -> (close, score, qual)
    if os.path.exists(HIST_CSV):
        old = pd.read_csv(HIST_CSV)
        for _, r in old.iterrows():
            store[(str(r["sym"]), str(r["week_ending"]))] = (float(r["close"]), int(r["score"]), bool(r["qual"]))
    for rec in records:
        for h in rec["hist"]:
            iso = dt.datetime.strptime(h["date"], "%d/%m/%Y").strftime("%Y-%m-%d")
            store[(rec["sym"], iso)] = (h["close"], h["score"], h["qual"])
    out = pd.DataFrame(
        [{"week_ending": iso, "sym": sym, "close": c, "score": s, "qual": q}
         for (sym, iso), (c, s, q) in store.items()]
    ).sort_values(["sym", "week_ending"])
    out.to_csv(HIST_CSV, index=False)

    per = {}
    for (sym, iso), (c, s, q) in store.items():
        per.setdefault(sym, []).append((iso, c, s, q))
    for sym in per:
        rows = sorted(per[sym])[-WEEKS_SHOWN:]
        per[sym] = [{"date": dt.datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y"),
                     "close": c, "score": s, "qual": q} for iso, c, s, q in rows]
    return per

# ============================================================================
# 4. HTML
# ============================================================================
def render_html(records, data_date, run_stamp):
    return (TEMPLATE
            .replace("/*__DATA__*/null", json.dumps(records, ensure_ascii=False))
            .replace("/*__COLOURS__*/null", json.dumps(SIGNAL_COLOURS))
            .replace("__DATADATE__", data_date)
            .replace("__STAMP__", run_stamp))

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ekaiva · Mid/Small/Micro-Cap 100-EMA Tracker</title>
<style>
 :root{--orange:#E8734A;--green:#4A8C5C;--olive:#8B7355;--cream:#F5F0EB;--dark:#2D2D2D;--line:#e3ddd4;--muted:#8a8276}
 *{box-sizing:border-box;margin:0;padding:0}
 body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:var(--cream);color:var(--dark);line-height:1.45}
 .wrap{max-width:1240px;margin:0 auto;padding:0 18px 60px}
 header{background:linear-gradient(135deg,#2D2D2D,#3a3a3a);color:#fff;padding:20px 0 18px;border-bottom:4px solid var(--orange)}
 header .wrap{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}
 .brand h1{font-size:20px;font-weight:800}.brand h1 span{color:var(--orange)}
 .brand p{font-size:12.5px;color:#cfc7bb;margin-top:3px}
 .stamp{text-align:right;font-size:12px;color:#cfc7bb}.stamp b{color:#fff;display:block;font-size:14px}
 .stamp .run{font-size:11px;opacity:.8;display:block;margin-top:4px}
 .toggle{display:flex;gap:8px;margin:18px 0 4px;flex-wrap:wrap}
 .toggle button{font:inherit;font-size:13px;font-weight:700;padding:8px 16px;border:1px solid var(--line);background:#fff;color:var(--dark);border-radius:8px;cursor:pointer}
 .toggle button.on{background:var(--orange);color:#fff;border-color:var(--orange)}
 .tiles{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin:14px 0 22px}
 .tile{background:#fff;border:1px solid var(--line);border-radius:10px;padding:11px 10px;text-align:center}
 .tile .n{font-size:22px;font-weight:800;line-height:1}.tile .t{font-size:10px;color:var(--muted);margin-top:5px;font-weight:600;text-transform:uppercase;letter-spacing:.3px}
 .qbox{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;margin-bottom:16px}
 .qbox h2{font-size:15px;font-weight:800;padding:13px 16px;color:#fff;background:var(--green);display:flex;justify-content:space-between}
 .qbox h2 span{font-weight:600;opacity:.9}
 .qbody{max-height:330px;overflow-y:auto;display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:0}
 .qrow{display:flex;align-items:center;gap:10px;padding:10px 16px;border-bottom:1px solid var(--line);border-right:1px solid var(--line);cursor:pointer}
 .qrow:hover{background:#faf7f2}.qrow .nm{flex:1;min-width:0;overflow:hidden}
 .qrow .sy{font-weight:800;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
 .qrow .cn{font-size:10.5px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
 .qrow .sbar{flex-shrink:0}.qrow .track{width:74px}
 .pill{font-size:9.5px;font-weight:800;padding:2px 6px;border-radius:5px;color:#fff}
 .pill.Midcap{background:#6c5ce7}.pill.Smallcap{background:#0984e3}.pill.Microcap{background:#00897b}
 .watch{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:26px}
 .panel{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden}
 .panel h3{font-size:13.5px;font-weight:800;padding:11px 16px;color:#fff;display:flex;justify-content:space-between}
 .panel.d h3{background:#f39c12}.panel.w h3{background:#0984e3}
 .panel .body{max-height:240px;overflow-y:auto}
 .prow{display:flex;align-items:center;gap:10px;padding:9px 16px;border-bottom:1px solid var(--line);cursor:pointer;font-size:12.5px}
 .prow:hover{background:#faf7f2}.prow .sy{font-weight:700}.prow .nm{flex:1;color:var(--muted);font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
 input[type=search],select{font:inherit;font-size:13px;padding:8px 11px;border:1px solid var(--line);border-radius:8px;background:#fff;color:var(--dark)}
 input[type=search]{width:200px}
 table{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden}
 thead th{background:#f1ece4;text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);padding:10px 14px;font-weight:700}
 thead th.r{text-align:right}thead th.c{text-align:center}
 tbody td{padding:10px 14px;border-top:1px solid var(--line);font-size:13px}
 tbody tr{cursor:pointer}tbody tr:hover{background:#faf7f2}
 td.sy{font-weight:800}td.nm{color:var(--muted);font-size:11.5px;max-width:230px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
 td.num{text-align:right;font-variant-numeric:tabular-nums;font-weight:600}
 td.chg{text-align:right;font-variant-numeric:tabular-nums;font-weight:600;font-size:12px}
 td.c{text-align:center;font-weight:800}.up{color:#1f9d3a}.down{color:#c0392b}
 .yes{color:#1f9d3a}.no{color:#c0392b;opacity:.5}
 .thinflag{font-size:9px;font-weight:800;color:#b07d00;background:#fff3cd;padding:1px 5px;border-radius:4px;margin-left:6px}
 .sbar{display:flex;flex-direction:column;gap:3px}.track{height:10px;background:#e8e3da;border-radius:6px;overflow:hidden;width:110px}
 .fill{height:100%;border-radius:6px}.slabel{font-size:10.5px;font-weight:800}
 .overlay{position:fixed;inset:0;background:rgba(20,18,15,.55);display:none;align-items:flex-start;justify-content:center;padding:40px 16px;z-index:50;overflow-y:auto}
 .overlay.open{display:flex}
 .modal{background:#fff;width:100%;max-width:760px;border-radius:14px;overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,.3)}
 .modal .head{background:linear-gradient(135deg,var(--orange),#d85f37);color:#fff;padding:16px 20px;display:flex;align-items:center;justify-content:space-between}
 .modal .head h3{font-size:16px;font-weight:800}.modal .head .sub{font-size:12px;opacity:.9;margin-top:2px}
 .x{background:rgba(255,255,255,.2);border:none;color:#fff;width:30px;height:30px;border-radius:50%;font-size:17px;cursor:pointer}
 .modal .mbody{max-height:62vh;overflow-y:auto}.modal table{border:none;border-radius:0}
 .qtag{font-size:10px;font-weight:800;padding:2px 7px;border-radius:5px}.qtag.y{background:#d8f0dd;color:#1f7a33}.qtag.n{background:#f1ece4;color:var(--muted)}
 .emapanel{padding:14px 20px;border-bottom:1px solid var(--line);background:#faf7f2}
 .emapanel .ttl{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:10px}
 .emagrid{display:grid;grid-template-columns:repeat(3,1fr);gap:9px 18px}
 .emaitem{display:flex;align-items:center;gap:8px;font-size:12.5px;border-bottom:1px solid #ece6dd;padding-bottom:6px}
 .emaitem .lab{font-weight:800}.emaitem .val{font-variant-numeric:tabular-nums;color:#555;margin-left:auto}
 .abtag{font-size:9.5px;font-weight:800;padding:2px 7px;border-radius:5px}
 .abtag.a{background:#d8f0dd;color:#1f7a33}.abtag.b{background:#fbe0dc;color:#c0392b}
 .note{font-size:11px;color:var(--muted);padding:10px 20px;border-top:1px solid var(--line);background:#faf7f2}
 footer{text-align:center;font-size:11.5px;color:var(--muted);margin-top:30px;line-height:1.7}footer b{color:var(--olive)}
 @media(max-width:820px){.tiles{grid-template-columns:repeat(3,1fr)}.watch{grid-template-columns:1fr}.stamp{text-align:left}}
</style></head><body>
<header><div class="wrap">
 <div class="brand"><h1>EKAIVA <span>·</span> Mid / Small / Micro-Cap 100-EMA Tracker</h1>
 <p>Headline box = close above 100-EMA on BOTH daily &amp; weekly · 0–6 score = daily price vs 5/10/20/50/100/200 EMA</p></div>
 <div class="stamp">Market data as of<b>__DATADATE__</b><span class="run">Generated __STAMP__</span></div>
</div></header>
<div class="wrap">
 <div class="toggle" id="toggle"></div>
 <div class="tiles" id="tiles"></div>
 <div class="qbox"><h2 style="background:var(--orange)">⚡ Crossed Today · newly closed above the 100-EMA (daily &amp; weekly) today <span id="ccount"></span></h2><div class="qbody" id="cbody"></div></div>
 <div class="qbox"><h2>▲ Qualified · above 100-EMA on daily AND weekly <span id="qcount"></span></h2><div class="qbody" id="qbody"></div></div>
 <div class="watch">
  <div class="panel d"><h3>Daily-100 only <span>(daily ✓, weekly ✗ — building)</span></h3><div class="body" id="donly"></div></div>
  <div class="panel w"><h3>Weekly-100 only <span>(weekly ✓, daily ✗ — cooling)</span></h3><div class="body" id="wonly"></div></div>
 </div>
 <footer><b>Ekaiva Wealth</b> · Internal market-breadth tool · ekaivawealth.com · +91 93766 98983 · ARN 305896<br>
  &quot;Crossed Today&quot; resets and refreshes every day: stocks that were NOT above both 100-EMAs yesterday and closed above both today. Click any stock for weekly history. Not investment advice.</footer>
</div>
<div class="overlay" id="overlay"><div class="modal">
 <div class="head"><div><h3 id="mTitle">Stock</h3><div class="sub" id="mSub"></div></div><button class="x" onclick="closeModal()">✕</button></div>
 <div class="emapanel" id="emaPanel"></div>
 <div class="mbody"><table><thead><tr><th>Week ending</th><th class="r">Close</th><th>Score</th><th class="c">Qualified?</th></tr></thead>
  <tbody id="histRows"></tbody></table></div>
 <div class="note">Weekly weekend snapshot (Friday close). Qualified = above 100-EMA on daily &amp; weekly that week. History accumulates each weekend.</div>
</div></div>
<script>
const ALL=/*__DATA__*/null;
const COLOURS=/*__COLOURS__*/null;
const colorFor=s=>COLOURS[s];
const fmt=n=>n.toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
let cap='All';
const CAPS=['All','Midcap','Smallcap','Microcap'];
function view(){return cap==='All'?ALL:ALL.filter(d=>d.cap===cap);}
function scoreBar(s){return `<div class="sbar"><div class="track"><div class="fill" style="width:${s/6*100}%;background:${colorFor(s)}"></div></div><span class="slabel" style="color:${colorFor(s)}">${s}/6</span></div>`;}
function thin(d){return d.thin?'<span class="thinflag">THIN</span>':'';}
(function(){document.getElementById('toggle').innerHTML=CAPS.map(c=>`<button data-c="${c}" class="${c==='All'?'on':''}">${c==='All'?'All caps':c}</button>`).join('');
 document.querySelectorAll('#toggle button').forEach(b=>b.onclick=()=>{cap=b.dataset.c;document.querySelectorAll('#toggle button').forEach(x=>x.classList.toggle('on',x===b));renderAll();});})();
function renderTiles(){const v=view();
 const cells=[['Crossed today',v.filter(d=>d.crossed).length],['Qualified',v.filter(d=>d.qual).length],
  ['Daily-100 only',v.filter(d=>d.d100&&!d.w100).length],['Weekly-100 only',v.filter(d=>d.w100&&!d.d100).length],
  ['Score 6/6',v.filter(d=>d.score===6).length]];
 tiles.innerHTML=cells.map(([t,n])=>`<div class="tile"><div class="n">${n}</div><div class="t">${t}</div></div>`).join('');}
function qcard(d){return `<div class="qrow" onclick="openModal('${d.sym}')"><span class="pill ${d.cap}">${d.cap[0]}</span>
  <span class="nm"><div class="sy">${d.sym}${thin(d)}</div><div class="cn">${d.name}</div></span>${scoreBar(d.score)}</div>`;}
function prow(d){return `<div class="prow" onclick="openModal('${d.sym}')"><span class="sy">${d.sym}</span><span class="nm">${d.name}</span><span class="slabel" style="color:${colorFor(d.score)}">${d.score}/6</span></div>`;}
function renderBoxes(){const v=view();
 const q=v.filter(d=>d.qual).sort((a,b)=>b.score-a.score||a.sym.localeCompare(b.sym));
 qcount.textContent=q.length+' stocks';
 qbody.innerHTML=q.length?q.map(qcard).join(''):'<div class="qrow"><span class="cn">No stock qualifies right now.</span></div>';
 const don=v.filter(d=>d.d100&&!d.w100).sort((a,b)=>b.score-a.score);
 const won=v.filter(d=>d.w100&&!d.d100).sort((a,b)=>b.score-a.score);
 donly.innerHTML=don.length?don.map(prow).join(''):'<div class="prow"><span class="nm">—</span></div>';
 wonly.innerHTML=won.length?won.map(prow).join(''):'<div class="prow"><span class="nm">—</span></div>';}
function ccard(d){return `<div class="qrow" onclick="openModal('${d.sym}')"><span class="pill ${d.cap}">${d.cap[0]}</span>
  <span class="nm"><div class="sy">${d.sym}${thin(d)}</div><div class="cn">${d.name}</div></span>
  <span style="text-align:right;min-width:86px;flex-shrink:0;white-space:nowrap"><div class="sy" style="font-size:12px">${fmt(d.close)}</div>
   <div class="cn" style="color:${d.chg>=0?'#1f9d3a':'#c0392b'}">${d.chg>=0?'+':''}${d.chg.toFixed(2)}%</div></span>${scoreBar(d.score)}</div>`;}
function renderCrossed(){const v=view();
 const c=v.filter(d=>d.crossed).sort((a,b)=>b.score-a.score||a.sym.localeCompare(b.sym));
 ccount.textContent=c.length+' stock'+(c.length===1?'':'s');
 cbody.innerHTML=c.length?c.map(ccard).join(''):'<div class="qrow"><span class="cn">No new crossovers today.</span></div>';}
function renderAll(){renderTiles();renderBoxes();renderCrossed();}
function openModal(sym){const d=ALL.find(x=>x.sym===sym);if(!d)return;mTitle.textContent=d.sym+' · '+d.name;
 mSub.innerHTML=`${d.cap} · Today ${d.score}/6 · Close ${fmt(d.close)} · D&gt;100 ${d.d100?'✓':'✗'} · W&gt;100 ${d.w100?'✓':'✗'}`;
 const order=[5,10,20,50,100,200];
 emaPanel.innerHTML=`<div class="ttl">Daily EMA vs Close (${fmt(d.close)})</div><div class="emagrid">`+
   order.map(p=>{const v=d.ema[p];const ab=d.close>=v;
     return `<div class="emaitem"><span class="lab">EMA ${p}</span><span class="val">${fmt(v)}</span><span class="abtag ${ab?'a':'b'}">${ab?'Above':'Below'}</span></div>`;}).join('')+`</div>`;
 histRows.innerHTML=[...d.hist].reverse().map(h=>`<tr><td style="font-weight:600">${h.date}</td><td class="num">${fmt(h.close)}</td>
  <td>${scoreBar(h.score)}</td><td class="c"><span class="qtag ${h.qual?'y':'n'}">${h.qual?'YES':'no'}</span></td></tr>`).join('');
 overlay.classList.add('open');}
function closeModal(){overlay.classList.remove('open');}
overlay.addEventListener('click',e=>{if(e.target===overlay)closeModal();});
renderAll();
</script></body></html>"""

# ============================================================================
# 5. MAIN
# ============================================================================
def main():
    end = dt.date.today()
    start = end - dt.timedelta(days=int(YEARS_BACK*365.25))
    print(f"Ekaiva Stock 100-EMA Tracker · {end}\n")

    uni = build_universe()
    print(f"\n  fetching ~{YEARS_BACK}y prices for {len(uni)} stocks ...")
    prices = fetch_prices(list(uni["symbol"]), start, end)

    records, skipped = [], 0
    for _, row in uni.iterrows():
        if row["symbol"] not in prices.columns:
            skipped += 1; continue
        rec = analyze(row["symbol"], row["name"], row["cap"], prices[row["symbol"]])
        if rec is None: skipped += 1; continue
        records.append(rec)

    if not records:
        print("\nNo records. Check the universe files and your network.")
        return

    per = update_weekly_history(records)
    for rec in records:
        rec["hist"] = per.get(rec["sym"], rec["hist"])

    records.sort(key=lambda r: (not r["qual"], -r["score"], r["sym"]))

    qdf = pd.DataFrame([{"symbol": r["sym"], "name": r["name"], "cap": r["cap"],
                         "close": r["close"], "score": r["score"]} for r in records if r["qual"]])
    qdf.to_csv(QUALIFIED_CSV, index=False)

    data_date = max(dt.datetime.strptime(r["date"], "%d/%m/%Y") for r in records).strftime("%d %b %Y")
    run_stamp = dt.datetime.now().strftime("%d %b %Y, %H:%M")
    with open(HTML_OUT, "w", encoding="utf-8") as f:
        f.write(render_html(records, data_date, run_stamp))

    nq = sum(r["qual"] for r in records)
    print(f"\n  analysed {len(records)} stocks · {nq} qualified · {skipped} skipped/no-data")
    print(f"  dashboard      -> {HTML_OUT}")
    print(f"  qualified list -> {QUALIFIED_CSV}")
    print(f"  weekly history -> {HIST_CSV}")
    print("\nDone. Open dashboard_stocks.html.")

if __name__ == "__main__":
    main()
