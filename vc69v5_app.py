#!/usr/bin/env python3
"""
VC69V5 - desktop app (Python port of the TradingView indicator by kheirandish_v)

* A settings window with the same inputs as the indicator.
* "Check history": scans the past N days of ONE coin (BTC / PEPE / ZEC / XAU) and lists
  every signal in a table, saves an Excel file (.xlsx) and (optionally) sends a summary
  to Telegram.
* "Start live": runs the indicator on BTC, PEPE, ZEC and XAU at the same time, using the
  same settings; signals go to Telegram and the log.
* Every crypto exchange contributes two sources: spot and USDT-perpetual futures.
* 23 CFD broker feeds (PEPPERSTONE, FXPRO, OANDA, ...) come from TradingView via the
  optional tvDatafeed package. They serve BTCUSD (asset BTC) and XAUUSD (asset XAU);
  their sources are labelled BROKER-C and counted as "CFD: x/y". XAU is CFD-only.
  Anonymous feeds hold ~4900 bars per request, so history depth depends on the lower
  timeframe: ~3.4 days at 1m, ~10 days at 3m, ~17 days at 5m, ~51 days at 15m (a free
  TradingView login raises the limits). Live CFD signals work with chart timeframes
  up to ~3h at LTF 1m and ~11h at LTF 3m.
* XAU (gold) is a session market: the CFD brokers close it over the weekend, from
  Friday 17:00 New York (21:00 UTC in the US summer, 22:00 UTC in the US winter)
  until Sunday 18:00 New York (22:00 UTC summer / 23:00 UTC winter). The app knows
  the closure: XAU evaluation pauses, the log shows the exact reopen time in Iran
  time and no weekend XAU signal can be produced. The volume SMA only averages
  really traded bars - like TradingView, whose XAU chart simply has no weekend
  bars - so the first bars after the reopen are not distorted by empty weekend
  bars. BTC CFDs: measured on Saturday 2026-11-07, 18 of the 23 feeds keep
  ticking 24/7 while 5 (CXM, DARWINEX, FPMARKETS, OANDA, FX) close their
  crypto CFD at Fri 17:00 New York - the engine drops such a stale source
  automatically, so weekend BTC signals simply use the remaining feeds.
* TradingView crypto exchange feeds (also via tvDatafeed): the volume-ranked
  exchange lists measured with the TV scanner on 2026-11-08. Spot feeds serve
  BTCUSDT (18 exchanges), PEPEUSDT (15) and ZECUSDT (14); perpetual feeds serve
  BTCUSDT.P (18), PEPEUSDT.P (11) and ZECUSDT.P (21). They form their OWN source
  families (labels EXCH-TVS / EXCH-TVP, counted as "TV-Spot"/"TV-Perp"), fully
  separate from the ccxt exchange list above - an exchange that appears in both
  stays two independent sources. Every symbol keeps its own list; lists are
  never merged. The Exchange menu in the menu bar shows each symbol's ranked
  list, and selecting an asset scans ALL of its symbols automatically:
  BTC -> BTC/USDT (ccxt) + BTCUSD (CFD) + BTCUSDT (TV spot) + BTCUSDT.P (TV perp).
  TV perpetual volume is normalized per source, because every exchange reports
  contracts in its own unit (OKX != Binance != Bybit).
* Every signal reports the confirmation triple x/y/z PER SYMBOL FAMILY:
  'BTCUSDT 2/15/40' = 2 exchanges confirmed, out of 15 checked (that really
  traded on the bar), and the confirming exchanges hold 40% of the checked
  volume. Families: BTCUSDT (ccxt spot + TV spot), BTCUSDT.P (ccxt futures
  + TV perpetual), BTCUSD (the 23 CFD brokers; XAU has only the CFD family).
  The triple also appears in the Excel history file and in the Telegram text.
* Headless mode for a free 24/7 server (vc69v5_headless.py, shipped ready for
  GitHub Actions): every 15 minutes it checks ALL chart bars closed since the
  previous run and pushes every new signal to Telegram - see gh_bot/README_fa.md.
* The lower timeframe (1m/3m/5m/15m) is fetched NATIVELY from every exchange that
  serves it; exchanges without native candles of that size fall back to 1m candles
  aggregated in-app, so any exchange works either way. Native candles multiply the
  history depth (3m = 3x the days of 1m per request). Native support is taken from
  ccxt's published timeframes plus a tiny live probe for unlisted sizes. Measured
  2026-11: MEXC, Deepcoin, HTX, LBank, BigOne and Poloniex have no native 3m
  candles; all 20 exchanges serve 1m/5m/15m.
* All signal times show the START (open) of the chart bar, in Iran time (UTC+3:30)
  - matching the bar labels of TradingView.

Install:   pip install ccxt aiohttp openpyxl
Optional:  pip install tvdatafeed        (enables the 23 CFD brokers + 97 crypto
           exchange feeds; a free TradingView account can be entered in the GUI
           for higher limits)
Run:       python vc69v5_app.py
(openpyxl is only needed for the .xlsx history file - without it a CSV is saved)
"""

import asyncio
import copy
import csv
import logging
import os
import queue
import re
import threading
import time
from array import array
from bisect import bisect_left
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except Exception:  # headless / tkinter missing: core logic still importable
    tk = ttk = messagebox = None

# =====================================================================
# CONSTANTS
# =====================================================================
ASSETS = ["BTC", "PEPE", "ZEC", "XAU"]   # XAU (gold) is served only by the CFD feeds
QUOTE = "USDT"
SOURCES = ["close", "open", "high", "low", "hl2", "hlc3", "ohlc4", "hlcc4"]
LOWER_TFS = ["1m", "3m", "5m", "15m"]
MIN_CHART_TF_MIN = 3

# Live-checked with ccxt on 2026-11-09 (see ccxt_exchange_rank.xlsx):
# TOKOCRYPTO removed - its spot candles are a 1:1 BINANCE mirror and it has
#   no futures at all (it just double-weighted BINANCE);
# BIGONE removed - near-dead (PEPE 210 / ZEC 648 USDT per 5 minutes, futures
#   fetch broken in ccxt);
# CRYPTOCOM added - independent top-10 venue (BTC/PEPE spot via ccxt, no ZEC
#   spot, no USDT swaps) - contributes spot votes only.
# Regional twins that just mirror the same market (binanceusdm, myokx, okxus,
# bybitid) are covered by their parent entry below, so they are not listed
# again. DIGIFINEX stays out on request: its futures volume is base*last-
# derived and looks inflated; POLONIEX takes its place.
EXCHANGES = [
    ("BINANCE", ["binance"]),          # 1  spot + USDM futures (one id serves both)
    ("OKX", ["okx"]),                  # 2
    ("GATE", ["gate", "gateio"]),      # 3
    ("BYBIT", ["bybit"]),              # 4
    ("TOOBIT", ["toobit"]),            # 5
    ("MEXC", ["mexc"]),                # 6
    ("BITGET", ["bitget"]),            # 7
    ("WHITEBIT", ["whitebit"]),        # 8
    ("XT", ["xt"]),                    # 9
    ("DEEPCOIN", ["deepcoin"]),        # 10
    ("BINGX", ["bingx"]),              # 11
    ("BITRUE", ["bitrue"]),            # 12
    ("HTX", ["htx", "huobi"]),         # 13
    ("LBANK", ["lbank"]),              # 14
    ("ASTER", ["aster"]),              # 15
    ("KUCOIN", ["kucoin"]),            # 16 (current ccxt serves spot AND futures)
    ("BLOFIN", ["blofin"]),            # 17
    ("POLONIEX", ["poloniex"]),        # 18
    ("CRYPTOCOM", ["cryptocom"]),      # 19 added 2026-11-09 (BTC/PEPE spot)
]

# ---------------------------------------------------------------------
# TradingView CFD broker feeds (via the optional tvDatafeed package)
# ---------------------------------------------------------------------
# 23 CFD feeds that carry BOTH BTCUSD and XAUUSD, ranked by their combined
# 24h tick volume (full scan of 2026-11-06, see brokers_24h_rank.xlsx).
# ICMARKETS is NOT listed: the feed "IC" below IS IC Markets (identical data).
# CFD sources are labelled BROKER-C; tick volume is rescaled so one CFD
# source weighs about as much as one average crypto source.
TV_BROKERS = ["CXM", "DARWINEX", "FUSIONMARKETS", "FXPRO", "CAPITALCOM",
              "OPO", "IC", "PEPPERSTONE", "SKILLING", "FPMARKETS", "VANTAGE",
              "DERIV", "ERRANTE", "GOMARKETS", "BLACKBULL", "TICKMILL",
              "OANDA", "BLUEBERRY", "EIGHTCAP", "THINKMARKETS", "GBEBROKERS",
              "CFI", "FX"]
TV_SYMBOLS = {"BTC": "BTCUSD", "XAU": "XAUUSD"}   # CFD feeds serve these assets
TV_CFD_SUFFIX = "-C"     # source label suffix, e.g. PEPPERSTONE-C
TV_REFRESH_MIN = 5       # live: refresh CFD candles every N minutes
TV_MAX_BARS = 4900       # anonymous TradingView cap per request (in BARS)
TV_FETCH_DELAY = 0.3     # pause between two TradingView requests (seconds)

# ---------------------------------------------------------------------
# TradingView crypto exchange feeds (spot + USDT perpetuals, tvDatafeed)
# ---------------------------------------------------------------------
# Measured with the TradingView scanner on 2026-11-08 (tv_usdt_pairs_rank.xlsx).
# Every symbol has its OWN exchange list, ranked by 24h volume; the lists are
# never merged with each other or with the ccxt EXCHANGES list above. Spot
# volumes = 24h traded value in USDT. Perpetual volumes = the raw figures
# TradingView reports (contract units that differ per exchange) - the engine
# normalizes each of those sources individually, see rescale_tv().
TV_CRYPTO_SPOT = {
    "BTCUSDT": [
        ("WOONETWORK", 539065301),
        ("BINANCE", 320821345),
        ("WHITEBIT", 210516237),
        ("KCEX", 189871144),
        ("COINW", 185126105),
        ("POLONIEX", 172961638),
        ("PHEMEX", 160445239),
        ("BYBIT", 159343824),
        ("BLOFIN", 148022342),
        ("OKX", 133795452),
        ("GATE", 121163688),
        ("MEXC", 99593277),
        ("BITGET", 52100038),
        ("KUCOIN", 42190251),
        ("BITRUE", 41376442),
        ("BITUNIX", 39672220),
        ("BINGX", 35644288),
        ("HTX", 35595686),
    ],
    "PEPEUSDT": [
        ("BINANCE", 15210906),
        ("OKX", 8481719),
        ("BLOFIN", 6835348),
        ("KUCOIN", 6259983),
        ("WHITEBIT", 5609196),
        ("POLONIEX", 5370328),
        ("GATE", 4580006),
        ("BYBIT", 3421520),
        ("PHEMEX", 2645496),
        ("COINW", 2109017),
        ("HTX", 813588),
        ("BITGET", 769958),
        ("WOONETWORK", 764854),
        ("MEXC", 473749),
        ("BINGX", 449868),
    ],
    "ZECUSDT": [
        ("BINANCE", 110508240),
        ("POLONIEX", 38837324),
        ("OKX", 38536327),
        ("WHITEBIT", 19433517),
        ("BITGET", 17349092),
        ("KUCOIN", 14556505),
        ("PHEMEX", 6852538),
        ("BLOFIN", 4827046),
        ("BINGX", 3428820),
        ("KCEX", 2962389),
        ("WEEX", 2473993),
        ("MEXC", 2387996),
        ("COINW", 2282626),
        ("HTX", 1066651),
    ],
}
TV_CRYPTO_PERP = {
    "BTCUSDT.P": [
        ("OKX", 1205079862),
        ("GATE", 100117719),
        ("MEXC", 73486599),
        ("BLOFIN", 35108239),
        ("XTCOM", 33489189),
        ("BYDFI", 19764021),
        ("DEEPCOIN", 5606700),
        ("KUCOIN", 770136),
        ("HTX", 191670),
        ("WHITEBIT", 79120),
        ("BTCC", 78685),
        ("BINANCE", 27432),
        ("COINW", 19645),
        ("TOOBIT", 17951),
        ("KCEX", 17826),
        ("WOONETWORK", 13281),
        ("PIONEX", 12982),
        ("BYBIT", 9378),
    ],
    "PEPEUSDT.P": [
        ("WHITEBIT", 4782976000000),
        ("KCEX", 2546850800000),
        ("BITGET", 2217005681000),
        ("LBANK", 1417590329971),
        ("PIONEX", 140097609794),
        ("OKX", 78896806),
        ("KUCOIN", 23258412),
        ("BLOFIN", 3554263),
        ("MEXC", 646717),
        ("GATE", 165229),
        ("HTX", 12240),
    ],
    "ZECUSDT.P": [
        ("OKX", 772964674),
        ("MEXC", 116016105),
        ("BLOFIN", 57205593),
        ("GATE", 12767465),
        ("KUCOIN", 7218864),
        ("XTCOM", 947905),
        ("BINANCE", 938060),
        ("COINW", 853968),
        ("BYDFI", 477099),
        ("BYBIT", 174309),
        ("COINSWITCH", 174302),
        ("ZOOMEX", 173634),
        ("WOONETWORK", 148138),
        ("LBANK", 121395),
        ("HTX", 102748),
        ("KCEX", 96534),
        ("WEEX", 80695),
        ("BITGET", 78565),
        ("BINGX", 75181),
        ("WHITEBIT", 54980),
        ("BITUNIX", 48272),
    ],
}
TV_CRYPTO_SUFFIX_S = "-TVS"   # source label suffix, e.g. BINANCE-TVS
TV_CRYPTO_SUFFIX_P = "-TVP"   # source label suffix, e.g. OKX-TVP
# One asset = ALL of its symbols scanned together (never merged lists:
# each symbol below keeps the exchange list it was measured with).
TV_CRYPTO_SYMBOLS = {"BTC": ["BTCUSDT", "BTCUSDT.P"],
                     "PEPE": ["PEPEUSDT", "PEPEUSDT.P"],
                     "ZEC": ["ZECUSDT", "ZECUSDT.P"]}
TV_SCAN_DATE = "2026-11-08"   # scan date shown in the Exchange menu

# XAU weekend closure, measured on the TradingView feeds (2026-11, 7 weekends,
# 6 brokers): the gold session ends Friday 17:00 New York and reopens Sunday
# 18:00 New York. In UTC: close 21:00 (US summer) / 22:00 (US winter), open
# 22:00 (US summer) / 23:00 (US winter) - i.e. in Iran time roughly Saturday
# 00:30/01:30 until Monday 01:30/02:30. BTC CFDs: 18 of 23 feeds trade 24/7;
# CXM, DARWINEX, FPMARKETS, OANDA and FX close their crypto CFD at Fri 17:00
# New York (measured Sat 2026-11-07) - stale CFD sources are dropped
# automatically, so no weekend signal can come from them.

TZ = timezone(timedelta(hours=3, minutes=30))   # Iran time, shown in all signal times
TZ_LABEL = "UTC+3:30"
MINUTE_MS = 60_000
DAY_MS = 86_400_000
FETCH_DELAY_SEC = 5
PAGE_LIMIT = 500
MAX_CONCURRENCY = 40
REQUEST_TIMEOUT_MS = 60000
LOAD_CONCURRENCY = 5
LOAD_RETRIES = 2
ZERO = (0.0, 0.0, 0.0, 0.0)

log = logging.getLogger("vc69v5")


# =====================================================================
# SETTINGS
# =====================================================================
def parse_tf_minutes(tf):
    tf = str(tf).strip().lower()
    try:
        n, unit = int(tf[:-1]), tf[-1]
        mult = {"m": 1, "h": 60, "d": 1440}[unit]
    except (ValueError, KeyError, IndexError):
        raise ValueError(f"Bad timeframe '{tf}' (use for example 3m, 2h, 1d)")
    if n <= 0:
        raise ValueError(f"Bad timeframe '{tf}'")
    return n * mult


# ---------------------------------------------------------------------
# Per-asset / per-chart-timeframe signal gates (Telegram bot + history).
# Each entry: (min spot exchanges, min futures exchanges, min CFD exchanges,
#              min combined Weight in percent).
# A group minimum applies only when that group really traded on the bar:
# PEPE and ZEC have no CFD feeds, and if TradingView is unreachable the CFD
# gate is skipped instead of blocking BTC forever. Chart timeframes that are
# not listed fall back to the global min_exchange_count / min_confirm.
SIGNAL_GATES = {
    "BTC":  {"3m": (6, 6, 10, 40), "15m": (6, 6, 10, 40),
             "30m": (4, 4, 4, 40), "45m": (4, 4, 4, 40)},
    "PEPE": {"3m": (6, 6, 0, 40), "15m": (6, 6, 0, 40),
             "30m": (4, 4, 0, 40), "45m": (4, 4, 0, 40)},   # no CFD feeds
    "ZEC":  {"3m": (4, 4, 0, 40), "15m": (2, 2, 0, 20),
             "30m": (2, 2, 0, 20), "45m": (2, 2, 0, 20)},   # no CFD feeds
}


@dataclass
class Params:
    chart_tfs: tuple = ("2h",)
    lower_tf: str = "3m"
    source: str = "close"
    different: int = 5
    min_vol: int = 69
    vol_sma_len: int = 20
    min_confirm: float = 0.1
    min_exchange_count: int = 2
    show_global: bool = False
    use_spot: bool = True
    use_futures: bool = True
    use_cfd: bool = True
    use_tv_spot: bool = True   # TradingView crypto spot feeds (BTCUSDT/PEPEUSDT/ZECUSDT)
    use_tv_perp: bool = True   # TradingView crypto perpetual feeds (USDT.P)
    exchanges: tuple = tuple(l for l, _ in EXCHANGES)
    history_asset: str = "BTC"
    history_days: int = 30
    tg_history: bool = True
    tg_token: str = ""
    tg_chat: str = ""
    proxy: str = ""
    tv_user: str = ""                      # optional free TradingView account
    tv_pass: str = ""

    def validate(self):
        if not self.chart_tfs:
            raise ValueError("Enter at least one chart timeframe")
        lm = parse_tf_minutes(self.lower_tf)
        for tf in self.chart_tfs:
            c = parse_tf_minutes(tf)
            if c < MIN_CHART_TF_MIN:
                raise ValueError(f"Chart timeframe {tf} is below the {MIN_CHART_TF_MIN}-minute minimum")
            if c <= lm:
                raise ValueError(f"Chart timeframe {tf} must be larger than the lower timeframe {self.lower_tf}")
            if c % lm:
                raise ValueError(f"Chart timeframe {tf} must be a multiple of the lower timeframe {self.lower_tf}")
        if self.source not in SOURCES:
            raise ValueError("Unknown source")
        if self.different < 1:
            raise ValueError("Different must be >= 1")
        if self.min_vol < 55:
            raise ValueError("Min Volume must be >= 55")
        if self.vol_sma_len < 1:
            raise ValueError("Volume SMA length must be >= 1")
        if not 0 <= self.min_confirm <= 1:
            raise ValueError("Minimum weighted confirmation must be between 0 and 1")
        if self.min_exchange_count < 1:
            raise ValueError("Minimum exchange count must be >= 1")
        if not (self.use_spot or self.use_futures or self.use_cfd
                or self.use_tv_spot or self.use_tv_perp):
            raise ValueError("Select at least one of spot/futures/CFD/TV spot/TV perpetual")
        if self.history_asset == "XAU" and not self.use_cfd:
            raise ValueError("XAU history needs the CFD brokers (TradingView) enabled")
        if not self.exchanges:
            raise ValueError("Select at least one exchange")
        if self.history_asset not in ASSETS:
            raise ValueError("Unknown history coin")
        if self.history_days < 1:
            raise ValueError("History days must be >= 1")


# =====================================================================
# DATA STORE  (lower-timeframe candles in compact arrays)
# =====================================================================
class Series:
    """Closed candles of one market on one exchange, bar_min minutes each
    (the native lower timeframe, or 1m for exchanges without native candles)."""

    def __init__(self, label, symbol, factor=1.0, ex_label=None, bar_min=1):
        self.label = label                    # BINANCE-S / BINANCE-F / OANDA-C
        self.ex_label = ex_label or label
        self.symbol = symbol
        self.factor = factor                  # volume -> coin units
        self.bar_min = bar_min                # candle size in minutes (1 = aggregated 1m)
        self.ts = array("q")
        self.o, self.h, self.l, self.c, self.v = (array("d") for _ in range(5))
        self.covered_from = None              # oldest ts ever received

    def merge(self, rows, now_ms):
        for r in rows:
            try:
                ts = int(r[0])
                o, h, l, c, v = r[1], r[2], r[3], r[4], r[5]
            except (TypeError, IndexError, ValueError):
                continue
            if o is None or h is None or l is None or c is None or v is None:
                continue
            if ts + self.bar_min * MINUTE_MS > now_ms:   # candle still forming
                continue
            vals = (float(o), float(h), float(l), float(c), float(v) * self.factor)
            n = len(self.ts)
            if n == 0 or ts > self.ts[-1]:
                self.ts.append(ts)
                for arr, x in zip((self.o, self.h, self.l, self.c, self.v), vals):
                    arr.append(x)
            else:
                i = bisect_left(self.ts, ts)
                if i < n and self.ts[i] == ts:
                    for arr, x in zip((self.o, self.h, self.l, self.c, self.v), vals):
                        arr[i] = x
                else:
                    self.ts.insert(i, ts)
                    for arr, x in zip((self.o, self.h, self.l, self.c, self.v), vals):
                        arr.insert(i, x)
            if self.covered_from is None or ts < self.covered_from:
                self.covered_from = ts

    def prune(self, min_ts):
        i = bisect_left(self.ts, min_ts)
        if i > 0:
            for arr in (self.ts, self.o, self.h, self.l, self.c, self.v):
                del arr[:i]


# =====================================================================
# INDICATOR LOGIC
# =====================================================================
def src_value(mode, o, h, l, c):
    if mode == "close":
        return c
    if mode == "open":
        return o
    if mode == "high":
        return h
    if mode == "low":
        return l
    if mode == "hl2":
        return (h + l) / 2
    if mode == "hlc3":
        return (h + l + c) / 3
    if mode == "ohlc4":
        return (o + h + l + c) / 4
    return (h + l + 2 * c) / 4   # hlcc4


def lower_bars(s, lo_ms, hi_ms, L_ms, mode):
    """Aggregate the 1m candles in [lo_ms, hi_ms) into lower-timeframe bars."""
    i = bisect_left(s.ts, lo_ms)
    j = bisect_left(s.ts, hi_ms)
    bts, bsrc, bvol = [], [], []
    cur = None
    o = h = l = c = vol = 0.0
    for k in range(i, j):
        b = s.ts[k] // L_ms * L_ms
        if b != cur:
            if cur is not None:
                bts.append(cur)
                bsrc.append(src_value(mode, o, h, l, c))
                bvol.append(vol)
            cur = b
            o, h, l, vol = s.o[k], s.h[k], s.l[k], 0.0
        else:
            if s.h[k] > h:
                h = s.h[k]
            if s.l[k] < l:
                l = s.l[k]
        c = s.c[k]
        vol += s.v[k]
    if cur is not None:
        bts.append(cur)
        bsrc.append(src_value(mode, o, h, l, c))
        bvol.append(vol)
    return bts, bsrc, bvol


def chart_stats(bts, bsrc, bvol, C_ms):
    """ltfStats() summed per chart bar. Returns {chart_bar_start: [buy, sell, buyA, sellA]}.
    A chart bar that has lower-bar data but no direction flips gets an all-zero entry;
    a chart bar without ANY lower-bar data gets NO entry (key absent) - that is how
    'traded flat' is told apart from 'market closed' (e.g. the XAU weekend)."""
    out = {}
    for j in range(1, len(bts)):
        key = bts[j] // C_ms * C_ms
        st = out.get(key)
        if st is None:
            st = out[key] = [0.0, 0.0, 0.0, 0.0]
        cur, prev = bsrc[j], bsrc[j - 1]
        if cur == prev:
            continue
        v = bvol[j]
        active = v > bvol[j - 1]
        if cur > prev:
            st[0] += v
            if active:
                st[2] += v
        else:
            st[1] += v
            if active:
                st[3] += v
    return out


def hist_from_stats(stats, B, C_ms, N):
    """Stats of the last N TRADED chart bars ending at B (oldest -> newest).
    Chart bars with no entry (market closed, e.g. the XAU weekend) are skipped,
    so the window reaches further back - matching TradingView, whose chart has
    no bars at all for a closed market. Returns None while fewer than N traded
    bars exist in the available history (source still warming up)."""
    out = []
    k = 0
    max_back = N * 6                     # enough to cross long session breaks
    while len(out) < N and k < max_back:
        k += 1
        st = stats.get(B - k * C_ms)
        if st is not None:
            out.append(st)
    if len(out) < N:
        return None
    out.reverse()
    return out


def source_hist(s, B, C_ms, L_ms, p):
    """Live path: history of one source for the chart bar that closed at B (None = warming up).
    Aggregates the whole stored array so the SMA window can reach over a session
    break (weekend); hist_from_stats then picks the last N traded chart bars."""
    if s.covered_from is None:
        return None
    bts, bsrc, bvol = lower_bars(s, 0, B, L_ms, p.source)
    return hist_from_stats(chart_stats(bts, bsrc, bvol, C_ms), B, C_ms, p.vol_sma_len)


def ex_of(label):
    """BINANCE-S / BINANCE-F -> BINANCE"""
    return label.rsplit("-", 1)[0]


def _pair_kind(label):
    """Which symbol family a source belongs to:
    SPOT = ccxt spot + TV spot (BTCUSDT), PERP = ccxt futures + TV perpetual
    (BTCUSDT.P), CFD = the broker feeds (BTCUSD / XAUUSD)."""
    if label.endswith(TV_CFD_SUFFIX):
        return "CFD"
    if label.endswith((TV_CRYPTO_SUFFIX_S, "-S")):
        return "SPOT"
    return "PERP"


def pair_name(asset, kind):
    """Display symbol of a family: BTC -> BTCUSDT / BTCUSDT.P / BTCUSD."""
    if kind == "SPOT":
        return f"{asset}{QUOTE}"
    if kind == "PERP":
        return f"{asset}{QUOTE}.P"
    return TV_SYMBOLS.get(asset, f"{asset}USD")


def pairs_text(asset, info):
    """'BTCUSDT 2/15/40, BTCUSDT.P 3/20/55, BTCUSD 1/18/12' - per symbol family:
    confirm count / checked count / confirming volume in percent of checked volume."""
    if not info or not info.get("pairs"):
        return ""
    return ", ".join(f"{pair_name(asset, k)} {t[0]}/{t[1]}/{t[2]}"
                     for k, t in info["pairs"].items())


def combine(per_ex, p, asset=None, tf=None):
    """per_ex: [(label, hist)], hist = N tuples (buy, sell, buyA, sellA), newest last.
    asset/tf pick the per-asset/per-TF entry of SIGNAL_GATES (None = global only).
    A hist entry may be None (= that chart bar had no data, e.g. a CFD weekend
    bar); None bars are skipped and the volume SMA averages only traded bars."""
    cur_vol = {lab: (h[-1][0] + h[-1][1] if h[-1] is not None else 0.0)
               for lab, h in per_ex}
    total_vol = sum(v for v in cur_vol.values() if v > 0)

    buy_counter = sell_counter = 0.0
    buy_names, sell_names = [], []
    g_buy = g_sell = g_buy_a = g_sell_a = 0.0
    active = []   # sources that traded on this bar (denominator of the "x/y" counts)

    for label, h in per_ex:
        cur = h[-1]
        if cur is None:               # this source did not trade on the current bar
            continue
        b, se, ba, sa = cur
        t_vol = b + se
        if t_vol <= 0:
            continue
        active.append(label)
        t_vol_a = ba + sa
        traded = [x for x in h if x is not None]
        vol_sma = sum(x[0] + x[1] for x in traded) / len(traded)
        r_buy = round(b / t_vol, 2) * 100
        r_sell = round(se / t_vol, 2) * 100
        r_buy_a = round(ba / t_vol_a, 2) * 100 if t_vol_a > 0 else 0
        r_sell_a = round(sa / t_vol_a, 2) * 100 if t_vol_a > 0 else 0

        is_buy = r_buy > r_buy_a + p.different and r_buy > p.min_vol and t_vol > vol_sma
        is_sell = r_sell > r_sell_a + p.different and r_sell > p.min_vol and t_vol > vol_sma
        weight = t_vol / total_vol
        if is_buy:
            buy_counter += weight
            buy_names.append((label, round(weight * 100)))
        if is_sell:
            sell_counter += weight
            sell_names.append((label, round(weight * 100)))
        g_buy += b
        g_sell += se
        g_buy_a += ba
        g_sell_a += sa

    res = {"buy": None, "sell": None, "global_buy": False, "global_sell": False,
           "used": len(per_ex)}

    def info(names, counter):
        # names: [(label, weight%)] -> shown as "BINANCE-S-15, BINANCE-F-20"
        # Counts are reported separately per market type, each as "x/y/z":
        # x = exchanges of that type confirming the signal,
        # y = exchanges of that type that traded on this bar,
        # z = the confirming volume as a percentage of that type's traded volume.
        def _ex(labels):
            return len({ex_of(n) for n in labels})
        nm_s = [n for n, _ in names if n.endswith("-S")]
        nm_f = [n for n, _ in names if n.endswith("-F")]
        nm_c = [n for n, _ in names if n.endswith(TV_CFD_SUFFIX)]
        nm_tvs = [n for n, _ in names if n.endswith(TV_CRYPTO_SUFFIX_S)]
        nm_tvp = [n for n, _ in names if n.endswith(TV_CRYPTO_SUFFIX_P)]
        act_s = [n for n in active if n.endswith("-S")]
        act_f = [n for n in active if n.endswith("-F")]
        act_c = [n for n in active if n.endswith(TV_CFD_SUFFIX)]
        act_tvs = [n for n in active if n.endswith(TV_CRYPTO_SUFFIX_S)]
        act_tvp = [n for n in active if n.endswith(TV_CRYPTO_SUFFIX_P)]

        def _volp(conf, grp):
            tv = sum(cur_vol.get(l, 0.0) for l in grp)
            if tv <= 0:
                return 0
            cv = sum(cur_vol.get(l, 0.0) for l in conf)
            return int(round(cv / tv * 100))

        return {"ex_s": _ex(nm_s),
                "ex_s_total": _ex(act_s),
                "ex_f": _ex(nm_f),
                "ex_f_total": _ex(act_f),
                "ex_c": _ex(nm_c),
                "ex_c_total": _ex(act_c),
                "ex_tvs": _ex(nm_tvs),
                "ex_tvs_total": _ex(act_tvs),
                "ex_tvp": _ex(nm_tvp),
                "ex_tvp_total": _ex(act_tvp),
                "volp_s": _volp(nm_s, act_s),
                "volp_f": _volp(nm_f, act_f),
                "volp_c": _volp(nm_c, act_c),
                "volp_tvs": _volp(nm_tvs, act_tvs),
                "volp_tvp": _volp(nm_tvp, act_tvp),
                "pct": round(counter * 100),
                "names": [f"{n}-{w}" for n, w in names]}

    def pair_triples(side_names):
        """Per symbol family (SPOT / PERP / CFD) the triple (x, y, z):
        x = exchanges confirming the signal, y = exchanges checked on this bar
        (i.e. that really traded) and z = the confirming exchanges' volume as
        a percentage of that family's total volume. Example '2/15/40' =
        2 of 15 exchanges confirmed and they hold 40% of the checked volume."""
        out = {}
        for kind in ("SPOT", "PERP", "CFD"):
            act_k = [l for l in active if _pair_kind(l) == kind]
            if not act_k:
                continue                      # e.g. PEPE has no CFD feeds at all
            y_ex = len({ex_of(l) for l in act_k})
            y_vol = sum(cur_vol.get(l, 0.0) for l in act_k)
            conf_k = [n for n, _w in side_names if _pair_kind(n) == kind]
            x_ex = len({ex_of(n) for n in conf_k})
            x_vol = sum(cur_vol.get(n, 0.0) for n in conf_k)
            z = int(round(x_vol / y_vol * 100)) if y_vol > 0 else 0
            out[kind] = (x_ex, y_ex, z)
        return out

    def gate_ok(names, counter):
        """Global thresholds first, then the per-asset/per-TF SIGNAL_GATES
        entry. A group minimum is only enforced when that group actually
        traded on this bar (e.g. PEPE/ZEC have no CFD sources)."""
        if len(names) < p.min_exchange_count or counter < p.min_confirm:
            return False
        g = SIGNAL_GATES.get(asset or "", {}).get(tf or "")
        if g is None:
            return True
        min_s, min_f, min_c, min_pct = g
        ns = nf = nc = 0
        for n, _w in names:
            if n.endswith("-S"):
                ns += 1
            elif n.endswith("-F"):
                nf += 1
            elif n.endswith(TV_CFD_SUFFIX):
                nc += 1
        act_s = any(n.endswith("-S") for n in active)
        act_f = any(n.endswith("-F") for n in active)
        act_c = any(n.endswith(TV_CFD_SUFFIX) for n in active)
        if min_s > 0 and act_s and ns < min_s:
            return False
        if min_f > 0 and act_f and nf < min_f:
            return False
        if min_c > 0 and act_c and nc < min_c:
            return False
        return counter * 100 >= min_pct

    if gate_ok(buy_names, buy_counter):
        res["buy"] = info(buy_names, buy_counter)
        res["buy"]["pairs"] = pair_triples(buy_names)
    if gate_ok(sell_names, sell_counter):
        res["sell"] = info(sell_names, sell_counter)
        res["sell"]["pairs"] = pair_triples(sell_names)

    res["used"] = len(active)         # sources that really traded on this bar

    if p.show_global and total_vol > 0:
        total_hist = []
        for k in range(len(per_ex[0][1])):
            total_hist.append(sum(v for v in (h[k][0] + h[k][1]
                                              for _, h in per_ex if h[k] is not None) if v > 0))
        traded_tot = [v for v in total_hist if v > 0]
        g_sma = sum(traded_tot) / len(traded_tot) if traded_tot else 0.0
        g_vol_a = g_buy_a + g_sell_a
        rgb = round(g_buy / total_vol, 2) * 100
        rgs = round(g_sell / total_vol, 2) * 100
        rgba = round(g_buy_a / g_vol_a, 2) * 100 if g_vol_a > 0 else 0
        rgsa = round(g_sell_a / g_vol_a, 2) * 100 if g_vol_a > 0 else 0
        res["global_buy"] = rgb > rgba + p.different and rgb > p.min_vol and total_vol > g_sma
        res["global_sell"] = rgs > rgsa + p.different and rgs > p.min_vol and total_vol > g_sma
    return res


def signals_from(res):
    out = []
    if res["buy"]:
        out.append(("BUY", res["buy"]))
    if res["sell"]:
        out.append(("SELL", res["sell"]))
    if res["global_buy"]:
        out.append(("GLOBAL LONG", None))
    if res["global_sell"]:
        out.append(("GLOBAL SHORT", None))
    return out


def fmt_time(ms):
    return datetime.fromtimestamp(ms / 1000, tz=TZ).strftime("%Y-%m-%d %H:%M")


def asset_quote(asset):
    """XAU is quoted in USD on the CFD feeds; everything else trades vs USDT."""
    return "USD" if asset == "XAU" else QUOTE


# ---------------------------------------------------------------------
# XAU weekend closure (the gold session: closed from Friday 17:00 New York
# until Sunday 18:00 New York - see the comment above TV_FETCH_DELAY)
# ---------------------------------------------------------------------
def _nth_sunday_utc(year, month, n):
    """Datetime of the n-th Sunday of a month, 00:00 UTC."""
    d = datetime(year, month, 1, tzinfo=timezone.utc)
    return d + timedelta(days=(6 - d.weekday()) % 7 + 7 * (n - 1))


def us_dst_active(dt):
    """US daylight saving: 2nd Sunday of March 07:00 UTC - 1st Sunday of November 06:00 UTC."""
    start = _nth_sunday_utc(dt.year, 3, 2).replace(hour=7)
    end = _nth_sunday_utc(dt.year, 11, 1).replace(hour=6)
    return start <= dt < end


def xau_weekend_closed(now_ms):
    """Is the XAU CFD session inside the weekend closure? -> (closed, reopen_ms).
    Friday close 21:00 UTC (US summer) / 22:00 UTC (winter), Sunday open
    22:00 UTC (summer) / 23:00 UTC (winter) - measured on the broker feeds."""
    dt = datetime.fromtimestamp(now_ms / 1000, tz=timezone.utc)

    def reopen_ms(sunday):
        # the open hour follows the DST regime valid on that Sunday (12:00 UTC
        # is always past the 06:00/07:00 UTC switch and before the evening open)
        h = 22 if us_dst_active(sunday.replace(hour=12)) else 23
        return int(sunday.replace(hour=h, minute=0, second=0,
                                  microsecond=0).timestamp() * 1000)

    wd = dt.weekday()                              # Mon=0 .. Sun=6
    if wd == 5:                                    # Saturday: reopens tomorrow evening
        return True, reopen_ms(dt + timedelta(days=1))
    if wd == 6:                                    # Sunday before the 18:00 NY open
        r = reopen_ms(dt)
        if now_ms < r:
            return True, r
        return False, 0
    if wd == 4:                                    # Friday after the 17:00 NY close
        close_h = 21 if us_dst_active(dt) else 22
        if dt >= dt.replace(hour=close_h, minute=0, second=0, microsecond=0):
            return True, reopen_ms(dt + timedelta(days=2))
    return False, 0


def counts_line(info):
    """'Spot: x/y/z | Futures: x/y/z | CFD: x/y/z | TV-Spot: x/y/z | TV-Perp: x/y/z'
    - only the market types that exist. z = the confirming exchanges' volume
    as a percentage of that market type's total traded volume on the bar
    (same method as the per-family triple in pairs_text)."""
    def _t(a, b, v):
        return f"{a}/{b}/{v}"
    parts = []
    if info.get("ex_s_total"):
        parts.append(f"Spot: {_t(info['ex_s'], info['ex_s_total'], info.get('volp_s', 0))}")
    if info.get("ex_f_total"):
        parts.append(f"Futures: {_t(info['ex_f'], info['ex_f_total'], info.get('volp_f', 0))}")
    if info.get("ex_c_total"):
        parts.append(f"CFD: {_t(info['ex_c'], info['ex_c_total'], info.get('volp_c', 0))}")
    if info.get("ex_tvs_total"):
        parts.append(f"TV-Spot: {_t(info['ex_tvs'], info['ex_tvs_total'], info.get('volp_tvs', 0))}")
    if info.get("ex_tvp_total"):
        parts.append(f"TV-Perp: {_t(info['ex_tvp'], info['ex_tvp_total'], info.get('volp_tvp', 0))}")
    return " | ".join(parts)


def format_signal(side, asset, tf, ltf, B, info):
    """Telegram text, exactly the requested shape (one market type per line,
    triplet x/y/z = confirm / checked / confirming volume % of that group):
        🟢 BUY  BTC/USDT  3m / 1m
        Spot: 1/14/27
        Futu: 5/14/79
        CFD: 3/22/16
        TVS: 1/18/11
        TVF: 6/18/26
        Weight: 19%
        2026-11-09 10:18
    The bottom time is the bar START in Tehran time; no timezone label."""
    head = {"BUY": "🟢 BUY", "SELL": "🔴 SELL", "GLOBAL LONG": "🔵 GLOBAL LONG",
            "GLOBAL SHORT": "🟡 GLOBAL SHORT"}[side]
    lines = [f"{head}  {asset}/{asset_quote(asset)}  {tf} / {ltf}"]
    if info:
        for name, key in (("Spot", "s"), ("Futu", "f"), ("CFD", "c"),
                          ("TVS", "tvs"), ("TVF", "tvp")):
            tot = info.get(f"ex_{key}_total", 0)
            if tot:
                lines.append(f"{name}: {info[f'ex_{key}']}/{tot}/{info.get(f'volp_{key}', 0)} ")
        lines.append(f"Weight: {info['pct']}%")
    # B is the bar close time; display the bar START (open) time,
    # matching TradingView's bar labels
    lines.append(fmt_time(B - parse_tf_minutes(tf) * MINUTE_MS) + " ")
    return "\n".join(lines)


def compute_history(sources, asset, p, now_ms):
    """Evaluate every closed chart bar of the last p.history_days days."""
    N = p.vol_sma_len
    L_ms = parse_tf_minutes(p.lower_tf) * MINUTE_MS
    rows = []
    for tf in p.chart_tfs:
        C_ms = parse_tf_minutes(tf) * MINUTE_MS
        stats = []
        for s in sources:
            if len(s.ts) == 0:
                continue
            bts, bsrc, bvol = lower_bars(s, 0, now_ms + 1, L_ms, p.source)
            stats.append((s, chart_stats(bts, bsrc, bvol, C_ms)))
        B_last = now_ms // C_ms * C_ms
        B_first = ((now_ms - p.history_days * DAY_MS) // C_ms + 1) * C_ms
        for B in range(B_first, B_last + 1, C_ms):
            per_ex = []
            for s, st in stats:
                if s.covered_from is None:
                    continue
                hh = hist_from_stats(st, B, C_ms, N)
                if hh is not None:             # None = fewer than N traded bars yet
                    per_ex.append((s.label, hh))
            if not per_ex:
                continue
            for side, info in signals_from(combine(per_ex, p, asset, tf)):
                rows.append({
                    # display the bar START (open) time; ts stays the close time
                    "time": fmt_time(B - C_ms), "ts": B, "asset": asset, "chart_tf": tf,
                    "lower_tf": p.lower_tf, "side": side,
                    "spot_count": f"{info['ex_s']}/{info['ex_s_total']}" if info else "",
                    "futures_count": f"{info['ex_f']}/{info['ex_f_total']}" if info else "",
                    "cfd_count": f"{info['ex_c']}/{info['ex_c_total']}" if info else "",
                    "tvspot_count": f"{info['ex_tvs']}/{info['ex_tvs_total']}" if info else "",
                    "tvperp_count": f"{info['ex_tvp']}/{info['ex_tvp_total']}" if info else "",
                    "pairs": pairs_text(asset, info) if info else "",
                    "weight": info["pct"] if info else "",
                    "exchanges": ", ".join(info["names"]) if info else "",
                })
    rows.sort(key=lambda r: (r["ts"], r["chart_tf"]))
    return rows


# =====================================================================
# EXCEL / CSV OUTPUT
# =====================================================================
# (header, row key, column width). First column is 20 wide, as requested.
EXCEL_COLS = [
    ("bar start (UTC+3:30)", "time", 20),
    ("asset", "asset", 9),
    ("chart_tf", "chart_tf", 11),
    ("lower_tf", "lower_tf", 11),
    ("side", "side", 13),
    ("spot_count", "spot_count", 14),
    ("futures_count", "futures_count", 16),
    ("cfd_count", "cfd_count", 14),
    ("tvspot_count", "tvspot_count", 14),
    ("tvperp_count", "tvperp_count", 16),
    ("pairs x/y/z", "pairs", 30),
    ("weight", "weight", 10),
    ("exchanges", "exchanges", 80),
]


def _counts_to_text(v):
    """'2/11' -> '2 of 11' so Excel cannot misread it as a date (CSV fallback only)."""
    return re.sub(r"^(\d+)/(\d+)$", r"\1 of \2", str(v))


def save_excel(rows, asset, p):
    """Save the history rows as a real Excel file (.xlsx) with centered cells and
    a first column 20 wide. Values are written as text, so Excel can no longer
    turn entries like 2/11 into dates. Falls back to CSV without openpyxl."""
    base = "history_%s_%s_%s" % (asset, "-".join(p.chart_tfs),
                                 datetime.now().strftime("%Y%m%d_%H%M%S"))
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
        from openpyxl.utils import get_column_letter
    except ImportError:
        fname = os.path.abspath(base + ".csv")
        with open(fname, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([head for head, _, _ in EXCEL_COLS])
            for r in rows:
                w.writerow([_counts_to_text(r.get(k, "")) for _, k, _ in EXCEL_COLS])
        log.warning("openpyxl is not installed - saved CSV instead "
                    "(pip install openpyxl for a real Excel file): %s", fname)
        return fname
    wb = Workbook()
    ws = wb.active
    ws.title = "history"
    center = Alignment(horizontal="center", vertical="center")
    for j, (head, key, width) in enumerate(EXCEL_COLS, 1):
        c = ws.cell(row=1, column=j, value=head)
        c.font = Font(bold=True)
        c.alignment = center
        ws.column_dimensions[get_column_letter(j)].width = width
    for i, r in enumerate(rows, 2):
        for j, (_, key, _) in enumerate(EXCEL_COLS, 1):
            c = ws.cell(row=i, column=j, value=r.get(key, ""))
            c.alignment = center
    fname = os.path.abspath(base + ".xlsx")
    wb.save(fname)
    log.info("Excel saved: %s", fname)
    return fname


# =====================================================================
# EXCHANGE HELPERS
# =====================================================================
def limit_markets(ex):
    """Load spot + swap markets only (no options) and skip the slow currencies list."""
    wanted = {"spot", "swap", "linear"}
    try:
        ex.has["fetchCurrencies"] = False
    except Exception:  # noqa: BLE001
        pass
    try:
        fm = ex.options.get("fetchMarkets")
        if isinstance(fm, list):
            kept = [t for t in fm if t in wanted]
            if kept:
                ex.options["fetchMarkets"] = kept
        elif isinstance(fm, dict) and isinstance(fm.get("types"), list):
            kept = [t for t in fm["types"] if t in wanted]
            if kept:
                fm["types"] = kept
    except Exception:  # noqa: BLE001
        pass


def _set_types(ex, wanted):
    fm = ex.options.get("fetchMarkets")
    if isinstance(fm, list):
        kept = [t for t in fm if t in wanted]
        if kept:
            ex.options["fetchMarkets"] = kept
            return True
    elif isinstance(fm, dict) and isinstance(fm.get("types"), list):
        kept = [t for t in fm["types"] if t in wanted]
        if kept:
            fm["types"] = kept
            return True
    return False


async def load_markets_resilient(ex, timeout):
    """load_markets(); if one host is blocked, load spot and derivatives separately and keep what works."""
    try:
        await asyncio.wait_for(ex.load_markets(), timeout=timeout)
        return ""
    except Exception as first:  # noqa: BLE001
        first_err = first
    merged, ok_parts = {}, []
    for name, wanted in (("spot", {"spot"}), ("futures", {"swap", "linear"})):
        saved = copy.deepcopy(ex.options.get("fetchMarkets"))
        try:
            if not _set_types(ex, wanted):
                continue
            await asyncio.wait_for(ex.load_markets(True), timeout=timeout)
            merged.update(ex.markets)
            ok_parts.append(name)
        except Exception:  # noqa: BLE001
            pass
        finally:
            ex.options["fetchMarkets"] = saved
    if not ok_parts:
        raise first_err
    ex.set_markets(list(merged.values()))
    return "partial: " + " + ".join(ok_parts) + " only"


def find_markets(ex, asset, use_spot=True, use_futures=True):
    """Spot market and USDT-linear perpetual of `asset`; factor converts volume to coin units."""
    out = {"spot": None, "futures": None}
    best_mult = None
    for sym, m in ex.markets.items():
        if m.get("active") is False or m.get("quote") != QUOTE:
            continue
        base = m.get("base") or ""
        if use_spot and m.get("spot") and base == asset and out["spot"] is None:
            out["spot"] = (sym, 1.0)
        elif use_futures and m.get("swap") and m.get("linear") is not False \
                and m.get("settle") in (QUOTE, None) and base.endswith(asset):
            pre = base[: len(base) - len(asset)]
            if pre == "" or (pre.isdigit() and int(pre) > 0):
                mult = int(pre) if pre else 1
                if best_mult is None or mult < best_mult:
                    best_mult = mult
                    out["futures"] = (sym, float(m.get("contractSize") or 1) * mult)
    return out


def make_exchange(ccxt, ids, proxy):
    ex_id = next((i for i in ids if i in ccxt.exchanges), None)
    if ex_id is None:
        return None
    ex = getattr(ccxt, ex_id)({"enableRateLimit": True, "timeout": REQUEST_TIMEOUT_MS})
    if proxy:
        for attr in ("aiohttp_proxy", "https_proxy"):
            try:
                setattr(ex, attr, proxy)
            except Exception:  # noqa: BLE001
                pass
    limit_markets(ex)
    return ex


async def probe_timeframes(ex):
    """Some ccxt exchanges do not publish every candle size they really serve
    (e.g. gate serves 3m but lists none). Probe the missing sizes once with a
    tiny request and merge the result into ex.timeframes - only ADDS keys, so
    internal ccxt mappings keep their original values."""
    published = dict(getattr(ex, "timeframes", None) or {})
    want = [tf for tf in ("1m", "3m", "5m", "15m") if tf not in published]
    if not want or not ex.markets:
        return
    sym = next((s for s in ex.markets if s.startswith("BTC/USDT:USDT")), None) \
        or next((s for s in ex.markets if s.startswith("BTC/USDT")), None)
    if sym is None:
        return
    added = {}
    for tf in want:
        try:
            if await ex.fetch_ohlcv(sym, tf, limit=2):
                added[tf] = tf
        except Exception:  # noqa: BLE001
            pass
        await asyncio.sleep(0.1)
    if added:
        ex.timeframes = {**published, **added}
        log.info("  probe: %s also serves native %s", ex.id, ", ".join(sorted(added)))


async def fetch_rows(ex, symbol, since, limit, sem, tf="1m"):
    last_err = None
    for attempt in range(3):
        async with sem:
            try:
                return await ex.fetch_ohlcv(symbol, tf, since=since, limit=limit)
            except Exception as e:  # noqa: BLE001
                last_err = e
                try:
                    return await ex.fetch_ohlcv(symbol, tf, since=since)
                except Exception as e2:  # noqa: BLE001
                    last_err = e2
        await asyncio.sleep(2 * (attempt + 1))
    raise last_err


async def backfill(ex, s, needed_min, sem):
    tf = f"{s.bar_min}m"                  # native lower TF, or 1m fallback
    tf_ms = s.bar_min * MINUTE_MS
    now_ms = int(time.time() * 1000)
    since = now_ms - needed_min * MINUTE_MS
    last = None
    empties = 0
    max_pages = needed_min // (PAGE_LIMIT * s.bar_min) * 2 + 60
    try:
        for _ in range(max_pages):
            rows = await fetch_rows(ex, s.symbol, since, PAGE_LIMIT, sem, tf)
            if not rows:
                empties += 1
                if empties > 40 or since + PAGE_LIMIT * tf_ms >= now_ms:
                    break
                since += PAGE_LIMIT * tf_ms
                continue
            s.merge(rows, now_ms)
            newest = max(int(r[0]) for r in rows)
            if last is not None and newest <= last:
                break
            last = newest
            if newest + 2 * tf_ms >= now_ms:
                break
            since = newest + tf_ms
    except Exception as e:  # noqa: BLE001
        log.warning("backfill failed %s %s: %s", s.label, s.symbol, str(e)[:100])
    days = (now_ms - s.covered_from) / DAY_MS if s.covered_from else 0
    log.info("  %-14s %-22s %7d candles (%.1f days)", s.label, s.symbol, len(s.ts), days)


async def poll(ex, s, sem):
    now_ms = int(time.time() * 1000)
    since = (s.ts[-1] - s.bar_min * MINUTE_MS) if len(s.ts) else None
    try:
        rows = await fetch_rows(ex, s.symbol, since, 20, sem, f"{s.bar_min}m")
        if rows:
            s.merge(rows, now_ms)
    except Exception as e:  # noqa: BLE001
        log.debug("poll failed %s %s: %s", s.label, s.symbol, e)


async def send_telegram(p, text):
    log.info("SIGNAL\n%s", text)
    if not p.tg_token or not p.tg_chat:
        return
    import aiohttp
    url = f"https://api.telegram.org/bot{p.tg_token}/sendMessage"
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json={"chat_id": p.tg_chat, "text": text[:4000]},
                                    proxy=p.proxy or None,
                                    timeout=aiohttp.ClientTimeout(total=30)) as r:
                if r.status != 200:
                    log.warning("Telegram error %s: %s", r.status, (await r.text())[:200])
    except Exception as e:  # noqa: BLE001
        log.warning("Telegram send failed: %s", e)


# =====================================================================
# TRADINGVIEW CFD FEEDS  (optional tvDatafeed package)
# =====================================================================
def tv_available():
    try:
        import tvDatafeed  # noqa: F401
        return True
    except ImportError:
        return False


def tv_df_rows(df):
    """tvDatafeed DataFrame -> [(ts_ms, o, h, l, c, v)]; ts = bar open (UTC)."""
    rows = []
    if df is None or len(df) == 0:
        return rows
    try:
        from pandas import isna
    except ImportError:                     # pragma: no cover
        isna = lambda x: x is None or x != x  # noqa: E731
    for ts, r in df.iterrows():
        try:
            v = r["volume"]
            v = 0.0 if isna(v) else float(v)
            ts = ts.to_pydatetime()
            if ts.tzinfo is None:
                ts = ts.replace(tzinfo=timezone.utc)
            rows.append((int(ts.timestamp() * 1000), float(r["open"]), float(r["high"]),
                         float(r["low"]), float(r["close"]), v))
        except Exception:  # noqa: BLE001
            continue
    return rows


TV_INTERVALS = None          # {bar_minutes: tvDatafeed Interval}, filled lazily


def tv_interval(bar_min):
    """Map a bar size in minutes to the tvDatafeed interval (default 1m)."""
    global TV_INTERVALS
    if TV_INTERVALS is None:
        from tvDatafeed import Interval
        TV_INTERVALS = {1: Interval.in_1_minute, 3: Interval.in_3_minute,
                        5: Interval.in_5_minute, 15: Interval.in_15_minute}
    return TV_INTERVALS.get(bar_min, TV_INTERVALS[1])


def tv_fetch_all(tv, series_list, n_bars):
    """Sequentially fetch every CFD series; runs inside an executor thread.
    Returns [(series, rows)] - merging happens on the engine loop thread.
    Each series is fetched in its own bar size (the native lower TF)."""
    out = []
    for s in series_list:
        try:
            df = tv.get_hist(symbol=s.symbol, exchange=s.ex_label,
                             interval=tv_interval(s.bar_min), n_bars=n_bars)
            out.append((s, tv_df_rows(df)))
        except Exception as e:  # noqa: BLE001
            log.debug("TV fetch failed %s %s: %s", s.ex_label, s.symbol, str(e)[:80])
        time.sleep(TV_FETCH_DELAY)
    return out


def rescale_tv(lst):
    """Bring the non-ccxt source volumes into one comparable scale.
    - CFD feeds report TICK volume: one group factor scales the whole CFD family
      so one CFD source weighs about as much as one average crypto source
      (ticks are comparable among brokers).
    - TV PERPETUAL feeds report contract volume in the exchange's OWN unit
      (contract sizes differ per exchange - OKX vs Binance vs Bybit), so each of
      those sources is scaled INDIVIDUALLY to one average crypto source.
    - TV SPOT feeds report coin volume exactly like ccxt - left untouched.
    XAU has no crypto sources -> its CFD ticks stay raw (consistent)."""
    def med(s):
        # median PER-MINUTE volume, so 1m-fallback and native-3m series compare fairly
        vs = sorted(x for x in s.v if x > 0)
        return vs[len(vs) // 2] / s.bar_min if vs else 0.0

    def is_cfd(s):
        return s.label.endswith(TV_CFD_SUFFIX)

    def is_tvp(s):
        return s.label.endswith(TV_CRYPTO_SUFFIX_P)

    cm = sorted(m for m in (med(s) for s in lst if not is_cfd(s) and not is_tvp(s)) if m > 0)
    cfd = [s for s in lst if is_cfd(s)]
    tvp = [s for s in lst if is_tvp(s)]
    if not cm:
        return
    ref = cm[len(cm) // 2]
    fm = sorted(m for m in (med(s) for s in cfd) if m > 0)
    if cfd and fm:
        f = ref / fm[len(fm) // 2]
        if 0.01 < f < 100.0:
            for s in cfd:
                s.v = array("d", (x * f for x in s.v))
                s.factor = f
            log.info("CFD tick volume rescaled x%.4g (1 CFD source ~= 1 average crypto source)", f)
        else:
            log.info("CFD rescale skipped (unusual ratio %.4g)", f)
    for s in tvp:
        m = med(s)
        if m <= 0:
            continue
        f = ref / m
        if not 1e-9 < f < 1e9:
            log.info("TV perp rescale skipped for %s (unusual ratio %.4g)", s.label, f)
            continue
        s.v = array("d", (x * f for x in s.v))
        s.factor = f
    if tvp:
        log.info("TV perpetual volume normalized per source (contract units differ per exchange)")


# =====================================================================
# ENGINE  (one background asyncio loop for history + live)
# =====================================================================
class Engine:
    def __init__(self, ui):
        self.ui = ui                      # ui(kind, payload) - thread safe
        self.loop = asyncio.new_event_loop()
        self.exchanges = {}               # label -> ccxt exchange with markets loaded
        self.tv = None                    # shared TradingView session (CFD feeds)
        self.tv_lock = asyncio.Lock()
        self.proxy = None
        self.live_task = None
        self.live_stop = None
        self.hist_running = False
        self.hist_task = None
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def submit(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self.loop)

    # ---- exchanges ---------------------------------------------------
    async def ensure_exchanges(self, labels, proxy):
        import ccxt.async_support as ccxt
        if self.proxy is not None and proxy != self.proxy:
            await self.close_exchanges()
        self.proxy = proxy
        todo = [(l, ids) for l, ids in EXCHANGES if l in labels and l not in self.exchanges]
        if not todo:
            return
        log.info("Loading markets of %d exchanges (first time takes a few minutes)...", len(todo))
        sem = asyncio.Semaphore(LOAD_CONCURRENCY)

        async def load(label, ids):
            ex = make_exchange(ccxt, ids, proxy)
            if ex is None:
                log.warning("%-10s not supported by this ccxt version - skipped", label)
                return
            try:
                async with sem:
                    t0 = time.time()
                    err = None
                    for _ in range(LOAD_RETRIES):
                        try:
                            note = await load_markets_resilient(ex, REQUEST_TIMEOUT_MS / 1000 * 4)
                            self.exchanges[label] = ex
                            log.info("%-10s markets loaded (%.0fs)%s", label, time.time() - t0,
                                     f"  [{note}]" if note else "")
                            try:
                                await probe_timeframes(ex)
                            except Exception:  # noqa: BLE001
                                pass          # probe is best-effort; 1m fallback stays safe
                            return
                        except Exception as e:  # noqa: BLE001
                            err = e
                    log.warning("%-10s load failed (%s) - skipped", label, str(err)[:100])
                    await ex.close()
            except asyncio.CancelledError:          # history stopped while loading
                try:
                    await ex.close()
                except Exception:  # noqa: BLE001
                    pass
                raise

        await asyncio.gather(*[load(l, ids) for l, ids in todo])

    async def close_exchanges(self):
        for ex in list(self.exchanges.values()):
            try:
                await ex.close()
            except Exception:  # noqa: BLE001
                pass
        self.exchanges.clear()

    # ---- TradingView CFD feeds ----------------------------------------
    async def ensure_tv(self, p):
        """One shared TradingView session, created lazily and reused."""
        async with self.tv_lock:
            if self.tv is not None or not (p.use_cfd or p.use_tv_spot
                                           or p.use_tv_perp) or not tv_available():
                return self.tv

            def _make():
                from io import StringIO
                from contextlib import redirect_stdout
                from tvDatafeed import TvDatafeed
                logging.getLogger().setLevel(logging.ERROR)
                buf = StringIO()
                with redirect_stdout(buf):
                    if p.tv_user and p.tv_pass:
                        return TvDatafeed(username=p.tv_user, password=p.tv_pass)
                    return TvDatafeed()

            try:
                self.tv = await self.loop.run_in_executor(None, _make)
                log.info("TradingView session ready (%s)",
                         "logged in" if p.tv_user and p.tv_pass else "anonymous")
            except Exception as e:  # noqa: BLE001
                log.warning("TradingView connection failed - CFD feeds disabled (%s)",
                            str(e)[:100])
                self.tv = None
            return self.tv

    async def tv_fill(self, tv_list, p, n_bars):
        """Backfill ALL TradingView series (CFD brokers + crypto spot/perp feeds)
        from TradingView (anonymous cap ~4900 bars)."""
        if not tv_list:
            return
        tv = await self.ensure_tv(p)
        if tv is None:
            return
        n = min(n_bars, TV_MAX_BARS)
        log.info("Backfilling %d TradingView sources (%d %s bars each)...",
                 len(tv_list), n, f"{tv_list[0].bar_min}m")
        async with self.tv_lock:
            fetched = await self.loop.run_in_executor(None, tv_fetch_all, tv, tv_list, n)
        now_ms = int(time.time() * 1000)
        got = 0
        for s, rows in fetched:
            if rows:
                s.merge(rows, now_ms)
                got += 1
        log.info("TradingView backfill done: %d/%d sources have data", got, len(tv_list))

    async def tv_refresh(self, tv_list, n_bars=40):
        """Refresh TradingView candles during live (every TV_REFRESH_MIN minutes)."""
        if self.tv is None or not tv_list:
            return
        async with self.tv_lock:
            fetched = await self.loop.run_in_executor(None, tv_fetch_all,
                                                      self.tv, tv_list, n_bars)
        now_ms = int(time.time() * 1000)
        for s, rows in fetched:
            if rows:
                s.merge(rows, now_ms)

    def build_sources(self, assets, p):
        out = {}
        L_min = parse_tf_minutes(p.lower_tf)
        if (p.use_cfd or p.use_tv_spot or p.use_tv_perp) and not tv_available():
            log.warning("tvdatafeed is not installed - the TradingView feeds "
                        "(%d CFD brokers + %d crypto exchange feeds) are skipped "
                        "(pip install tvdatafeed)", len(TV_BROKERS),
                        sum(len(v) for v in TV_CRYPTO_SPOT.values())
                        + sum(len(v) for v in TV_CRYPTO_PERP.values()))
        for asset in assets:
            lst = []
            no_native = []
            for label in p.exchanges:
                ex = self.exchanges.get(label)
                if ex is None or not ex.has.get("fetchOHLCV"):
                    continue
                if L_min > 1:
                    native = p.lower_tf in (getattr(ex, "timeframes", None) or {})
                else:
                    native = True
                if not native:
                    no_native.append(label)
                f = find_markets(ex, asset, p.use_spot, p.use_futures)
                if f["spot"]:
                    lst.append(Series(label + "-S", f["spot"][0], f["spot"][1], label,
                                      bar_min=L_min if native else 1))
                if f["futures"]:
                    lst.append(Series(label + "-F", f["futures"][0], f["futures"][1], label,
                                      bar_min=L_min if native else 1))
            if p.use_cfd and asset in TV_SYMBOLS and tv_available():
                for broker in TV_BROKERS:
                    lst.append(Series(broker + TV_CFD_SUFFIX, TV_SYMBOLS[asset], 1.0,
                                      broker, bar_min=L_min))
            if tv_available():
                # TradingView crypto exchange feeds: every symbol keeps its own
                # volume-ranked exchange list (never merged, separate from ccxt)
                for sym in TV_CRYPTO_SYMBOLS.get(asset, []):
                    if sym.endswith(".P"):
                        if not p.use_tv_perp or sym not in TV_CRYPTO_PERP:
                            continue
                        feeds, suffix = TV_CRYPTO_PERP[sym], TV_CRYPTO_SUFFIX_P
                    else:
                        if not p.use_tv_spot or sym not in TV_CRYPTO_SPOT:
                            continue
                        feeds, suffix = TV_CRYPTO_SPOT[sym], TV_CRYPTO_SUFFIX_S
                    for exch, _vol in feeds:
                        lst.append(Series(exch + suffix, sym, 1.0, exch, bar_min=L_min))
            out[asset] = lst
            if no_native:
                log.info("%s: no native %s candles on %s - those sources use "
                         "1m aggregation instead", asset, p.lower_tf,
                         ", ".join(no_native))
            n_s = sum(1 for x in lst if x.label.endswith("-S"))
            n_tvs = sum(1 for x in lst if x.label.endswith(TV_CRYPTO_SUFFIX_S))
            n_tvp = sum(1 for x in lst if x.label.endswith(TV_CRYPTO_SUFFIX_P))
            n_c = sum(1 for x in lst if x.label.endswith(TV_CFD_SUFFIX))
            n_native = sum(1 for x in lst if x.bar_min > 1)
            log.info("%s/%s: %d sources (spot %d, futures %d, CFD %d, TV-spot %d, "
                     "TV-perp %d; %d in native %s)",
                     asset, asset_quote(asset), len(lst), n_s,
                     len(lst) - n_s - n_c - n_tvs - n_tvp, n_c, n_tvs, n_tvp,
                     n_native, p.lower_tf)
            for x in lst:
                log.info("    %-14s %-22s %-4s volume factor %g", x.label, x.symbol,
                         f"{x.bar_min}m", x.factor)
        return out

    # ---- history -----------------------------------------------------
    def start_history(self, p):
        if self.hist_running:
            return False
        self.hist_running = True
        self.hist_task = self.submit(self._history_wrapper(p))
        return True

    def stop_history(self):
        """Cancel the running history scan (safe from any thread)."""
        f = self.hist_task
        if f is not None and not f.done():
            f.cancel()          # cancels the asyncio task on the engine loop
            return True
        return False

    async def _history_wrapper(self, p):
        self.ui("hist_state", True)
        try:
            await self.run_history(p)
        except asyncio.CancelledError:
            log.info("=== HISTORY stopped by user ===")
        except Exception as e:  # noqa: BLE001
            log.exception("history failed: %s", e)
        finally:
            self.hist_running = False
            self.ui("hist_state", False)

    async def run_history(self, p):
        asset = p.history_asset
        log.info("=== HISTORY %s | chart %s | LTF %s | last %d days ===",
                 asset, ",".join(p.chart_tfs), p.lower_tf, p.history_days)
        await self.ensure_exchanges(p.exchanges, p.proxy)
        lst = self.build_sources([asset], p)[asset]
        if not lst:
            log.error("No sources found for %s", asset)
            return
        L_min = parse_tf_minutes(p.lower_tf)
        c_max = max(parse_tf_minutes(t) for t in p.chart_tfs)
        needed_min = p.history_days * 1440 + (p.vol_sma_len + 1) * c_max + L_min + 5
        log.info("Downloading %s history (%d minutes per source; sources without "
                 "native %s candles use 1m aggregation)...", p.lower_tf, needed_min,
                 p.lower_tf)
        sem = asyncio.Semaphore(MAX_CONCURRENCY)
        ccxt_lst = [s for s in lst if s.label.endswith(("-S", "-F"))]
        tv_lst = [s for s in lst if s.label.endswith((TV_CFD_SUFFIX, TV_CRYPTO_SUFFIX_S,
                                                     TV_CRYPTO_SUFFIX_P))]
        if ccxt_lst:
            await asyncio.gather(*[backfill(self.exchanges[s.ex_label], s, needed_min, sem)
                                   for s in ccxt_lst])
        if tv_lst:
            # the anonymous cap counts BARS: 4900 x lower-TF bars per request
            await self.tv_fill(tv_lst, p, -(-needed_min // L_min))
            if not ccxt_lst:
                days = TV_MAX_BARS * L_min / 1440
                log.info("Note: anonymous TradingView feeds hold only ~%d %s bars - "
                         "history for %s covers the last ~%.1f days.", TV_MAX_BARS,
                         p.lower_tf, asset, days)
        rescale_tv(lst)
        now_ms = int(time.time() * 1000)
        log.info("Computing signals...")
        rows = await self.loop.run_in_executor(None, compute_history, lst, asset, p, now_ms)
        n_buy = sum(1 for r in rows if r["side"] == "BUY")
        n_sell = sum(1 for r in rows if r["side"] == "SELL")
        n_glob = len(rows) - n_buy - n_sell
        log.info("History done: %d signals (BUY %d, SELL %d, global %d)", len(rows), n_buy, n_sell, n_glob)
        fname = save_excel(rows, asset, p)
        self.ui("hist_rows", rows)
        if p.tg_history:
            lines = [f"📊 VC69V5 history | {asset}/{asset_quote(asset)}",
                     f"chart {','.join(p.chart_tfs)} / LTF {p.lower_tf} | last {p.history_days} days | times {TZ_LABEL}",
                     f"BUY {n_buy} | SELL {n_sell} | global {n_glob}"]
            if rows:
                lines.append("Latest signals:")
                for r in rows[-10:][::-1]:
                    parts = []
                    if r["spot_count"]:
                        parts.append(f"spot {r['spot_count']}")
                    if r["futures_count"]:
                        parts.append(f"fut {r['futures_count']}")
                    if r["cfd_count"]:
                        parts.append(f"cfd {r['cfd_count']}")
                    if r["tvspot_count"]:
                        parts.append(f"tvspot {r['tvspot_count']}")
                    if r["tvperp_count"]:
                        parts.append(f"tvperp {r['tvperp_count']}")
                    if r["weight"] != "":
                        parts.append(f"{r['weight']}%")
                    w_txt = (" " + " | ".join(parts)) if parts else ""
                    lines.append(f"{r['time']} {r['chart_tf']} {r['side']}{w_txt}")
            lines.append(f"Excel: {os.path.basename(fname)}")
            await send_telegram(p, "\n".join(lines))

    # ---- live --------------------------------------------------------
    def start_live(self, p):
        if self.live_task is not None and not self.live_task.done():
            return False
        self.submit(self._start_live(p))
        return True

    async def _start_live(self, p):
        self.live_stop = asyncio.Event()
        self.live_task = asyncio.create_task(self._live_wrapper(p))

    async def _live_wrapper(self, p):
        self.ui("live_state", True)
        try:
            await self.run_live(p, self.live_stop)
        except Exception as e:  # noqa: BLE001
            log.exception("live failed: %s", e)
        finally:
            log.info("=== LIVE stopped ===")
            self.ui("live_state", False)

    def stop_live(self):
        if self.live_stop is not None:
            self.loop.call_soon_threadsafe(self.live_stop.set)

    async def run_live(self, p, stop):
        log.info("=== LIVE %s | chart %s | LTF %s ===", ",".join(ASSETS), ",".join(p.chart_tfs), p.lower_tf)
        await self.ensure_exchanges(p.exchanges, p.proxy)
        series_by = self.build_sources(ASSETS, p)
        all_series = [s for lst in series_by.values() for s in lst]
        ccxt_series = [s for s in all_series if s.label.endswith(("-S", "-F"))]
        tv_series = [s for s in all_series if s.label.endswith((TV_CFD_SUFFIX,
                                                               TV_CRYPTO_SUFFIX_S,
                                                               TV_CRYPTO_SUFFIX_P))]
        cfd_series = [s for s in tv_series if s.label.endswith(TV_CFD_SUFFIX)]
        crypto_tv_series = [s for s in tv_series
                            if s.label.endswith((TV_CRYPTO_SUFFIX_S, TV_CRYPTO_SUFFIX_P))]
        if not all_series:
            log.error("No sources found")
            return
        tf_list = [(tf, parse_tf_minutes(tf)) for tf in p.chart_tfs]
        L_min = parse_tf_minutes(p.lower_tf)
        L_ms = L_min * MINUTE_MS
        c_max = max(m for _, m in tf_list)
        needed_min = (p.vol_sma_len + 1) * c_max + L_min + 5
        keep_min = needed_min + 60
        sem = asyncio.Semaphore(MAX_CONCURRENCY)
        log.info("Backfilling %d minutes for %d sources...", needed_min, len(all_series))
        if ccxt_series:
            await asyncio.gather(*[backfill(self.exchanges[s.ex_label], s, needed_min, sem)
                                   for s in ccxt_series])
        if tv_series:
            # ask for the maximum the anonymous feed can give (4900 BARS = ~10 days
            # at LTF 3m): after the weekend closure the volume SMA needs traded bars
            # from BEFORE the break
            await self.tv_fill(tv_series, p, TV_MAX_BARS)
            for lst in series_by.values():
                rescale_tv(lst)
        last_b = int(time.time() * 1000) // MINUTE_MS * MINUTE_MS
        log.info("Live is running. Waiting for chart bars to close...")

        async def sleep_or_stop(sec):
            try:
                await asyncio.wait_for(stop.wait(), timeout=sec)
                return True
            except asyncio.TimeoutError:
                return False

        async def tv_refresher():
            tick = 0
            while not stop.is_set():
                if await sleep_or_stop(TV_REFRESH_MIN * 60):
                    break
                tick += 1
                try:
                    closed, _ = xau_weekend_closed(int(time.time() * 1000))
                    xau_cfd = [s for s in cfd_series if s.symbol == TV_SYMBOLS["XAU"]]
                    todo_cfd = cfd_series if not closed else \
                        [s for s in cfd_series if s.symbol != TV_SYMBOLS["XAU"]]
                    # crypto exchange feeds are 24/7 markets - refresh them on every
                    # second cycle to keep the anonymous TradingView request rate sane
                    todo_crypto = crypto_tv_series if tick % 2 == 0 else []
                    if todo_cfd:
                        await self.tv_refresh(todo_cfd, n_bars=40)
                    if todo_crypto:
                        await self.tv_refresh(todo_crypto, n_bars=80)
                    extra = " - XAU skipped (market closed)" if closed and xau_cfd else ""
                    log.info("TradingView candles refreshed (CFD %d, TV-crypto %d%s)",
                             len(todo_cfd), len(todo_crypto), extra)
                except asyncio.CancelledError:
                    raise
                except Exception as e:  # noqa: BLE001
                    log.warning("TradingView refresh failed: %s", str(e)[:100])

        refresher = asyncio.create_task(tv_refresher()) if tv_series else None
        xau_paused = False
        while not stop.is_set():
            now = time.time()
            nxt = (int(now // 60) + 1) * 60
            if await sleep_or_stop(max(0.0, nxt + FETCH_DELAY_SEC - now)):
                break
            # ccxt sources are polled every minute; ALL TradingView sources (CFD +
            # crypto spot/perp) are refreshed separately by tv_refresher()
            await asyncio.gather(*[poll(self.exchanges[s.ex_label], s, sem)
                                   for s in ccxt_series])
            now_ms = int(time.time() * 1000)
            b_now = now_ms // MINUTE_MS * MINUTE_MS
            start_b = max(last_b + MINUTE_MS, b_now - 120 * MINUTE_MS)
            # weekend closure of the gold session - checked once per minute tick
            xau_closed, xau_reopen = xau_weekend_closed(now_ms)
            if xau_closed and not xau_paused:
                log.info("XAU market closed (weekend) - XAU signals pause until %s (%s)",
                         fmt_time(xau_reopen), TZ_LABEL)
                xau_paused = True
            elif not xau_closed and xau_paused:
                log.info("XAU market reopened - XAU signals resumed")
                xau_paused = False
            for B in range(start_b, b_now + MINUTE_MS, MINUTE_MS):
                for tf, tf_min in tf_list:
                    if (B // MINUTE_MS) % tf_min:
                        continue
                    C_ms = tf_min * MINUTE_MS
                    for asset, lst in series_by.items():
                        if asset == "XAU" and xau_closed:
                            continue          # no XAU evaluation while the market is shut
                        per_ex = []
                        for s in lst:
                            h = source_hist(s, B, C_ms, L_ms, p)
                            if h is not None:
                                per_ex.append((s.label, h))
                        if not per_ex:
                            continue
                        res = combine(per_ex, p, asset, tf)
                        log.info("%s %s bar start %s | %d sources", asset, tf, fmt_time(B - C_ms), res["used"])
                        for side, info in signals_from(res):
                            await send_telegram(p, format_signal(side, asset, tf, p.lower_tf, B, info))
            last_b = b_now
            for s in all_series:
                if s.label.endswith(TV_CFD_SUFFIX):
                    # never prune the pre-weekend bars a CFD source still needs
                    s.prune(b_now - (TV_MAX_BARS * s.bar_min + 120) * MINUTE_MS)
                else:
                    s.prune(b_now - keep_min * MINUTE_MS)
        if refresher is not None:
            refresher.cancel()

    async def shutdown(self):
        self.stop_live()
        await asyncio.sleep(0.2)
        await self.close_exchanges()
        tv, self.tv = self.tv, None
        if tv is not None:
            for meth in ("close", "disconnect"):
                try:
                    getattr(tv, meth)()
                except Exception:  # noqa: BLE001
                    pass


# =====================================================================
# GUI
# =====================================================================
class QueueLogHandler(logging.Handler):
    def __init__(self, put):
        super().__init__()
        self.put = put

    def emit(self, record):
        try:
            self.put("log", self.format(record))
        except Exception:  # noqa: BLE001
            pass


DARK = {"bg": "#1e1f22", "panel": "#2b2d31", "btn": "#34373d", "btn_hover": "#44484f",
        "fg": "#e6e6e6", "muted": "#8b8f97", "border": "#3a3d42", "select": "#2f5fb3",
        "accent": "#3b82f6"}


def apply_dark_theme(root):
    c = DARK
    root.configure(bg=c["bg"])
    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure(".", background=c["bg"], foreground=c["fg"], fieldbackground=c["panel"],
                 bordercolor=c["border"], lightcolor=c["bg"], darkcolor=c["bg"],
                 troughcolor=c["bg"], focuscolor=c["bg"])
    st.configure("TFrame", background=c["bg"])
    st.configure("TLabel", background=c["bg"], foreground=c["fg"])
    st.configure("TLabelframe", background=c["bg"], bordercolor=c["border"])
    st.configure("TLabelframe.Label", background=c["bg"], foreground=c["fg"])
    st.configure("TCheckbutton", background=c["bg"], foreground=c["fg"])
    st.map("TCheckbutton", background=[("active", c["bg"])],
           indicatorcolor=[("selected", c["accent"]), ("!selected", c["panel"])],
           foreground=[("disabled", c["muted"])])
    st.configure("TButton", background=c["btn"], foreground=c["fg"], bordercolor=c["border"], padding=5)
    st.map("TButton", background=[("active", c["btn_hover"]), ("disabled", c["panel"])],
           foreground=[("disabled", c["muted"])])
    st.configure("TEntry", fieldbackground=c["panel"], foreground=c["fg"], insertcolor=c["fg"],
                 bordercolor=c["border"])
    st.configure("TCombobox", fieldbackground=c["panel"], background=c["btn"], foreground=c["fg"],
                 arrowcolor=c["fg"], bordercolor=c["border"], selectbackground=c["panel"],
                 selectforeground=c["fg"])
    st.map("TCombobox", fieldbackground=[("readonly", c["panel"])], foreground=[("readonly", c["fg"])],
           selectbackground=[("readonly", c["panel"])], selectforeground=[("readonly", c["fg"])])
    st.configure("TSpinbox", fieldbackground=c["panel"], background=c["btn"], foreground=c["fg"],
                 arrowcolor=c["fg"], insertcolor=c["fg"], bordercolor=c["border"])
    st.configure("TNotebook", background=c["bg"], bordercolor=c["border"])
    st.configure("TNotebook.Tab", background=c["panel"], foreground=c["fg"], padding=[12, 4])
    st.map("TNotebook.Tab", background=[("selected", c["btn"])], foreground=[("selected", "#ffffff")])
    st.configure("Treeview", background=c["panel"], fieldbackground=c["panel"], foreground=c["fg"],
                 bordercolor=c["border"], rowheight=22)
    st.map("Treeview", background=[("selected", c["select"])], foreground=[("selected", "#ffffff")])
    st.configure("Treeview.Heading", background=c["btn"], foreground=c["fg"], relief="flat")
    st.map("Treeview.Heading", background=[("active", c["btn_hover"])])
    st.configure("Vertical.TScrollbar", background=c["btn"], troughcolor=c["bg"], arrowcolor=c["fg"],
                 bordercolor=c["border"])
    st.configure("Horizontal.TScrollbar", background=c["btn"], troughcolor=c["bg"], arrowcolor=c["fg"],
                 bordercolor=c["border"])
    # drop-down list of the comboboxes (plain tk widget)
    root.option_add("*TCombobox*Listbox.background", c["panel"])
    root.option_add("*TCombobox*Listbox.foreground", c["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", c["select"])
    root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")


class App:
    def __init__(self):
        self.q = queue.Queue()
        self.root = tk.Tk()
        self.root.title("VC69V5")
        self.root.geometry("1180x800")
        apply_dark_theme(self.root)
        self.engine = Engine(lambda kind, payload: self.q.put((kind, payload)))
        self._setup_logging()
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(150, self.poll_queue)

    def _setup_logging(self):
        log.setLevel(logging.INFO)
        h = QueueLogHandler(lambda k, v: self.q.put((k, v)))
        h.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
        log.addHandler(h)
        sh = logging.StreamHandler()
        sh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(sh)

    def _build(self):
        d = Params()
        pad = {"padx": 6, "pady": 3}

        # --- menu bar with the Exchange submenu (volume-ranked lists) -----
        mb = tk.Menu(self.root)
        mex = tk.Menu(mb, tearoff=0)
        mex.add_command(label="TradingView crypto exchange feeds (ranked by 24h volume)",
                        state="disabled")
        mex.add_separator()
        for sym in TV_CRYPTO_SPOT:
            mex.add_command(label=f"{sym} - spot ({len(TV_CRYPTO_SPOT[sym])} exchanges)",
                            command=lambda s=sym: self.show_exchange_list(s))
        mex.add_separator()
        mex.add_command(label="TradingView perpetual feeds (USDT.P)", state="disabled")
        mex.add_separator()
        for sym in TV_CRYPTO_PERP:
            mex.add_command(label=f"{sym} - perpetual ({len(TV_CRYPTO_PERP[sym])} exchanges)",
                            command=lambda s=sym: self.show_exchange_list(s))
        mb.add_cascade(label="Exchange", menu=mex)
        self.root.configure(menu=mb)

        # --- settings column inside a vertically scrollable canvas -------
        outer = ttk.Frame(self.root)
        outer.pack(side="left", fill="y", padx=8, pady=8)
        canvas = tk.Canvas(outer, bg=DARK["bg"], highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.canvas = canvas
        left = ttk.Frame(canvas)
        win = canvas.create_window((0, 0), window=left, anchor="nw")

        def _on_left_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        left.bind("<Configure>", _on_left_configure)

        def _on_canvas_configure(event):
            # keep the settings frame as wide as the visible canvas
            canvas.itemconfigure(win, width=max(event.width, left.winfo_reqwidth()))

        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            up = event.num == 4 or getattr(event, "delta", 0) > 0
            down = event.num == 5 or getattr(event, "delta", 0) < 0
            if up:
                canvas.yview_scroll(-1, "units")
            elif down:
                canvas.yview_scroll(1, "units")
            return "break"   # keep spinboxes/comboboxes from eating the wheel

        right = ttk.Frame(self.root)
        right.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)

        # --- indicator settings
        f = ttk.LabelFrame(left, text="Indicator settings")
        f.pack(fill="x")
        self.v_chart = tk.StringVar(value=",".join(d.chart_tfs))
        self.v_lower = tk.StringVar(value=d.lower_tf)
        self.v_source = tk.StringVar(value=d.source)
        self.v_diff = tk.StringVar(value=str(d.different))
        self.v_minvol = tk.StringVar(value=str(d.min_vol))
        self.v_sma = tk.StringVar(value=str(d.vol_sma_len))
        self.v_conf = tk.StringVar(value=str(d.min_confirm))
        self.v_cnt = tk.StringVar(value=str(d.min_exchange_count))
        self.v_global = tk.BooleanVar(value=d.show_global)
        self.v_spot = tk.BooleanVar(value=True)
        self.v_fut = tk.BooleanVar(value=True)
        self.v_cfd = tk.BooleanVar(value=d.use_cfd)
        self.v_tvs = tk.BooleanVar(value=d.use_tv_spot)
        self.v_tvp = tk.BooleanVar(value=d.use_tv_perp)
        rows = [
            ("Chart timeframe(s)", ttk.Entry(f, textvariable=self.v_chart, width=14)),
            ("Lower timeframe", ttk.Combobox(f, textvariable=self.v_lower, values=LOWER_TFS, width=11, state="readonly")),
            ("Source (volume direction)", ttk.Combobox(f, textvariable=self.v_source, values=SOURCES, width=11, state="readonly")),
            ("Different", ttk.Spinbox(f, textvariable=self.v_diff, from_=1, to=100, width=12)),
            ("Min Volume", ttk.Spinbox(f, textvariable=self.v_minvol, from_=55, to=100, width=12)),
            ("Volume SMA length", ttk.Spinbox(f, textvariable=self.v_sma, from_=1, to=200, width=12)),
            ("Min weighted confirmation", ttk.Spinbox(f, textvariable=self.v_conf, from_=0.0, to=1.0, increment=0.1, width=12)),
            ("Minimum Exchange Count", ttk.Spinbox(f, textvariable=self.v_cnt, from_=1, to=120, width=12)),
        ]
        for i, (label, w) in enumerate(rows):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", **pad)
            w.grid(row=i, column=1, sticky="w", **pad)
        ttk.Checkbutton(f, text="Show Global Signals (Long/Short)", variable=self.v_global).grid(
            row=len(rows), column=0, columnspan=2, sticky="w", **pad)
        ttk.Checkbutton(f, text="Spot", variable=self.v_spot).grid(row=len(rows) + 1, column=0, sticky="w", **pad)
        ttk.Checkbutton(f, text="Futures (USDT perpetual)", variable=self.v_fut).grid(
            row=len(rows) + 1, column=1, sticky="w", **pad)
        ttk.Checkbutton(f, text=f"CFD brokers - TradingView ({len(TV_BROKERS)})",
                        variable=self.v_cfd).grid(
            row=len(rows) + 2, column=0, columnspan=2, sticky="w", **pad)
        ttk.Checkbutton(f, text=f"TV crypto spot ({sum(len(v) for v in TV_CRYPTO_SPOT.values())})",
                        variable=self.v_tvs).grid(
            row=len(rows) + 3, column=0, sticky="w", **pad)
        ttk.Checkbutton(f, text=f"TV crypto perpetual ({sum(len(v) for v in TV_CRYPTO_PERP.values())})",
                        variable=self.v_tvp).grid(
            row=len(rows) + 3, column=1, sticky="w", **pad)
        ttk.Label(f, text="Chart timeframes: comma separated, e.g. 2h or 15m,2h", foreground="gray").grid(
            row=len(rows) + 4, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(f, text="XAU weekend closure: Fri 21/22:00 - Sun 22/23:00 UTC "
                          "(signals pause)", foreground="gray").grid(
            row=len(rows) + 5, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(f, text="Exchange menu (menu bar) = volume-ranked exchange lists",
                  foreground="gray").grid(
            row=len(rows) + 6, column=0, columnspan=2, sticky="w", **pad)

        # --- exchanges
        fe = ttk.LabelFrame(left, text="Exchanges (ccxt spot + futures)")
        fe.pack(fill="x", pady=(8, 0))
        self.v_ex = {}
        for i, (label, _) in enumerate(EXCHANGES):
            v = tk.BooleanVar(value=True)
            self.v_ex[label] = v
            ttk.Checkbutton(fe, text=label, variable=v).grid(row=i // 3, column=i % 3, sticky="w", padx=6, pady=1)
        bar = ttk.Frame(fe)
        bar.grid(row=7, column=0, columnspan=3, sticky="w", pady=3)
        ttk.Button(bar, text="All", width=6, command=lambda: [v.set(True) for v in self.v_ex.values()]).pack(side="left", padx=4)
        ttk.Button(bar, text="None", width=6, command=lambda: [v.set(False) for v in self.v_ex.values()]).pack(side="left")

        # --- history
        fh = ttk.LabelFrame(left, text="History (one coin)")
        fh.pack(fill="x", pady=(8, 0))
        self.v_hcoin = tk.StringVar(value="BTC")
        self.v_hdays = tk.StringVar(value=str(d.history_days))
        self.v_hsend = tk.BooleanVar(value=True)
        ttk.Label(fh, text="Coin").grid(row=0, column=0, sticky="w", **pad)
        ttk.Combobox(fh, textvariable=self.v_hcoin, values=ASSETS, width=10, state="readonly").grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(fh, text="Days back").grid(row=1, column=0, sticky="w", **pad)
        ttk.Spinbox(fh, textvariable=self.v_hdays, from_=1, to=365, width=12).grid(row=1, column=1, sticky="w", **pad)
        ttk.Checkbutton(fh, text="Send summary to Telegram", variable=self.v_hsend).grid(row=2, column=0, columnspan=2, sticky="w", **pad)
        hbtns = ttk.Frame(fh)
        hbtns.grid(row=3, column=0, columnspan=2, sticky="we", **pad)
        self.btn_hist = ttk.Button(hbtns, text="Check history", command=self.on_history)
        self.btn_hist.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.btn_hist_stop = ttk.Button(hbtns, text="Stop", command=self.on_hist_stop,
                                        state="disabled", width=7)
        self.btn_hist_stop.pack(side="left")

        # --- live
        fl = ttk.LabelFrame(left, text="Live (BTC, PEPE, ZEC, XAU at the same time)")
        fl.pack(fill="x", pady=(8, 0))
        self.btn_live = ttk.Button(fl, text="Start live", command=self.on_live)
        self.btn_live.pack(side="left", padx=6, pady=6)
        self.btn_stop = ttk.Button(fl, text="Stop live", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6, pady=6)

        # --- telegram / proxy
        ft = ttk.LabelFrame(left, text="Telegram / proxy")
        ft.pack(fill="x", pady=(8, 0))
        self.v_tok = tk.StringVar(value=os.getenv("TELEGRAM_BOT_TOKEN", ""))
        self.v_chat = tk.StringVar(value=os.getenv("TELEGRAM_CHAT_ID", ""))
        self.v_proxy = tk.StringVar(value=os.getenv("PROXY", ""))
        for i, (label, var, show) in enumerate([("Bot token", self.v_tok, "*"), ("Chat ID", self.v_chat, ""),
                                                ("Proxy", self.v_proxy, "")]):
            ttk.Label(ft, text=label).grid(row=i, column=0, sticky="w", **pad)
            ttk.Entry(ft, textvariable=var, width=26, show=show).grid(row=i, column=1, sticky="w", **pad)

        # --- tradingview (optional login for higher CFD rate limits)
        fv = ttk.LabelFrame(left, text=f"TradingView login (optional, {len(TV_BROKERS)} CFD "
                                       f"+ {sum(len(v) for v in TV_CRYPTO_SPOT.values()) + sum(len(v) for v in TV_CRYPTO_PERP.values())} crypto feeds)")
        fv.pack(fill="x", pady=(8, 0))
        self.v_tvuser = tk.StringVar(value=d.tv_user)
        self.v_tvpass = tk.StringVar(value=d.tv_pass)
        ttk.Label(fv, text="TV username").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(fv, textvariable=self.v_tvuser, width=26).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(fv, text="TV password").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(fv, textvariable=self.v_tvpass, width=26, show="*").grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(fv, text="free account = higher TradingView rate limits",
                  foreground="gray").grid(row=2, column=0, columnspan=2, sticky="w", **pad)

        # --- right side: log + history table
        self.status = tk.StringVar(value="Ready")
        ttk.Label(right, textvariable=self.status).pack(anchor="w")
        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)
        tab_log = ttk.Frame(nb)
        tab_hist = ttk.Frame(nb)
        nb.add(tab_log, text="Log")
        nb.add(tab_hist, text="History results")
        self.nb = nb
        self.tab_hist = tab_hist

        self.txt = tk.Text(tab_log, wrap="none", height=30, font=("Consolas", 9),
                           bg=DARK["panel"], fg=DARK["fg"], insertbackground=DARK["fg"],
                           selectbackground=DARK["select"], selectforeground="#ffffff",
                           relief="flat", borderwidth=0, highlightthickness=1,
                           highlightbackground=DARK["border"], highlightcolor=DARK["border"])
        sb = ttk.Scrollbar(tab_log, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True)

        cols = ("time", "tf", "side", "spot", "futures", "cfd", "tvspot", "tvperp",
                "weight", "names")
        self.tree = ttk.Treeview(tab_hist, columns=cols, show="headings")
        for c, w in zip(cols, (190, 60, 110, 80, 90, 80, 80, 80, 70, 1800)):
            self.tree.heading(c, text=f"bar start ({TZ_LABEL})" if c == "time" else c)
            self.tree.column(c, width=w, anchor="w")
        for tag, color in (("buy", "#4ade80"), ("sell", "#f87171"), ("long", "#60a5fa"), ("short", "#facc15")):
            self.tree.tag_configure(tag, foreground=color)
        # horizontal scrollbar: lets the user scroll to see every exchange name
        xsb = ttk.Scrollbar(tab_hist, orient="horizontal", command=self.tree.xview)
        sb2 = ttk.Scrollbar(tab_hist, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb2.set, xscrollcommand=xsb.set)
        xsb.pack(side="bottom", fill="x")      # packed before sb2 -> spans full width
        sb2.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # --- mouse-wheel scrolling for the settings column:
        # bind the handler on the canvas and every widget inside it
        def _bind_wheel(w):
            w.bind("<MouseWheel>", _on_mousewheel)
            w.bind("<Button-4>", _on_mousewheel)
            w.bind("<Button-5>", _on_mousewheel)
            for c in w.winfo_children():
                _bind_wheel(c)

        _bind_wheel(canvas)

    # ---- settings ------------------------------------------------------
    def collect(self):
        try:
            p = Params(
                chart_tfs=tuple(t.strip() for t in self.v_chart.get().split(",") if t.strip()),
                lower_tf=self.v_lower.get().strip(),
                source=self.v_source.get(),
                different=int(self.v_diff.get()),
                min_vol=int(self.v_minvol.get()),
                vol_sma_len=int(self.v_sma.get()),
                min_confirm=float(self.v_conf.get()),
                min_exchange_count=int(self.v_cnt.get()),
                show_global=self.v_global.get(),
                use_spot=self.v_spot.get(),
                use_futures=self.v_fut.get(),
                use_cfd=self.v_cfd.get(),
                use_tv_spot=self.v_tvs.get(),
                use_tv_perp=self.v_tvp.get(),
                exchanges=tuple(l for l, _ in EXCHANGES if self.v_ex[l].get()),
                history_asset=self.v_hcoin.get(),
                history_days=int(self.v_hdays.get()),
                tg_history=self.v_hsend.get(),
                tg_token=self.v_tok.get().strip(),
                tg_chat=self.v_chat.get().strip(),
                proxy=self.v_proxy.get().strip(),
                tv_user=self.v_tvuser.get().strip(),
                tv_pass=self.v_tvpass.get().strip(),
            )
            p.validate()
            return p
        except ValueError as e:
            messagebox.showerror("Invalid settings", str(e))
            return None

    # ---- buttons -------------------------------------------------------
    def on_history(self):
        p = self.collect()
        if p is None:
            return
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.engine.start_history(p):
            self.nb.select(0)

    def on_live(self):
        p = self.collect()
        if p is None:
            return
        if self.engine.start_live(p):
            self.nb.select(0)

    def on_stop(self):
        self.engine.stop_live()
        self.status.set("Stopping live...")

    def on_hist_stop(self):
        if self.engine.stop_history():
            self.status.set("Stopping history...")

    # ---- Exchange menu --------------------------------------------------
    @staticmethod
    def _compact_vol(v):
        for div, tag in ((1e12, "T"), (1e9, "B"), (1e6, "M"), (1e3, "K")):
            if v >= div:
                return f"{v / div:.2f}{tag}"
        return f"{v:,.0f}"

    def show_exchange_list(self, sym):
        """Exchange submenu: show the volume-ranked exchange list of ONE symbol.
        Every symbol has its own list - lists are never merged."""
        if sym in TV_CRYPTO_SPOT:
            feeds, kind = TV_CRYPTO_SPOT[sym], "spot"

            def vol_txt(v):
                return "$" + self._compact_vol(v)
            note = ("volume = 24h traded value in USDT (TradingView scanner, "
                    f"{TV_SCAN_DATE}) - this is the source order used by the engine")
        else:
            feeds, kind = TV_CRYPTO_PERP[sym], "perpetual"

            def vol_txt(v):
                return self._compact_vol(v)
            note = ("volume = 24h volume as reported by TradingView, in contract "
                    f"units (scanner {TV_SCAN_DATE}) - units differ per exchange, so "
                    "the engine normalizes each of these sources individually")
        win = tk.Toplevel(self.root)
        win.title(f"{sym} - TradingView {kind} feeds, ranked by 24h volume")
        win.configure(bg=DARK["bg"])
        win.geometry("520x560")
        cols = ("rank", "exchange", "symbol", "volume")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c, w, t in zip(cols, (50, 190, 120, 110),
                           ("#", "Exchange (TradingView feed)", "Symbol", "24h volume")):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w" if c == "exchange" else "center")
        for i, (exch, vol) in enumerate(feeds, 1):
            tree.insert("", "end", values=(i, exch, sym, vol_txt(vol)),
                        tags=("top",) if i <= 3 else ())
        tree.tag_configure("top", foreground="#4ade80")
        sb = ttk.Scrollbar(win, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=8, pady=(8, 0))
        ttk.Label(win, text=note, foreground=DARK["muted"], wraplength=490).pack(
            anchor="w", padx=8, pady=6)

    def on_close(self):
        self.engine.stop_history()
        try:
            self.engine.submit(self.engine.shutdown()).result(timeout=8)
        except Exception:  # noqa: BLE001
            pass
        self.root.destroy()

    # ---- queue from engine ----------------------------------------------
    def poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.txt.insert("end", payload + "\n")
                    if int(self.txt.index("end-1c").split(".")[0]) > 3000:
                        self.txt.delete("1.0", "500.0")
                    self.txt.see("end")
                elif kind == "hist_state":
                    self.btn_hist.configure(state="disabled" if payload else "normal")
                    self.btn_hist_stop.configure(state="normal" if payload else "disabled")
                    if payload:
                        self.status.set("History running...")
                    else:
                        self.status.set("Live running" if self.btn_stop["state"] == "normal" else "Ready")
                elif kind == "live_state":
                    self.btn_live.configure(state="disabled" if payload else "normal")
                    self.btn_stop.configure(state="normal" if payload else "disabled")
                    self.status.set("Live running" if payload else "Ready")
                elif kind == "hist_rows":
                    for r in reversed(payload):   # newest first
                        side = r["side"]
                        self.tree.insert("", "end", values=(
                            r["time"], r["chart_tf"], side, r["spot_count"], r["futures_count"],
                            r["cfd_count"], r["tvspot_count"], r["tvperp_count"],
                            f"{r['weight']}%" if r["weight"] != "" else "",
                            (str(r["exchanges"]) + ("  |  " + r["pairs"] if r.get("pairs") else ""))),
                            tags=({"BUY": "buy", "SELL": "sell", "GLOBAL LONG": "long",
                                   "GLOBAL SHORT": "short"}.get(side, ""),))
                    self.nb.select(1)
        except queue.Empty:
            pass
        self.root.after(150, self.poll_queue)

    def run(self):
        self.root.mainloop()


def main():
    if tk is None:
        raise SystemExit("tkinter is not available in this Python installation.")
    App().run()


if __name__ == "__main__":
    main()
