#!/usr/bin/env python3
"""
vc69v5 headless runner - one check per invocation, made for GitHub Actions.

Every run:
  1. (optional) waits until FIRST_RUN_TEHRAN ("2026-11-08 13:30" by default)
  2. downloads the recent 1m candles of every source (ccxt + tvDatafeed)
  3. evaluates EVERY chart bar closed since the previous run (state file),
     so a late GitHub trigger never misses or repeats a signal - bar times
     stay exactly on their timeframe boundaries
  4. sends every new signal to Telegram with the per-family triple
     "BTCUSDT 2/15/40" (confirm / checked / volume%)

Configuration (environment variables, all optional):
  TG_TOKEN / TG_CHAT      Telegram bot token and chat id (secrets in GitHub)
  TV_USER / TV_PASS       free TradingView account - strongly recommended on
                          servers (fewer 429s than anonymous access)
  FIRST_RUN_TEHRAN        "YYYY-MM-DD HH:MM" - do nothing before this Iran time
  CHART_TFS               default "3m,15m,30m,45m"
  LOWER_TF                default "1m"
  ASSETS                  default "BTC,PEPE,ZEC"   (XAU is excluded by design)
  DIFFERENT / MIN_VOL / VOL_SMA_LEN / MIN_CONFIRM / MIN_EXCHANGES
                          indicator inputs (defaults 5 / 69 / 20 / 0.1 / 2;
                          MIN_EXCHANGES/MIN_CONFIRM apply to assets WITHOUT
                          a per-asset rule)
                          NOTE: the per-asset signal rules live in
                          SIGNAL_MIN_COUNTS inside vc69v5_app.py - a signal
                          fires when ANY ONE market family (spot/futures/
                          CFD/TV-spot/TV-perp) on its own has BTC & PEPE:
                          6 confirming exchanges AND 40% of that family's
                          volume, ZEC: 3 AND 40%.
  SIGNAL_MIN_COUNTS       optional override of that rule list, format
                          "ASSET=COUNT[/WEIGHT%]", e.g.
                          "BTC=6/40,PEPE=6/40,ZEC=3/40". A signal fires when
                          ANY ONE market family on its own reaches BOTH
                          numbers; WEIGHT defaults to 40. Setting the
                          variable REPLACES the whole rule list; the value
                          "none" disables every per-asset rule (all assets
                          then use MIN_EXCHANGES/MIN_CONFIRM). The effective
                          rules are printed in the first "settings:" log line.
  SHOW_GLOBAL             "1" to also emit GLOBAL LONG/SHORT (default off)
  FIRST_WINDOW_MIN        baseline window of the very first run (default 15)
  MAX_LOOKBACK_MIN        cap when the state is stale (default 360)

Local test without Telegram:   python vc69v5_headless.py --dry-run
"""

import argparse
import asyncio
import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vc69v5_app as app  # noqa: E402

log = logging.getLogger("vc69v5.headless")

TEHRAN = app.TZ
MINUTE_MS = app.MINUTE_MS
STATE_VERSION = 1

# import-time snapshot of the per-asset rules coded in vc69v5_app.py -
# apply_signal_rules() restores it whenever the env variable is unset
_DEFAULT_SIGNAL_MIN_COUNTS = dict(app.SIGNAL_MIN_COUNTS)


def env(name, default):
    v = os.environ.get(name)
    return v if v not in (None, "") else default


def build_params():
    apply_signal_rules()
    tfs = tuple(t.strip() for t in env("CHART_TFS", "3m,15m,30m,45m").split(",") if t.strip())
    p = app.Params(
        chart_tfs=tfs,
        lower_tf=env("LOWER_TF", "1m"),
        different=int(env("DIFFERENT", "5")),
        min_vol=int(env("MIN_VOL", "69")),
        vol_sma_len=int(env("VOL_SMA_LEN", "20")),
        min_confirm=float(env("MIN_CONFIRM", "0.1")),
        min_exchange_count=int(env("MIN_EXCHANGES", "2")),
        show_global=env("SHOW_GLOBAL", "0") == "1",
        tg_token=env("TG_TOKEN", ""),
        tg_chat=env("TG_CHAT", ""),
        tv_user=env("TV_USER", ""),
        tv_pass=env("TV_PASS", ""),
    )
    p.validate()
    return p


# --------------------------------------------------- per-asset rule override
def rules_text():
    """'BTC=6/40 PEPE=6/40 ZEC=3/40' (for the settings log line) - or
    'none(global-rule)' when the per-asset list is empty."""
    return " ".join(f"{a}={n}/{w}"
                    for a, (n, w) in app.SIGNAL_MIN_COUNTS.items()) \
        or "none(global-rule)"


def apply_signal_rules():
    """Optional env override for SIGNAL_MIN_COUNTS in vc69v5_app.py.

    SIGNAL_MIN_COUNTS="BTC=6/40,PEPE=6/40,ZEC=3/40"
      ASSET=COUNT[/WEIGHT%]: a signal fires when ANY ONE market family
      (spot / futures / CFD / TV-spot / TV-perp) ON ITS OWN has >= COUNT
      confirming exchanges AND >= WEIGHT% of that family's traded volume;
      WEIGHT defaults to 40 when omitted. Setting the variable REPLACES the
      whole rule list; "none" clears it (every asset falls back to the
      global MIN_EXCHANGES/MIN_CONFIRM rule). Unset = the defaults coded in
      vc69v5_app.py stay active.
    """
    raw = env("SIGNAL_MIN_COUNTS", "")
    if not raw.strip():
        app.SIGNAL_MIN_COUNTS.clear()
        app.SIGNAL_MIN_COUNTS.update(_DEFAULT_SIGNAL_MIN_COUNTS)
        return
    if raw.strip().lower() == "none":
        app.SIGNAL_MIN_COUNTS.clear()
        log.info("SIGNAL_MIN_COUNTS env = none - every asset uses the "
                 "global MIN_EXCHANGES/MIN_CONFIRM rule")
        return
    rules = {}
    for part in raw.split(","):
        part = part.strip()
        m = re.fullmatch(r"([A-Za-z]{2,6})\s*=\s*(\d+)(?:\s*/\s*(\d+))?", part)
        if not m:
            raise SystemExit(f"Bad SIGNAL_MIN_COUNTS item '{part}' - use "
                             "ASSET=COUNT[/WEIGHT], e.g. "
                             "'BTC=6/40,PEPE=6/40,ZEC=3/40'")
        a, n, w = m.group(1).upper(), int(m.group(2)), m.group(3)
        if a not in app.ASSETS:
            raise SystemExit(f"SIGNAL_MIN_COUNTS: unknown asset '{a}' - "
                             f"choose from {', '.join(app.ASSETS)}")
        if n < 1:
            raise SystemExit(f"SIGNAL_MIN_COUNTS: COUNT for {a} must be >= 1")
        w = int(w) if w is not None else 40
        if not 1 <= w <= 100:
            raise SystemExit(f"SIGNAL_MIN_COUNTS: WEIGHT for {a} must be "
                             "1..100")
        rules[a] = (n, w)
    app.SIGNAL_MIN_COUNTS.clear()
    app.SIGNAL_MIN_COUNTS.update(rules)
    log.info("SIGNAL_MIN_COUNTS env applied: %s", rules_text())


# --------------------------------------------------------------------- state
def load_state(path):
    try:
        with open(path, encoding="utf-8") as f:
            st = json.load(f)
        if isinstance(st, dict) and st.get("v") == STATE_VERSION \
                and isinstance(st.get("last_bar"), dict):
            return st
        log.warning("state file has an unknown format - starting fresh")
    except FileNotFoundError:
        log.info("no state file yet - first run")
    except Exception as e:  # noqa: BLE001
        log.warning("cannot read state (%s) - starting fresh", e)
    return {"v": STATE_VERSION, "last_bar": {}, "first_done": False,
            "last_run_utc": None}


def save_state(path, st):
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, path)
    log.info("state saved: %s", path)


# ------------------------------------------------------------------ telegram
async def tg_send(p, text, dry):
    log.info("SIGNAL\n%s", text)
    if dry or not p.tg_token or not p.tg_chat:
        return True
    import aiohttp
    url = f"https://api.telegram.org/bot{p.tg_token}/sendMessage"
    for attempt in range(3):
        try:
            async with aiohttp.ClientSession() as s:
                async with s.post(url, json={"chat_id": p.tg_chat, "text": text[:4000]},
                                  proxy=p.proxy or None,
                                  timeout=aiohttp.ClientTimeout(total=30)) as r:
                    if r.status == 200:
                        return True
                    log.warning("Telegram HTTP %s: %s", r.status, (await r.text())[:120])
        except Exception as e:  # noqa: BLE001
            log.warning("Telegram send failed (try %d/3): %s", attempt + 1, e)
        await asyncio.sleep(2 * (attempt + 1))
    return False


# ----------------------------------------------------------------- main pass
async def one_pass(eng, p, state_path, dry):
    now_ms = int(time.time() * 1000)

    fr = env("FIRST_RUN_TEHRAN", "")
    if fr:
        try:
            fr_ms = int(datetime.strptime(fr, "%Y-%m-%d %H:%M")
                        .replace(tzinfo=TEHRAN).timestamp() * 1000)
        except ValueError:
            raise SystemExit(f"Bad FIRST_RUN_TEHRAN '{fr}' - use 'YYYY-MM-DD HH:MM'")
        if now_ms < fr_ms:
            log.info("Before the configured first run (%s Tehran) - nothing to do.", fr)
            return

    assets = [a.strip().upper() for a in env("ASSETS", "BTC,PEPE,ZEC").split(",")
              if a.strip()]
    bad = [a for a in assets if a not in app.ASSETS]
    if bad:
        raise SystemExit(f"Unknown asset(s) {bad} - choose from {app.ASSETS}")

    state = load_state(state_path)
    L_min = app.parse_tf_minutes(p.lower_tf)
    L_ms = L_min * MINUTE_MS
    c_max = max(app.parse_tf_minutes(t) for t in p.chart_tfs)
    base_min = (p.vol_sma_len + 1) * c_max + L_min + 5

    # how far back do we need candles? (oldest unprocessed bar + SMA window)
    oldest_last = now_ms
    for tf in p.chart_tfs:
        C_ms = app.parse_tf_minutes(tf) * MINUTE_MS
        for a in assets:
            last = state["last_bar"].get(f"{a}|{tf}")
            if last is None:
                last = (now_ms - int(env("FIRST_WINDOW_MIN", "15")) * MINUTE_MS) \
                    // C_ms * C_ms
            oldest_last = min(oldest_last, int(last))
    needed_min = base_min + max(0, (now_ms - oldest_last) // MINUTE_MS)

    log.info("=== CHECK assets=%s chart=%s LTF=%s | lookback ~%d min ===",
             ",".join(assets), ",".join(p.chart_tfs), p.lower_tf, needed_min)
    await eng.ensure_exchanges(p.exchanges, p.proxy)
    series_by = eng.build_sources(assets, p)
    sem = asyncio.Semaphore(app.MAX_CONCURRENCY)

    for a in assets:
        lst = series_by.get(a, [])
        if not lst:
            log.warning("%s: no sources - skipped", a)
            continue
        ccxt_lst = [s for s in lst if s.label.endswith(("-S", "-F"))]
        tv_lst = [s for s in lst if s.label.endswith((app.TV_CFD_SUFFIX,
                                                      app.TV_CRYPTO_SUFFIX_S,
                                                      app.TV_CRYPTO_SUFFIX_P))]
        if ccxt_lst:
            await asyncio.gather(*[app.backfill(eng.exchanges[s.ex_label], s,
                                                needed_min, sem) for s in ccxt_lst])
        if tv_lst:
            n_bars = min(app.TV_MAX_BARS, max(500, -(-needed_min // L_min) + 80))
            await eng.tv_fill(tv_lst, p, n_bars)
        app.rescale_tv(lst)

    total_bars = total_signals = 0
    for a in assets:
        lst = series_by.get(a, [])
        for tf in p.chart_tfs:
            C_min = app.parse_tf_minutes(tf)
            C_ms = C_min * MINUTE_MS
            b_now = now_ms // C_ms * C_ms            # close time of the newest closed bar
            key = f"{a}|{tf}"
            last = state["last_bar"].get(key)
            if last is None:                          # first run: only the last quarter hour
                last = (now_ms - int(env("FIRST_WINDOW_MIN", "15")) * MINUTE_MS) \
                    // C_ms * C_ms
            last = int(last)
            cap = (now_ms - int(env("MAX_LOOKBACK_MIN", "360")) * MINUTE_MS) \
                // C_ms * C_ms
            if last < cap:                            # stale state - never flood
                last = cap
            n_bars = n_sig = 0
            for B in range(last + C_ms, b_now + 1, C_ms):
                per_ex = []
                for s in lst:
                    h = app.source_hist(s, B, C_ms, L_ms, p)
                    if h is not None:
                        per_ex.append((s.label, h))
                if not per_ex:
                    continue
                res = app.combine(per_ex, p, a, tf)
                n_bars += 1
                log.info("%s %s bar start %s | %d sources checked",
                         a, tf, app.fmt_time(B - C_ms), res["used"])
                for side, info in app.signals_from(res):
                    text = app.format_signal(side, a, tf, p.lower_tf, B, info)
                    if await tg_send(p, text, dry):
                        n_sig += 1
                    else:
                        log.error("Telegram delivery FAILED for the signal above")
            log.info("CHECKED %s %s: %d closed bar(s), %d signal(s)",
                     a, tf, n_bars, n_sig)
            total_bars += n_bars
            total_signals += n_sig
            state["last_bar"][key] = b_now

    state["first_done"] = True
    state["last_run_utc"] = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    save_state(state_path, state)
    log.info("DONE: %d bar(s) evaluated, %d signal(s).", total_bars, total_signals)


def main():
    ap = argparse.ArgumentParser(
        description="vc69v5 headless signal checker (one pass, GitHub-Actions ready)")
    ap.add_argument("--state", default=env("STATE_FILE", "state/state.json"),
                    help="state file (default: state/state.json)")
    ap.add_argument("--dry-run", action="store_true",
                    help="log signals but do NOT send them to Telegram")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)-7s %(message)s",
                        datefmt="%H:%M:%S")
    # the app silences the ROOT logger when it connects to TradingView
    # (GUI quirk) - keep our loggers at INFO explicitly
    logging.getLogger("vc69v5").setLevel(logging.INFO)
    logging.getLogger("vc69v5.headless").setLevel(logging.INFO)
    p = build_params()
    log.info("settings: different=%d min_vol=%d vol_sma_len=%d min_confirm=%.2f "
             "min_exchanges=%d | signal_rules: %s | telegram=%s | tv_login=%s",
             p.different, p.min_vol, p.vol_sma_len, p.min_confirm,
             p.min_exchange_count, rules_text(),
             "on" if (p.tg_token and p.tg_chat) and not args.dry_run else "OFF",
             "yes" if p.tv_user and p.tv_pass else "anonymous")

    eng = app.Engine(ui=lambda kind, payload: None)
    try:
        eng.submit(one_pass(eng, p, args.state, args.dry_run)).result()
    except KeyboardInterrupt:
        log.warning("interrupted")
    finally:
        try:
            eng.submit(eng.shutdown()).result(timeout=60)
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    main()
