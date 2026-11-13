#!/usr/bin/env python3
"""
vc69v5 headless runner - one check per invocation, made for GitHub Actions.

Every run:
  1. reads settings.toml (ALL signal settings live there - edit it on GitHub)
  2. first run only: sends a Telegram TEST message (connection check)
  3. downloads recent candles of every source (ccxt; TradingView optional)
  4. evaluates EVERY chart bar closed since the previous run (state file),
     so a late GitHub trigger never misses or repeats a signal
  5. sends every new signal to Telegram, one line per market:
        Spot: 2/15/40/50
        2 = confirming exchanges in that market
        15 = exchanges of that market that traded on the bar (with data)
        40 = % of that market's volume held by the confirming exchanges
        50 = % of the volume of ALL markets held by all confirming sources

Secrets (GitHub -> Settings -> Secrets and variables -> Actions):
  TG_TOKEN / TG_CHAT      Telegram bot token and chat id
  TV_USER / TV_PASS       optional TradingView login (only if [tradingview] enabled)

Local tests:
  python vc69v5_headless.py --check-settings     validate settings.toml and print it
  python vc69v5_headless.py --dry-run            run once, log signals, send nothing
"""

import argparse
import asyncio
import json
import logging
import os
import sys
import threading
import time
import tomllib
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vc69v5_app as app  # noqa: E402

log = logging.getLogger("vc69v5.headless")

TEHRAN = app.TZ
MINUTE_MS = app.MINUTE_MS
STATE_VERSION = 1
FAMILIES = ("SPOT", "PERP", "CFD")
FAM_NAME = {"SPOT": "Spot", "PERP": "Futu", "CFD": "CFD"}
MARKET_KEYS = {"spot": "SPOT", "futures": "PERP", "cfd": "CFD"}
RESERVED = {"telegram", "runtime", "tradingview", "exchanges", "defaults"}

BUILTIN_ASSET = {
    "enabled": True,
    "chart_timeframes": ["3m", "15m", "30m", "45m"],
    "lower_timeframe": "1m",
    "different": 5,
    "min_volume": 69,
    "volume_sma_length": 20,
    "show_global_signals": False,
    "spot": {"exchanges": 3, "weight": 20},
    "futures": {"exchanges": 3, "weight": 20},
    "cfd": {"exchanges": 0, "weight": 0},
    "total_weight": 0,
}
BUILTIN_RUNTIME = {
    "first_run_tehran": "",
    "first_window_min": 15,
    "max_lookback_min": 360,
    "max_run_minutes": 10,
    "exchange_load_seconds": 150,
    "backfill_seconds": 150,
    "tv_seconds": 120,
}


class SettingsError(Exception):
    pass


# =====================================================================
# settings.toml
# =====================================================================
def _num(v, where, lo, hi, as_int=False):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        raise SettingsError(f"{where}: must be a number (got {v!r})")
    if as_int and int(v) != v:
        raise SettingsError(f"{where}: must be a whole number (got {v!r})")
    if not lo <= v <= hi:
        raise SettingsError(f"{where}: must be between {lo} and {hi} (got {v!r})")
    return int(v) if as_int else float(v)


def _market(m, where):
    if not isinstance(m, dict):
        raise SettingsError(f"{where}: use  {{ exchanges = 3, weight = 20 }}")
    extra = set(m) - {"exchanges", "weight"}
    if extra:
        raise SettingsError(f"{where}: unknown key(s) {sorted(extra)} - allowed: exchanges, weight")
    return (_num(m.get("exchanges", 0), f"{where}.exchanges", 0, 100, as_int=True),
            _num(m.get("weight", 0), f"{where}.weight", 0, 100))


def resolve_asset(name, defaults, block):
    """built-in defaults <- [defaults] <- [ASSET]; returns the validated config."""
    allowed = set(BUILTIN_ASSET)
    for src, label in ((defaults, "[defaults]"), (block, f"[{name}]")):
        unknown = set(src) - allowed
        if unknown:
            raise SettingsError(f"{label}: unknown setting(s) {sorted(unknown)} - "
                                f"allowed: {', '.join(sorted(allowed))}")
    m = {**BUILTIN_ASSET, **defaults, **block}
    where = f"[{name}]"
    tfs = m["chart_timeframes"]
    if not isinstance(tfs, list) or not tfs or not all(isinstance(t, str) for t in tfs):
        raise SettingsError(f'{where}.chart_timeframes: use a list like ["3m", "15m", "1h"]')
    if not isinstance(m["lower_timeframe"], str):
        raise SettingsError(f'{where}.lower_timeframe: use a text like "1m"')
    rules = {}
    for key, fam in MARKET_KEYS.items():
        rules[fam] = _market(m[key], f"{where}.{key}")
    if all(rules[f][0] == 0 for f in FAMILIES):
        raise SettingsError(f"{where}: at least one market (spot / futures / cfd) needs exchanges >= 1")
    rules["overall"] = _num(m["total_weight"], f"{where}.total_weight", 0, 100)
    return {
        "name": name,
        "enabled": bool(m["enabled"]),
        "chart_tfs": tuple(t.strip() for t in tfs),
        "lower_tf": m["lower_timeframe"].strip(),
        "different": _num(m["different"], f"{where}.different", 1, 100, as_int=True),
        "min_vol": _num(m["min_volume"], f"{where}.min_volume", 55, 100, as_int=True),
        "sma_len": _num(m["volume_sma_length"], f"{where}.volume_sma_length", 1, 200, as_int=True),
        "show_global": bool(m["show_global_signals"]),
        "rules": rules,
    }


def load_settings(path):
    try:
        with open(path, "rb") as f:
            raw = tomllib.load(f)
    except FileNotFoundError:
        raise SettingsError(f"settings file not found: {path}")
    except tomllib.TOMLDecodeError as e:
        raise SettingsError(f"settings.toml has a syntax error: {e}")

    tg = raw.get("telegram", {})
    test_id = tg.get("test_message_id", 1)
    if isinstance(test_id, bool) or not isinstance(test_id, int):
        raise SettingsError("[telegram].test_message_id must be a whole number")

    rt = {**BUILTIN_RUNTIME, **raw.get("runtime", {})}
    unknown = set(rt) - set(BUILTIN_RUNTIME)
    if unknown:
        raise SettingsError(f"[runtime]: unknown setting(s) {sorted(unknown)}")
    for k in ("first_window_min", "max_lookback_min", "max_run_minutes",
              "exchange_load_seconds", "backfill_seconds", "tv_seconds"):
        rt[k] = _num(rt[k], f"[runtime].{k}", 1, 100000, as_int=True)

    tv_enabled = bool(raw.get("tradingview", {}).get("enabled", False))
    exclude = {str(x).upper() for x in raw.get("exchanges", {}).get("exclude", [])}
    known = {l for l, _ in app.EXCHANGES}
    bad = exclude - known
    if bad:
        raise SettingsError(f"[exchanges].exclude: unknown exchange(s) {sorted(bad)} - "
                            f"known: {', '.join(sorted(known))}")

    defaults = raw.get("defaults", {})
    assets = {}
    for name, block in raw.items():
        if name in RESERVED:
            continue
        if not isinstance(block, dict):
            raise SettingsError(f"unexpected top-level setting '{name}' - asset sections look like [BTC]")
        if name not in app.ASSETS or name == "XAU":
            raise SettingsError(f"[{name}]: unknown asset - choose from "
                                f"{', '.join(a for a in app.ASSETS if a != 'XAU')}")
        cfg = resolve_asset(name, defaults, block)
        if cfg["enabled"]:
            assets[name] = cfg
    if not assets:
        raise SettingsError("no enabled asset section - add [BTC], [PEPE] or [ZEC]")

    labels = tuple(l for l, _ in app.EXCHANGES if l not in exclude)
    for name, cfg in assets.items():
        p = app.Params(chart_tfs=cfg["chart_tfs"], lower_tf=cfg["lower_tf"],
                       different=cfg["different"], min_vol=cfg["min_vol"],
                       vol_sma_len=cfg["sma_len"], show_global=cfg["show_global"],
                       use_cfd=tv_enabled, use_tv_spot=tv_enabled, use_tv_perp=tv_enabled,
                       exchanges=labels)
        try:
            p.validate()
        except ValueError as e:
            raise SettingsError(f"[{name}]: {e}")
        cfg["params"] = p
        if cfg["rules"]["CFD"][0] > 0 and not tv_enabled:
            log.warning("[%s] cfd rule is set but [tradingview] is disabled - "
                        "the CFD market never trades, so it can never fire", name)
    return {"telegram": {"test_message_id": test_id}, "runtime": rt, "tv_enabled": tv_enabled,
            "labels": labels, "assets": assets}


def rules_line(cfg):
    r = cfg["rules"]
    parts = [f"{FAM_NAME[f]} >= {r[f][0]} ex & >= {r[f][1]:g}%" for f in FAMILIES if r[f][0] > 0]
    tot = f"total >= {r['overall']:g}%" if r["overall"] > 0 else "total: off"
    return "; ".join(parts) + "; " + tot


def settings_text(settings):
    lines = []
    for name, cfg in settings["assets"].items():
        lines.append(f"{name}: charts {','.join(cfg['chart_tfs'])} / LTF {cfg['lower_tf']} | "
                     f"{rules_line(cfg)}")
    return "\n".join(lines)


# =====================================================================
# state
# =====================================================================
def load_state(path):
    try:
        with open(path, encoding="utf-8") as f:
            st = json.load(f)
        if isinstance(st, dict) and st.get("v") == STATE_VERSION \
                and isinstance(st.get("last_bar"), dict):
            st.setdefault("test_id", None)
            return st
        log.warning("state file has an unknown format - starting fresh")
    except FileNotFoundError:
        log.info("no state file yet - first run")
    except Exception as e:  # noqa: BLE001
        log.warning("cannot read state (%s) - starting fresh", e)
    return {"v": STATE_VERSION, "last_bar": {}, "first_done": False,
            "last_run_utc": None, "test_id": None}


def save_state(path, st):
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, path)


# =====================================================================
# telegram
# =====================================================================
async def tg_send(token, chat, text, dry, proxy=""):
    """True only when Telegram really accepted the message (or in --dry-run)."""
    if dry:
        log.info("[dry-run] would send:\n%s", text)
        return True
    if not token or not chat:
        log.error("Telegram NOT configured: secrets TG_TOKEN / TG_CHAT are empty - "
                  "message not sent")
        return False
    import aiohttp
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    for attempt in range(3):
        try:
            async with aiohttp.ClientSession() as s:
                async with s.post(url, json={"chat_id": chat, "text": text[:4000]},
                                  proxy=proxy or None,
                                  timeout=aiohttp.ClientTimeout(total=30)) as r:
                    if r.status == 200:
                        return True
                    log.warning("Telegram HTTP %s: %s", r.status, (await r.text())[:160])
                    if r.status in (400, 401, 403, 404):   # wrong token / chat id: retry is useless
                        return False
        except Exception as e:  # noqa: BLE001
            log.warning("Telegram send failed (try %d/3): %s", attempt + 1, e)
        await asyncio.sleep(2 * (attempt + 1))
    return False


def format_signal(side, asset, tf, ltf, B, info):
    """Spot: 2/15/40/50  (confirming / checked / market weight % / total weight %)."""
    head = {"BUY": "🟢 BUY", "SELL": "🔴 SELL", "GLOBAL LONG": "🔵 GLOBAL LONG",
            "GLOBAL SHORT": "🟡 GLOBAL SHORT"}[side]
    lines = [f"{head}  {asset}/{app.asset_quote(asset)}  {tf} / {ltf}"]
    if info:
        total = info["pct"]
        for kind in FAMILIES:
            t = info.get("pairs", {}).get(kind)
            if t:
                lines.append(f"{FAM_NAME[kind]}: {t[0]}/{t[1]}/{t[2]}/{total}")
    lines.append(f"Bar open: {app.fmt_time(B - app.parse_tf_minutes(tf) * MINUTE_MS)} (Tehran)")
    return "\n".join(lines)


def probe_text(res):
    """Near-miss numbers of every bar (for tuning settings.toml)."""
    out = []
    for side in ("buy", "sell"):
        pairs, total = res["probe"][side]
        fam = " ".join(f"{FAM_NAME[k]} {t[0]}/{t[1]}/{t[2]}" for k, t in pairs.items())
        out.append(f"{side.upper()} {fam} all {total}%")
    return " | ".join(out)


async def maybe_send_test(settings, state, tg, dry):
    want = settings["telegram"]["test_message_id"]
    if state.get("test_id") == want:
        return
    now = datetime.now(TEHRAN).strftime("%Y-%m-%d %H:%M")
    text = ("✅ VC69V5 bot - Telegram test\n"
            f"Connection OK ({now}, Tehran time)\n\n"
            "Active settings:\n" + settings_text(settings) + "\n\n"
            "Signal lines look like:  Spot: 2/15/40/50\n"
            "2 = confirming exchanges, 15 = exchanges with data in that market,\n"
            "40 = their share of that market's volume %, 50 = share of all markets %.\n"
            "This test is sent once. To send it again, change test_message_id in settings.toml.")
    ok = await tg_send(tg["token"], tg["chat"], text, dry, tg["proxy"])
    if ok and not dry:
        state["test_id"] = want
        log.info("Telegram test message sent")
    elif not ok:
        log.error("Telegram TEST FAILED - check the TG_TOKEN / TG_CHAT secrets "
                  "(chat id: write to the bot first; groups/channels need the bot as member/admin)")


# =====================================================================
# one pass
# =====================================================================
async def timed(coro, seconds, what):
    try:
        return await asyncio.wait_for(coro, timeout=max(5, seconds))
    except asyncio.TimeoutError:
        log.warning("%s exceeded %ds - continuing with what was loaded", what, seconds)


async def one_pass(eng, settings, state_path, dry, tg):
    t0 = time.time()
    rt = settings["runtime"]
    deadline = t0 + rt["max_run_minutes"] * 60
    now_ms = int(time.time() * 1000)

    fr = rt["first_run_tehran"]
    if fr:
        try:
            fr_ms = int(datetime.strptime(fr, "%Y-%m-%d %H:%M")
                        .replace(tzinfo=TEHRAN).timestamp() * 1000)
        except ValueError:
            raise SettingsError(f"[runtime].first_run_tehran '{fr}': use 'YYYY-MM-DD HH:MM'")
        if now_ms < fr_ms:
            log.info("Before the configured first run (%s Tehran) - nothing to do.", fr)
            return

    state = load_state(state_path)
    await maybe_send_test(settings, state, tg, dry)
    save_state(state_path, state)

    for name, cfg in settings["assets"].items():
        log.info("settings %s: charts %s / LTF %s | %s", name, ",".join(cfg["chart_tfs"]),
                 cfg["lower_tf"], rules_line(cfg))
    await timed(eng.ensure_exchanges(settings["labels"], tg["proxy"]),
                rt["exchange_load_seconds"], "Loading exchanges")
    log.info("%d exchange(s) ready: %s", len(eng.exchanges), ", ".join(sorted(eng.exchanges)))

    total_bars = total_signals = 0
    for a, cfg in settings["assets"].items():
        if time.time() > deadline:
            log.warning("%s: run time limit (%d min) reached - skipped, it is picked up next run",
                        a, rt["max_run_minutes"])
            continue
        p = cfg["params"]
        p.tg_token, p.tg_chat, p.proxy = tg["token"], tg["chat"], tg["proxy"]
        p.tv_user, p.tv_pass = tg["tv_user"], tg["tv_pass"]
        L_min = app.parse_tf_minutes(p.lower_tf)
        L_ms = L_min * MINUTE_MS
        c_max = max(app.parse_tf_minutes(t) for t in p.chart_tfs)

        def first_last(tf):
            C = app.parse_tf_minutes(tf) * MINUTE_MS
            last = state["last_bar"].get(f"{a}|{tf}")
            if last is None:                     # first run: only the last few minutes
                last = (now_ms - rt["first_window_min"] * MINUTE_MS) // C * C
            cap = (now_ms - rt["max_lookback_min"] * MINUTE_MS) // C * C
            return max(int(last), cap)           # stale state never floods Telegram

        oldest_last = min(first_last(tf) for tf in p.chart_tfs)
        needed_min = (p.vol_sma_len + 1) * c_max + L_min + 5 \
            + max(0, (now_ms - oldest_last) // MINUTE_MS)

        lst = eng.build_sources([a], p).get(a, [])
        if not lst:
            log.warning("%s: no sources - skipped", a)
            continue
        ccxt_lst = [s for s in lst if s.label.endswith(("-S", "-F"))]
        tv_lst = [s for s in lst if s not in ccxt_lst]
        sem = asyncio.Semaphore(app.MAX_CONCURRENCY)
        log.info("%s: downloading ~%d min of candles for %d sources...", a, needed_min, len(ccxt_lst))
        await timed(asyncio.gather(*[app.backfill(eng.exchanges[s.ex_label], s, needed_min, sem)
                                     for s in ccxt_lst]),
                    rt["backfill_seconds"], f"{a} candle download")
        if tv_lst:
            n_bars = min(app.TV_MAX_BARS, max(500, -(-needed_min // L_min) + 80))
            await timed(eng.tv_fill(tv_lst, p, n_bars), rt["tv_seconds"], f"{a} TradingView download")
        app.rescale_tv(lst)

        for tf in p.chart_tfs:
            C_ms = app.parse_tf_minutes(tf) * MINUTE_MS
            b_now = now_ms // C_ms * C_ms        # close time of the newest closed bar
            key = f"{a}|{tf}"
            last = first_last(tf)
            n_bars = n_sig = 0
            for B in range(last + C_ms, b_now + 1, C_ms):
                per_ex = []
                for s in lst:
                    h = app.source_hist(s, B, C_ms, L_ms, p)
                    if h is not None:
                        per_ex.append((s.label, h))
                if not per_ex:
                    continue
                res = app.combine(per_ex, p, a, tf, rules=cfg["rules"])
                n_bars += 1
                log.info("%s %s bar %s | %d sources | %s", a, tf,
                         app.fmt_time(B - C_ms), res["used"], probe_text(res))
                for side, info in app.signals_from(res):
                    text = format_signal(side, a, tf, p.lower_tf, B, info)
                    log.info("SIGNAL\n%s", text)
                    if await tg_send(tg["token"], tg["chat"], text, dry, tg["proxy"]):
                        n_sig += 1
                    else:
                        log.error("Telegram delivery FAILED for the signal above")
            log.info("CHECKED %s %s: %d closed bar(s), %d signal(s)", a, tf, n_bars, n_sig)
            total_bars += n_bars
            total_signals += n_sig
            state["last_bar"][key] = b_now
            save_state(state_path, state)        # progress is kept even if a later asset fails

    state["first_done"] = True
    state["last_run_utc"] = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    save_state(state_path, state)
    log.info("DONE in %.0fs: %d bar(s) evaluated, %d signal(s).", time.time() - t0,
             total_bars, total_signals)


# =====================================================================
def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description="vc69v5 headless signal checker (one pass)")
    ap.add_argument("--state", default=os.environ.get("STATE_FILE", "state/state.json"))
    ap.add_argument("--settings", default=os.environ.get("SETTINGS_FILE",
                                                         os.path.join(here, "settings.toml")))
    ap.add_argument("--dry-run", action="store_true", help="log signals, send nothing")
    ap.add_argument("--check-settings", action="store_true",
                    help="validate settings.toml, print it and exit")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s",
                        datefmt="%H:%M:%S")
    logging.getLogger("vc69v5").setLevel(logging.INFO)
    logging.getLogger("vc69v5.headless").setLevel(logging.INFO)

    tg = {"token": os.environ.get("TG_TOKEN", "").strip(), "chat": os.environ.get("TG_CHAT", "").strip(),
          "proxy": "", "tv_user": os.environ.get("TV_USER", ""), "tv_pass": os.environ.get("TV_PASS", "")}
    try:
        settings = load_settings(args.settings)
    except SettingsError as e:
        log.error("SETTINGS ERROR: %s", e)
        if not args.dry_run and not args.check_settings and tg["token"] and tg["chat"]:
            asyncio.run(tg_send(tg["token"], tg["chat"],
                                f"⚠️ VC69V5: settings.toml error - the bot did not run.\n{e}", False))
        sys.exit(2)

    log.info("settings OK (%s) | telegram=%s | tradingview=%s | exchanges=%d",
             os.path.basename(args.settings),
             "on" if tg["token"] and tg["chat"] and not args.dry_run else "OFF",
             "on" if settings["tv_enabled"] else "off", len(settings["labels"]))
    for line in settings_text(settings).splitlines():
        log.info("  %s", line)
    if args.check_settings:
        return

    # fail fast on blocked hosts instead of waiting minutes (GitHub runners are in the US)
    app.REQUEST_TIMEOUT_MS = 20000
    app.LOAD_RETRIES = 1

    code = {"v": 0}

    def watchdog():
        log.error("WATCHDOG: the run exceeded %d minutes - stopping (progress saved so far is kept)",
                  settings["runtime"]["max_run_minutes"] + 1)
        logging.shutdown()
        os._exit(3)

    wd = threading.Timer((settings["runtime"]["max_run_minutes"] + 1) * 60, watchdog)
    wd.daemon = True
    wd.start()

    eng = app.Engine(ui=lambda kind, payload: None)
    try:
        eng.submit(one_pass(eng, settings, args.state, args.dry_run, tg)).result()
    except SettingsError as e:
        log.error("SETTINGS ERROR: %s", e)
        code["v"] = 2
    except KeyboardInterrupt:
        log.warning("interrupted")
        code["v"] = 130
    except Exception:  # noqa: BLE001
        log.exception("run failed")
        code["v"] = 1
    finally:
        try:
            eng.submit(eng.shutdown()).result(timeout=20)
        except Exception:  # noqa: BLE001
            pass
    wd.cancel()
    logging.shutdown()
    sys.stdout.flush()
    os._exit(code["v"])    # do not wait for background download threads


if __name__ == "__main__":
    main()
