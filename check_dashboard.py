"""Regenerate the dashboard from Poker standings.xlsx and check it against the spreadsheet.

Usage: python check_dashboard.py [--json reference.json]

Reads the spreadsheet directly (independently of the dashboard code), regenerates
poker-dashboard.html and index.html with update_dashboard.py, then checks:
  - every spreadsheet row made it into the generated data, unchanged
  - both generated files carry identical data
  - every pot reconciles (buy-ins = payouts)
  - every photo the dashboard points at exists
and prints reference figures (net, ROI, wins, best net vs biggest payout, drawdown,
rolling averages, ranking eligibility) to compare with what the page shows.
The spreadsheet is never written to.
"""
import argparse
import json
import re
import subprocess
import sys
import warnings
from collections import defaultdict
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
PATTERN = re.compile(r"/\*DATA\*/(.*?)/\*END DATA\*/", re.S)
MIN_RANKED = 3
failures = []


def check(ok, msg):
    print(("  ok    " if ok else "  FAIL  ") + msg)
    if not ok:
        failures.append(msg)


def sheet_rows():
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(HERE / "Poker standings.xlsx", data_only=True)
    rows = []
    for r in wb["Data"].iter_rows(min_row=4, values_only=True):
        _, date, num, player, buyin, won, _, place = r[:8]
        if date is None or player is None:
            continue
        rows.append({"num": num, "date": date.strftime("%Y-%m-%d"), "p": str(player).strip(),
                     "buyin": buyin or 0, "won": won or 0, "place": place or 1})
    hosts = {}
    if "Host" in wb.sheetnames:
        for r in wb["Host"].iter_rows(min_row=2, values_only=True):
            if r and r[0] is not None and len(r) > 1 and r[1]:
                hosts[r[0]] = str(r[1]).strip()
    return rows, hosts


def data_of(path):
    return json.loads(PATTERN.search(path.read_text(encoding="utf-8")).group(1))


def drawdown(nets):
    cum = peak = dd = 0
    for n in nets:
        cum += n
        peak = max(peak, cum)
        dd = max(dd, peak - cum)
    return dd


def player_stats(games, name):
    rows = [(g, e) for g in games for e in g["entries"] if e["p"] == name]
    nets = [e["won"] - e["buyin"] for _, e in rows]
    buy = sum(e["buyin"] for _, e in rows)
    paid = sum(e["won"] for _, e in rows)
    best = max(rows, key=lambda r: r[1]["won"] - r[1]["buyin"])
    big = max(rows, key=lambda r: r[1]["won"])
    rolling = [sum(nets[max(0, i - 2):i + 1]) / len(nets[max(0, i - 2):i + 1]) for i in range(len(nets))]
    shared = sum(1 for g, e in rows if e["place"] == 1 and sum(1 for o in g["entries"] if o["place"] == 1) > 1)
    return {
        "games": len(rows), "buyins": buy, "payouts": paid, "net": paid - buy,
        "roi_pct": round((paid - buy) / buy * 100) if buy else 0,
        "wins": sum(1 for _, e in rows if e["place"] == 1), "shared_wins": shared,
        "profitable": sum(1 for n in nets if n > 0), "payout_nights": sum(1 for _, e in rows if e["won"] > 0),
        "recovered": sum(1 for n in nets if n >= 0),
        "best_net": [best[0]["num"], nets[rows.index(best)]], "biggest_payout": [big[0]["num"], big[1]["won"]],
        "worst_net": min(nets), "max_drawdown": drawdown(nets),
        "rolling3_first3": [round(x, 2) for x in rolling[:3]], "rolling3_last": round(rolling[-1], 2),
        "avg_net": round((paid - buy) / len(rows), 2),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", help="also write the reference figures to this file")
    args = ap.parse_args()

    print("1. Regenerating from the spreadsheet")
    run = subprocess.run([sys.executable, str(HERE / "update_dashboard.py")], capture_output=True, text=True)
    check(run.returncode == 0, "update_dashboard.py ran" + ("" if run.returncode == 0 else f": {run.stderr.strip()}"))

    rows, hosts = sheet_rows()
    page, site = data_of(HERE / "poker-dashboard.html"), data_of(HERE / "index.html")
    games = sorted(page["games"], key=lambda g: (g["date"], g["num"]))

    print("2. Spreadsheet -> dashboard data")
    check(page == site, "poker-dashboard.html and index.html carry identical data")
    check(site_ok := (HERE / "index.html").read_text(encoding="utf-8").lstrip().lower().startswith("<!doctype html>"), "index.html is a standalone page")
    flat = sorted((g["num"], g["date"], e["p"], e["buyin"], e["won"], e["place"]) for g in games for e in g["entries"])
    src = sorted((r["num"], r["date"], r["p"], r["buyin"], r["won"], r["place"]) for r in rows)
    check(flat == src, f"all {len(src)} player results match the Data tab exactly")
    check(all(g.get("host", "") == hosts.get(g["num"], "") for g in games), "hosts match the Host tab")

    print("3. Pots and files")
    for g in games:
        pot, paid = sum(e["buyin"] for e in g["entries"]), sum(e["won"] for e in g["entries"])
        check(pot == paid, f"game {g['num']}: buy-ins ${pot} = payouts ${paid}")
    photos = [g["photo"] for g in games if g.get("photo")] + [p["photo"] for p in page.get("players", {}).values() if p.get("photo")]
    missing = [p for p in photos if not (HERE / p).is_file()]
    check(not missing, f"{len(photos)} photo references all exist" + (f" (missing: {missing})" if missing else ""))
    tracked = subprocess.run(["git", "ls-files", "photos"], cwd=HERE, capture_output=True, text=True).stdout.split()
    untracked = [p for p in photos if p not in tracked]
    check(not untracked, "every referenced photo is committed to git" + (f" (not yet: {untracked})" if untracked else ""))

    print("4. Reference figures")
    players = []
    for g in games:
        for e in g["entries"]:
            if e["p"] not in players:
                players.append(e["p"])
    total_buy = sum(e["buyin"] for g in games for e in g["entries"])
    print(f"  {len(games)} games, {len(src)} player results, ${total_buy:,} total buy-ins")
    ref = {"games": len(games), "results": len(src), "buyins": total_buy, "players": {}}
    seasons = sorted({g["date"][:4] for g in games})
    for p in players:
        s = player_stats(games, p)
        s["ranked_now"] = s["games"] >= MIN_RANKED
        s["season_net"] = {y: sum(e["won"] - e["buyin"] for g in games if g["date"][:4] == y for e in g["entries"] if e["p"] == p) for y in seasons}
        ref["players"][p] = s
        print(f"  {p:8} {s['games']:>2} games  net {s['net']:+5}  ROI {s['roi_pct']:+4}%  wins {s['wins']} ({s['shared_wins']} shared)"
              f"  best net {s['best_net'][1]:+} (G{s['best_net'][0]})  biggest payout ${s['biggest_payout'][1]} (G{s['biggest_payout'][0]})"
              f"  drawdown ${s['max_drawdown']}  {'ranked' if s['ranked_now'] else 'unranked'}")
    check(sum(s["net"] for s in ref["players"].values()) == 0, "everyone's net results sum to $0")

    if len(games) == 15 and games[-1]["num"] == 15:
        print("5. Original 15-game snapshot")
        w = ref["players"].get("Wendell", {})
        check(len(src) == 68, f"68 player results (got {len(src)})")
        check(total_buy == 3700, f"$3,700 total buy-ins (got ${total_buy:,})")
        check(w.get("net") == 360 and w.get("roi_pct") == 48, f"Wendell +$360 all-time, 48% ROI (got {w.get('net'):+}, {w.get('roi_pct')}%)")
        check(w.get("season_net", {}).get("2026") == 50, f"Wendell +$50 in 2026 (got {w.get('season_net', {}).get('2026'):+})")
    else:
        print("5. Snapshot checks skipped: the spreadsheet has moved on from the original 15 games")

    if args.json:
        Path(args.json).write_text(json.dumps(ref, indent=1), encoding="utf-8")
        print(f"Reference figures written to {args.json}")
    print("\nAll checks passed." if not failures else f"\n{len(failures)} check(s) failed.")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
