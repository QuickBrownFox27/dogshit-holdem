"""Rebuild the results baked into poker-dashboard.html (and index.html for GitHub Pages)
from Poker standings.xlsx.

Usage: python update_dashboard.py [--next YYYY-MM-DD] [--host Name] [--name "Comp name"]
Settings (next game date, host, comp name) are kept from the last run unless given.
"""
import argparse
import json
import re
import warnings
from collections import OrderedDict
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
XLSX = HERE / "Poker standings.xlsx"
HTML = HERE / "poker-dashboard.html"
SITE = HERE / "index.html"  # standalone copy for GitHub Pages
SITE_HEAD = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<style>body{margin:0}img{max-width:100%}</style>
</head>
<body>
"""
SITE_TAIL = """
</body>
</html>
"""
PATTERN = re.compile(r"/\*DATA\*/(.*?)/\*END DATA\*/", re.S)


def read_games():
    warnings.simplefilter("ignore")
    ws = openpyxl.load_workbook(XLSX)["Data"]
    games = OrderedDict()
    for row in ws.iter_rows(min_row=4, values_only=True):
        _, date, num, player, buyin, won, _, place = row[:8]
        if date is None or player is None:
            continue
        g = games.setdefault(num, {"num": num, "date": date.strftime("%Y-%m-%d"), "entries": []})
        g["entries"].append({"p": str(player).strip(), "buyin": buyin or 0, "won": won or 0, "place": place or 1})
    # A photo for a game is picked up from photos/game<N>.<ext> ("game 13.jpg" gets renamed to "game13.jpg")
    photos = HERE / "photos"
    if photos.is_dir():
        for f in photos.iterdir():
            m = re.fullmatch(r"game[ _-]+(\d+)\.(webp|jpe?g|png)", f.name, re.I)
            if m:
                f.rename(photos / f"game{m.group(1)}.{m.group(2).lower()}")
    for g in games.values():
        for ext in ("webp", "jpg", "jpeg", "png"):
            path = photos / f"game{g['num']}.{ext}"
            if path.exists():
                g["photo"] = f"photos/game{g['num']}.{ext}"
                try:
                    from PIL import Image
                    with Image.open(path) as im:
                        g["photoW"], g["photoH"] = im.size
                except Exception:
                    pass
                break
    return list(games.values())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--next", help="next game date, YYYY-MM-DD (use '' to clear)")
    ap.add_argument("--name", help="comp name")
    ap.add_argument("--host", help="who is hosting the next game (use '' to clear)")
    args = ap.parse_args()

    html = HTML.read_text(encoding="utf-8")
    m = PATTERN.search(html)
    settings = json.loads(m.group(1)).get("settings", {})
    if args.next is not None:
        settings["nextGame"] = args.next
    if args.host is not None:
        settings["host"] = args.host
    if args.name is not None:
        settings["name"] = args.name

    games = read_games()
    data = json.dumps({"settings": settings, "games": games}, separators=(",", ":"))
    html = html[: m.start(1)] + data + html[m.end(1):]
    HTML.write_text(html, encoding="utf-8")
    SITE.write_text(SITE_HEAD + html + SITE_TAIL, encoding="utf-8")

    for g in games:
        pot = sum(e["buyin"] for e in g["entries"])
        paid = sum(e["won"] for e in g["entries"])
        flag = "" if pot == paid else f"  <-- payouts {paid} vs pot {pot}"
        print(f"Game {g['num']:>2}  {g['date']}  {len(g['entries'])} players  pot {pot}{flag}")
    print("settings:", settings)


if __name__ == "__main__":
    main()
