# HHGA Website

Static site for the High Handicapper Golf Association's annual golf trip
(2000–present). Vanilla HTML/CSS/JS, no build step — the pages fetch
`data/hhga.json` and render everything client-side.

## Pages

- `index.html` — reigning champion, all-time stats, champions wall
- `years.html` — per-year leaderboards, Mist-Weed results, hole-by-hole scorecards
- `players.html` — player profiles with finish-history charts

## Running locally

`fetch()` needs a server, so from the repo root:

```
python3 -m http.server
```

then open http://localhost:8000.

## Yearly update

1. Drop the final scores workbook into `HH/<year>/` (the `HH/` archive is
   local-only, never committed).
2. Add an `extract_<year>()` function to `scripts/extract.py` and register
   it in `EXTRACTORS` (recent years mostly reuse `_extract_oc_era`).
3. Run:
   ```
   python3 scripts/extract.py
   python3 scripts/summary.py
   ```
4. Review `data/years/<year>.json`, verify the champion, commit, push.

The site needs no changes — it renders whatever is in `data/hhga.json`.

## Data notes

- `data/aliases.json` maps every name spelling ever used in the workbooks
  to a canonical player id. Players marked `"phantom": true` (e.g. the
  Deadman placeholder) are excluded from rosters.
- Champion = lowest cumulative net across the trip (gross for 2000, the
  one year with no handicap data). Every champion 2000–2025 was verified
  by Ryan against the extracted data.
