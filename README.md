# Daily Sales Reporting — leadership one-pager

Turns the manually-maintained sales sheets into one clean, decision-ready
picture — an **interactive HTML page** and a **WhatsApp-ready PNG** — instead of
a screenshot of a coloured spreadsheet.

It answers three questions at a glance:
1. **How much did we sell?** MTD GMV, by channel and total.
2. **Are we growing?** Like-for-like growth vs the same dates last month.
3. **Where will the month land, and how far is that from today?** Run-rate
   estimate and the remaining gap.

## Quick start

```
pip install openpyxl pillow
cd report_tool
python main.py            # builds HTML + PNG into ./output (+ runs QC)
python main.py --no-image # HTML only
python main.py --strict   # abort (no PNG) if QC fails
python qc.py              # run the QC self-check on its own
```

Outputs land in `report_tool/output/index.{html,png}` (fixed location, overwritten each run).
PNG rendering uses the Chrome/Edge already installed on Windows — no download.
The reporting month is resolved automatically (rolls July → August when
August data lands; falls back to the last populated month). See GUIDE.md.

## How the numbers are defined (mirrors the sheet)

- **GMV (MTD)** = sum of daily `Value` for days that have data.
- **Growth** = `(GMV − LM GMV) / LM GMV`, like-for-like over the **same dates**
  (so lagged/blank days never read as a drop).
- **Run-rate estimate** = `GMV_MTD ÷ days_elapsed × days_in_month`
  (`days_elapsed` = calendar day of the last data date — matches the sheet's
  Shopify `/20`, Amazon `/19`, Minutes `/19`).
- **Gap** = `estimate − GMV_MTD` (still to come at the current pace).

## Data handling rules

- **Daily channels** (Blinkit, Big Basket, Swiggy, Zepto, Amazon) get the date grid.
- **Amazon** arrives on a **T-2 lag** — its latest 1–2 days are pending, not a drop.
- **Monthly channels** (Amazon Core, Amazon NOW+Fresh, Flipkart, Minutes,
  First Club) are shown as MTD totals only — no invented daily split.
- **Shopify** is a daily D2C channel: GMV = `Value` (gross); `Total Sale` (net, after
  cancellations/discounts) and `Dolchi` are parsed but not shown.
- **Ad Sales** is shown in the summary and the per-platform daily tables, for the channels that
  report it (Blinkit, Big Basket, Swiggy, Zepto, Amazon). Blinkit & Amazon ad are T-2.
- Day-wise view = one table per platform (Date · Units · GMV · LM GMV · Ad Sales · Growth), GMV
  cell shaded by that day's growth, with Total + run-rate Estimate rows.
- The stray `xx` column and the sheet's hand-typed `AD39/AD40` cells are ignored.
- **No data bluffing** — only what the sheets contain is displayed.

## Files

| File | Role |
|------|------|
| `config.py` | period, sheet IDs, channel→cadence map, aliases |
| `data_source.py` | **read-only** parse of the workbooks (dynamic column mapping) |
| `transform.py` | MTD / growth / estimate / gap, per-channel as-of date |
| `render_html.py` | self-contained HTML one-pager |
| `render_image.py` | HTML → WhatsApp PNG (headless Chrome/Edge + autocrop) |
| `main.py` | orchestrates resolve → transform → render → QC |
| `qc.py` | 3-layer self-check: sheet→model, identities, model→HTML |
| `gsheets.py` | optional **read-only** Google Sheets → local snapshot |

## Switching to live Google Sheets (read-only)

Today the pipeline reads the local `.xlsx`. To read the live sheets instead:

```
pip install google-api-python-client google-auth
cd "…\Sales & Estimation\report_tool"                   # must be in this folder
python -c "import gsheets; gsheets.fetch_snapshots()"   # read-only Sheets-API pull
python main.py                                          # auto-uses snapshots if present
```

`gsheets.py` uses the **read-only Google Sheets API** (`values.get`) — the sheets
have "download disabled for viewers", which blocks Drive export even for a viewer,
but the Sheets API data read is allowed. It writes local snapshots and **never
edits the source sheets**. Share both sheets with the service account as **Viewer**;
see the header of `gsheets.py` and GUIDE.md §5 for the one-time setup.

## Why this beats a spreadsheet screenshot (for leadership)

- **Consistent & clean** daily — no stray colours, no half-cut columns, legible.
- **One source of truth** — numbers come straight from the sheet, no re-typing.
- **Faster** — one command regenerates both formats; no manual crop/paste.
- **Actionable first** — growth and pace-to-estimate up top, detail below.
- **Archivable** — every file is dated; the HTML supports detail the image can't.

Easy future add-ons: PDF one-pager, a hosted always-latest dashboard link, and
scheduled auto-generation/sending.
