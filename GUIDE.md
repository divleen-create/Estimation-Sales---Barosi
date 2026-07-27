# GUIDE — how to update the report and get the screenshot

This is the day-to-day runbook. For what the numbers mean and the file layout,
see [README.md](README.md).

---

## TL;DR (the one command)

```powershell
cd "C:\Users\Divleen Chaudhary\Downloads\Sales & Estimation\report_tool"
python main.py
```

This does everything in order:
1. Resolves the reporting month automatically (July now, August later — see §4).
2. Builds/updates the HTML one-pager.
3. Runs the QC self-check (must say `QC: N/N checks passed ✔`).
4. Renders the WhatsApp PNG.

Then send the PNG:

```
report_tool\output\index.png
```

That PNG is the image to drop into WhatsApp. Open the `.html` when someone wants
the fuller view on a laptop.

---

## 1. One-time setup

```powershell
pip install openpyxl pillow
```

The PNG step uses the Chrome/Edge already installed on Windows — nothing else to
install. (Google-Sheets live mode needs extra packages; see §5.)

---

## 2. The normal daily run

1. Make sure the latest source data is in place:
   - **Local mode (default):** the two workbooks sit in the project folder
     (`Daily Sales Reporting.xlsx`, `Daily Sales Reporting-(New -Platform + D2C+ Amazon).xlsx`).
     Save the freshest copies there.
   - **Live Sheets mode:** run the read-only snapshot pull first (see §5).
2. Run:
   ```powershell
   cd "C:\Users\Divleen Chaudhary\Downloads\Sales & Estimation\report_tool"
   python main.py
   ```
3. Read the console output:
   - `HTML : …` and `PNG : …` are the files that were written/updated.
   - `QC: N/N checks passed ✔` means the figures in the sheet, the model, and the
     HTML all agree. If it says `✗ FAILURES ABOVE`, see §6 — **do not send** yet.
   - The `Headline —` block echoes the key numbers so you can eyeball them.
4. Attach `output\index.png` in WhatsApp and send.

Want HTML only (skip the image)? `python main.py --no-image`
Want it to stop instead of producing a PNG when QC fails? `python main.py --strict`

---

## 3. What the QC self-check verifies

`qc.py` runs automatically inside `main.py` (or on its own: `python qc.py`). Three
independent layers, so a mistake in one place is caught by another:

- **A — sheet → model:** re-reads the sheet and recomputes every channel's GMV,
  LM, units, growth, estimate and gap; they must match what the report computed.
- **B — identities:** growth/estimate/gap formulas hold; section totals = sum of
  channels; grand total = sum of sections.
- **C — model → HTML:** the HTML actually contains the headline figures, every
  channel's GMV, and every daily date — and does **not** show excluded channels.

You want `QC: N/N checks passed ✔` before sending.

---

## 4. Month rollover (July → August) — automatic

You don't change anything. Each run picks the **latest month tab that is
populated in both workbooks**:

- While only July has data → the report shows **July**.
- As soon as an **August** tab exists *and* has data in both sheets → the next
  run shows **August** by itself.
- If an August tab exists but is still **empty**, the report **falls back to
  July** and shows a note: *"August 2026 is not populated yet — showing
  July 2026 (last data available)."*

Tab names are matched flexibly (`July26`, `Aug26`, `August26`, `FEB 26`, …), so
whatever the team names the new tab will be recognised.

To **pin** a specific month (e.g. to re-issue June), edit `config.py`:
```python
FORCE_PERIOD = (2026, 6)   # set back to None for automatic
```

---

## 5. Switching to live Google Sheets (read-only, via service account)

The tool never edits the sheets — it only reads them. Because these sheets have
**download/export disabled for viewers**, the pull uses the **Google Sheets API**
(`values.get`), which a Viewer service account is allowed to read.

One-time setup:
1. `pip install google-api-python-client google-auth`
2. Google Cloud console → new project → enable **Google Sheets API**.
3. Create a **service account**, add a **JSON key**, download it.
4. **Share both sheets** with the service-account email (…iam.gserviceaccount.com) as **Viewer**.
5. Point the tool at the key:
   ```powershell
   $env:GOOGLE_APPLICATION_CREDENTIALS="C:\path\to\key.json"
   ```

Then, each refresh — **you must be inside the `report_tool` folder** (that's where
`gsheets.py` lives; running from `C:\Users\…>` gives `No module named 'gsheets'`):
```powershell
cd "C:\Users\Divleen Chaudhary\Downloads\Sales & Estimation\report_tool"
python -c "import gsheets; gsheets.fetch_snapshots()"   # read-only pull -> snapshots\
python main.py                                          # auto-uses the snapshots
```
Or just run `run_report.bat` (it `cd`s itself, so it works from any directory).

`fetch_snapshots()` pulls the recent month tabs of both sheets into
`report_tool\snapshots\*.xlsx` (values only — the estimate day-count is derived
from values, so the freshness check still works). Delete the `snapshots\` folder
to go back to the local files. Full details in the header of `gsheets.py`.

> Note: the two sheets name their July tab differently (`July'26` with an
> apostrophe on the live sheets vs `July26` in older local files) — the tool
> matches both. Reading via the Sheets API avoids the gviz/link pitfalls
> (silent wrong-tab fallback, blocked downloads) found when trying the raw link.

---

## 5b. Automate the daily refresh (Windows Task Scheduler)

`run_report.bat` does the whole thing in one shot: pull live (read-only) → build
`output\index.html` + `index.png` + QC summary. Double-click it to test, then
schedule it.

**Prerequisites for a *live* auto-run:** the service account (§5) must be set up
**and** its key must be visible to the scheduled task. Easiest is to set the env
var at the machine level (once, elevated PowerShell):
```powershell
setx GOOGLE_APPLICATION_CREDENTIALS "C:\path\to\key.json" /M
```
(Without it, the .bat still runs but reports on the local snapshot.)

**Register a daily 09:30 run** (one line, run once):
```powershell
schtasks /create /tn "BarosiDailySalesReport" /sc daily /st 09:30 ^
  /tr "\"C:\Users\Divleen Chaudhary\Downloads\Sales & Estimation\report_tool\run_report.bat\""
```
- Change `/st 09:30` to your preferred time. The PC must be on/awake then.
- Check it: `schtasks /query /tn "BarosiDailySalesReport"` · run now: `schtasks /run /tn "BarosiDailySalesReport"` · remove: `schtasks /delete /tn "BarosiDailySalesReport" /f`.

**Note on delivery:** this produces the PNG automatically; **sending it to WhatsApp**
still needs a human (or a separate WhatsApp Business API integration, which is a
different setup). The scheduled task leaves `output\index.png` ready to forward.

## 6. If QC fails or something looks off

- **QC `✗` on layer A/B:** a calculation regressed — re-run; if it persists, the
  sheet layout may have changed (a new column/channel). Check `config.py` aliases.
- **QC `✗` on layer C:** the HTML didn't contain an expected figure — usually a
  formatting change; re-run `python main.py`.
- **PNG is blank / clipped at the bottom:** increase `height` in
  `render_image.py` (`html_to_png(..., height=…)`); it auto-crops the extra.
- **`₹` looks wrong in the console:** cosmetic only (Windows console encoding);
  the HTML and PNG are correct.
- **A channel is missing / mis-named:** add its header spelling to
  `CHANNEL_ALIASES` / `FIELD_ALIASES` in `config.py`.
- **Amazon's last 1–2 days look empty:** expected — Amazon GMV is T-2; it is not
  a drop and is handled as such.

---

## 7. Cheat sheet

| I want to… | Command |
|---|---|
| Build HTML + PNG + QC | `python main.py` |
| HTML only | `python main.py --no-image` |
| Fail hard if QC fails | `python main.py --strict` |
| Run only the QC check | `python qc.py` |
| Inspect parsed data / resolved month | `python data_source.py` |
| Pull live Sheets (read-only) | `python -c "import gsheets; gsheets.fetch_snapshots()"` |

---

## 8. Cloud automation — the live web report (current production setup)

The daily **web** report is fully automated in the cloud and needs **no PC on**.
Windows Task Scheduler (§5b) is now only for generating the **WhatsApp PNG** locally.

**The flow (runs every day, hands-off):**
```
cron-job.org  ──11:30 IST──▶  GitHub API (workflow_dispatch)
                                     │
                                     ▼
   GitHub Actions runs daily-report.yml:
     live Sheets pull (required) → python main.py --no-image --strict
     → commit index.html + output/ → push → GitHub Pages redeploys
                                     │
                                     ▼
   Report live at: https://divleen-create.github.io/Estimation-Sales---Barosi/
```

### The pieces
- **Repo:** `github.com/divleen-create/Estimation-Sales---Barosi` — code, snapshots,
  and workflow all at the **root** of the `main` branch. Published via GitHub Pages.
- **Local ↔ remote:** this `report_tool/` folder is a git repo on branch `master`
  that pushes to `main` (`git push origin master:main`). The Actions bot commits
  "Auto-update daily report" to `main` on every run — so your local is always a
  step behind. **Always `git pull --rebase origin main` before you push.**
- **Trigger:** **cron-job.org** is the *sole* trigger (fires punctually at 11:30 IST
  via the `workflow_dispatch` API). GitHub's own `schedule:` cron was removed on
  purpose — it fired hours late and caused surprise mid-day re-publishes.
- **Live data:** the service-account JSON is stored as the GitHub repo secret
  **`GCP_SERVICE_ACCOUNT_KEY`**. The SA email must stay shared as **Viewer** on both
  Sheets and the **Sheets API** enabled. Pull is a **required** step — a bad key
  fails the run loudly instead of publishing stale data.
- **Trigger token:** fine-grained PAT `barosi-sales-cron-job-trigger`, scoped to this
  repo only, Actions: Read+write, stored in the cron-job.org job's `Authorization:
  Bearer …` header. **⚠ Expires 2027-07-24** — regenerate then (GitHub → Settings →
  Developer settings → Fine-grained tokens) and paste the new value into cron-job.org,
  or the daily trigger stops.

### Data-integrity guarantee
The build runs `--strict`: if the numbers don't reconcile against the sheet's own
Total row (QC layers A/B/C), the run **fails and does not publish** — the last good
report stays up and GitHub emails you the failure. So a *published* report has always
reconciled with the sheet. (The estimate-freshness/day-count check stays advisory and
never blocks.)

### Common tasks
- **Change the daily time:** edit the schedule in the **cron-job.org** job (time is in
  UTC there; 06:00 UTC = 11:30 IST). Don't add a GitHub `schedule:` back — it's unreliable.
- **Trigger a run now:** cron-job.org → your job → **Test run** (expect HTTP **204**),
  or GitHub → Actions → Daily Sales Report → **Run workflow**.
- **Edit the tool code:** change files here → `git pull --rebase origin main` →
  `git add … && git commit -m "…"` → `git push origin master:main`.
- **Page looks stale after a run:** GitHub Pages CDN cache — hard-refresh with
  **Ctrl+Shift+R**. ("Data as of" = latest date *in the sheet*; it only moves when the
  sheet gets a new day. "Generated … IST" = when the report was built.)
- **A run is red (failed):** open Actions → the run → the failed step. Usual causes:
  live-pull key/JSON problem, SA not shared on a sheet, or `--strict` caught a
  data mismatch — check the sheet before trusting numbers.

### Month slicer (history)
The report has a **Month dropdown** (top-right). It defaults to the current month; picking an
older month re-renders every KPI and table and updates the "Compared to <previous month>" caption.
- **Range:** `data_source.list_periods()` returns the **union** of months across both sheets,
  newest first (back to July 2023). A month present in only one sheet (Quick Commerce history
  predates the Marketplace & D2C sheet) renders only the section(s) that have data, with a note.
- **Comparison:** the current month uses the sheet's own **LM column** (QC ground truth). Older
  months' LM column is blank, so `transform.link_previous()` fills each month's baseline from the
  **actual previous month's data** we already built — so month-over-month growth works all the way
  back (earliest month shows "–", nothing before it). Current month is never altered by linking.
- **Build:** `build_report(period)` builds each month; the **current month is the only one QC/
  `--strict` gates**; historical months are best-effort (`SKIP history …` on a bad tab, never
  breaks the publish). `render_html.render_multi()` embeds one hidden pane per month + a tiny
  toggle script.
- **Cloud:** `gsheets.fetch_snapshots()` now pulls **all** month tabs (old 4-tab cap removed).
  ~37 months ≈ 1.3 MB `index.html` — fine for a static page (gzips small over the wire).

### Report UX (in each month pane)
- **Data notes · freshness** card (under the KPIs): per-platform latest data date, flagging any
  behind its expected lag (`config.FRESHNESS_LAG_DAYS`: T-1 default, Amazon T-2; `MONTHLY_PLATFORMS`
  = Flipkart, never flagged). Reference is "today" for the current month, month-end for past months
  — so it changes with the month picker. Built by `render_html._notes_card`.
- **Gap to estimate** is green (+ve) / red (−ve) everywhere via `render_html._gap_span`.
- **Date-wise detail** is one table + a **platform dropdown** spanning both sections
  (`render_html._daily_filter_card`, JS `__showPlat`). Every platform uses the same 7-column schema
  (ad columns show `–` when none); Amazon shows its Core/NOW+Fresh contribution strip.

### Trend chart (year-over-year)
Between Platform view and Date-wise detail sits a global **Trend** card: an inline-SVG line chart
(vanilla JS, no libraries) with a **KPI** selector (GMV/Units/Growth/Ad Sales/Estimation/Gap) and a
**Year** selector. It plots the selected year (blue solid `--tr-cur #12608a`) vs the year before it
(amber dashed `--tr-prev #d9a441`); the prior year is omitted if absent. Data (`{kpi:{year:[12]}}`
from each month's `.grand`) is embedded once as `#trendData` JSON; `render_html._TREND_JS` draws it.
To place one global chart between the two blocks, each month is split into a `_platform_pane` and a
`_daily_pane` (both `.month-pane`, toggled together by the month dropdown) with the Trend between.

### Still manual
- **WhatsApp PNG** — the cloud runner has no browser, so it skips the image
  (`--no-image`). Generate it locally when needed: `python main.py` → `output\index.png`.
