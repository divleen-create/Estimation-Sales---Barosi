"""One command: read (read-only) -> transform -> render HTML + WhatsApp PNG.

    python main.py            # build both HTML and PNG from local .xlsx
    python main.py --no-image # HTML only (skip the browser screenshot)

Data source is selected in data_source.load_channels(); it reads the local
.xlsx snapshots today and never writes to any source.
"""
from __future__ import annotations
import argparse
import datetime as _dt
import sys

# Windows consoles default to cp1252 and choke on '₹'; force UTF-8 output.
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import config
import qc
from transform import build_report
from data_source import load_sheet_diagnostics, list_periods
from render_html import write_html_multi
from render_image import html_to_png
import fmt


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-image", action="store_true", help="skip PNG render")
    ap.add_argument("--strict", action="store_true",
                    help="abort (no PNG) if any QC check fails")
    args = ap.parse_args()

    # Stamp in IST (UTC+5:30) so the "Generated" time is correct no matter where
    # this runs — the GitHub runner's clock is UTC.
    IST = _dt.timezone(_dt.timedelta(hours=5, minutes=30))
    generated = _dt.datetime.now(IST)
    # Build the CURRENT month once (honors FORCE_PERIOD) — this is the month QC
    # gates and leadership acts on. Thread its model/diags through render + QC.
    model = build_report()
    diags = load_sheet_diagnostics()
    freshness = qc.compute_freshness(model, diags)
    if model.pending_note:
        print(f"NOTE : {model.pending_note}")

    # Build every OTHER available month (best-effort) for the history slicer.
    # A malformed old tab must never break the daily publish — skip and log it.
    cur_ym = (model.year, model.month)
    models = [model]
    for p in list_periods():
        if (p.year, p.month) == cur_ym:
            continue
        try:
            models.append(build_report(p))
        except Exception as e:  # noqa: BLE001 — historical month is non-critical
            print(f"SKIP history {p.label}: {e}")
    print(f"MONTHS: {', '.join(m.month_label for m in models)}")

    # 1) build/update the HTML (current month is the default pane)
    html_path = write_html_multi(models, generated, freshness=freshness)
    print(f"HTML : {html_path}")

    # 2) QC self-check: sheet→model, identities, model→html (+ estimate audit)
    passed, results, freshness = qc.run(html_path, model=model, diags=diags)
    qc.print_report(results)
    qc.print_freshness(freshness)
    summary_path = html_path.with_name(html_path.stem + "_qc_summary.txt")
    summary_path.write_text(qc.summary_text(results, freshness), encoding="utf-8")
    print(f"\nQC summary saved: {summary_path}")
    if not passed and args.strict:
        raise SystemExit("QC failed and --strict set; PNG not generated.")

    # 3) render the WhatsApp PNG
    if not args.no_image:
        png_path = html_path.with_suffix(".png")
        html_to_png(html_path, png_path)
        print(f"\nPNG  : {png_path}")

    g = model.grand
    print(f"\nHeadline — {model.month_label} (as of {model.as_of}):")
    print(f"  MTD GMV   : {fmt.gmv_auto(g.gmv_mtd)}  ({fmt.pct(g.growth)} vs LM)")
    print(f"  Estimate  : {fmt.gmv_auto(g.estimate)}  (run-rate, {model.days_in_month}-day month)")
    print(f"  Gap       : {fmt.gmv_auto(g.gap)}  still to come at this pace")


if __name__ == "__main__":
    main()
