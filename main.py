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
from transform import build_report, link_previous
from data_source import load_sheet_diagnostics, list_periods, source_fingerprint
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
    # Fingerprint the source workbooks BEFORE reading, so QC can prove the
    # read-only rule held (nothing was written to a source during the run).
    fp_before = source_fingerprint()
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
    # Fill month-over-month comparison for older months whose in-tab LM column is
    # blank, using the actual prior month's data (leaves the current month as-is).
    link_previous(models)
    print(f"MONTHS: {', '.join(m.month_label for m in models)}")

    # 1) build/update the HTML (current month is the default pane)
    html_path = write_html_multi(models, generated, freshness=freshness)
    print(f"HTML : {html_path}")

    # Cross-sheet GMV mismatches (both sheets have the channel/month but disagree).
    # Values were NOT summed — the richer series is kept — but flag it for an alert.
    conflicts = [c for m in models for c in m.merge_conflicts]
    conflict_file = config.OUTPUT_DIR / "sheet_conflicts.txt"
    if conflicts:
        print("\n⚠ CROSS-SHEET GMV MISMATCHES (QC vs MKT sheet):")
        for c in conflicts:
            print(f"   {c}")
            print(f"::warning title=Sheet mismatch::{c}")  # surfaces in GitHub Actions
        conflict_file.write_text("\n".join(conflicts), encoding="utf-8")
    elif conflict_file.exists():
        conflict_file.unlink()

    # 2) QC self-check: A sheet→model, B identities, C model→html, D parsing,
    #    E business edge cases, F page structure (+ advisory audits).
    ctx = qc.build_context(fp_before=fp_before)
    passed, results, freshness, notes = qc.run(html_path, model=model, diags=diags,
                                               models=models, ctx=ctx)
    qc.print_report(results)
    qc.print_freshness(freshness)
    qc.print_advisories(notes)
    summary_path = html_path.with_name(html_path.stem + "_qc_summary.txt")
    summary_path.write_text(
        qc.summary_text(results, freshness, conflicts=conflicts,
                        month_label=model.month_label, model=model, diags=diags,
                        notes=notes, generated=generated), encoding="utf-8")
    # HTML twin of the same validation — this is the daily email body.
    mail_path = html_path.with_name(html_path.stem + "_qc_summary.html")
    mail_path.write_text(
        qc.summary_html(results, freshness, conflicts=conflicts,
                        month_label=model.month_label, model=model, diags=diags,
                        notes=notes, generated=generated), encoding="utf-8")
    print(f"\nQC summary saved: {summary_path}\nQC mail body  : {mail_path}")
    if model.derived_parents:
        for d in model.derived_parents:
            print(f"DERIVED: {d}")
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
