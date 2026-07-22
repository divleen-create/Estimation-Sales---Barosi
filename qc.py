"""In-code QC self-check + Excel estimate-freshness audit.

Correctness layers (drive pass/fail):
  A. SHEET → MODEL   Every channel's GMV / LM / units must equal the sheet's
                     OWN Total-row cells (independent ground truth).
  B. IDENTITIES      Model is internally consistent — growth/estimate/gap
                     formulas hold; section totals = Σ channels; grand = Σ sections.
  C. MODEL → HTML    The HTML contains the headline + per-channel figures + daily
                     dates, and does NOT show excluded channels.

Advisory audit (does NOT fail the run):
  D. ESTIMATE FRESHNESS  We always compute the run-rate estimate from the days
                     of data we actually have (per platform: divisor = day-of-month
                     of that platform's latest data). We then compare against the
                     sheet's estimate cell and the day-count baked into its formula
                     — so you can see whether someone forgot to update the "days"
                     in Excel. Reported as text; re-run after data updates and it
                     re-checks automatically.

Usage:
    ok, results, freshness = qc.run(html_path)
"""
from __future__ import annotations
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import config
import fmt
from data_source import load_sheet_diagnostics
from transform import build_report

REL_TOL = 0.001  # 0.1%


def _close(a, b) -> bool:
    a, b = (a or 0), (b or 0)
    return abs(a - b) <= max(1.0, abs(b) * REL_TOL)


# --- correctness layers -----------------------------------------------------
def _layer_a_sheet_vs_model(model, diags, results):
    for s in model.sections:
        for c in s.channels:
            d = diags.get(c.name)
            if not d or d.total_gmv is None:
                results.append(("A sheet→model", f"{c.name}", False, "no sheet Total row found"))
                continue
            checks = [("GMV", c.gmv_mtd, d.total_gmv), ("LM", c.lm_mtd, d.total_lm),
                      ("units", c.units_mtd, d.total_units)]
            if c.has_ad and d.total_ad is not None:
                checks.append(("ad sales total", c.ad_mtd, d.total_ad))
            for key, mine, sheet in checks:
                results.append(("A sheet→model", f"{c.name} {key}", _close(mine, sheet),
                                f"model={mine:,.1f} sheet={(sheet or 0):,.1f}"))
            # Ad-contribution vs the sheet's own Ad/Value ratio, where it exists.
            if c.has_ad and d.ad_contri_excel is not None and c.ad_contri is not None:
                results.append(("A sheet→model", f"{c.name} ad-contri %",
                                _close(c.ad_contri, d.ad_contri_excel),
                                f"model={c.ad_contri*100:.2f}% excel={d.ad_contri_excel*100:.2f}%"))


def _layer_b_identities(model, results):
    for s in model.sections:
        for key, tot, chsum in [
            ("Σunits", s.totals.units_mtd, sum(c.units_mtd for c in s.channels)),
            ("ΣGMV", s.totals.gmv_mtd, sum(c.gmv_mtd for c in s.channels)),
            ("ΣLM", s.totals.lm_mtd, sum(c.lm_mtd for c in s.channels)),
            ("Σest", s.totals.estimate, sum(c.estimate for c in s.channels)),
        ]:
            results.append(("B identities", f"{s.name} {key}", _close(tot, chsum),
                            f"{tot:,.1f} vs {chsum:,.1f}"))
        for c in s.channels:
            results.append(("B identities", f"{c.name} gap=est-gmv",
                            _close(c.gap, c.estimate - c.gmv_mtd), ""))
            if c.growth is not None and c.lm_mtd:
                results.append(("B identities", f"{c.name} growth=(g-lm)/lm",
                                _close(c.growth, (c.gmv_mtd - c.lm_mtd) / c.lm_mtd), ""))
            if c.has_ad and c.ad_contri is not None and c.gmv_mtd:
                results.append(("B identities", f"{c.name} ad-contri=ad/gmv",
                                _close(c.ad_contri, c.ad_mtd / c.gmv_mtd), ""))
    for key, g, ssum in [
        ("GMV", model.grand.gmv_mtd, sum(s.totals.gmv_mtd for s in model.sections)),
        ("estimate", model.grand.estimate, sum(s.totals.estimate for s in model.sections)),
    ]:
        results.append(("B identities", f"grand {key}", _close(g, ssum), f"{g:,.1f} vs {ssum:,.1f}"))


def _layer_c_html(model, html_text, results):
    def present(label, s):
        results.append(("C model→html", label, s in html_text, f"looking for '{s}'"))
    present("headline GMV", fmt.gmv_auto(model.grand.gmv_mtd))
    present("headline estimate", fmt.gmv_auto(model.grand.estimate))
    present("headline gap", fmt.gmv_auto(model.grand.gap))
    present("Ad Sales column", "Ad Sales")
    present("Ad Contri column", "Ad Contri")
    any_ad = any(c.has_ad for s in model.sections for c in s.channels)
    if any_ad:  # at least one Ad Sales MTD value should be rendered
        c_ad = next(c for s in model.sections for c in s.channels if c.has_ad)
        present(f"{c_ad.name} Ad value", fmt.gmv_auto(c_ad.ad_mtd))
    for s in model.sections:
        for c in s.channels:
            present(f"{c.name} channel row", f'class="chan">{c.name}')
            present(f"{c.name} GMV cell", fmt.gmv_auto(c.gmv_mtd))
        for d in s.dates:
            present(f"{s.name} date {d:%d %b}", d.strftime("%d %b"))
    for ex in config.EXCLUDED_CHANNELS:
        results.append(("C model→html", f"{ex} excluded",
                        f'class="chan">{ex}' not in html_text, "must not be a channel row"))


# --- advisory: estimate freshness ------------------------------------------
@dataclass
class Freshness:
    channel: str
    data_days: int              # divisor our estimate uses (platform's latest day)
    our_estimate: float
    excel_divisor: Optional[int]
    excel_mult: Optional[int]
    excel_estimate: Optional[float]
    status: str                 # OK | STALE | NO_EXCEL | MISMATCH
    message: str


def _freshness(model, diags, days_in_month) -> list[Freshness]:
    out: list[Freshness] = []
    for s in model.sections:
        for c in s.channels:
            data_days = c.last_date.day if c.last_date else 0
            d = diags.get(c.name)
            if not d or d.est_divisor is None:
                out.append(Freshness(c.name, data_days, c.estimate, None, None, None,
                                     "NO_EXCEL",
                                     f"no Excel estimate cell; ours = {fmt.gmv_auto(c.estimate)} "
                                     f"(÷{data_days} days)"))
                continue
            div_ok = d.est_divisor == data_days
            mult_ok = d.est_mult == days_in_month
            val_ok = _close(d.est_value, c.estimate)
            if div_ok and mult_ok and val_ok:
                out.append(Freshness(c.name, data_days, c.estimate, d.est_divisor, d.est_mult,
                                     d.est_value, "OK",
                                     f"Excel updated & correct — ÷{d.est_divisor}×{d.est_mult}, "
                                     f"matches ({fmt.gmv_auto(c.estimate)})"))
            else:
                bits = []
                if not div_ok:
                    bits.append(f"Excel divides by {d.est_divisor} but data supports {data_days} days")
                if not mult_ok:
                    bits.append(f"Excel projects to {d.est_mult} days, month has {days_in_month}")
                bits.append(f"Excel {fmt.gmv_auto(d.est_value)} vs correct {fmt.gmv_auto(c.estimate)}")
                out.append(Freshness(c.name, data_days, c.estimate, d.est_divisor, d.est_mult,
                                     d.est_value, "STALE", "; ".join(bits)))
    return out


# --- runner -----------------------------------------------------------------
def compute_freshness(model, diags=None) -> list[Freshness]:
    """Estimate-freshness audit for a given model (loads sheet diagnostics)."""
    if diags is None:
        diags = load_sheet_diagnostics()
    return _freshness(model, diags, model.days_in_month)


def run(html_path: Optional[Path] = None, model=None, diags=None):
    model = model or build_report()
    diags = diags if diags is not None else load_sheet_diagnostics()
    results: list[tuple[str, str, bool, str]] = []
    _layer_a_sheet_vs_model(model, diags, results)
    _layer_b_identities(model, results)
    if html_path and Path(html_path).exists():
        _layer_c_html(model, Path(html_path).read_text(encoding="utf-8"), results)
    freshness = _freshness(model, diags, model.days_in_month)
    passed = all(ok for _, _, ok, _ in results)
    return passed, results, freshness


def print_report(results) -> None:
    layers: dict[str, list] = {}
    for layer, name, ok, detail in results:
        layers.setdefault(layer, []).append((name, ok, detail))
    for layer, items in layers.items():
        fails = [i for i in items if not i[1]]
        print(f"\n[{layer}]  {len(items) - len(fails)}/{len(items)} passed")
        for name, ok, detail in fails:
            print(f"   XX {name:32} {detail}")
    total, ok = len(results), sum(1 for _, _, o, _ in results if o)
    print(f"\nQC: {ok}/{total} checks passed", "✔" if ok == total else "✗ FAILURES ABOVE")


def print_freshness(freshness: list[Freshness]) -> None:
    icon = {"OK": "✔", "STALE": "⚠", "MISMATCH": "⚠", "NO_EXCEL": "·"}
    print("\n----- ESTIMATE FRESHNESS (our data-driven run-rate vs Excel's estimate) -----")
    for f in freshness:
        print(f"  {icon.get(f.status,'?')} {f.channel:16} {f.message}")
    stale = [f.channel for f in freshness if f.status in ("STALE", "MISMATCH")]
    if stale:
        print(f"\n  ➜ EXCEL NEEDS UPDATE (day-count not refreshed): {', '.join(stale)}")
    else:
        print("\n  ➜ Excel estimates are all up to date and match. ✔")


def summary_text(results, freshness: list[Freshness]) -> str:
    """Plain-text summary (QC result + estimate freshness) for saving/sharing."""
    total, ok = len(results), sum(1 for _, _, o, _ in results if o)
    lines = [f"QC self-check: {ok}/{total} checks passed "
             f"{'(all good)' if ok == total else '- FAILURES, see console'}",
             "",
             "Estimate freshness (our data-driven run-rate vs Excel's estimate cell):"]
    tag = {"OK": "[OK]", "STALE": "[STALE]", "MISMATCH": "[STALE]", "NO_EXCEL": "[--]"}
    for f in freshness:
        lines.append(f"  {tag.get(f.status,'[?]')} {f.channel}: {f.message}")
    stale = [f.channel for f in freshness if f.status in ("STALE", "MISMATCH")]
    lines += ["", ("EXCEL NEEDS UPDATE (day-count not refreshed): " + ", ".join(stale))
              if stale else "Excel estimates are all up to date and match."]
    return "\n".join(lines)


if __name__ == "__main__":
    html = config.OUTPUT_DIR / "index.html"
    passed, results, freshness = run(html if html.exists() else None)
    print_report(results)
    print_freshness(freshness)
    raise SystemExit(0 if passed else 1)
