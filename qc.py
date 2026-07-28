"""In-code QC self-check: Excel-vs-output reconciliation, identities, HTML,
parsing, business edge cases, page structure — plus advisory audits.

Correctness layers (drive pass/fail; `--strict` blocks the publish on any fail):
  A. SHEET → MODEL   Every channel's GMV / LM / units / ad must equal the sheet's
                     OWN Total-row cells (independent ground truth), and ad-contri
                     must match the sheet's own Ad/Value ratio where it exists.
  B. IDENTITIES      Model is internally consistent — growth / estimate / gap /
                     ad-contri formulas hold; section totals = Σ channels;
                     grand = Σ sections (GMV, LM, units, ad, estimate, gap).
  C. MODEL → HTML    The HTML contains the headline + per-channel figures + daily
                     dates, and does NOT show excluded channels.
  D. PARSING         Every channel with activity has a GMV column parsed (catches
                     month-labelled headers like 'March GMV' being dropped → GMV 0).
  E. EDGE CASES      The agreed business rules, each as a check: sub-channel
                     roll-up (Amazon ≈ Core + NOW+Fresh), T-2 lag never read as
                     degrowth, ad-lag blanks stay blank, Shopify gross-not-net,
                     dropped columns really dropped, two-sheet merge overwrote
                     (never summed), conflicted channels withheld, no bluffing,
                     date-grid sanity, no stray 'xx'/'Total' channel.
  F. STRUCTURE       Page integrity: one pane + one dropdown option per month,
                     one daily table per daily platform (sub-channels get none),
                     Amazon contribution strip, trend dataset well-formed and
                     agreeing with the model, freshness card, pending-month note.

Advisory audits (reported, never block):
  * ESTIMATE FRESHNESS  Our run-rate (divisor = the platform's own latest data
                     day) vs the sheet's estimate cell and the day-count baked
                     into its formula — shows when Excel's "days" is stale.
  * SANITY ADVISORIES  Per-platform data freshness vs the expected lag, ad-lag
                     tails entered as 0 instead of blank, ad-contri and growth
                     outliers, negative daily cells, and the read-only guarantee
                     (source workbooks unchanged by the run).

Usage:
    ok, results, freshness = qc.run(html_path)                    # figures only
    ok, results, freshness, advisories = qc.run(html_path, model=m, diags=d,
                                                models=all_months, full=True)
"""
from __future__ import annotations
import calendar
import datetime as _dt
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import config
import fmt
from data_source import (cross_sheet_gmv, load_sheet_diagnostics, parsed_columns,
                         raw_headers, source_fingerprint)
from transform import build_report

REL_TOL = 0.001  # 0.1%
IST = _dt.timezone(_dt.timedelta(hours=5, minutes=30))


def _close(a, b, tol: float = REL_TOL) -> bool:
    a, b = (a or 0), (b or 0)
    return abs(a - b) <= max(1.0, abs(b) * tol)


def _today() -> _dt.date:
    """'Today' in IST — the report's reference clock, wherever it runs."""
    return _dt.datetime.now(IST).date()


# --- correctness layers -----------------------------------------------------
def _layer_a_sheet_vs_model(model, diags, results):
    # A parent we had to BUILD from its sub-channels has no Total-row cell of its
    # own to reconcile against — layer E checks it against Σ parts instead.
    built = {r.parent for r in model.rollups if r.derived}
    for s in model.sections:
        for c in s.channels:
            if c.name in built:
                results.append(("A sheet→model", f"{c.name} (derived)", True,
                                "no sheet column — built from sub-channels, checked in E"))
                continue
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
    dim = model.days_in_month
    for s in model.sections:
        for key, tot, chsum in [
            ("Σunits", s.totals.units_mtd, sum(c.units_mtd for c in s.channels)),
            ("ΣGMV", s.totals.gmv_mtd, sum(c.gmv_mtd for c in s.channels)),
            ("ΣLM", s.totals.lm_mtd, sum(c.lm_mtd for c in s.channels)),
            ("Σad", s.totals.ad_mtd, sum(c.ad_mtd for c in s.channels)),
            ("Σest", s.totals.estimate, sum(c.estimate for c in s.channels)),
            ("Σgap", s.totals.gap, sum(c.gap for c in s.channels)),
        ]:
            results.append(("B identities", f"{s.name} {key}", _close(tot, chsum),
                            f"{tot:,.1f} vs {chsum:,.1f}"))
        for c in s.channels:
            results.append(("B identities", f"{c.name} gap=est-gmv",
                            _close(c.gap, c.estimate - c.gmv_mtd), ""))
            # Run-rate uses the platform's OWN latest data day as the divisor
            # (Amazon T-2 → ÷N-2 while quick commerce → ÷N), not a global day count.
            if c.last_date:
                want = c.gmv_mtd / c.last_date.day * dim
                results.append(("B identities", f"{c.name} est=gmv/{c.last_date.day}x{dim}",
                                _close(c.estimate, want),
                                f"{c.estimate:,.1f} vs {want:,.1f}"))
            # MTD figures must equal the sum of the daily cells that back them —
            # catches a daily grid that drifts from the summary card.
            sums = [("GMV", c.gmv_mtd, [d.gmv for d in c.daily]),
                    ("units", c.units_mtd, [d.units for d in c.daily]),
                    ("ad", c.ad_mtd, [d.ad_spend for d in c.daily])]
            if not model.lm_linked:   # a linked baseline has no per-day LM for every day
                sums.append(("LM", c.lm_mtd, [d.lm_gmv for d in c.daily]))
            for key, mtd, cells in sums:
                results.append(("B identities", f"{c.name} MTD {key}=Σdaily",
                                _close(mtd, sum(v for v in cells if v is not None)), ""))
            if c.growth is not None and c.lm_mtd:
                results.append(("B identities", f"{c.name} growth=(g-lm)/lm",
                                _close(c.growth, (c.gmv_mtd - c.lm_mtd) / c.lm_mtd), ""))
            if c.has_ad and c.ad_contri is not None and c.gmv_mtd:
                results.append(("B identities", f"{c.name} ad-contri=ad/gmv",
                                _close(c.ad_contri, c.ad_mtd / c.gmv_mtd), ""))
    for key, g, ssum in [
        ("GMV", model.grand.gmv_mtd, sum(s.totals.gmv_mtd for s in model.sections)),
        ("LM", model.grand.lm_mtd, sum(s.totals.lm_mtd for s in model.sections)),
        ("units", model.grand.units_mtd, sum(s.totals.units_mtd for s in model.sections)),
        ("ad", model.grand.ad_mtd, sum(s.totals.ad_mtd for s in model.sections)),
        ("estimate", model.grand.estimate, sum(s.totals.estimate for s in model.sections)),
        ("gap", model.grand.gap, sum(s.totals.gap for s in model.sections)),
    ]:
        results.append(("B identities", f"grand {key}", _close(g, ssum), f"{g:,.1f} vs {ssum:,.1f}"))
    if model.grand.lm_mtd:
        want = (model.grand.gmv_mtd - model.grand.lm_mtd) / model.grand.lm_mtd
        results.append(("B identities", "grand growth=(g-lm)/lm",
                        _close(model.grand.growth, want), f"{model.grand.growth} vs {want}"))


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


# --- layer E: business edge cases ------------------------------------------
@dataclass
class EdgeContext:
    """Extra sheet-level facts layer E needs (columns actually parsed, raw headers,
    each sheet's own per-channel GMV before the merge, source fingerprint)."""
    columns: dict = field(default_factory=dict)
    headers: dict = field(default_factory=dict)
    cross: dict = field(default_factory=dict)
    fp_before: Optional[dict] = None
    today: Optional[_dt.date] = None


def build_context(period=None, fp_before: Optional[dict] = None) -> EdgeContext:
    return EdgeContext(columns=parsed_columns(period), headers=raw_headers(period),
                       cross=cross_sheet_gmv(period), fp_before=fp_before,
                       today=_today())


def _conflicted_names(model) -> set:
    """Channel names withheld this month because the two sheets disagreed.
    Conflict lines read '<month> · <channel>: QC ₹x vs MKT ₹y'."""
    out = set()
    for conf in model.merge_conflicts:
        tail = str(conf).split("·")[-1]
        out.add(tail.split(":")[0].strip())
    return out


def _layer_e_edges(model, ctx: EdgeContext, results):
    """The agreed business rules, each as a pass/fail check. Everything here is
    guaranteed by our own code or config, so a failure means a real regression
    (layout change, rule broken) — not merely late data."""
    def chk(name, ok, detail=""):
        results.append(("E edge cases", name, bool(ok), detail))

    chans = {c.name: c for s in model.sections for c in s.channels}
    today = ctx.today or _today()
    # Parents built from their sub-channels have no column of their own, so the
    # sheet-column rules (ad column present, merge = one sheet) don't apply to them.
    built = {r.parent for r in model.rollups if r.derived}

    # 1) Sub-channel roll-up, both directions:
    #    - parent column ENTERED in the sheet -> it is the number we report; the
    #      parts only give the split, and parent ≈ Σ parts within ROLLUP_TOL.
    #    - parent column ABSENT -> it was built by summing the parts, so parent
    #      must equal Σ parts exactly.
    #    Either way the parts stay sub-channels (no day-wise table of their own).
    for r in model.rollups:
        tol = REL_TOL if r.derived else config.ROLLUP_TOL
        how = "derived from parts" if r.derived else f"entered; tol {tol:.0%}"
        for key, pv, sv in (("GMV", r.entered_gmv, r.parts_gmv),
                            ("LM", r.entered_lm, r.parts_lm)):
            chk(f"{r.parent} {'=' if r.derived else '≈'} Σ sub-channels {key}",
                _close(sv, pv, tol),
                f"parts={sv:,.0f} parent={pv:,.0f} ({how})")
        shares = [s for _, _, s in r.parts if s is not None]
        chk(f"{r.parent} contribution shares are sane",
            all(0 <= s <= 1.0001 for s in shares),
            "; ".join(f"{n} {s*100:.1f}%" for n, _, s in r.parts if s is not None))
        if r.derived:
            # Built parent: every DAY must be the sum of that day's parts, and a
            # day where all parts are blank must stay blank (no phantom zeros).
            p = chans.get(r.parent)
            part_ch = [chans[n] for n, _, _ in r.parts if n in chans]
            bad_day, phantom = [], []
            for cell in (p.daily if p else []):
                vals = [d.gmv for x in part_ch for d in x.daily
                        if d.date == cell.date and d.gmv is not None]
                if vals and not _close(cell.gmv, sum(vals)):
                    bad_day.append(f"{cell.date:%d %b}")
                if not vals and cell.gmv is not None:
                    phantom.append(f"{cell.date:%d %b}")
            chk(f"{r.parent} daily = Σ parts per day", not bad_day,
                f"off on: {', '.join(bad_day[:5])}")
            chk(f"{r.parent} blank when all parts blank", not phantom,
                f"phantom value on: {', '.join(phantom[:5])}")
    for parent, subs in config.SUBCHANNELS.items():
        for x in [chans[s] for s in subs if s in chans]:
            chk(f"{x.name} is a sub-channel (no own daily table)", x.cadence == "sub",
                f"cadence={x.cadence}")

    # 2) Per-channel invariants.
    for name, c in chans.items():
        gmv_cells = [d for d in c.daily if d.gmv is not None]
        # A blank day must never read as a drop (Amazon T-2 / any pending day).
        blank_growth = [d.date.strftime("%d %b") for d in c.daily
                        if d.gmv is None and d.growth is not None]
        chk(f"{name} blank day never shows growth", not blank_growth,
            f"offending: {', '.join(blank_growth[:5])}")
        chk(f"{name} days_with_data = GMV cells",
            c.days_with_data == len(gmv_cells), f"{c.days_with_data} vs {len(gmv_cells)}")
        # No bluffing: no figure may exist without a day of data behind it.
        chk(f"{name} no figure without data",
            c.days_with_data > 0 or (c.gmv_mtd == 0 and c.estimate == 0),
            f"days={c.days_with_data} gmv={c.gmv_mtd:,.0f} est={c.estimate:,.0f}")
        chk(f"{name} MTD figures non-negative",
            min(c.gmv_mtd, c.units_mtd, c.ad_mtd, c.lm_mtd) >= 0,
            f"gmv={c.gmv_mtd:,.0f} units={c.units_mtd:,.0f} ad={c.ad_mtd:,.0f}")
        # Date grid: unique, ascending, inside the reporting month, never ahead of today.
        ds = [d.date for d in c.daily]
        chk(f"{name} dates unique & ascending", ds == sorted(ds) and len(ds) == len(set(ds)),
            f"{len(ds)} rows, {len(set(ds))} distinct")
        chk(f"{name} dates inside {model.month_label}",
            all(d.year == model.year and d.month == model.month for d in ds),
            f"first={ds[0] if ds else '-'} last={ds[-1] if ds else '-'}")
        future = [d.date.strftime("%d %b") for d in gmv_cells if d.date > today]
        chk(f"{name} no future-dated GMV", not future, f"{', '.join(future[:5])}")
        chk(f"{name} last data day ≤ month length",
            not c.last_date or c.last_date.day <= model.days_in_month,
            f"last={c.last_date} dim={model.days_in_month}")

    # 3) Lag channels (Amazon GMV T-2): the pending tail must stay BLANK, not 0.
    for name, lag in config.LAG_DAYS.items():
        c = chans.get(name)
        if not c or not c.last_date:
            continue
        tail = [d for d in c.daily if d.date > c.last_date]
        zeros = [d.date.strftime("%d %b") for d in tail if d.gmv is not None]
        chk(f"{name} T-{lag} tail is pending, not zero", not zeros,
            f"days after {c.last_date:%d %b} carrying a value: {', '.join(zeros[:5])}")

    # 4) Ad columns: every channel expected to carry Ad Sales still does. Losing
    #    one silently shows Ad Sales 0 — the same failure mode as the GMV bug.
    for name in sorted(config.AD_CHANNELS & set(chans) - built):
        chk(f"{name} Ad Sales column present", chans[name].has_ad,
            f"ad_mtd={chans[name].ad_mtd:,.0f}")

    # 5) Dropped source columns really are dropped (Shopify 'Total Sale' = net,
    #    'Dolchi' = sub-brand split): GMV must stay the gross 'Value' column.
    for sheet, chmap in ctx.columns.items():
        for name, flds in chmap.items():
            bad = sorted(set(flds) & config.DROPPED_FIELDS)
            chk(f"{name} [{sheet}] no dropped column parsed", not bad,
                f"leaked: {bad}; parsed={sorted(flds)}")
    for sheet in ("MKT", "QC"):
        sh = ctx.columns.get(sheet, {}).get("Shopify")
        if sh:
            chk("Shopify GMV = gross 'Value' column", "gmv" in sh, f"parsed={sorted(sh)}")
            break

    # 6) Two-sheet merge OVERWROTE — the model equals one sheet, never their sum.
    conflicted = _conflicted_names(model)
    for name, c in chans.items():
        pair = ctx.cross.get(name)
        if not pair or name in built:      # a built parent is a sum by design
            continue
        sides = {k: v for k, v in pair.items() if v is not None}
        if not sides:
            continue
        matches_one = any(_close(c.gmv_mtd, v) for v in sides.values())
        is_sum = len(sides) == 2 and _close(c.gmv_mtd, sum(sides.values()))
        chk(f"{name} merge kept one sheet (never summed)", matches_one and not is_sum,
            "model={:,.0f} · {}".format(
                c.gmv_mtd, " · ".join(f"{k}={v:,.0f}" for k, v in sides.items())))
    # 7) A conflicted channel must be WITHHELD, and nothing else may go missing.
    for name in conflicted:
        chk(f"conflicted '{name}' withheld from output", name not in chans,
            "two sheets disagreed → must be blank")
    for name, pair in ctx.cross.items():
        if any(v for v in pair.values() if v):
            chk(f"{name} reached the output", name in chans or name in conflicted,
                f"sheet has GMV {pair} but the channel is absent")

    # 8) No stray group became a channel ('Total', 'xx', excluded names).
    allowed = set(config.CHANNEL_ALIASES.values())
    for name in chans:
        chk(f"'{name}' is a recognised channel",
            name in allowed and name not in config.EXCLUDED_CHANNELS
            and name.strip().lower() not in config.CHANNEL_SKIP, "")

    # 9) Month/period integrity.
    chk("days_in_month matches the calendar",
        model.days_in_month == calendar.monthrange(model.year, model.month)[1],
        f"{model.days_in_month}")
    latest = max((c.last_date for c in chans.values() if c.last_date), default=None)
    chk("'Data as of' = latest data date anywhere", model.as_of == latest,
        f"as_of={model.as_of} latest={latest}")
    chk("current month is not LM-linked (uses the sheet's own LM column)",
        not model.lm_linked, "LM must come from the sheet for the reported month")


# --- layer F: page structure ------------------------------------------------
def _layer_f_structure(models, html_text, results, complete: bool = True):
    """The rendered page matches the model set: one pane and one dropdown option
    per month, one daily table per daily platform (sub-channels get none), the
    Amazon contribution strip, a well-formed trend dataset, freshness card.

    `complete` says whether `models` is the FULL set the page was rendered from;
    when it is not (e.g. `python qc.py` builds only the current month), the
    count-based checks are skipped rather than failed."""
    def chk(name, ok, detail=""):
        results.append(("F structure", name, bool(ok), detail))

    if not models or not html_text:
        return
    m0 = models[0]
    if complete:
        n_panes = html_text.count('class="month-pane"')
        chk("month panes = 2 per month (platform + daily)", n_panes == 2 * len(models),
            f"found {n_panes} for {len(models)} months")
        want_daily = sum(1 for m in models for s in m.sections
                         for c in s.channels if c.cadence == "daily")
        n_tbl = html_text.count('class="ptbl-pane"')
        chk("one daily table per platform-month", n_tbl == want_daily,
            f"{n_tbl} panes vs {want_daily} daily platforms")
    for i, m in enumerate(models):
        seen = html_text.count('data-idx="%d"' % i)
        chk(f"{m.month_label} pane pair", seen == 2, f"found {seen}")
        chk(f"{m.month_label} in the month dropdown",
            f">{m.month_label}</option>" in html_text, "")
    chk("current month is the default pane",
        bool(re.search(r'<option value="0" data-prev="[^"]*" selected>', html_text)), "")
    opts = re.findall(r'<select id="platPick"[^>]*>(.*?)</select>', html_text, re.S)
    for sub in [s for v in config.SUBCHANNELS.values() for s in v]:
        chk(f"{sub} not offered as its own daily table",
            all(f">{sub}<" not in o for o in opts), "sub-channels roll into the parent")
    for parent, subs in config.SUBCHANNELS.items():
        p = next((c for s in m0.sections for c in s.channels if c.name == parent), None)
        if p and p.gmv_mtd and any(c.name in subs for s in m0.sections for c in s.channels):
            chk(f"{parent} contribution strip rendered",
                f"Contribution to {parent} GMV" in html_text, "")
    chk("freshness / data-notes card rendered", "Data notes · freshness" in html_text, "")
    if m0.pending_note:
        chk("pending-month note rendered", m0.pending_note in html_text, "")

    blob = re.search(r'<script id="trendData" type="application/json">(.*?)</script>',
                     html_text, re.S)
    chk("trend dataset embedded", bool(blob), "")
    if not blob:
        return
    try:
        data = json.loads(blob.group(1))
    except Exception as e:  # noqa: BLE001
        chk("trend dataset is valid JSON", False, str(e))
        return
    chk("trend dataset is valid JSON", True, "")
    chk("trend has the consolidated 'All' series", "All" in data, "")
    malformed = [f"{p}/{k}/{y}" for p, kp in data.items() for k, ys in kp.items()
                 for y, arr in ys.items() if len(arr) != 12]
    chk("every trend series has 12 month slots", not malformed,
        f"{len(malformed)} malformed: {malformed[:3]}")
    missing = [c.name for s in m0.sections for c in s.channels
               if c.cadence == "daily" and c.name not in data]
    chk("every daily platform has a trend series", not missing, f"missing: {missing}")
    for kpi, attr in (("gmv", "gmv_mtd"), ("units", "units_mtd"), ("ad", "ad_mtd"),
                      ("estimate", "estimate"), ("gap", "gap")):
        v = (data.get("All", {}).get(kpi, {}).get(str(m0.year)) or [None] * 12)[m0.month - 1]
        chk(f"trend '{kpi}' current month = model", _close(v, getattr(m0.grand, attr)),
            f"chart={v} model={getattr(m0.grand, attr)}")


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


def _layer_d_parsing(model, results):
    """Parsing completeness: any channel with activity must have a GMV column
    parsed. Catches month-labelled headers ('March GMV', 'Total Value', 'Last
    Month GMV') being silently dropped -> units present but zero GMV days."""
    for s in model.sections:
        for c in s.channels:
            ok = not (c.units_mtd and c.days_with_data == 0)
            results.append(("D parsing", f"{c.name} GMV parsed",
                            ok, f"units={c.units_mtd:,.0f} gmv_days={c.days_with_data}"))


# --- advisory: sanity checks that must never block the publish --------------
@dataclass
class Advisory:
    level: str      # OK | WARN | INFO
    name: str
    message: str


def advisories(model, models=None, ctx: Optional[EdgeContext] = None) -> list[Advisory]:
    """Things worth knowing that are NOT our bug and must not stop the report:
    late data, sheet oddities, distorted growth from a lagging tail, outliers,
    and the read-only guarantee."""
    out: list[Advisory] = []
    ctx = ctx or EdgeContext(today=_today())
    today = ctx.today or _today()
    chans = {c.name: c for s in model.sections for c in s.channels}

    # 1) Data freshness per platform (same rule as the page's notes card).
    is_current = (model.year, model.month) == (today.year, today.month)
    ref = today if is_current else _dt.date(model.year, model.month, model.days_in_month)
    behind = []
    for name, c in chans.items():
        if c.cadence != "daily" or name in config.MONTHLY_PLATFORMS:
            continue
        lag = config.FRESHNESS_LAG_DAYS.get(name, config.FRESHNESS_LAG_DAYS["_default"])
        expected = ref - _dt.timedelta(days=lag)
        if c.last_date is None:
            behind.append(f"{name}: no data")
        elif c.last_date < expected:
            behind.append(f"{name}: {c.last_date:%d %b} (expected {expected:%d %b}, T-{lag})")
    out.append(Advisory("WARN" if behind else "OK", "Data freshness (expected lag)",
                        "; ".join(behind) if behind
                        else f"all platforms current as of {ref:%d %b} (T-1, Amazon T-2)"))

    # 2) Lag tail vs growth: totals are full-column (matching the sheet's Total
    #    row), so a T-2 platform's LM includes days whose GMV has not landed —
    #    report the like-for-like number beside it so nobody reads a false drop.
    for name in config.LAG_DAYS:
        c = chans.get(name)
        if not c or not c.lm_mtd:
            continue
        lfl_lm = sum(d.lm_gmv for d in c.daily if d.gmv is not None and d.lm_gmv is not None)
        if not lfl_lm:
            continue
        full = c.growth
        lfl = (c.gmv_mtd - lfl_lm) / lfl_lm
        gap_days = len([d for d in c.daily if d.gmv is None and d.lm_gmv is not None])
        out.append(Advisory("INFO" if abs((full or 0) - lfl) > 0.02 else "OK",
                            f"{name} T-{config.LAG_DAYS[name]} growth view",
                            f"reported {fmt.pct(full)} (full column, = sheet Total) vs "
                            f"{fmt.pct(lfl)} like-for-like on the {c.days_with_data} days that "
                            f"have GMV; {gap_days} pending day(s) carry LM but no GMV"))

    # 3) Ad-lag tail entered as 0 rather than left blank (reads as "no spend").
    zero_tail = []
    for name, lag in config.AD_LAG_DAYS.items():
        c = chans.get(name)
        if not c or not c.has_ad or not c.last_date:
            continue
        cutoff = c.last_date - _dt.timedelta(days=lag - 1)
        zeros = [f"{d.date:%d %b}" for d in c.daily
                 if d.date >= cutoff and d.ad_spend == 0]
        if zeros:
            zero_tail.append(f"{name}: {', '.join(zeros)}")
    out.append(Advisory("WARN" if zero_tail else "OK", "Ad T-2 tail blank (not 0)",
                        "; ".join(zero_tail) + " — a typed 0 reads as 'no spend'"
                        if zero_tail else "pending ad days are blank, shown as '–'"))

    # 4) Outliers worth a human glance.
    hot = [f"{n}: {c.ad_contri*100:.1f}%" for n, c in chans.items()
           if c.ad_contri and c.ad_contri > 0.25]
    out.append(Advisory("WARN" if hot else "OK", "Ad contribution outliers (>25%)",
                        "; ".join(hot) if hot else "all channels within 25% of GMV"))
    wild = [f"{n} {d.date:%d %b} {d.growth*100:+.0f}%" for n, c in chans.items()
            for d in c.daily if d.growth is not None and abs(d.growth) > 3]
    out.append(Advisory("INFO" if wild else "OK", "Daily growth outliers (>±300%)",
                        "; ".join(wild[:8]) + (f" (+{len(wild)-8} more)" if len(wild) > 8 else "")
                        if wild else "no day swings beyond ±300%"))
    neg = [f"{n} {d.date:%d %b}" for n, c in chans.items() for d in c.daily
           if (d.gmv or 0) < 0 or (d.units or 0) < 0 or (d.ad_spend or 0) < 0]
    out.append(Advisory("WARN" if neg else "OK", "Negative daily cells",
                        "; ".join(neg[:8]) if neg else "none"))

    # 5) A channel that gained an ad column (not in config.AD_CHANNELS).
    extra = sorted(n for n, c in chans.items() if c.has_ad and n not in config.AD_CHANNELS)
    out.append(Advisory("INFO" if extra else "OK", "New ad-carrying channel",
                        (", ".join(extra) + " — add to config.AD_CHANNELS to make it expected")
                        if extra else "ad channels match config"))

    # 5b) A parent built from its parts inherits only what the parts carry.
    for r in getattr(model, "rollups", []):
        if not r.derived:
            continue
        p = chans.get(r.parent)
        out.append(Advisory("WARN", f"{r.parent} built from sub-channels",
                            f"the sheets carry no {r.parent} column, so {r.parent} = "
                            f"{' + '.join(n for n, _, _ in r.parts)} "
                            f"({fmt.gmv_auto(r.parts_gmv)}); contribution split "
                            + " · ".join(f"{n} {s*100:.1f}%" for n, _, s in r.parts
                                         if s is not None)
                            + ("" if (p and p.has_ad) else
                               f" — no Ad Sales for {r.parent} (the parts have none)")))

    # 6) Channels with nothing yet this month (reported as blank, not zero).
    idle = sorted(n for n, c in chans.items() if c.days_with_data == 0)
    out.append(Advisory("INFO" if idle else "OK", "Channels with no data this month",
                        ", ".join(idle) if idle else "every channel has at least one day"))

    # 7) Read-only guarantee: the source workbooks must be untouched by the run.
    if ctx.fp_before:
        after = source_fingerprint()
        changed = [p for p, v in ctx.fp_before.items() if after.get(p) != v]
        out.append(Advisory("WARN" if changed else "OK", "Sources unchanged (read-only)",
                            ("MODIFIED during the run: " + ", ".join(changed)) if changed
                            else f"{len(ctx.fp_before)} source file(s) byte-identical before/after"))
    else:
        out.append(Advisory("INFO", "Sources unchanged (read-only)",
                            "fingerprint not captured for this run"))

    # 8) The deliberately-dropped headers still exist in the sheet (rule is live).
    if ctx.headers:
        seen = {f for hdrs in ctx.headers.values() for _, f in hdrs}
        dropped_hdrs = sorted(h for h, canon in config.FIELD_ALIASES.items()
                              if canon in config.DROPPED_FIELDS and h in seen)
        out.append(Advisory("OK" if dropped_hdrs else "INFO", "Dropped columns present in sheet",
                            (", ".join(dropped_hdrs) + " — read and intentionally excluded")
                            if dropped_hdrs else "none found in this month's tabs"))

    # 9) History coverage for the month slicer.
    if models:
        out.append(Advisory("OK", "Months rendered",
                            f"{len(models)} month(s): {models[0].month_label} … "
                            f"{models[-1].month_label}"))
    return out


# --- runner -----------------------------------------------------------------
def compute_freshness(model, diags=None) -> list[Freshness]:
    """Estimate-freshness audit for a given model (loads sheet diagnostics)."""
    if diags is None:
        diags = load_sheet_diagnostics()
    return _freshness(model, diags, model.days_in_month)


def run(html_path: Optional[Path] = None, model=None, diags=None, models=None,
        ctx: Optional[EdgeContext] = None):
    """Run every correctness layer (A-F) plus the advisory audits.

    Returns (passed, results, freshness, advisories). `models` (all months) and
    `ctx` (sheet-level facts) enable layers E/F; without them the check falls
    back to the figure-level layers only.
    """
    model = model or build_report()
    diags = diags if diags is not None else load_sheet_diagnostics()
    ctx = ctx if ctx is not None else build_context()
    results: list[tuple[str, str, bool, str]] = []
    _layer_a_sheet_vs_model(model, diags, results)
    _layer_b_identities(model, results)
    html_text = ""
    if html_path and Path(html_path).exists():
        html_text = Path(html_path).read_text(encoding="utf-8")
        _layer_c_html(model, html_text, results)
    _layer_d_parsing(model, results)
    _layer_e_edges(model, ctx, results)
    _layer_f_structure(models or [model], html_text, results, complete=bool(models))
    freshness = _freshness(model, diags, model.days_in_month)
    notes = advisories(model, models, ctx)
    passed = all(ok for _, _, ok, _ in results)
    return passed, results, freshness, notes


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


def print_advisories(notes: list[Advisory]) -> None:
    icon = {"OK": "✔", "WARN": "⚠", "INFO": "·"}
    print("\n----- SANITY ADVISORIES (never block the publish) -----")
    for a in notes:
        print(f"  {icon.get(a.level,'?')} {a.name:34} {a.message}")


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


_LAYER_TITLES = [
    ("A", "Excel Total-row vs output (reconciliation)"),
    ("B", "Internal identities (growth / run-rate / gap / totals)"),
    ("C", "Output HTML carries every figure"),
    ("D", "GMV column parsed for every active channel"),
    ("E", "Business edge cases (lag, roll-up, merge, no-bluffing)"),
    ("F", "Page structure (months, daily tables, trend)"),
]


def _recon_table(model, diags) -> list[str]:
    """Side-by-side Excel (the sheet's own Total row) vs output, per channel —
    the numbers check, in the daily mail."""
    hdr = (f"  {'Channel':<18}{'metric':<7}{'EXCEL (sheet Total)':>21}"
           f"{'OUTPUT (report)':>18}{'diff':>12}  verdict")
    lines = [hdr, "  " + "-" * (len(hdr) - 2)]
    for s in model.sections:
        for c in s.channels:
            d = diags.get(c.name)
            rows = [("GMV", c.gmv_mtd, d.total_gmv if d else None),
                    ("LM", c.lm_mtd, d.total_lm if d else None),
                    ("units", c.units_mtd, d.total_units if d else None)]
            if c.has_ad:
                rows.append(("ad", c.ad_mtd, d.total_ad if d else None))
            for metric, mine, sheet in rows:
                if sheet is None:
                    lines.append(f"  {c.name:<18}{metric:<7}{'(no cell in sheet)':>21}"
                                 f"{mine:>18,.0f}{'-':>12}  n/a")
                    continue
                diff = mine - sheet
                lines.append(f"  {c.name:<18}{metric:<7}{sheet:>21,.0f}{mine:>18,.0f}"
                             f"{diff:>12,.0f}  {'MATCH' if _close(mine, sheet) else 'MISMATCH'}")
    g = model.grand
    lines += ["  " + "-" * (len(hdr) - 2),
              f"  {'GRAND TOTAL':<18}{'GMV':<7}{'':>21}{g.gmv_mtd:>18,.0f}",
              f"  {'':<18}{'LM':<7}{'':>21}{g.lm_mtd:>18,.0f}",
              f"  {'':<18}{'est':<7}{'':>21}{g.estimate:>18,.0f}",
              f"  {'':<18}{'gap':<7}{'':>21}{g.gap:>18,.0f}"]
    return lines


def summary_text(results, freshness: list[Freshness], conflicts=None,
                 month_label: str = "", model=None, diags=None,
                 notes: "list[Advisory] | None" = None, generated=None) -> str:
    """Plain-text QC report for saving / the daily email. Covers, in order:
    the verdict, per-layer tallies, the Excel-vs-output number reconciliation,
    any failures, the edge-case rules that were verified, cross-sheet
    mismatches, estimate freshness, and the sanity advisories."""
    total, ok = len(results), sum(1 for _, _, o, _ in results if o)
    layers: dict[str, list] = {}
    for layer, name, okk, detail in results:
        layers.setdefault(layer, []).append((name, okk, detail))

    def tally(prefix):
        items = [i for lyr, it in layers.items() if lyr.startswith(prefix) for i in it]
        return sum(1 for _, o, _ in items if o), len(items)

    head = "ALL GOOD" if ok == total else "FAILURES PRESENT — DO NOT TRUST NUMBERS"
    lines = [
        f"DAILY QC VALIDATION{(' — ' + month_label) if month_label else ''}",
        (f"Built {generated:%d %b %Y, %H:%M} IST" if generated else ""),
        f"Overall: {ok}/{total} checks passed  [{head}]",
        "",
        "WHAT WAS CHECKED",
    ]
    for prefix, title in _LAYER_TITLES:
        l_ok, l_n = tally(prefix)
        if l_n:
            lines.append(f"  [{'OK' if l_ok == l_n else 'FAIL'}] {title}: {l_ok}/{l_n}")
    if model is not None:
        g = model.grand
        lines += ["",
                  f"HEADLINE  MTD GMV {fmt.gmv_auto(g.gmv_mtd)} · {fmt.pct(g.growth)} vs LM · "
                  f"estimate {fmt.gmv_auto(g.estimate)} · gap {fmt.gmv_auto(g.gap)}"
                  + (f" · data as of {model.as_of:%d %b %Y}" if model.as_of else "")]

    fails = [(lyr, n, dt) for lyr, its in layers.items() for n, o, dt in its if not o]
    if fails:
        lines += ["", "FAILURES (must be fixed before the numbers are trusted):"]
        lines += [f"  [{lyr}] {n} — {dt}" for lyr, n, dt in fails]

    if model is not None and diags is not None:
        lines += ["", "NUMBERS — EXCEL vs OUTPUT (figures in ₹; Excel = the sheet's own "
                  "Total row)"] + _recon_table(model, diags)

    # Edge cases: name the rules that were verified, so the mail is evidence.
    e_items = layers.get("E edge cases", [])
    f_items = layers.get("F structure", [])
    if e_items or f_items:
        e_ok, e_n = tally("E")
        f_ok, f_n = tally("F")
        lines += ["", f"EDGE CASES ({e_ok}/{e_n}) and PAGE STRUCTURE ({f_ok}/{f_n}) — rules verified:",
                  "  - Sub-channel roll-up: Amazon ~ Amazon Core + Amazon NOW+Fresh (GMV & LM); "
                  "the parts get no daily table of their own, only the contribution strip.",
                  "  - T-2 lag (Amazon GMV): the pending tail stays BLANK, never 0, and a blank "
                  "day never renders a growth figure — so a lag is never read as a drop.",
                  "  - Ad Sales: every channel expected to carry an ad column still does "
                  f"({', '.join(sorted(config.AD_CHANNELS))}); Blinkit/Amazon ad is T-2.",
                  "  - Shopify GMV = gross 'Value'; 'Total Sale' (net) and 'Dolchi' are read "
                  "and deliberately dropped, never mixed into GMV.",
                  "  - Two-sheet merge OVERWRITES, never sums: each channel's GMV equals one "
                  "sheet's figure; a disagreement withholds the channel instead of guessing.",
                  "  - No bluffing: no channel shows a figure without a day of data behind it; "
                  "no negative MTD, no future-dated GMV, dates unique/ascending/inside the month.",
                  "  - Run-rate divisor is each platform's OWN latest data day (Amazon T-2 "
                  "divides by fewer days than quick commerce).",
                  "  - Stray sheet groups ('Total', 'xx') never become channels; every channel "
                  "the sheets carry reaches the output (or is a declared conflict).",
                  "  - Page: one pane + one dropdown option per month, one daily table per "
                  "daily platform, freshness card, and a trend series that agrees with the model."]

    if model is not None and model.rollups:
        lines += ["", "SUB-CHANNEL ROLL-UP (parent vs the sum of its parts)"]
        for r in model.rollups:
            src = ("BUILT from its parts (no parent column in the sheets)" if r.derived
                   else "ENTERED in the sheet (the entered figure is what we report)")
            lines += [f"  {r.parent}: {src}",
                      f"    entered/reported : {r.entered_gmv:>15,.0f}",
                      f"    Σ parts          : {r.parts_gmv:>15,.0f}  "
                      f"({' + '.join(n for n, _, _ in r.parts)})",
                      f"    GAP (entered-Σ)  : {r.gap:>15,.0f}"
                      + (f"  ({r.gap_pct*100:+.2f}% of {r.parent})" if r.gap_pct is not None else ""),
                      f"    LM gap           : {r.lm_gap:>15,.0f}",
                      "    split: " + " · ".join(
                          f"{n} {fmt.gmv_auto(g)} = {s*100:.1f}%" if s is not None else f"{n} –"
                          for n, g, s in r.parts)]

    lines += ["", "CROSS-SHEET CHECK (Quick Commerce sheet vs Marketplace & D2C sheet, same month):"]
    if conflicts:
        lines += [f"  MISMATCH (left blank in output, needs reconcile): {c}" for c in conflicts]
    else:
        lines += ["  No mismatches — where both sheets carry a channel, values agree."]

    lines += ["", "ESTIMATE FRESHNESS (our data-driven run-rate vs Excel's estimate cell):"]
    tag = {"OK": "[OK]", "STALE": "[STALE]", "MISMATCH": "[STALE]", "NO_EXCEL": "[--]"}
    for f in freshness:
        lines.append(f"  {tag.get(f.status,'[?]')} {f.channel}: {f.message}")
    stale = [f.channel for f in freshness if f.status in ("STALE", "MISMATCH")]
    lines += [("  => EXCEL NEEDS UPDATE (day-count not refreshed): " + ", ".join(stale))
              if stale else "  => Excel estimates are all up to date and match."]

    if notes:
        lines += ["", "SANITY ADVISORIES (informational — never block the publish):"]
        tg = {"OK": "[OK]", "WARN": "[WARN]", "INFO": "[INFO]"}
        lines += [f"  {tg.get(a.level,'[?]')} {a.name}: {a.message}" for a in notes]
    return "\n".join(l for l in lines if l is not None)


# --- HTML version of the daily QC mail --------------------------------------
_C = {  # (text, background) per state — inline styles, for email clients
    "ok":   ("#0b6b34", "#e7f6ec"),
    "fail": ("#a51f26", "#fdecec"),
    "warn": ("#8a5a00", "#fff6e0"),
    "info": ("#3f4b5b", "#eef1f5"),
}
_BRAND, _INK, _MUTED, _LINE = "#12608a", "#1a2430", "#5b6774", "#dfe5ea"
_REPORT_URL = "https://divleen-create.github.io/Estimation-Sales---Barosi/"


def _esc(s) -> str:
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _chip(text, kind="ok") -> str:
    fg, bg = _C.get(kind, _C["info"])
    return (f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:{bg};color:{fg};font:600 11px/1.6 Segoe UI,Arial,sans-serif;'
            f'white-space:nowrap">{_esc(text)}</span>')


def _tbl(head, rows, widths=None) -> str:
    """Simple email-safe table: inline-styled th/td, zebra rows."""
    ws = widths or [""] * len(head)
    th = "".join(
        f'<th style="text-align:{"right" if i else "left"};padding:7px 10px;'
        f'background:{_BRAND};color:#fff;font:600 12px Segoe UI,Arial,sans-serif;'
        f'border:0"{f" width={w}" if w else ""}>{_esc(h)}</th>'
        for i, (h, w) in enumerate(zip(head, ws)))
    body = []
    for n, r in enumerate(rows):
        bg = "#ffffff" if n % 2 else "#f7f9fb"
        tds = "".join(
            f'<td style="text-align:{"right" if i else "left"};padding:6px 10px;'
            f'border-bottom:1px solid {_LINE};color:{_INK};'
            f'font:{"600" if i == 0 else "400"} 12px Segoe UI,Arial,sans-serif">{cell}</td>'
            for i, cell in enumerate(r))
        body.append(f'<tr style="background:{bg}">{tds}</tr>')
    return (f'<table cellpadding="0" cellspacing="0" border="0" width="100%" '
            f'style="border-collapse:collapse;margin:6px 0 16px">'
            f'<thead><tr>{th}</tr></thead><tbody>{"".join(body)}</tbody></table>')


def _h2(text) -> str:
    return (f'<div style="margin:22px 0 2px;font:700 14px Segoe UI,Arial,sans-serif;'
            f'color:{_INK};border-left:4px solid {_BRAND};padding-left:8px">{_esc(text)}</div>')


def _note(text) -> str:
    return (f'<div style="font:400 11px Segoe UI,Arial,sans-serif;color:{_MUTED};'
            f'margin:0 0 6px">{text}</div>')


def summary_html(results, freshness: list[Freshness], conflicts=None, month_label: str = "",
                 model=None, diags=None, notes: "list[Advisory] | None" = None,
                 generated=None) -> str:
    """The same daily validation as summary_text(), formatted as an HTML email:
    headings, tables and colour-coded verdicts/highlights."""
    total, ok = len(results), sum(1 for _, _, o, _ in results if o)
    allgood = ok == total
    layers: dict[str, list] = {}
    for layer, name, okk, detail in results:
        layers.setdefault(layer, []).append((name, okk, detail))

    def tally(prefix):
        items = [i for lyr, it in layers.items() if lyr.startswith(prefix) for i in it]
        return sum(1 for _, o, _ in items if o), len(items)

    fg, bg = _C["ok"] if allgood else _C["fail"]
    out = [f'<div style="font-family:Segoe UI,Arial,sans-serif;max-width:900px;color:{_INK}">',
           f'<div style="background:{_BRAND};color:#fff;padding:14px 16px;border-radius:6px 6px 0 0">'
           f'<div style="font:700 18px Segoe UI,Arial,sans-serif">Daily Sales Report — QC validation</div>'
           f'<div style="font:400 12px Segoe UI,Arial,sans-serif;opacity:.9">'
           f'{_esc(month_label)}'
           + (f' · built {generated:%d %b %Y, %H:%M} IST' if generated else '') + '</div></div>',
           f'<div style="background:{bg};color:{fg};padding:12px 16px;'
           f'font:700 15px Segoe UI,Arial,sans-serif;border:1px solid {_LINE};border-top:0">'
           f'{"✔" if allgood else "✗"} {ok}/{total} checks passed — '
           f'{"ALL GOOD" if allgood else "FAILURES PRESENT · DO NOT TRUST THE NUMBERS"}</div>',
           f'<div style="border:1px solid {_LINE};border-top:0;border-radius:0 0 6px 6px;'
           f'padding:4px 16px 16px">']

    # Headline KPIs
    if model is not None:
        g = model.grand
        kpis = [("MTD GMV", fmt.gmv_auto(g.gmv_mtd)),
                ("vs last month", fmt.pct(g.growth)),
                ("Run-rate estimate", fmt.gmv_auto(g.estimate)),
                ("Gap to estimate", fmt.gmv_auto(g.gap)),
                ("Data as of", f"{model.as_of:%d %b %Y}" if model.as_of else "–")]
        cells = "".join(
            f'<td style="padding:10px 12px;border:1px solid {_LINE};background:#f7f9fb">'
            f'<div style="font:600 10px Segoe UI,Arial,sans-serif;color:{_MUTED};'
            f'text-transform:uppercase;letter-spacing:.4px">{_esc(k)}</div>'
            f'<div style="font:700 16px Segoe UI,Arial,sans-serif;color:'
            f'{_C["ok"][0] if (k == "vs last month" and (g.growth or 0) >= 0) else _INK}">'
            f'{_esc(v)}</div></td>' for k, v in kpis)
        out.append(f'<table cellpadding="0" cellspacing="0" border="0" width="100%" '
                   f'style="border-collapse:collapse;margin:14px 0 4px"><tr>{cells}</tr></table>')

    # What was checked
    out.append(_h2("What was checked"))
    out.append(_note("All six layers are blocking — the cloud build refuses to publish "
                     "if any of them fails."))
    rows = []
    for prefix, title in _LAYER_TITLES:
        l_ok, l_n = tally(prefix)
        if l_n:
            rows.append([_esc(f"{prefix} — {title}"), f"{l_ok}/{l_n}",
                         _chip("PASS" if l_ok == l_n else "FAIL",
                               "ok" if l_ok == l_n else "fail")])
    out.append(_tbl(["Layer", "Checks", "Result"], rows, ["", "90", "90"]))

    # Failures first when present
    fails = [(lyr, n, dt) for lyr, its in layers.items() for n, o, dt in its if not o]
    if fails:
        out.append(_h2("Failures — fix before trusting the numbers"))
        out.append(_tbl(["Check", "Layer", "Detail"],
                        [[_esc(n), _chip(lyr.split()[0], "fail"), _esc(dt)]
                         for lyr, n, dt in fails]))

    # Excel vs output
    if model is not None and diags is not None:
        out.append(_h2("Numbers — Excel vs output"))
        out.append(_note('“Excel” is the source sheet’s <b>own Total row</b>; “Output” is what '
                         'the report published. Figures in ₹.'))
        rows = []
        for s in model.sections:
            for c in s.channels:
                d = diags.get(c.name)
                metrics = [("GMV", c.gmv_mtd, d.total_gmv if d else None),
                           ("LM", c.lm_mtd, d.total_lm if d else None),
                           ("Units", c.units_mtd, d.total_units if d else None)]
                if c.has_ad:
                    metrics.append(("Ad Sales", c.ad_mtd, d.total_ad if d else None))
                for i, (metric, mine, sheet) in enumerate(metrics):
                    name = _esc(c.name) if i == 0 else ""
                    if sheet is None:
                        rows.append([name, metric, "—", f"{mine:,.0f}", "—",
                                     _chip("derived", "info")])
                    else:
                        good = _close(mine, sheet)
                        rows.append([name, metric, f"{sheet:,.0f}", f"{mine:,.0f}",
                                     f"{mine - sheet:,.0f}",
                                     _chip("MATCH" if good else "MISMATCH",
                                           "ok" if good else "fail")])
        g = model.grand
        rows.append([f'<b>GRAND TOTAL</b>', "GMV", "", f"<b>{g.gmv_mtd:,.0f}</b>", "",
                     _chip(fmt.pct(g.growth), "ok" if (g.growth or 0) >= 0 else "fail")])
        rows.append(["", "Estimate", "", f"{g.estimate:,.0f}", "", ""])
        rows.append(["", "Gap", "", f"{g.gap:,.0f}", "", ""])
        out.append(_tbl(["Channel", "Metric", "Excel (sheet Total)", "Output (report)",
                         "Diff", "Verdict"], rows))

    # Sub-channel roll-up
    if model is not None and model.rollups:
        out.append(_h2("Sub-channel roll-up — parent vs the sum of its parts"))
        out.append(_note("When the sheet carries the parent column, the <b>entered</b> figure is "
                         "what we report and the parts only give the split; the difference is "
                         "shown as the gap. When the sheet has <b>only the parts</b>, the parent "
                         "is built by summing them (gap 0 by construction)."))
        rows = []
        for r in model.rollups:
            src = _chip("built from parts", "warn") if r.derived else _chip("entered in sheet", "ok")
            big = r.gap_pct is not None and abs(r.gap_pct) > config.ROLLUP_TOL
            rows.append([_esc(r.parent), src, f"{r.entered_gmv:,.0f}", f"{r.parts_gmv:,.0f}",
                         f"<b>{r.gap:,.0f}</b>",
                         _chip(f"{r.gap_pct*100:+.2f}%" if r.gap_pct is not None else "–",
                               "fail" if big else ("info" if r.gap else "ok"))])
            for n, gv, share in r.parts:
                rows.append([f'&nbsp;&nbsp;↳ {_esc(n)}', "", f"{gv:,.0f}", "",
                             _chip(f"{share*100:.1f}% of {r.parent}" if share is not None else "–",
                                   "info"), ""])
        out.append(_tbl([f"Parent / part", "Source", "Entered / parent", "Σ parts",
                         "Gap (entered − Σ)", "Gap %"], rows))
        for r in model.rollups:
            if r.lm_gap:
                out.append(_note(f"{_esc(r.parent)} LM gap: <b>{r.lm_gap:,.0f}</b> "
                                 f"(LM entered {r.entered_lm:,.0f} vs Σ parts {r.parts_lm:,.0f})."))

    # Cross-sheet
    out.append(_h2("Cross-sheet check"))
    if conflicts:
        out.append(_tbl(["Mismatch (channel left blank in the report — reconcile the sheets)"],
                        [[_esc(c)] for c in conflicts]))
    else:
        out.append(f'<div>{_chip("No mismatches", "ok")} '
                   f'<span style="font:400 12px Segoe UI,Arial,sans-serif;color:{_MUTED}">'
                   f'where both sheets carry a channel, the values agree.</span></div>')

    # Estimate freshness
    out.append(_h2("Estimate freshness — our run-rate vs Excel’s estimate cell"))
    kind = {"OK": "ok", "STALE": "warn", "MISMATCH": "warn", "NO_EXCEL": "info"}
    label = {"OK": "up to date", "STALE": "EXCEL STALE", "MISMATCH": "EXCEL STALE",
             "NO_EXCEL": "no Excel cell"}
    out.append(_tbl(["Channel", "Status", "Detail"],
                    [[_esc(f.channel), _chip(label.get(f.status, "?"), kind.get(f.status, "info")),
                      _esc(f.message)] for f in freshness]))
    stale = [f.channel for f in freshness if f.status in ("STALE", "MISMATCH")]
    if stale:
        fgw, bgw = _C["warn"]
        out.append(f'<div style="background:{bgw};color:{fgw};padding:10px 12px;border-radius:4px;'
                   f'font:600 12px Segoe UI,Arial,sans-serif">⚠ Excel day-count needs updating: '
                   f'{_esc(", ".join(stale))}</div>')

    # Advisories
    if notes:
        out.append(_h2("Sanity advisories — informational, never block the publish"))
        kmap = {"OK": "ok", "WARN": "warn", "INFO": "info"}
        out.append(_tbl(["Check", "Status", "Detail"],
                        [[_esc(a.name), _chip(a.level, kmap.get(a.level, "info")),
                          _esc(a.message)] for a in notes]))

    out.append(f'<div style="margin-top:18px;font:400 11px Segoe UI,Arial,sans-serif;'
               f'color:{_MUTED};border-top:1px solid {_LINE};padding-top:10px">'
               f'Report: <a href="{_REPORT_URL}" style="color:{_BRAND}">{_REPORT_URL}</a><br>'
               f'Source sheets are read-only — this run never writes to them. '
               f'The plain-text version of this validation is attached.</div>')
    out.append("</div></div>")
    return "".join(out)


if __name__ == "__main__":
    html = config.OUTPUT_DIR / "index.html"
    passed, results, freshness, notes = run(html if html.exists() else None)
    print_report(results)
    print_freshness(freshness)
    print_advisories(notes)
    raise SystemExit(0 if passed else 1)
