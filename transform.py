"""Turn parsed channels into the leadership report model.

All business math lives here and mirrors what the sheet already does:
  growth   = (GMV_mtd - LM_GMV_mtd) / LM_GMV_mtd     (like-for-like, same dates)
  estimate = GMV_mtd / days_with_data * days_in_month  (run-rate projection)
  gap      = estimate - GMV_mtd                        (still to come at this pace)

Like-for-like: LM is summed only over dates that actually have current GMV,
so Amazon's blank T-2 days never read as a drop and never inflate the base.
"""
from __future__ import annotations
import datetime as _dt
from dataclasses import dataclass, field
from typing import Optional

import config
from data_source import Channel, DailyRecord, load_channels


@dataclass
class DailyCell:
    date: _dt.date
    units: Optional[float]
    gmv: Optional[float]
    lm_gmv: Optional[float]
    growth: Optional[float]  # cell-wise growth vs LM
    ad_spend: Optional[float] = None
    ad_contri: Optional[float] = None  # ad_spend / gmv


@dataclass
class ChannelSummary:
    name: str
    section: str
    cadence: str                # "daily" | "monthly"
    units_mtd: float
    gmv_mtd: float
    lm_mtd: float
    growth: Optional[float]
    days_with_data: int
    last_date: Optional[_dt.date]
    estimate: float
    gap: float
    ad_mtd: float = 0.0
    ad_contri: Optional[float] = None  # ad_mtd / gmv_mtd
    has_ad: bool = False
    daily: list[DailyCell] = field(default_factory=list)


@dataclass
class Totals:
    units_mtd: float
    gmv_mtd: float
    lm_mtd: float
    growth: Optional[float]
    estimate: float
    gap: float
    ad_mtd: float = 0.0
    ad_contri: Optional[float] = None


@dataclass
class Section:
    name: str
    channels: list[ChannelSummary]
    totals: Totals
    dates: list[_dt.date]        # union of daily dates in this section
    last_date: Optional[_dt.date]


@dataclass
class Rollup:
    """Parent-vs-parts reconciliation for a channel that has sub-channels.

    Two cases, both handled:
      * the sheet carries the parent column   -> `derived` False: the ENTERED figure
        is the report's number, the parts only give the contribution split, and
        `gap` (entered - Σparts) is reported so the difference is visible.
      * the sheet has only the parts          -> `derived` True: the parent was
        summed from them, so the gap is 0 by construction.
    """
    parent: str
    derived: bool
    entered_gmv: float          # the parent figure the report uses
    parts_gmv: float            # Σ sub-channels
    gap: float                  # entered - parts (₹)
    gap_pct: Optional[float]    # gap / entered
    entered_lm: float
    parts_lm: float
    lm_gap: float
    parts: list = field(default_factory=list)   # [(name, gmv, share_of_parent)]


@dataclass
class ReportModel:
    month_label: str
    year: int
    month: int
    days_in_month: int
    as_of: Optional[_dt.date]
    sections: list[Section]
    grand: Totals
    best_channel: Optional[str]
    worst_channel: Optional[str]
    pending_note: Optional[str] = None
    merge_conflicts: list = field(default_factory=list)
    rollups: list = field(default_factory=list)     # Rollup per parent channel
    derived_parents: list = field(default_factory=list)  # notes from data_source
    # True when link_previous() filled this month's LM baseline from the actual
    # previous month (its own in-tab LM column was blank). The current month is
    # never linked, so QC's "LM = Σ daily LM cells" identity only applies there.
    lm_linked: bool = False


def _safe_growth(gmv: float, lm: float) -> Optional[float]:
    return (gmv - lm) / lm if lm else None


def _summarise_channel(ch: Channel, section: str, cadence: str,
                       days_in_month: int) -> ChannelSummary:
    days = [r for r in ch.records if r.gmv is not None]
    # Sum each metric over the FULL column (matches the sheet's Total row).
    # Some channels (e.g. Minutes) record LM GMV on days that have no current
    # GMV and vice-versa, so summing only over GMV-days undercounts LM.
    gmv_mtd = sum(r.gmv for r in ch.records if r.gmv is not None)
    lm_mtd = sum(r.lm_gmv for r in ch.records if r.lm_gmv is not None)
    units_mtd = sum(r.units for r in ch.records if r.units is not None)
    ad_mtd = sum(r.ad_spend for r in ch.records if r.ad_spend is not None)
    has_ad = any(r.ad_spend is not None for r in ch.records)
    n = len(days)
    last = max((r.date for r in days), default=None)
    # Run-rate uses DAYS ELAPSED (calendar day-of-month of the last data date),
    # not the count of non-empty rows. This matches the sheet's own formula
    # (Shopify /20, Amazon /19, Minutes /19) and avoids sparse-channel blow-ups.
    days_elapsed = last.day if last else 0
    estimate = gmv_mtd / days_elapsed * days_in_month if days_elapsed else 0.0

    cells = []
    for r in ch.records:
        ac = (r.ad_spend / r.gmv) if (r.ad_spend is not None and r.gmv) else None
        cells.append(DailyCell(
            date=r.date, units=r.units, gmv=r.gmv, lm_gmv=r.lm_gmv,
            growth=_safe_growth(r.gmv, r.lm_gmv) if (r.gmv is not None and r.lm_gmv) else None,
            ad_spend=r.ad_spend, ad_contri=ac,
        ))
    return ChannelSummary(
        name=ch.name, section=section, cadence=cadence,
        units_mtd=units_mtd, gmv_mtd=gmv_mtd, lm_mtd=lm_mtd,
        growth=_safe_growth(gmv_mtd, lm_mtd), days_with_data=n, last_date=last,
        estimate=estimate, gap=estimate - gmv_mtd, ad_mtd=ad_mtd,
        ad_contri=(ad_mtd / gmv_mtd if (has_ad and gmv_mtd) else None), has_ad=has_ad,
        daily=cells,
    )


def _totals(channels: list[ChannelSummary]) -> Totals:
    gmv = sum(c.gmv_mtd for c in channels)
    lm = sum(c.lm_mtd for c in channels)
    ad = sum(c.ad_mtd for c in channels)
    return Totals(
        units_mtd=sum(c.units_mtd for c in channels),
        gmv_mtd=gmv, lm_mtd=lm, growth=_safe_growth(gmv, lm),
        estimate=sum(c.estimate for c in channels),
        gap=sum(c.gap for c in channels),
        ad_mtd=ad, ad_contri=(ad / gmv if gmv else None),
    )


def build_report(period=None) -> ReportModel:
    chans, period = load_channels(period)
    channels = {c.name: c for c in chans}
    dim = period.days_in_month
    summaries: list[ChannelSummary] = []
    subs = {s for members in config.SUBCHANNELS.values() for s in members}

    def build_section(section_name):
        out = []
        for name in config.SECTIONS[section_name]["channels"]:
            if name in channels:
                cadence = "sub" if name in subs else "daily"
                s = _summarise_channel(channels[name], section_name, cadence, dim)
                out.append(s); summaries.append(s)
        return out

    qc = build_section("Quick Commerce")
    mkt = build_section("Marketplace & D2C")

    def section(name, chans):
        daily_dates = sorted({c.date for ch in chans if ch.cadence == "daily"
                               for c in ch.daily if c.gmv is not None})
        last = max((c.last_date for c in chans if c.last_date), default=None)
        return Section(name=name, channels=chans, totals=_totals(chans),
                       dates=daily_dates, last_date=last)

    # Drop a section with no channels (e.g. older months that predate the
    # Marketplace & D2C sheet show only Quick Commerce).
    sections = [s for s in (section("Quick Commerce", qc),
                            section("Marketplace & D2C", mkt)) if s.channels]
    grand = _totals(summaries)

    ranked = [c for c in summaries if c.growth is not None]
    best = max(ranked, key=lambda c: c.growth, default=None)
    worst = min(ranked, key=lambda c: c.growth, default=None)
    as_of = max((s.last_date for s in sections if s.last_date), default=None)

    return ReportModel(
        month_label=period.label, year=period.year, month=period.month,
        days_in_month=period.days_in_month, as_of=as_of, sections=sections,
        grand=grand, best_channel=best.name if best else None,
        worst_channel=worst.name if worst else None,
        pending_note=period.pending_note,
        merge_conflicts=list(period.conflicts),
        rollups=_rollups(summaries, period.derived),
        derived_parents=list(period.derived),
    )


def _rollups(summaries: list[ChannelSummary], derived_notes) -> list[Rollup]:
    """Parent vs Σ sub-channels for every parent in config.SUBCHANNELS, plus each
    part's contribution share of the parent (the denominator is always the parent
    figure the report uses — entered when the sheet has it, else the derived sum)."""
    by = {c.name: c for c in summaries}
    was_derived = {str(n).split("·")[-1].split("=")[0].strip() for n in derived_notes}
    out: list[Rollup] = []
    for parent, subs in config.SUBCHANNELS.items():
        p, parts = by.get(parent), [by[s] for s in subs if s in by]
        if not p or not parts:
            continue
        pg = sum(x.gmv_mtd for x in parts)
        pl = sum(x.lm_mtd for x in parts)
        out.append(Rollup(
            parent=parent, derived=parent in was_derived,
            entered_gmv=p.gmv_mtd, parts_gmv=pg,
            gap=p.gmv_mtd - pg,
            gap_pct=((p.gmv_mtd - pg) / p.gmv_mtd) if p.gmv_mtd else None,
            entered_lm=p.lm_mtd, parts_lm=pl, lm_gap=p.lm_mtd - pl,
            parts=[(x.name, x.gmv_mtd, (x.gmv_mtd / p.gmv_mtd) if p.gmv_mtd else None)
                   for x in parts],
        ))
    return out


def _prev_ym(year: int, month: int) -> tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def link_previous(models: list[ReportModel]) -> None:
    """Fill month-over-month comparison (LM) for months whose in-tab LM column is
    empty, using the ACTUAL previous month's data (we build every month, so the
    prior month's GMV is available). Mutates the models in place.

    Months that already carry an in-tab LM total (e.g. the current month, whose
    figures QC validates against the sheet) are left untouched — the guard below
    only links months whose LM is effectively zero.
    """
    by_ym = {(m.year, m.month): m for m in models}
    for m in models:
        if (m.grand.lm_mtd or 0) > 1:          # already has in-tab LM -> keep as-is
            continue
        prev = by_ym.get(_prev_ym(m.year, m.month))
        if not prev:                           # earliest month -> no baseline
            continue
        m.lm_linked = True
        prev_ch = {c.name: c for s in prev.sections for c in s.channels}
        prev_day = {(c.name, d.date.day): d.gmv
                    for s in prev.sections for c in s.channels
                    for d in c.daily if d.gmv is not None}
        for s in m.sections:
            for c in s.channels:
                pc = prev_ch.get(c.name)
                if pc is not None:
                    c.lm_mtd = pc.gmv_mtd
                    c.growth = _safe_growth(c.gmv_mtd, c.lm_mtd)
                for cell in c.daily:
                    pv = prev_day.get((c.name, cell.date.day))
                    if pv is not None:
                        cell.lm_gmv = pv
                        cell.growth = (_safe_growth(cell.gmv, pv)
                                       if (cell.gmv is not None and pv) else None)
            s.totals = _totals(s.channels)
        m.grand = _totals([c for s in m.sections for c in s.channels])
        ranked = [c for s in m.sections for c in s.channels if c.growth is not None]
        m.best_channel = max(ranked, key=lambda c: c.growth).name if ranked else None
        m.worst_channel = min(ranked, key=lambda c: c.growth).name if ranked else None


if __name__ == "__main__":
    m = build_report()
    print(f"{m.month_label} | as of {m.as_of} | {m.days_in_month} days")
    for s in m.sections:
        t = s.totals
        g = f"{t.growth*100:+.1f}%" if t.growth is not None else "n/a"
        print(f"\n== {s.name} == GMV {t.gmv_mtd:,.0f} | LM {t.lm_mtd:,.0f} | {g} "
              f"| est {t.estimate:,.0f} | gap {t.gap:,.0f}")
        for c in s.channels:
            cg = f"{c.growth*100:+.1f}%" if c.growth is not None else "n/a"
            print(f"   {c.name:18} {c.cadence:7} GMV {c.gmv_mtd:>12,.0f} "
                  f"LM {c.lm_mtd:>12,.0f} {cg:>8} est {c.estimate:>12,.0f} "
                  f"gap {c.gap:>12,.0f} (as of {c.last_date})")
    g = f"{m.grand.growth*100:+.1f}%"
    print(f"\n## GRAND GMV {m.grand.gmv_mtd:,.0f} | {g} | est {m.grand.estimate:,.0f} "
          f"| gap {m.grand.gap:,.0f} | best {m.best_channel} | worst {m.worst_channel}")
