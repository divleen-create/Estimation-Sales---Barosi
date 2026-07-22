"""Read-only data access + parsing.

Reads the local .xlsx snapshots today. A Google Sheets reader (read-only
scope) can be dropped in behind `load_channels()` without changing anything
downstream. This module NEVER writes to any source.
"""
from __future__ import annotations
import calendar
import datetime as _dt
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.utils import range_boundaries

import config

# Workbook load cache (keyed by path + mode + mtime) so repeated reads within a
# run don't re-parse the large workbooks. Read-only; never mutated.
_WB_CACHE: dict[tuple, tuple[float, object]] = {}


def _load_wb(path, data_only: bool = False, read_only: bool = False):
    key = (str(path), data_only, read_only)
    mt = os.path.getmtime(path)
    hit = _WB_CACHE.get(key)
    if hit and hit[0] == mt:
        return hit[1]
    wb = openpyxl.load_workbook(path, data_only=data_only, read_only=read_only)
    _WB_CACHE[key] = (mt, wb)
    return wb


@dataclass
class DailyRecord:
    date: _dt.date
    units: Optional[float] = None
    gmv: Optional[float] = None
    lm_gmv: Optional[float] = None
    ad_spend: Optional[float] = None


@dataclass
class Channel:
    name: str
    records: list[DailyRecord] = field(default_factory=list)


def _norm(s) -> str:
    """Lower-case, collapse whitespace, strip. Handles the stray leading
    spaces in headers like '                Shopify' and ' Minutes'."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _group_map_from_row1(ws) -> dict[int, str]:
    """Map each column index -> group (channel) name, expanding merged cells
    in row 1 so every column under a merged header inherits its name."""
    col_group: dict[int, str] = {}
    # Merged ranges give the true span of each group header.
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(rng))
        if min_row == 1:
            name = ws.cell(1, min_col).value
            for c in range(min_col, max_col + 1):
                col_group[c] = name
    # Any non-merged, non-empty row-1 cells (single-column groups).
    for c in range(1, ws.max_column + 1):
        if c not in col_group:
            v = ws.cell(1, c).value
            if v not in (None, ""):
                col_group[c] = v
    # Forward-fill: a column with no row-1 header belongs to the channel to its
    # left. Fixes blocks whose merge is narrower than the block, e.g. Amazon's
    # header merges only G:I but its 'AD sales' column J is un-merged.
    last = None
    for c in range(1, ws.max_column + 1):
        val = col_group.get(c)
        if val not in (None, ""):
            last = val
        elif last is not None:
            col_group[c] = last
    return col_group


def _column_map(ws) -> dict[str, dict[str, int]]:
    """channel display name -> {field_name: column_index}, from rows 1-2.
    Skips Total/xx/excluded channels and keeps only units/gmv/lm_gmv/ad_spend."""
    col_group = _group_map_from_row1(ws)
    channels: dict[str, dict[str, int]] = {}
    for c in range(2, ws.max_column + 1):  # col 1 is DATE
        grp = _norm(col_group.get(c))
        if grp in config.CHANNEL_SKIP:
            continue
        display = config.CHANNEL_ALIASES.get(grp)
        if display is None or display in config.EXCLUDED_CHANNELS:
            continue
        fld = config.FIELD_ALIASES.get(_norm(ws.cell(2, c).value))
        if fld in (None, "dolchi", "total_sale"):
            continue
        channels.setdefault(display, {})[fld] = c
    return channels


def _date_rows(ws) -> list[tuple[int, _dt.date]]:
    """Row index + date for each daily row (row 3 down while col A is a date)."""
    rows: list[tuple[int, _dt.date]] = []
    for r in range(3, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, _dt.datetime):
            rows.append((r, v.date()))
        elif isinstance(v, _dt.date):
            rows.append((r, v))
        elif rows:
            break  # first non-date after data begins => end of daily block
    return rows


def parse_tab(path, tab: str) -> list[Channel]:
    """Parse one workbook tab into a list of real Channels (skips Total/xx
    and excluded channels). Column mapping is derived from rows 1-2 so it
    survives layout differences between the two workbooks."""
    ws = _load_wb(path, data_only=True)[tab]
    channels = _column_map(ws)
    date_rows = _date_rows(ws)

    result: list[Channel] = []
    for display, fld_cols in channels.items():
        ch = Channel(name=display)
        for r, d in date_rows:
            rec = DailyRecord(date=d)
            for fld, c in fld_cols.items():
                val = ws.cell(r, c).value
                if isinstance(val, (int, float)):
                    setattr(rec, fld, float(val))
            ch.records.append(rec)
        result.append(ch)
    return result


@dataclass
class SheetChannelDiag:
    total_gmv: Optional[float] = None
    total_lm: Optional[float] = None
    total_units: Optional[float] = None
    total_ad: Optional[float] = None
    est_value: Optional[float] = None      # computed value in the estimate cell
    est_divisor: Optional[int] = None      # days the sheet divided by
    est_mult: Optional[int] = None         # days it projected to
    est_cell: Optional[str] = None
    ad_contri_excel: Optional[float] = None  # sheet's own Ad/Value ratio, if present


_EST_RE = re.compile(r"^=\s*[A-Za-z]+\d+\s*/\s*(\d+)\s*\*\s*(\d+)\s*$")
# Ad-contribution ratio the sheet sometimes stores in the ad column: '=M34/K34'.
_RATIO_RE = re.compile(r"^=[A-Za-z]+\d+/[A-Za-z]+\d+$")


def sheet_diagnostics(path, tab: str) -> dict[str, SheetChannelDiag]:
    """Read the sheet's OWN Total-row values and estimate-row formulas per
    channel — ground truth that QC compares against (independent of our math).
    The estimate divisor is parsed from the formula (e.g. '=H35/19*31' -> 19)
    so we can tell whether the sheet's day-count is stale."""
    wsv = _load_wb(path, data_only=True)[tab]
    wsf = _load_wb(path, data_only=False)[tab]
    channels = _column_map(wsv)
    rows = _date_rows(wsv)
    if not rows:
        return {}
    last = rows[-1][0]
    ym = parse_period_from_name(tab)
    dim = calendar.monthrange(*ym)[1] if ym else 31  # days in month (estimate mult)

    # Totals row = the row just below the daily block with the most numeric
    # GMV cells across channels (Total/SUM row).
    best_row, best_hits = None, 0
    for r in range(last + 1, min(last + 8, wsv.max_row) + 1):
        hits = sum(1 for cols in channels.values()
                   if "gmv" in cols and isinstance(wsv.cell(r, cols["gmv"]).value, (int, float)))
        if hits > best_hits:
            best_hits, best_row = hits, r

    out: dict[str, SheetChannelDiag] = {}
    for name, cols in channels.items():
        d = SheetChannelDiag()
        if best_row and "gmv" in cols:
            d.total_gmv = wsv.cell(best_row, cols["gmv"]).value
            if "lm_gmv" in cols:
                d.total_lm = wsv.cell(best_row, cols["lm_gmv"]).value
            if "units" in cols:
                d.total_units = wsv.cell(best_row, cols["units"]).value
            if "ad_spend" in cols:
                d.total_ad = wsv.cell(best_row, cols["ad_spend"]).value
        # estimate: prefer the formula divisor (e.g. '=H35/19*31' -> 19); if the
        # snapshot has no formulas (e.g. a live Sheets-API pull), derive the
        # divisor from values instead: divisor = round(total_gmv * dim / est).
        if "gmv" in cols:
            gcol = cols["gmv"]
            for r in range(last + 1, min(last + 12, wsf.max_row) + 1):
                f = wsf.cell(r, gcol).value
                if isinstance(f, str):
                    m = _EST_RE.match(f.replace(" ", ""))
                    if m:
                        d.est_divisor, d.est_mult = int(m.group(1)), int(m.group(2))
                        d.est_value = wsv.cell(r, gcol).value
                        d.est_cell = wsv.cell(r, gcol).coordinate
                        break
            if d.est_divisor is None and d.total_gmv:
                for r in range(last + 1, min(last + 12, wsv.max_row) + 1):
                    v = wsv.cell(r, gcol).value
                    if isinstance(v, (int, float)) and v > d.total_gmv * 1.02:
                        d.est_value, d.est_mult = v, dim
                        d.est_divisor = round(d.total_gmv * dim / v)
                        d.est_cell = wsv.cell(r, gcol).coordinate
                        break
        # ad-contribution: the sheet's Ad/Value ratio, from formula ('=M34/K34')
        # or, in a values-only snapshot, a small ratio value in the ad column.
        if "ad_spend" in cols:
            acol = cols["ad_spend"]
            for r in range(last + 1, min(last + 6, wsf.max_row) + 1):
                f = wsf.cell(r, acol).value
                if isinstance(f, str) and _RATIO_RE.match(f.replace(" ", "")):
                    d.ad_contri_excel = wsv.cell(r, acol).value
                    break
            if d.ad_contri_excel is None:
                for r in range(last + 1, min(last + 6, wsv.max_row) + 1):
                    v = wsv.cell(r, acol).value
                    if isinstance(v, (int, float)) and 0 < v < 3:
                        d.ad_contri_excel = v
                        break
        out[name] = d
    return out


# --- Reporting-period resolution (auto-rolling month) -----------------------
_MONTHS = {m[:3].lower(): i for i, m in enumerate(calendar.month_name) if m}
_MONTHS["sept"] = 9  # sheets use both 'Sep' and 'Sept'


@dataclass
class Period:
    year: int
    month: int
    days_in_month: int
    label: str          # e.g. "July 2026"
    tab_qc: str
    tab_mkt: str
    pending_note: Optional[str] = None  # e.g. newer month exists but is empty


def parse_period_from_name(name: str) -> Optional[tuple[int, int]]:
    """'July26'->(2026,7), 'FEB 26'->(2026,2), 'Sept25'->(2025,9),
    'July-2023'->(2023,7). Non-month tabs ('working','Trends') -> None."""
    n = re.sub(r"[\s\-']", "", str(name)).lower()
    m = re.match(r"^([a-z]+)(\d{2,4})$", n)
    if not m:
        return None
    mon_txt, yr_txt = m.group(1), m.group(2)
    mon = _MONTHS.get("sept" if mon_txt.startswith("sept") else mon_txt[:3])
    if not mon:
        return None
    yr = int(yr_txt)
    if yr < 100:
        yr += 2000
    return (yr, mon)


def _month_map(path) -> dict[tuple[int, int], str]:
    """Map (year, month) -> sheet/tab name for a workbook."""
    wb = _load_wb(path, data_only=True, read_only=True)
    out: dict[tuple[int, int], str] = {}
    for name in wb.sheetnames:
        ym = parse_period_from_name(name)
        if ym and ym not in out:
            out[ym] = name
    return out


def _tab_has_data(path, tab: str) -> bool:
    try:
        chans = parse_tab(path, tab)
    except Exception:
        return False
    return any(r.gmv is not None for ch in chans for r in ch.records)


def resolve_period(qc_path, mkt_path) -> Period:
    """Pick the latest month present AND populated in both workbooks.

    Rolls forward automatically (July -> August once August has data) and
    falls back to the last populated month if the newer tab is still empty.
    Set config.FORCE_PERIOD to pin a specific (year, month).
    """
    qc_map, mkt_map = _month_map(qc_path), _month_map(mkt_path)
    common = sorted(set(qc_map) & set(mkt_map), reverse=True)  # newest first
    if not common:
        raise RuntimeError("No month tab is present in both workbooks.")

    def make(ym: tuple[int, int], note: Optional[str] = None) -> Period:
        y, mth = ym
        return Period(y, mth, calendar.monthrange(y, mth)[1],
                      _dt.date(y, mth, 1).strftime("%B %Y"),
                      qc_map[ym], mkt_map[ym], note)

    if config.FORCE_PERIOD and tuple(config.FORCE_PERIOD) in set(common):
        return make(tuple(config.FORCE_PERIOD))

    populated = [ym for ym in common
                 if _tab_has_data(qc_path, qc_map[ym]) and _tab_has_data(mkt_path, mkt_map[ym])]
    if not populated:
        # Nothing populated yet -> show the newest common month as-is (fallback).
        return make(common[0], "No data populated yet — showing the latest tab as-is.")

    chosen = populated[0]
    note = None
    if common[0] != chosen:  # a newer month tab exists but isn't populated yet
        ny, nm = common[0]
        note = (f"{_dt.date(ny, nm, 1).strftime('%B %Y')} is not populated yet — "
                f"showing {make(chosen).label} (last data available).")
    return make(chosen, note)


# --- Loading ---------------------------------------------------------------
def _sources() -> tuple[Path, Path]:
    """Prefer read-only Google Sheets snapshots when present, else local .xlsx."""
    snap = Path(__file__).resolve().parent / "snapshots"
    qc_src, mkt_src = snap / "qc_snapshot.xlsx", snap / "mkt_snapshot.xlsx"
    if qc_src.exists() and mkt_src.exists():
        return qc_src, mkt_src
    return config.WORKBOOK_QC, config.WORKBOOK_MKT


def load_channels() -> tuple[list[Channel], Period]:
    """Return (channels from both workbooks, resolved reporting Period)."""
    qc_src, mkt_src = _sources()
    period = resolve_period(qc_src, mkt_src)
    qc = parse_tab(qc_src, period.tab_qc)
    mkt = parse_tab(mkt_src, period.tab_mkt)
    return qc + mkt, period


def load_sheet_diagnostics() -> dict[str, SheetChannelDiag]:
    """Sheet Total-row values + estimate formulas for every channel, for the
    resolved period. Used by QC as ground truth."""
    qc_src, mkt_src = _sources()
    period = resolve_period(qc_src, mkt_src)
    out: dict[str, SheetChannelDiag] = {}
    out.update(sheet_diagnostics(qc_src, period.tab_qc))
    out.update(sheet_diagnostics(mkt_src, period.tab_mkt))
    return out


if __name__ == "__main__":
    chans, period = load_channels()
    print(f"Period: {period.label} | tabs: {period.tab_qc} / {period.tab_mkt}"
          f" | {period.days_in_month} days")
    if period.pending_note:
        print("Note:", period.pending_note)
    for ch in chans:
        days = [r for r in ch.records if r.gmv is not None]
        last = max((r.date for r in days), default=None)
        total = sum(r.gmv or 0 for r in ch.records)
        print(f"  {ch.name:18} days_with_gmv={len(days):2}  last={last}  MTD_gmv={total:,.0f}")
