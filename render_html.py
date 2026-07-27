"""Render the ReportModel into a self-contained HTML one-pager.

Pure-Python string building (no external template engine). Inline CSS only,
no network calls, so the file opens offline and screenshots cleanly.
"""
from __future__ import annotations
import datetime as _dt
import html
from pathlib import Path

import config
import fmt
from transform import ReportModel, Section, ChannelSummary, build_report

# sub-channel -> parent (e.g. "Amazon Core" -> "Amazon")
_PARENT_OF = {sub: parent for parent, subs in config.SUBCHANNELS.items() for sub in subs}

CSS = """
:root{
  --bg:#f4f6f8; --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e5e9f0;
  --brand:#0b3d5c; --brand2:#12608a;
  --pos2:#0f9d58; --pos1:#d6f0e0; --pos1t:#0b7a43;
  --neg2:#d93025; --neg1:#fbe0dd; --neg1t:#b1271b; --neu:#f1f4f8;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  -webkit-font-smoothing:antialiased;font-size:13px;line-height:1.35}
.wrap{max-width:1180px;margin:0 auto;padding:22px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;
  padding:18px 20px;margin-bottom:16px;box-shadow:0 1px 2px rgba(15,23,42,.04)}
.head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;
  padding-bottom:14px;border-bottom:2px solid var(--brand);margin-bottom:2px}
.title{font-size:22px;font-weight:800;color:var(--brand);letter-spacing:-.3px}
.sub{color:var(--muted);font-size:12px;margin-top:3px}
.asof{text-align:right;color:var(--muted);font-size:12px}
.asof b{color:var(--ink);font-size:13px}

.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:14px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.kpi .lab{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.5px}
.kpi .val{font-size:22px;font-weight:800;margin-top:4px;letter-spacing:-.4px}
.kpi .note{font-size:11px;color:var(--muted);margin-top:3px}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-weight:700;font-size:12px}
.badge.up{background:var(--pos1);color:var(--pos1t)}
.badge.down{background:var(--neg1);color:var(--neg1t)}

.sec-h{display:flex;justify-content:space-between;align-items:baseline;
  margin:2px 0 10px;gap:10px;flex-wrap:wrap}
.sec-h .name{font-size:16px;font-weight:800;color:var(--brand2)}
.sec-h .stat{color:var(--muted);font-size:12px}
.sec-h .stat b{color:var(--ink)}

table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}
th,td{padding:6px 8px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
th:first-child,td:first-child{text-align:left}
thead th{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.4px;font-weight:700}
tbody tr:last-child td{border-bottom:none}
.chan{font-weight:700}
.tag{font-size:9px;color:var(--muted);border:1px solid var(--line);border-radius:6px;
  padding:1px 5px;margin-left:6px;text-transform:uppercase;letter-spacing:.4px}
.tot td{font-weight:800;border-top:2px solid var(--line);background:#fbfcfe}
.pos{color:var(--pos1t);font-weight:700}
.neg{color:var(--neg1t);font-weight:700}
.gap0{color:var(--ink);font-weight:700}
.na{color:var(--muted)}

.grid-note{color:var(--muted);font-size:11px;margin:14px 0 6px}
.grid td,.grid th{padding:4px 6px;font-size:11.5px}
/* per-platform day-wise tables */
.ptbl{margin:12px 0 4px;break-inside:avoid}
.ptitle{font-size:13px;font-weight:800;color:var(--brand2);margin:10px 0 4px}
.ptitle .pmeta{font-weight:500;color:var(--muted);font-size:11px;margin-left:6px}
table.daily{width:100%;table-layout:fixed}
.daily th:first-child,.daily td:first-child{width:12%}
.daily td,.daily th{padding:5px 10px;font-size:12px}
.daily tbody tr:hover td{background:#f8fafc}
.daily .tot td{background:#eef2f7}
.daily .tot:hover td{background:#eef2f7}
.daily .est td{font-weight:700;color:var(--brand);border-top:1px solid var(--line);background:#fff}
.daily .est td:first-child{color:var(--muted);font-weight:700}
.contrib{font-size:11.5px;color:var(--ink);background:#eef4fb;border:1px solid #dbe7f4;
  border-radius:8px;padding:7px 10px;margin:10px 0 2px}
.contrib .cp{display:inline-block;margin-right:14px;white-space:nowrap}
.contrib .cp b{color:var(--brand2)}
.g-pos2{background:#bfe8cf}.g-pos1{background:var(--pos1)}
.g-neu{background:var(--neu)}.g-neg1{background:var(--neg1)}.g-neg2{background:#f6c9c4}
.g-na{background:#fff;color:var(--muted)}
.grow{font-size:9px;display:block;color:#334155}
.latest{outline:2px solid var(--brand2);outline-offset:-2px}
.foot{color:var(--muted);font-size:11px;line-height:1.6}
.foot b{color:var(--ink)}
.tw{overflow-x:auto;-webkit-overflow-scrolling:touch}
.section-label{font-size:12px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;
  color:var(--muted);margin:2px 0 12px}
/* responsive / mobile */
@media (max-width:720px){
  .wrap{padding:12px}
  .card{padding:14px 12px;border-radius:12px}
  .title{font-size:19px}
  .head{flex-direction:column;align-items:flex-start;gap:6px}
  .asof{text-align:left}
  .kpis{grid-template-columns:1fr 1fr;gap:8px}
  .kpi .val{font-size:19px}
  .sec-h{flex-direction:column;align-items:flex-start;gap:2px}
  .sec-h .stat{font-size:11px}
  table.daily{min-width:520px}      /* keep columns legible; scroll inside .tw */
  table{min-width:560px}
  .ptitle{font-size:12px}
}
@media (max-width:420px){ .kpis{grid-template-columns:1fr} }
/* month slicer */
.picker{display:flex;flex-direction:column;align-items:flex-end;gap:4px}
.picker label{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.5px}
.picker select{font:inherit;font-size:14px;font-weight:700;color:var(--brand);background:var(--card);
  border:1px solid var(--line);border-radius:8px;padding:6px 12px;cursor:pointer}
.cmp{color:var(--muted);font-size:11px}
.cmp b{color:var(--brand2)}
.mtitle{font-size:16px;font-weight:800;color:var(--brand2);letter-spacing:-.2px}
.month-pane[hidden]{display:none}
/* data-freshness notes */
.notes .nrow{display:flex;flex-wrap:wrap;gap:6px}
.nchip{font-size:11px;padding:3px 9px;border-radius:999px;border:1px solid var(--line);white-space:nowrap}
.nchip.ok{background:var(--pos1);color:var(--pos1t);border-color:#bfe3cf}
.nchip.warn{background:var(--neg1);color:var(--neg1t);border-color:#f3cfca}
.nchip.mo{background:var(--neu);color:var(--muted)}
/* platform filter */
.platfilter{flex-direction:row;align-items:center;gap:8px}
.ptbl-pane[hidden]{display:none}
@media (max-width:720px){ .picker{align-items:flex-start;width:100%} .platfilter{flex-direction:row;width:auto} }
"""


def _e(s) -> str:
    return html.escape(str(s))


def _growth_span(g):
    p = (g or 0) * 100
    cls = "na" if (g is None or round(p, 1) == 0) else ("pos" if p > 0 else "neg")
    return f'<span class="{cls}">{fmt.pct(g)}</span>'


def _gap_span(v):
    """Gap to estimate = estimate − MTD. A positive gap means we're still SHORT
    of the run-rate estimate → red; MTD at/above estimate → green; ~zero → black.
    Rounds to display resolution so a tiny residual reads '₹0.0 L', not '₹-0.0 L'."""
    if v is None:
        return '<span class="na">–</span>'
    if round(v / 1e5, 1) == 0:            # rounds to zero at 0.1 L display resolution
        return '<span class="gap0">0</span>'
    cls = "neg" if v > 0 else "pos"       # short of estimate = red; at/over = green
    return f'<span class="{cls}">{fmt.gmv_auto(v)}</span>'


def _kpi_block(m: ReportModel) -> str:
    gr = m.grand
    up = (gr.growth or 0) >= 0
    badge = f'<span class="badge {"up" if up else "down"}">{fmt.pct(gr.growth)} vs LM</span>'
    return f"""
    <div class="kpis">
      <div class="kpi"><div class="lab">Run-rate estimate (month)</div>
        <div class="val">{fmt.gmv_auto(gr.estimate)}</div>
        <div class="note">MTD ÷ days elapsed × {m.days_in_month}</div></div>
      <div class="kpi"><div class="lab">MTD GMV</div>
        <div class="val">{fmt.gmv_auto(gr.gmv_mtd)}</div>
        <div class="note">{badge}</div></div>
      <div class="kpi"><div class="lab">Gap to estimate</div>
        <div class="val">{_gap_span(gr.gap)}</div>
        <div class="note">still to come at this pace</div></div>
      <div class="kpi"><div class="lab">Momentum</div>
        <div class="val" style="font-size:15px">▲ {_e(m.best_channel)}</div>
        <div class="note">▼ watch: {_e(m.worst_channel)}</div></div>
    </div>"""


def _summary_table(sec: Section) -> str:
    rows = []
    for c in sec.channels:
        tag = (f'<span class="tag">of {_e(_PARENT_OF[c.name])}</span>'
               if c.name in _PARENT_OF else '')
        ad = fmt.gmv_auto(c.ad_mtd) if c.has_ad else '<span class="na">–</span>'
        adc = fmt.pct_plain(c.ad_contri) if c.has_ad else '<span class="na">–</span>'
        rows.append(f"""
      <tr>
        <td class="chan">{_e(c.name)}{tag}</td>
        <td>{fmt.indian_group(c.units_mtd)}</td>
        <td>{fmt.gmv_auto(c.gmv_mtd)}</td>
        <td>{fmt.gmv_auto(c.lm_mtd)}</td>
        <td>{_growth_span(c.growth)}</td>
        <td>{ad}</td>
        <td>{adc}</td>
        <td>{fmt.gmv_auto(c.estimate)}</td>
        <td>{_gap_span(c.gap)}</td>
        <td class="na">{c.last_date.strftime('%d %b') if c.last_date else '–'}</td>
      </tr>""")
    t = sec.totals
    rows.append(f"""
      <tr class="tot">
        <td>Total</td>
        <td>{fmt.indian_group(t.units_mtd)}</td>
        <td>{fmt.gmv_auto(t.gmv_mtd)}</td>
        <td>{fmt.gmv_auto(t.lm_mtd)}</td>
        <td>{_growth_span(t.growth)}</td>
        <td>{fmt.gmv_auto(t.ad_mtd)}</td>
        <td>{fmt.pct_plain(t.ad_contri)}</td>
        <td>{fmt.gmv_auto(t.estimate)}</td>
        <td>{_gap_span(t.gap)}</td>
        <td></td>
      </tr>""")
    return f"""
    <div class="tw"><table>
      <thead><tr>
        <th>Channel</th><th>Units</th><th>GMV · MTD</th><th>GMV · LM</th><th>Growth</th>
        <th>Ad Sales</th><th>Ad Contri</th><th>Est · month</th><th>Gap</th><th>As of</th>
      </tr></thead>
      <tbody>{''.join(rows)}</tbody>
    </table></div>"""


def _platform_daily_table(c: ChannelSummary, dim: int) -> str:
    """Full day-wise KPI table for one platform. Uniform 7-column schema for
    EVERY platform — Date | Units | GMV | GMV·LM | Growth | Ad Sales | Ad Contri
    — so the platform filter swaps tables seamlessly; platforms with no ad data
    show '–' in the ad columns. The GMV cell is shaded by that day's growth vs LM."""
    cells = [x for x in c.daily
             if any(v is not None for v in (x.gmv, x.lm_gmv, x.ad_spend, x.units))]
    if not cells:
        return ""
    latest = c.last_date
    na = '<span class="na">–</span>'

    body = []
    for x in cells:
        gcls = fmt.growth_class(x.growth)
        ad_td = (f"<td>{fmt.money_l(x.ad_spend)}</td><td>{fmt.pct_plain(x.ad_contri)}</td>"
                 if c.has_ad else f"<td>{na}</td><td>{na}</td>")
        datecls = ' class="latest"' if x.date == latest else ''
        body.append(
            f'<tr><td{datecls}>{x.date.strftime("%d %b")}</td>'
            f'<td>{fmt.indian_group(x.units) if x.units is not None else "–"}</td>'
            f'<td class="{gcls}">{fmt.money_l(x.gmv)}</td>'
            f'<td>{fmt.money_l(x.lm_gmv)}</td>'
            f'<td>{_growth_span(x.growth)}</td>'
            f'{ad_td}</tr>')

    ad_tot = (f"<td>{fmt.gmv_auto(c.ad_mtd)}</td><td>{fmt.pct_plain(c.ad_contri)}</td>"
              if c.has_ad else f"<td>{na}</td><td>{na}</td>")
    total_row = (f'<tr class="tot"><td>Total</td>'
                 f'<td>{fmt.indian_group(c.units_mtd)}</td>'
                 f'<td>{fmt.gmv_auto(c.gmv_mtd)}</td>'
                 f'<td>{fmt.gmv_auto(c.lm_mtd)}</td>'
                 f'<td>{_growth_span(c.growth)}</td>{ad_tot}</tr>')
    div = latest.day if latest else 0
    est_row = (f'<tr class="est"><td>Est · month</td>'
               f'<td>{fmt.gmv_auto(c.estimate)}</td><td colspan="5"></td></tr>')

    asof = latest.strftime("%d %b") if latest else "–"
    return f"""
    <div class="ptbl">
      <div class="ptitle">{_e(c.name)} <span class="pmeta">daily · data to {asof}
        · estimate ÷{div}×{dim}</span></div>
      <div class="tw"><table class="daily">
        <thead><tr><th>Date</th><th>Units</th><th>GMV</th><th>GMV · LM</th><th>Growth</th><th>Ad Sales</th><th>Ad Contri</th></tr></thead>
        <tbody>{''.join(body)}{total_row}{est_row}</tbody>
      </table></div>
    </div>"""


def _contrib_strip(sec: Section, parent: str) -> str:
    """A one-line contribution glance for a parent's sub-channels (share of the
    parent's GMV), shown at the top of the parent's day-wise table."""
    subs = [c for c in sec.channels
            if _PARENT_OF.get(c.name) == parent and c.name in config.SUBCHANNELS.get(parent, [])]
    p = next((c for c in sec.channels if c.name == parent), None)
    if not subs or not p or not p.gmv_mtd:
        return ""
    parts = "".join(
        f'<span class="cp"><b>{_e(s.name)}</b> {fmt.gmv_auto(s.gmv_mtd)} '
        f'· <span class="pos">{fmt.pct_plain(s.gmv_mtd / p.gmv_mtd)}</span></span>'
        for s in subs)
    return (f'<div class="contrib">Contribution to {_e(parent)} GMV '
            f'({fmt.gmv_auto(p.gmv_mtd)}): {parts}</div>')


def _daily_tables(sec: Section, dim: int) -> str:
    daily = [c for c in sec.channels if c.cadence == "daily"]
    if not daily:
        return ""
    note = ('<div class="grid-note">Day-wise KPIs per platform — money in ₹ (Lakh); '
            'the GMV cell is shaded by that day’s growth vs the same day last month. '
            'Latest day outlined. Amazon Core &amp; Amazon NOW+Fresh are shown as a '
            'contribution split on the Amazon table (Amazon ≈ Core + NOW+Fresh).</div>')
    out = [note]
    for c in daily:
        strip = _contrib_strip(sec, c.name) if c.name in config.SUBCHANNELS else ""
        out.append(strip + _platform_daily_table(c, dim))
    return "".join(out)


def _section_header(sec: Section) -> str:
    t = sec.totals
    up = (t.growth or 0) >= 0
    stat = (f'GMV <b>{fmt.gmv_auto(t.gmv_mtd)}</b> · '
            f'<span class="{"pos" if up else "neg"}">{fmt.pct(t.growth)}</span> vs LM · '
            f'Ad <b>{fmt.gmv_auto(t.ad_mtd)}</b> · '
            f'Est <b>{fmt.gmv_auto(t.estimate)}</b> · Gap {_gap_span(t.gap)}')
    return (f'<div class="sec-h"><div class="name">{_e(sec.name)}</div>'
            f'<div class="stat">{stat}</div></div>')


def _summary_card(sec: Section) -> str:
    """Platform-level view for one section (summary table only)."""
    return f'<div class="card">{_section_header(sec)}{_summary_table(sec)}</div>'


def _daily_card(sec: Section, dim: int) -> str:
    """Date-wise view for one section (per-platform daily tables)."""
    tables = _daily_tables(sec, dim)
    if not tables.strip():
        return ""
    return (f'<div class="card"><div class="sec-h"><div class="name">{_e(sec.name)}</div></div>'
            f'{tables}</div>')


def _freshness_note(freshness) -> str:
    """Small footer line flagging Excel estimate day-counts that need refreshing."""
    if not freshness:
        return ""
    stale = [f for f in freshness if f.status in ("STALE", "MISMATCH")]
    if not stale:
        return ('<div class="foot" style="margin-top:8px;color:var(--pos1t)">'
                '✔ Excel estimate day-counts are all current.</div>')
    items = ", ".join(f"{_e(f.channel)} (÷{f.excel_divisor}→{f.data_days})" for f in stale)
    return ('<div class="foot" style="margin-top:8px;border-left:4px solid #b8860b;'
            'padding-left:10px;color:#7a5c00">'
            f'⚠ <b>Excel estimate day-counts to refresh:</b> {items}. '
            'The report already uses the correct data-driven day-counts; this flags the '
            'source sheet.</div>')


def _notes_card(m: ReportModel, generated: _dt.datetime) -> str:
    """Data-freshness note: each platform's latest data date, flagging any that
    are behind their expected lag (T-1 default, Amazon T-2, Flipkart monthly).
    Reference is 'today' (generated) for the current month, else the month end."""
    chans = [c for s in m.sections for c in s.channels if c.cadence == "daily"]
    if not chans:
        return ""
    is_current = (m.year == generated.year and m.month == generated.month)
    ref = generated.date() if is_current else _dt.date(m.year, m.month, m.days_in_month)
    chips, behind = [], []
    for c in chans:
        if c.name in config.MONTHLY_PLATFORMS:
            chips.append(f'<span class="nchip mo"><b>{_e(c.name)}</b> · monthly</span>')
            continue
        lag = config.FRESHNESS_LAG_DAYS.get(c.name, config.FRESHNESS_LAG_DAYS["_default"])
        expected = ref - _dt.timedelta(days=lag)
        last = c.last_date
        if last is None:
            behind.append(c.name)
            chips.append(f'<span class="nchip warn"><b>{_e(c.name)}</b> · no data</span>')
        elif last >= expected:
            chips.append(f'<span class="nchip ok"><b>{_e(c.name)}</b> · '
                         f'{last.strftime("%d %b")} ✓</span>')
        else:
            behind.append(c.name)
            chips.append(f'<span class="nchip warn"><b>{_e(c.name)}</b> · '
                         f'{last.strftime("%d %b")} ⚠ exp {expected.strftime("%d %b")}</span>')
    summary = (f'<span class="neg">⚠ Behind: {_e(", ".join(behind))}</span> — latest data older '
               f'than expected.' if behind
               else '<span class="pos">✓ All platforms current</span> — data up to the expected date.')
    return (f'<div class="card notes">'
            f'<div class="sec-h"><div class="name" style="font-size:13px">Data notes · freshness</div></div>'
            f'<div class="grid-note" style="margin:2px 0 8px">{summary} '
            f'Expected: T-1 for all · Amazon T-2 · Flipkart monthly.</div>'
            f'<div class="nrow">{"".join(chips)}</div></div>')


def _daily_filter_card(m: ReportModel, dim: int) -> str:
    """One Date-wise table with a platform dropdown spanning BOTH sections. Each
    platform's table (+ contribution strip for parents like Amazon) is pre-rendered
    as a hidden pane; the dropdown toggles which shows (first by default)."""
    sec_of = {c.name: s for s in m.sections for c in s.channels}
    chans = [c for s in m.sections for c in s.channels if c.cadence == "daily"]
    if not chans:
        return ""
    note = ('<div class="grid-note">Day-wise KPIs for the selected platform — money in ₹ (Lakh); '
            'the GMV cell is shaded by that day’s growth vs the same day last month. Latest day '
            'outlined. Amazon shows its Core / NOW+Fresh contribution split.</div>')
    opts, panes = [], []
    for i, c in enumerate(chans):
        opts.append(f'<option value="{i}"{" selected" if i == 0 else ""}>{_e(c.name)}</option>')
        strip = (_contrib_strip(sec_of[c.name], c.name) if c.name in config.SUBCHANNELS else "")
        hidden = "" if i == 0 else " hidden"
        panes.append(f'<div class="ptbl-pane" data-plat="{i}"{hidden}>{strip}'
                     f'{_platform_daily_table(c, dim)}</div>')
    picker = (f'<div class="picker platfilter"><label for="platPick">Platform</label>'
              f'<select id="platPick" class="platpick" onchange="__showPlat(this)">'
              f'{"".join(opts)}</select></div>')
    return (f'<div class="card">'
            f'<div class="sec-h"><div class="name">Daily KPIs by platform</div>{picker}</div>'
            f'{note}{"".join(panes)}</div>')


def _prev_month_label(year: int, month: int) -> str:
    """Calendar month before (year, month), e.g. (2025, 7) -> 'June 2025'."""
    y, mth = (year - 1, 12) if month == 1 else (year, month - 1)
    return _dt.date(y, mth, 1).strftime("%B %Y")


def _month_pane(m: ReportModel, idx: int, generated: _dt.datetime) -> str:
    """One month's full content (header + KPIs + notes + platform + daily) as a
    toggleable pane. idx 0 is visible; the rest start hidden."""
    dim = m.days_in_month
    asof = m.as_of.strftime("%d %b %Y") if m.as_of else "–"
    prev = _prev_month_label(m.year, m.month)
    notes = _notes_card(m, generated)
    platform_block = ('<div class="section-label">Platform view · month-to-date</div>'
                      + "".join(_summary_card(s) for s in m.sections))
    daily_block = ('<div class="section-label">Date-wise detail</div>'
                   + _daily_filter_card(m, dim))
    pending = (f'<div class="card" style="border-left:4px solid #b8860b;color:#7a5c00">'
               f'⚠ {_e(m.pending_note)}</div>') if m.pending_note else ""
    present = {s.name for s in m.sections}
    missing = [n for n in config.SECTIONS if n not in present]
    miss_note = (f'<div class="card" style="border-left:4px solid var(--brand2)">'
                 f'<div class="grid-note" style="margin:0">This month tracks '
                 f'<b>{_e(" / ".join(present))}</b> only — {_e(" & ".join(missing))} '
                 f'data begins in a later month.</div></div>') if missing else ""
    hidden = "" if idx == 0 else " hidden"
    return f"""<section class="month-pane" data-idx="{idx}"{hidden}>
  <div class="card">
    <div class="head">
      <div><div class="mtitle">{_e(m.month_label)}</div>
        <div class="sub">month-to-date · compared to {_e(prev)}</div></div>
      <div class="asof">Data as of <b>{asof}</b></div>
    </div>
    {_kpi_block(m)}
  </div>
  {notes}{miss_note}{pending}
  {platform_block}{daily_block}
</section>"""


def render(m: ReportModel, generated: _dt.datetime, freshness=None) -> str:
    asof = m.as_of.strftime("%d %b %Y") if m.as_of else "–"
    dim = m.days_in_month
    # Structure: all PLATFORM-LEVEL views together, then all DATE-WISE views together.
    platform_block = ('<div class="section-label">Platform view · month-to-date</div>'
                      + "".join(_summary_card(s) for s in m.sections))
    daily_block = ('<div class="section-label">Date-wise detail</div>'
                   + "".join(_daily_card(s, dim) for s in m.sections))
    body = platform_block + daily_block
    pending = (f'<div class="card" style="border-left:4px solid #b8860b;color:#7a5c00">'
               f'⚠ {_e(m.pending_note)}</div>') if m.pending_note else ""
    fresh_note = _freshness_note(freshness)
    foot = (
        "<b>Run-rate estimate</b> = MTD GMV ÷ days elapsed × days in month (the team's own "
        "method). <b>Growth</b> compares like-for-like same dates vs last month (LM GMV). "
        "<b>Ad Contri</b> = Ad Sales ÷ current GMV. "
        "<b>Amazon</b> GMV (and Blinkit ad spend) arrive on a T-2 lag, so the latest 1–2 days are "
        "pending, not a drop. <b>Amazon Core, NOW+Fresh, Flipkart, Minutes, First Club</b> are "
        "reviewed monthly, shown as MTD totals (no daily split). <b>Shopify</b> is D2C: GMV = gross "
        "sales. Figures read from a read-only snapshot of the source sheets; no data is altered or invented."
    )
    return f"""<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Daily Sales Report — {_e(m.month_label)}</title><style>{CSS}</style></head>
<body><div class="wrap">
  <div class="card">
    <div class="head">
      <div><div class="title">Daily Sales Report</div>
        <div class="sub">{_e(m.month_label)} · month-to-date</div></div>
      <div class="asof">Data as of <b>{asof}</b><br>Generated {generated.strftime('%d %b %Y, %H:%M')} IST</div>
    </div>
    {_kpi_block(m)}
  </div>
  {pending}
  {body}
  <div class="card foot">{foot}{fresh_note}</div>
</div></body></html>"""


def write_html(generated: _dt.datetime | None = None, model: ReportModel | None = None,
               freshness=None) -> Path:
    generated = generated or _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=5, minutes=30)))
    m = model or build_report()
    if freshness is None:
        import qc  # local import avoids any import-order coupling
        freshness = qc.compute_freshness(m)
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = config.OUTPUT_DIR / "index.html"   # single fixed location, overwritten each run
    out.write_text(render(m, generated, freshness), encoding="utf-8")
    return out


def render_multi(models: "list[ReportModel]", generated: _dt.datetime, freshness=None) -> str:
    """Multi-month page: a Month dropdown swaps between pre-rendered per-month
    panes. models[0] is the current month and is shown by default; the freshness
    footer reflects that current month."""
    if not models:
        return render(build_report(), generated, freshness)
    opts, panes = [], []
    for i, m in enumerate(models):
        prev = _prev_month_label(m.year, m.month)
        sel = " selected" if i == 0 else ""
        opts.append(f'<option value="{i}" data-prev="{_e(prev)}"{sel}>{_e(m.month_label)}</option>')
        panes.append(_month_pane(m, i, generated))
    first_prev = _prev_month_label(models[0].year, models[0].month)
    fresh_note = _freshness_note(freshness)
    foot = (
        "<b>Run-rate estimate</b> = MTD GMV ÷ days elapsed × days in month (the team's own "
        "method). <b>Growth</b> compares like-for-like same dates vs last month (LM GMV). "
        "<b>Ad Contri</b> = Ad Sales ÷ current GMV. "
        "<b>Amazon</b> GMV (and Blinkit ad spend) arrive on a T-2 lag, so the latest 1–2 days are "
        "pending, not a drop. <b>Amazon Core, NOW+Fresh, Flipkart, Minutes, First Club</b> are "
        "reviewed monthly, shown as MTD totals (no daily split). <b>Shopify</b> is D2C: GMV = gross "
        "sales. Figures read from a read-only snapshot of the source sheets; no data is altered or invented."
    )
    picker = (f'<div class="picker"><label for="monthPicker">Month</label>'
              f'<select id="monthPicker" onchange="__showMonth(this.value)">{"".join(opts)}</select>'
              f'<div class="cmp">Compared to <b><span id="cmpMonth">{_e(first_prev)}</span></b></div></div>')
    script = ("<script>function __showMonth(v){var s=document.getElementById('monthPicker');"
              "var o=s.options[s.selectedIndex];var p=document.querySelectorAll('.month-pane');"
              "for(var k=0;k<p.length;k++){p[k].hidden=(p[k].getAttribute('data-idx')!==v);}"
              "document.getElementById('cmpMonth').textContent=o.getAttribute('data-prev');}"
              "function __showPlat(sel){var pane=sel.closest('.month-pane');var v=sel.value;"
              "var p=pane.querySelectorAll('.ptbl-pane');"
              "for(var k=0;k<p.length;k++){p[k].hidden=(p[k].getAttribute('data-plat')!==v);}}</script>")
    return f"""<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Daily Sales Report — {_e(models[0].month_label)}</title><style>{CSS}</style></head>
<body><div class="wrap">
  <div class="card">
    <div class="head">
      <div><div class="title">Daily Sales Report</div>
        <div class="sub">Generated {generated.strftime('%d %b %Y, %H:%M')} IST</div></div>
      {picker}
    </div>
  </div>
  {''.join(panes)}
  <div class="card foot">{foot}{fresh_note}</div>
  {script}
</div></body></html>"""


def write_html_multi(models: "list[ReportModel]", generated: _dt.datetime | None = None,
                     freshness=None) -> Path:
    """Write the multi-month report to the single fixed output/index.html."""
    generated = generated or _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=5, minutes=30)))
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = config.OUTPUT_DIR / "index.html"
    out.write_text(render_multi(models, generated, freshness), encoding="utf-8")
    return out


if __name__ == "__main__":
    # Deterministic timestamp when run directly (avoids Date.now-style nondeterminism in tests).
    p = write_html(_dt.datetime(2026, 7, 21, 9, 0))
    print("wrote", p)
