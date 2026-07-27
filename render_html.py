"""Render the ReportModel into a self-contained HTML one-pager.

Pure-Python string building (no external template engine). Inline CSS only,
no network calls, so the file opens offline and screenshots cleanly.
"""
from __future__ import annotations
import datetime as _dt
import html
import json
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
  --tr-cur:#12608a; --tr-prev:#d9a441;
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
/* consolidated total */
.consol{display:flex;justify-content:space-between;align-items:center;gap:14px;flex-wrap:wrap;
  background:#eef4fb;border:1px solid #d3e3f2}
.consol .ct-lab{font-size:11px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;color:var(--brand2)}
.consol .ct-gmv{font-size:26px;font-weight:800;color:var(--brand);letter-spacing:-.4px;margin-top:3px;
  display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.consol .ct-gmv .badge{font-size:12px}
.consol .ct-meta{font-size:13px;color:var(--muted);text-align:right}
.consol .ct-meta b{color:var(--ink)}
@media (max-width:720px){ .consol{align-items:flex-start} .consol .ct-meta{text-align:left} .consol .ct-gmv{font-size:22px} }
/* trend chart */
.trend-head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:10px}
.trend-head .name{font-size:16px;font-weight:800;color:var(--brand2)}
.trend-ctrls{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.trend-ctrls label{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-left:6px}
.trend-ctrls select{font:inherit;font-size:13px;font-weight:700;color:var(--brand);background:var(--card);
  border:1px solid var(--line);border-radius:8px;padding:5px 10px;cursor:pointer}
.trend .tr-legend{display:flex;gap:16px;font-size:11px;color:var(--muted);margin:2px 0 8px}
.trend .tr-legend .lg{display:inline-flex;align-items:center;gap:6px}
.trend .tr-legend .sw{width:18px;border-top:3px solid var(--tr-cur)}
.trend .tr-legend .sw.prev{border-top:2px dashed var(--tr-prev)}
.trend svg{width:100%;height:auto;display:block}
@media (max-width:720px){ .trend-head{align-items:flex-start} .trend-ctrls label{margin-left:0} }
.tr-tip{position:fixed;z-index:30;background:var(--card);border:1px solid var(--line);border-radius:8px;
  box-shadow:0 4px 14px rgba(15,23,42,.14);padding:6px 9px;font-size:11px;color:var(--ink);
  pointer-events:none;white-space:nowrap}
.tr-tip b{color:var(--brand2)}
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


def _summary_table(sec: Section, total_label: str = "Total") -> str:
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
        <td>{_e(total_label)}</td>
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


def _summary_card(sec: Section, total_label: str = "Total") -> str:
    """Platform-level view for one section (summary table only)."""
    return f'<div class="card">{_section_header(sec)}{_summary_table(sec, total_label)}</div>'


def _consolidated_total(m: ReportModel) -> str:
    """Highlighted consolidated total across all sections (Quick Commerce +
    Marketplace & D2C). Shown only when the month has more than one section."""
    if len(m.sections) < 2:
        return ""
    g = m.grand
    up = (g.growth or 0) >= 0
    badge = f'<span class="badge {"up" if up else "down"}">{fmt.pct(g.growth)} vs LM</span>'
    names = " + ".join(s.name for s in m.sections)
    return (f'<div class="card consol">'
            f'<div><div class="ct-lab">Consolidated · {_e(names)}</div>'
            f'<div class="ct-gmv">{fmt.gmv_auto(g.gmv_mtd)} {badge}</div></div>'
            f'<div class="ct-meta">LM <b>{fmt.gmv_auto(g.lm_mtd)}</b> · '
            f'Ad <b>{fmt.gmv_auto(g.ad_mtd)}</b> · '
            f'Est <b>{fmt.gmv_auto(g.estimate)}</b> · Gap {_gap_span(g.gap)}</div></div>')


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


def _platform_pane(m: ReportModel, idx: int, generated: _dt.datetime) -> str:
    """Top pane for a month: header + KPIs + freshness notes + Platform view
    (summary cards + consolidated total). idx 0 visible; the rest start hidden.
    The Date-wise detail is a separate pane (_daily_pane) so the global Trend
    card can sit between the platform and daily blocks."""
    asof = m.as_of.strftime("%d %b %Y") if m.as_of else "–"
    prev = _prev_month_label(m.year, m.month)
    notes = _notes_card(m, generated)
    multi = len(m.sections) > 1
    total_label = "Sub total" if multi else "Total"
    platform_block = ('<div class="section-label">Platform view · month-to-date</div>'
                      + "".join(_summary_card(s, total_label) for s in m.sections)
                      + _consolidated_total(m))
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
  {platform_block}
</section>"""


def _daily_pane(m: ReportModel, idx: int) -> str:
    """Bottom pane for a month: the platform-filtered Date-wise detail. Shares
    the month-pane class + data-idx so the month dropdown toggles it in step with
    the platform pane; the global Trend card sits between the two pane groups."""
    hidden = "" if idx == 0 else " hidden"
    daily_block = ('<div class="section-label">Date-wise detail</div>'
                   + _daily_filter_card(m, m.days_in_month))
    return f'<section class="month-pane" data-idx="{idx}"{hidden}>{daily_block}</section>'


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


# Vanilla-JS trend chart (no libs). Draws two SVG polylines (current year solid
# blue, last year dashed amber) with axes, dots, legend, headline + hover tooltip.
_TREND_JS = (
    "var __MONTHS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];"
    "var __KIND={gmv:'money',units:'count',growth:'pct',ad:'money',estimate:'money',gap:'money'};"
    "function __fmtT(v,k){if(v==null)return'–';"
    "if(k=='pct')return(v*100).toFixed(1)+'%';"
    "if(k=='count')return Math.round(v).toLocaleString('en-IN');"
    "if(Math.abs(v)>=1e7)return'₹'+(v/1e7).toFixed(2)+' Cr';return'₹'+(v/1e5).toFixed(1)+' L';}"
    "function __drawTrend(){var el=document.getElementById('trendSvg');if(!el)return;"
    "var DATA=JSON.parse(document.getElementById('trendData').textContent);"
    "var plat=document.getElementById('trendPlat').value;"
    "var kpi=document.getElementById('trendKpi').value,year=document.getElementById('trendYear').value,kind=__KIND[kpi];"
    "var kd=(DATA[plat]||{})[kpi]||{};var cur=kd[year]||[];var py=(parseInt(year,10)-1)+'';var prev=kd[py]||null;"
    "var maxYear=Math.max.apply(null,Object.keys(kd).map(Number));"
    "var vals=[];for(var i=0;i<12;i++){if(cur[i]!=null)vals.push(cur[i]);if(prev&&prev[i]!=null)vals.push(prev[i]);}"
    "var lo=vals.length?Math.min.apply(null,vals):0,hi=vals.length?Math.max.apply(null,vals):1;"
    "var yMin=Math.min(0,lo),yMax=hi;if(yMax===yMin){yMax=yMin+1;}"
    "var span=yMax-yMin,pad=span*0.08;yMax+=pad;if(yMin<0)yMin-=pad;"
    "var W=760,H=300,pL=58,pR=16,pT=14,pB=30,pw=W-pL-pR,ph=H-pT-pB,baseY=(pT+ph).toFixed(1);"
    "function X(i){return pL+pw*(i/11);}function Y(v){return pT+ph*(1-(v-yMin)/(yMax-yMin));}"
    "var lastI=-1;for(var i=0;i<12;i++){if(cur[i]!=null){lastI=i;}}"
    "var s='<defs><linearGradient id=\"trendGrad\" x1=\"0\" y1=\"0\" x2=\"0\" y2=\"1\">'"
    "+'<stop offset=\"0\" stop-color=\"#12608a\" stop-opacity=\"0.22\"/>'"
    "+'<stop offset=\"1\" stop-color=\"#12608a\" stop-opacity=\"0\"/></linearGradient></defs>';"
    "for(var g=0;g<5;g++){var gv=yMin+(yMax-yMin)*g/4,gy=Y(gv);"
    "s+='<line x1=\"'+pL+'\" y1=\"'+gy.toFixed(1)+'\" x2=\"'+(W-pR)+'\" y2=\"'+gy.toFixed(1)+'\" stroke=\"#e5e9f0\"/>';"
    "s+='<text x=\"'+(pL-6)+'\" y=\"'+(gy+3).toFixed(1)+'\" text-anchor=\"end\" font-size=\"10\" fill=\"#64748b\">'+__fmtT(gv,kind)+'</text>';}"
    "for(var i=0;i<12;i++){s+='<text x=\"'+X(i).toFixed(1)+'\" y=\"'+(H-10)+'\" text-anchor=\"middle\" font-size=\"10\" fill=\"#64748b\">'+__MONTHS[i]+'</text>';}"
    "var ad='',apen=false,firstX=null;for(var i=0;i<12;i++){if(cur[i]!=null){var x=X(i).toFixed(1),y=Y(cur[i]).toFixed(1);"
    "if(!apen){ad+='M'+x+' '+baseY+' L'+x+' '+y+' ';firstX=x;apen=true;}else{ad+='L'+x+' '+y+' ';}}}"
    "if(apen){ad+='L'+X(lastI).toFixed(1)+' '+baseY+' Z';s+='<path d=\"'+ad+'\" fill=\"url(#trendGrad)\" stroke=\"none\"/>';}"
    "function line(arr,color,w,dash){if(!arr)return'';var d='',pen=false,dots='';"
    "for(var i=0;i<12;i++){if(arr[i]!=null){var x=X(i).toFixed(1),y=Y(arr[i]).toFixed(1);d+=(pen?'L':'M')+x+' '+y+' ';pen=true;"
    "dots+='<circle cx=\"'+x+'\" cy=\"'+y+'\" r=\"2.6\" fill=\"'+color+'\"/>';}else{pen=false;}}"
    "var ln=d?'<path d=\"'+d+'\" fill=\"none\" stroke=\"'+color+'\" stroke-width=\"'+w+'\" stroke-linejoin=\"round\" stroke-linecap=\"round\"'+(dash?' stroke-dasharray=\"6 4\"':'')+'/>':'';return ln+dots;}"
    "if(prev)s+=line(prev,'#d9a441',1.5,true);s+=line(cur,'#12608a',2,false);"
    "if((+year)===maxYear&&lastI>=0){s+='<circle cx=\"'+X(lastI).toFixed(1)+'\" cy=\"'+Y(cur[lastI]).toFixed(1)+'\" r=\"3.5\" fill=\"#ffffff\" stroke=\"#12608a\" stroke-width=\"2\"/>';}"
    "el.innerHTML=s;"
    "var lg='<span class=\"lg\"><span class=\"sw\"></span>'+year+'</span>';if(prev)lg+='<span class=\"lg\"><span class=\"sw prev\"></span>'+py+'</span>';"
    "document.getElementById('trendLegend').innerHTML=lg;"
    "var tip=document.getElementById('trendTip');"
    "el.onmousemove=function(ev){var r=el.getBoundingClientRect();var rel=(ev.clientX-r.left)/r.width*W;"
    "var i=Math.round((rel-pL)/pw*11);if(i<0)i=0;if(i>11)i=11;var c=cur[i],p=prev?prev[i]:null;"
    "if(c==null&&p==null){tip.hidden=true;return;}var t='<b>'+__MONTHS[i]+'</b> '+year+': '+__fmtT(c,kind);"
    "if(prev)t+=' · '+py+': '+__fmtT(p,kind);tip.innerHTML=t;tip.hidden=false;"
    "tip.style.left=(ev.clientX+12)+'px';tip.style.top=(ev.clientY-10)+'px';};"
    "el.onmouseleave=function(){document.getElementById('trendTip').hidden=true;};}"
    "__drawTrend();"
)


# KPI key -> (label, grand-attr). Order = dropdown order.
_TREND_KPIS = [
    ("gmv", "GMV", "gmv_mtd"),
    ("units", "Units", "units_mtd"),
    ("growth", "Growth", "growth"),
    ("ad", "Ad Sales", "ad_mtd"),
    ("estimate", "Estimation", "estimate"),
    ("gap", "Gap", "gap"),
]


def _trend_dataset(models):
    """{platform: {kpi: {year(str): [12 monthly values, null]}}} + years + platforms.
    'All' = grand totals; each daily-cadence channel keyed by its name."""
    years = sorted({m.year for m in models})
    order = [n for sec in config.SECTIONS.values() for n in sec["channels"]]
    seen = {c.name for m in models for s in m.sections for c in s.channels if c.cadence == "daily"}
    platforms = ["All"] + [n for n in order if n in seen]

    def blank():
        return {k: {str(y): [None] * 12 for y in years} for k, _, _ in _TREND_KPIS}

    data = {p: blank() for p in platforms}
    for m in models:
        yr, mi = str(m.year), m.month - 1
        for k, _, attr in _TREND_KPIS:
            data["All"][k][yr][mi] = getattr(m.grand, attr)
        for s in m.sections:
            for c in s.channels:
                if c.name in data:
                    for k, _, attr in _TREND_KPIS:
                        data[c.name][k][yr][mi] = getattr(c, attr)
    return data, years, platforms


def _trend_card(models) -> str:
    """Global YoY trend chart (inline SVG drawn by JS). Own KPI + Year selectors;
    plots the selected year vs the year before it (omitted when absent)."""
    data, years, platforms = _trend_dataset(models)
    if not years:
        return ""
    years_desc = sorted(years, reverse=True)
    plat_opts = "".join(f'<option value="{_e(p)}"{" selected" if p == "All" else ""}>{_e(p)}</option>'
                        for p in platforms)
    kpi_opts = "".join(f'<option value="{k}"{" selected" if k == "gmv" else ""}>{_e(lbl)}</option>'
                       for k, lbl, _ in _TREND_KPIS)
    year_opts = "".join(f'<option value="{y}"{" selected" if i == 0 else ""}>{y}</option>'
                        for i, y in enumerate(years_desc))
    blob = json.dumps(data, separators=(",", ":"))
    return (
        '<div class="card trend">'
        '<div class="trend-head"><div class="name">Trend · year over year</div>'
        '<div class="trend-ctrls">'
        '<label for="trendPlat">Platform</label>'
        f'<select id="trendPlat" onchange="__drawTrend()">{plat_opts}</select>'
        '<label for="trendKpi">KPI</label>'
        f'<select id="trendKpi" onchange="__drawTrend()">{kpi_opts}</select>'
        '<label for="trendYear">Year</label>'
        f'<select id="trendYear" onchange="__drawTrend()">{year_opts}</select></div></div>'
        '<div class="tr-legend" id="trendLegend"></div>'
        '<div class="tw"><svg id="trendSvg" viewBox="0 0 760 300" '
        'preserveAspectRatio="xMidYMid meet" role="img" aria-label="Monthly trend"></svg></div>'
        '<div class="grid-note" style="margin-top:6px">Monthly totals across the year for the chosen '
        'platform; the latest month is month-to-date (hollow point). Current year solid, last year dashed.</div>'
        f'<script id="trendData" type="application/json">{blob}</script></div>'
        '<div class="tr-tip" id="trendTip" hidden></div>'
    )


def render_multi(models: "list[ReportModel]", generated: _dt.datetime, freshness=None) -> str:
    """Multi-month page: a Month dropdown swaps between pre-rendered per-month
    panes. models[0] is the current month and is shown by default; the freshness
    footer reflects that current month."""
    if not models:
        return render(build_report(), generated, freshness)
    opts, plat_panes, day_panes = [], [], []
    for i, m in enumerate(models):
        prev = _prev_month_label(m.year, m.month)
        sel = " selected" if i == 0 else ""
        opts.append(f'<option value="{i}" data-prev="{_e(prev)}"{sel}>{_e(m.month_label)}</option>')
        plat_panes.append(_platform_pane(m, i, generated))
        day_panes.append(_daily_pane(m, i))
    trend = _trend_card(models)
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
              "for(var k=0;k<p.length;k++){p[k].hidden=(p[k].getAttribute('data-plat')!==v);}}"
              + _TREND_JS + "</script>")
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
  {''.join(plat_panes)}
  {trend}
  {''.join(day_panes)}
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
