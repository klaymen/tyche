#!/usr/bin/env python3
"""Generate a manloading dashboard HTML from sample_input.xlsx."""

from __future__ import annotations

VERSION = "1.0.3"

import argparse
from datetime import datetime
from pathlib import Path
import json

import pandas as pd

# ── Config ────────────────────────────────────────────────────────────────────
_parser = argparse.ArgumentParser(description="Generate a manloading dashboard HTML.")
_parser.add_argument(
    "filename",
    nargs="?",
    default=str(Path(__file__).parent / "sample_input.xlsx"),
    help="Path to the input Excel file (default: sample_input.xlsx next to this script)",
)
_args = _parser.parse_args()

INPUT  = Path(_args.filename)
OUTPUT = Path(__file__).parent / "output" / "manloading_report.html"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

# ── Load data ─────────────────────────────────────────────────────────────────
raw = pd.read_excel(INPUT, header=1)          # row 0 = "Months" label, row 1 = real headers
raw.columns = [str(c).strip() for c in raw.columns]

# Month columns are everything after Employee, Project, Project Group, Probability, Hourly Rate
MONTH_COLS     = [c for c in raw.columns if c not in ("Employee", "Project", "Project Group", "Probability", "Hourly Rate")]
EMPLOYEES      = sorted(raw["Employee"].unique().tolist())
PROJECTS       = sorted(raw["Project"].unique().tolist())
PROJECT_GROUPS = sorted(raw["Project Group"].unique().tolist()) if "Project Group" in raw.columns else []

# Probability column (default 1.0 if missing)
if "Probability" not in raw.columns:
    raw["Probability"] = 1.0
raw["Probability"] = raw["Probability"].fillna(1.0)

# Project group lookup: {project: group}
_proj_group: dict[str, str] = (
    raw.drop_duplicates("Project").set_index("Project")["Project Group"].to_dict()
    if "Project Group" in raw.columns else {}
)

# Probability-weighted FTE per row
for _m in MONTH_COLS:
    raw[f"_pw_{_m}"] = raw[_m] * raw["Probability"]
_PW_COLS = [f"_pw_{m}" for m in MONTH_COLS]

# Per-project probability-weighted FTE per month
proj_prob_month = raw.groupby("Project")[_PW_COLS].sum()
proj_prob_month.columns = MONTH_COLS   # rename back

# Per-employee probability-weighted FTE per month
emp_pw_month = raw.groupby("Employee")[_PW_COLS].sum()
emp_pw_month.columns = MONTH_COLS

# Per-project average probability
proj_avg_prob: dict[str, float] = raw.groupby("Project")["Probability"].mean().to_dict()

# ── Charge bases (multiply by monthly hours at display time)
# data-base      = FTE × Probability × Hourly Rate  (all factors on)
# data-base-fte  = FTE only
# data-base-prob = FTE × Probability
# data-base-rate = FTE × Hourly Rate (no probability)
_hr = raw["Hourly Rate"].fillna(0) if "Hourly Rate" in raw.columns else 0
for _m in MONTH_COLS:
    raw[f"_cb_{_m}"]      = raw[_m] * raw["Probability"] * _hr   # full base
    raw[f"_cbf_{_m}"]     = raw[_m]                               # FTE only
    raw[f"_cbp_{_m}"]     = raw[_m] * raw["Probability"]          # FTE × Prob
    raw[f"_cbr_{_m}"]     = raw[_m] * _hr                         # FTE × Rate
_CB_COLS  = [f"_cb_{m}"  for m in MONTH_COLS]
_CBF_COLS = [f"_cbf_{m}" for m in MONTH_COLS]
_CBP_COLS = [f"_cbp_{m}" for m in MONTH_COLS]
_CBR_COLS = [f"_cbr_{m}" for m in MONTH_COLS]

def _make_charge_tables(groupby_col, cb, cbf, cbp, cbr):
    """Return (base, base_fte, base_prob, base_rate) DataFrames grouped by col."""
    def _agg(cols):
        t = raw.groupby(groupby_col)[cols].sum()
        t.columns = MONTH_COLS
        return t
    return _agg(cb), _agg(cbf), _agg(cbp), _agg(cbr)

proj_charge_base, proj_charge_fte, proj_charge_prob, proj_charge_rate = \
    _make_charge_tables("Project", _CB_COLS, _CBF_COLS, _CBP_COLS, _CBR_COLS)

if "Project Group" in raw.columns:
    grp_charge_base, grp_charge_fte, grp_charge_prob, grp_charge_rate = \
        _make_charge_tables("Project Group", _CB_COLS, _CBF_COLS, _CBP_COLS, _CBR_COLS)
else:
    grp_charge_base = grp_charge_fte = grp_charge_prob = grp_charge_rate = None

emp_charge_base, emp_charge_fte, emp_charge_prob, emp_charge_rate = \
    _make_charge_tables("Employee", _CB_COLS, _CBF_COLS, _CBP_COLS, _CBR_COLS)

def _slug(s: str) -> str:
    """URL-safe id slug for a project name."""
    return s.lower().replace(" ", "-").replace(".", "").replace("(", "").replace(")", "")

# ── Computed metrics ──────────────────────────────────────────────────────────
# Per-employee total FTE per month (sum across all their projects)
emp_month = (
    raw.groupby("Employee")[MONTH_COLS].sum()
)

# Per-project total FTE per month
proj_month = (
    raw.groupby("Project")[MONTH_COLS].sum()
)

# Summary KPIs
total_employees    = len(EMPLOYEES)
total_projects     = len(PROJECTS)
total_proj_groups  = len(PROJECT_GROUPS)
total_months       = len(MONTH_COLS)
total_fte_months   = raw[MONTH_COLS].values.sum()
total_exp_fte      = raw[_PW_COLS].values.sum()   # probability-weighted total
avg_fte_per_month  = raw.groupby("Employee")[MONTH_COLS].sum().mean(axis=1).mean()

# Overallocation: any employee+month where FTE > 1.05
over_count    = int((emp_month    > 1.05).values.sum())  # raw
over_count_pw = int((emp_pw_month > 1.05).values.sum())  # prob-weighted

# ── Colour helpers ────────────────────────────────────────────────────────────
def fte_color(v: float) -> str:
    """CSS class for a FTE utilisation value."""
    if v == 0:    return "fte-zero"
    if v < 0.5:   return "fte-low"   # yellow
    if v < 1.0:   return "fte-mid"   # blue
    if v <= 1.0:  return "fte-full"  # green (exactly 1.0)
    return "fte-over"                 # red

# ── HTML helpers ──────────────────────────────────────────────────────────────
def pct_bar(value: float, max_val: float, color: str = "#2e5fb0") -> str:
    pct = min(value / max_val * 100, 100) if max_val else 0
    return (
        f'<div class="bar-bg">'
        f'<div class="bar-fill" style="width:{pct:.1f}%;background:{color}"></div>'
        f'</div>'
    )

def fmt(v: float) -> str:
    return f"{v:.2f}"

# ── Build sections ────────────────────────────────────────────────────────────
def stat_cards() -> str:
    cards = [
        ("Employees",           total_employees,           "var(--accent)",
         "Total unique employees with allocations in this report."),
        ("Projects",            total_projects,            "var(--accent)",
         "Total unique projects being tracked."),
        ("Project Groups",      total_proj_groups,         "var(--accent)",
         "Number of distinct project groups (portfolios) in this report."),
        ("Total FTE\u00b7months",    f"{total_fte_months:.1f}", "var(--accent)",
         "Sum of all FTE allocations across every employee, project, and month."),
        ("Expected FTE\u00b7months", f"{total_exp_fte:.1f}",  "var(--ok)",
         "Probability-weighted sum of FTE allocations (FTE \u00d7 probability)."),
        ("Avg FTE / employee",  f"{avg_fte_per_month:.2f}", "var(--accent)",
         "Average total FTE per employee averaged across all months."),
        ("Overallocations (Raw)",      over_count,    "var(--danger)" if over_count    else "var(--ok)",
         "Employee\u2009\u00d7\u2009month cells where raw allocated FTE exceeds 1.05 (ignores probability)."),
        ("Overallocations (Expected)",  over_count_pw, "var(--danger)" if over_count_pw else "var(--ok)",
         "Employee\u2009\u00d7\u2009month cells where probability-weighted FTE exceeds 1.05."),
    ]
    html = ""
    for label, val, color, tip in cards:
        html += f"""
        <div class="stat-card">
          <div class="stat-card-header">
            <h3>{label}</h3>
            <span class="tip" tabindex="0" data-tip="{tip}" aria-label="{label} help">?</span>
          </div>
          <div class="stat-value" style="color:{color}">{val}</div>
        </div>"""
    return html

# -- Per-employee project breakdown lookup (for heatmap tooltips) -------------
_emp_proj_month: dict = {}  # {emp: {month: [(proj, fte, prob), ...]}}
for _, _r in raw.iterrows():
    _e, _p = _r["Employee"], _r["Project"]
    _prob  = float(_r.get("Probability", 1.0))
    _emp_proj_month.setdefault(_e, {})
    for _m in MONTH_COLS:
        _emp_proj_month[_e].setdefault(_m, []).append((_p, float(_r[_m]), _prob))


def _backward_cum_row(month_totals, label_cell: str, extra_cell: str = "") -> str:
    """Return a <tr> showing backward-cumulative FTE: each cell = sum(vals[i:])."""
    vals = [float(v) for v in month_totals]
    cells = "".join(
        f'<td class="fte-cell-cum">{fmt(sum(vals[i:]))}</td>'
        for i in range(len(vals))
    )
    return (
        f'<tr class="cum-row">'
        f'{label_cell}'
        f'{cells}'
        f'{extra_cell}'
        f'</tr>'
    )


def _util_tooltip(emp: str, month: str, total: float) -> str:
    """Build tooltip string: each project line (with prob if <100%) + total."""
    breakdown = _emp_proj_month.get(emp, {}).get(month, [])
    lines = []
    for p, v, prob in breakdown:
        if v > 0:
            prob_str = f" ({prob:.0%})" if prob < 1.0 else ""
            lines.append(f"{p}{prob_str}: {v:.2f}")
    return "\n".join(lines)


def utilisation_heatmap() -> str:
    """Employee × Month heatmap coloured by probability-weighted utilisation."""
    # Month header row
    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    rows = f"""
    <thead>
      <tr>
        <th class="th-emp">Employee</th>
        {th_months}
      </tr>
    </thead>"""
    rows = '<table class="heatmap-table" style="width:100%;min-width:600px">' + rows

    rows += "<tbody>"
    for emp in EMPLOYEES:
        if emp not in emp_pw_month.index:
            continue
        vals     = emp_pw_month.loc[emp]
        raw_vals = emp_month.loc[emp] if emp in emp_month.index else vals
        avg  = vals.mean()
        def _cell(m, v, _raw=raw_vals, _emp=emp):
            raw_v  = float(_raw[m])
            marker = ('<span class="fte-over-marker" title="Raw FTE: ' + fmt(raw_v) + '">'
                      '<svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">'
                      '<circle cx="12" cy="12" r="10"/>'
                      '<polyline points="12 6 12 12 16 14"/>'
                      '</svg></span>') if raw_v > 1.05 else ""
            return f'<td class="fte-cell {fte_color(v)}" data-cell-tip="{_util_tooltip(_emp, m, v)}">{marker}{fmt(v)}</td>'
        cells = "".join(_cell(m, v) for m, v in zip(MONTH_COLS, vals))
        avg_cls = fte_color(avg)
        rows += (
            f'<tr>'
            f'<td class="td-emp">{emp}</td>'
            f'{cells}'
            f'</tr>'
        )
    # Totals row
    month_totals = emp_pw_month.sum()
    total_cells = "".join(
        f'<td class="fte-cell-total">{fmt(v)}</td>'
        for v in month_totals
    )
    rows += (
        f'<tr class="heatmap-total-row">'
        f'<td class="td-emp total-label">TOTAL</td>'
        f'{total_cells}'
        f'</tr>'
    )
    rows += _backward_cum_row(month_totals, '<td class="td-emp cum-label">REMAINING</td>')
    rows += "</tbody></table>"
    return rows

def project_summary_rows() -> str:
    """Project × Month table, probability-weighted FTE per project."""
    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    html = f"""
    <thead>
      <tr>
        <th class="th-emp">Project</th>
        {th_months}
        <th class="th-month">Total</th>
      </tr>
    </thead><tbody>"""

    color_map = {
        "Project Alpha": "#3b5998",
        "Project Beta":  "#059669",
        "Project Gamma": "#d97706",
    }

    for proj in PROJECTS:
        if proj not in proj_prob_month.index:
            continue
        vals      = proj_prob_month.loc[proj, MONTH_COLS]
        row_total = float(vals.sum())
        color = color_map.get(proj, "#5f6978")
        cells = "".join(
            f'<td class="td-num">{fmt(v)}</td>'
            for v in vals
        )
        html += (
            f'<tr>'
            f'<td class="td-proj"><span class="proj-dot" style="background:{color}"></span>{proj}</td>'
            f'{cells}'
            f'<td class="td-total" style="color:{color}">{fmt(row_total)}</td>'
            f'</tr>'
        )

    # Totals row
    month_totals = proj_prob_month[MONTH_COLS].sum()
    grand_total  = float(month_totals.sum())
    total_cells  = "".join(f'<td>{fmt(v)}</td>' for v in month_totals)
    html += (
        f'<tr class="proj-table-total">'
        f'<td>Total</td>'
        f'{total_cells}'
        f'<td>{fmt(grand_total)}</td>'
        f'</tr>'
    )
    html += _backward_cum_row(month_totals, '<td class="cum-label">REMAINING</td>', f'<td class="fte-cell-cum">{fmt(grand_total)}</td>')
    html += "</tbody>"
    return html

# Stable colour palette for project groups (cycles if >6 groups)
_GROUP_COLORS = ["#7c3aed", "#0369a1", "#b45309", "#be123c", "#047857", "#1d4ed8"]
_GROUP_CHART_BG = [
    "rgba(124,58,237,0.15)", "rgba(3,105,161,0.15)", "rgba(180,83,9,0.15)",
    "rgba(190,18,60,0.15)",  "rgba(4,120,87,0.15)",  "rgba(29,78,216,0.15)",
]
_group_color_map: dict[str, str] = {
    g: _GROUP_COLORS[i % len(_GROUP_COLORS)] for i, g in enumerate(PROJECT_GROUPS)
}

# Module-level prob-weighted FTE per group per month
grp_pw_month = raw.groupby("Project Group")[_PW_COLS].sum()
grp_pw_month.columns = MONTH_COLS

_group_chart_datasets = []
for _gi, _g in enumerate(PROJECT_GROUPS):
    if _g not in grp_pw_month.index:
        continue
    _group_chart_datasets.append({
        "label": _g,
        "data": [round(float(v), 4) for v in grp_pw_month.loc[_g, MONTH_COLS]],
        "borderColor": _GROUP_COLORS[_gi % len(_GROUP_COLORS)],
        "backgroundColor": _GROUP_CHART_BG[_gi % len(_GROUP_CHART_BG)],
        "borderWidth": 2,
        "pointRadius": 4,
        "pointHoverRadius": 6,
        "fill": True,
        "tension": 0.35,
    })
group_chart_data_json = json.dumps({"labels": MONTH_COLS, "datasets": _group_chart_datasets})

def group_fte_demand_rows() -> str:
    """Project Group × Month table, probability-weighted FTE aggregated per group."""
    grp_pw = grp_pw_month

    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    html = f"""
    <thead>
      <tr>
        <th class="th-emp">Project Group</th>
        {th_months}
        <th class="th-month">Total</th>
      </tr>
    </thead><tbody>"""

    for group in PROJECT_GROUPS:
        if group not in grp_pw.index:
            continue
        vals      = grp_pw.loc[group, MONTH_COLS]
        row_total = float(vals.sum())
        color     = _group_color_map.get(group, "#5f6978")
        cells     = "".join(f'<td class="td-num">{fmt(v)}</td>' for v in vals)
        proj_count = len([p for p in PROJECTS if _proj_group.get(p) == group])
        html += (
            f'<tr>'
            f'<td class="td-proj">'
            f'<span class="proj-dot" style="background:{color}"></span>'
            f'{group} <span style="color:#8fa3bc;font-size:.78rem">({proj_count} project{"s" if proj_count != 1 else ""})</span>'
            f'</td>'
            f'{cells}'
            f'<td class="td-total" style="color:{color}">{fmt(row_total)}</td>'
            f'</tr>'
        )

    month_totals = grp_pw[MONTH_COLS].sum()
    grand_total  = float(month_totals.sum())
    total_cells  = "".join(f'<td>{fmt(v)}</td>' for v in month_totals)
    html += (
        f'<tr class="proj-table-total">'
        f'<td>Total</td>'
        f'{total_cells}'
        f'<td>{fmt(grand_total)}</td>'
        f'</tr>'
    )
    html += _backward_cum_row(month_totals, '<td class="cum-label">REMAINING</td>', f'<td class="fte-cell-cum">{fmt(grand_total)}</td>')
    html += "</tbody>"
    return html

def employee_detail_rows() -> str:
    """Flat Employee × Project detail table."""
    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    html = f"""
    <thead>
      <tr>
        <th class="th-emp sortable">Employee</th>
        <th class="th-emp sortable">Project</th>
        <th class="th-emp sortable">Group</th>
        <th class="th-month sortable">Prob.</th>
        {th_months}
        <th class="th-month sortable">TOTAL</th>
      </tr>
    </thead><tbody>"""

    prev_emp = None
    for _, row in raw.sort_values(["Employee", "Project"]).iterrows():
        emp   = row["Employee"]
        proj  = row["Project"]
        group = row.get("Project Group", "")
        prob  = float(row.get("Probability", 1.0))
        vals  = [row[m] for m in MONTH_COLS]
        total = sum(vals)
        shade = "tr-shade" if emp != prev_emp and (EMPLOYEES.index(emp) % 2 == 0) else ""
        prev_emp = emp

        cells = "".join(
            f'<td class="td-num">{fmt(v)}</td>'
            for v in vals
        )
        prob_cls = "prob-high" if prob >= 0.85 else ("prob-mid" if prob >= 0.60 else "prob-low")
        html += (
            f'<tr class="{shade}" data-emp="{emp}" data-proj="{proj}" data-group="{group}">'
            f'<td class="td-emp">{emp}</td>'
            f'<td class="td-proj">{proj}</td>'
            f'<td class="td-group">{group}</td>'
            f'<td class="td-prob"><span class="prob-badge {prob_cls}">{prob:.0%}</span></td>'
            f'{cells}'
            f'<td class="td-total"><strong>{fmt(total)}</strong></td>'
            f'</tr>'
        )
    html += "</tbody>"
    # Grand total row in tfoot (always visible, unaffected by filter/sort)
    month_totals = raw[MONTH_COLS].sum()
    grand_total  = month_totals.sum()
    total_cells  = "".join(
        f'<td>{fmt(v)}</td>'
        for v in month_totals
    )
    cum_cells = "".join(
        f'<td>{fmt(sum(float(v) for v in month_totals.iloc[i:]))}</td>'
        for i in range(len(month_totals))
    )
    html += (
        f'<tfoot>'
        f'<tr class="proj-table-total">'
        f'<td colspan="4">Total</td>'
        f'{total_cells}'
        f'<td>{fmt(grand_total)}</td>'
        f'</tr>'
        f'<tr class="cum-row">'
        f'<td colspan="4" class="cum-label">REMAINING</td>'
        f'{cum_cells}'
        f'<td>{fmt(grand_total)}</td>'
        f'</tr>'
        f'</tfoot>'
    )
    return html

def project_heatmaps() -> str:
    """One heatmap section per project — rows = employees on that project."""
    html = ""
    for proj in PROJECTS:
        proj_rows = raw[raw["Project"] == proj].groupby("Employee")[MONTH_COLS].sum()
        emps = sorted(proj_rows.index.tolist())
        if not emps:
            continue

        th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
        table = (
            f'<table class="heatmap-table" style="width:100%;min-width:600px">'
            f'<thead><tr><th class="th-emp">Employee</th>{th_months}'
            f'</tr></thead><tbody>'
        )
        for emp in emps:
            vals = proj_rows.loc[emp, MONTH_COLS]
            cells = "".join(
                f'<td class="fte-cell-proj {fte_color(float(v))}">{fmt(float(v))}</td>'
                for m, v in zip(MONTH_COLS, vals)
            )
            table += (
                f'<tr>'
                f'<td class="td-emp">{emp}</td>'
                f'{cells}'
                f'</tr>'
            )
        # Totals row
        month_totals = proj_rows[MONTH_COLS].sum()
        total_cells = "".join(
            f'<td class="fte-cell-total">{fmt(v)}</td>'
            for v in month_totals
        )
        table += (
            f'<tr class="heatmap-total-row">'
            f'<td class="td-emp total-label">TOTAL</td>'
            f'{total_cells}'
            f'</tr>'
        )
        table += _backward_cum_row(month_totals, '<td class="td-emp cum-label">REMAINING</td>')
        table += "</tbody></table>"

        html += f"""
  <section class="surface" id="sec-proj-{_slug(proj)}">
    <div class="card-header">
      <div>
        <p class="section-title">{proj} · Employee Heatmap</p>
        <p class="section-sub">FTE per employee per month on this project</p>
      </div>
    </div>
    <div class="table-wrap">
      <div class="table-scroll">{table}</div>
    </div>
  </section>"""
    return html


def project_group_heatmaps() -> str:
    """One heatmap per project group — only when >1 group exists and the group has >1 project."""
    if len(PROJECT_GROUPS) <= 1:
        return ""

    html = ""
    for group in PROJECT_GROUPS:
        group_projs = [p for p in PROJECTS if _proj_group.get(p) == group]
        if len(group_projs) <= 1:
            continue

        group_data      = raw[raw["Project Group"] == group]
        emp_group_month = group_data.groupby("Employee")[MONTH_COLS].sum()
        emps = sorted(emp_group_month.index.tolist())
        if not emps:
            continue

        th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
        table = (
            f'<table class="heatmap-table" style="width:100%;min-width:600px">'
            f'<thead><tr><th class="th-emp">Employee</th>{th_months}'
            f'</tr></thead><tbody>'
        )
        for emp in emps:
            vals  = emp_group_month.loc[emp, MONTH_COLS]
            cells = "".join(
                f'<td class="fte-cell {fte_color(v)}">{fmt(v)}</td>'
                for m, v in zip(MONTH_COLS, vals)
            )
            table += f'<tr><td class="td-emp">{emp}</td>{cells}</tr>'

        month_totals = emp_group_month[MONTH_COLS].sum()
        total_cells  = "".join(f'<td class="fte-cell-total">{fmt(v)}</td>' for v in month_totals)
        table += (
            f'<tr class="heatmap-total-row">'
            f'<td class="td-emp total-label">TOTAL</td>'
            f'{total_cells}'
            f'</tr>'
        )
        table += _backward_cum_row(month_totals, '<td class="td-emp cum-label">REMAINING</td>')
        table += "</tbody></table>"

        proj_list = ", ".join(group_projs)
        html += f"""
  <section class="surface" id="sec-group-{_slug(group)}">
    <div class="card-header">
      <div>
        <p class="section-title">{group} · Employee Heatmap</p>
        <p class="section-sub">Combined FTE per employee across {len(group_projs)} projects: {proj_list}</p>
      </div>
    </div>
    <div class="table-wrap">
      <div class="table-scroll">{table}</div>
    </div>
  </section>"""
    return html


_DEFAULT_MONTHLY_HOURS = 145

def charge_tables_html() -> str:
    """Project × Month and Group × Month charge tables (base values; JS multiplies by monthly hours)."""
    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)

    def _cell_attrs(base_df, fte_df, prob_df, rate_df, key, m):
        b  = float(base_df.loc[key, m])
        bf = float(fte_df.loc[key, m])
        bp = float(prob_df.loc[key, m])
        br = float(rate_df.loc[key, m])
        return f'data-base="{b:.4f}" data-base-fte="{bf:.4f}" data-base-prob="{bp:.4f}" data-base-rate="{br:.4f}"'

    def _total_attrs(base_df, fte_df, prob_df, rate_df, key):
        b  = float(base_df.loc[key, MONTH_COLS].sum())
        bf = float(fte_df.loc[key, MONTH_COLS].sum())
        bp = float(prob_df.loc[key, MONTH_COLS].sum())
        br = float(rate_df.loc[key, MONTH_COLS].sum())
        return f'data-base="{b:.4f}" data-base-fte="{bf:.4f}" data-base-prob="{bp:.4f}" data-base-rate="{br:.4f}"', b

    def _month_totals_attrs(base_df, fte_df, prob_df, rate_df):
        rows = []
        for i, m in enumerate(MONTH_COLS):
            b  = float(base_df[m].sum())
            bf = float(fte_df[m].sum())
            bp = float(prob_df[m].sum())
            br = float(rate_df[m].sum())
            rows.append((b, bf, bp, br))
        return rows

    # ── Project table ─────────────────────────────────────────────
    proj_html = f"""<table id="charge-proj-table" class="charge-table filterable-table">
    <thead><tr>
      <th class="th-emp">Project</th>{th_months}<th class="th-month">Total</th>
    </tr></thead><tbody>"""

    for i, proj in enumerate(PROJECTS):
        if proj not in proj_charge_base.index:
            continue
        color = PROJECT_COLORS[i % len(PROJECT_COLORS)]["border"]
        cells = "".join(
            f'<td class="td-num charge-cell" {_cell_attrs(proj_charge_base, proj_charge_fte, proj_charge_prob, proj_charge_rate, proj, m)}>'
            f'{float(proj_charge_base.loc[proj, m]) * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
            for m in MONTH_COLS
        )
        t_attrs, row_base = _total_attrs(proj_charge_base, proj_charge_fte, proj_charge_prob, proj_charge_rate, proj)
        proj_html += (
            f'<tr>'
            f'<td class="td-proj"><span class="proj-dot" style="background:{color}"></span>{proj}</td>'
            f'{cells}'
            f'<td class="td-total charge-total-cell" {t_attrs} style="color:{color}">'
            f'{row_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
            f'</tr>'
        )

    mt = _month_totals_attrs(proj_charge_base, proj_charge_fte, proj_charge_prob, proj_charge_rate)
    grand_base = sum(r[0] for r in mt)
    grand_fte  = sum(r[1] for r in mt)
    grand_prob = sum(r[2] for r in mt)
    grand_rate = sum(r[3] for r in mt)
    total_cells = "".join(
        f'<td data-base="{r[0]:.4f}" data-base-fte="{r[1]:.4f}" data-base-prob="{r[2]:.4f}" data-base-rate="{r[3]:.4f}">'
        f'{r[0] * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
        for r in mt
    )
    cum_cells = "".join(
        f'<td class="fte-cell-cum charge-cell"'
        f' data-base="{sum(r[0] for r in mt[i:]):.4f}"'
        f' data-base-fte="{sum(r[1] for r in mt[i:]):.4f}"'
        f' data-base-prob="{sum(r[2] for r in mt[i:]):.4f}"'
        f' data-base-rate="{sum(r[3] for r in mt[i:]):.4f}">'
        f'{sum(r[0] for r in mt[i:]) * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
        for i in range(len(mt))
    )
    proj_html += (
        f'<tr class="proj-table-total"><td>Total</td>{total_cells}'
        f'<td data-base="{grand_base:.4f}" data-base-fte="{grand_fte:.4f}" data-base-prob="{grand_prob:.4f}" data-base-rate="{grand_rate:.4f}">'
        f'{grand_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td></tr>'
        f'<tr class="cum-row"><td class="cum-label">REMAINING</td>{cum_cells}'
        f'<td class="fte-cell-cum" data-base="{grand_base:.4f}" data-base-fte="{grand_fte:.4f}" data-base-prob="{grand_prob:.4f}" data-base-rate="{grand_rate:.4f}">'
        f'{grand_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td></tr>'
    )
    proj_html += "</tbody></table>"

    # ── Group table ───────────────────────────────────────────────
    grp_html = ""
    if grp_charge_base is not None and len(PROJECT_GROUPS) > 1:
        grp_html = f"""<table id="charge-grp-table" class="charge-table filterable-table">
    <thead><tr>
      <th class="th-emp">Project Group</th>{th_months}<th class="th-month">Total</th>
    </tr></thead><tbody>"""

        for gi, group in enumerate(PROJECT_GROUPS):
            if group not in grp_charge_base.index:
                continue
            color = _GROUP_COLORS[gi % len(_GROUP_COLORS)]
            cells = "".join(
                f'<td class="td-num charge-cell" {_cell_attrs(grp_charge_base, grp_charge_fte, grp_charge_prob, grp_charge_rate, group, m)}>'
                f'{float(grp_charge_base.loc[group, m]) * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
                for m in MONTH_COLS
            )
            t_attrs, row_base = _total_attrs(grp_charge_base, grp_charge_fte, grp_charge_prob, grp_charge_rate, group)
            grp_html += (
                f'<tr>'
                f'<td class="td-proj"><span class="proj-dot" style="background:{color}"></span>{group}</td>'
                f'{cells}'
                f'<td class="td-total charge-total-cell" {t_attrs} style="color:{color}">'
                f'{row_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
                f'</tr>'
            )

        gmt = _month_totals_attrs(grp_charge_base, grp_charge_fte, grp_charge_prob, grp_charge_rate)
        gg_base = sum(r[0] for r in gmt)
        gg_fte  = sum(r[1] for r in gmt)
        gg_prob = sum(r[2] for r in gmt)
        gg_rate = sum(r[3] for r in gmt)
        grp_total_cells = "".join(
            f'<td data-base="{r[0]:.4f}" data-base-fte="{r[1]:.4f}" data-base-prob="{r[2]:.4f}" data-base-rate="{r[3]:.4f}">'
            f'{r[0] * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
            for r in gmt
        )
        grp_cum_cells = "".join(
            f'<td class="fte-cell-cum charge-cell"'
            f' data-base="{sum(r[0] for r in gmt[i:]):.4f}"'
            f' data-base-fte="{sum(r[1] for r in gmt[i:]):.4f}"'
            f' data-base-prob="{sum(r[2] for r in gmt[i:]):.4f}"'
            f' data-base-rate="{sum(r[3] for r in gmt[i:]):.4f}">'
            f'{sum(r[0] for r in gmt[i:]) * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
            for i in range(len(gmt))
        )
        grp_html += (
            f'<tr class="proj-table-total"><td>Total</td>{grp_total_cells}'
            f'<td data-base="{gg_base:.4f}" data-base-fte="{gg_fte:.4f}" data-base-prob="{gg_prob:.4f}" data-base-rate="{gg_rate:.4f}">'
            f'{gg_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td></tr>'
            f'<tr class="cum-row"><td class="cum-label">REMAINING</td>{grp_cum_cells}'
            f'<td class="fte-cell-cum" data-base="{gg_base:.4f}" data-base-fte="{gg_fte:.4f}" data-base-prob="{gg_prob:.4f}" data-base-rate="{gg_rate:.4f}">'
            f'{gg_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td></tr>'
        )
        grp_html += "</tbody></table>"

    # ── Employee table ────────────────────────────────────────────
    emp_th = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    emp_html = f"""<table id="charge-emp-table" class="charge-table filterable-table">
    <thead><tr>
      <th class="th-emp">Employee</th>{emp_th}<th class="th-month">Total</th>
    </tr></thead><tbody>"""

    for emp in EMPLOYEES:
        if emp not in emp_charge_base.index:
            continue
        cells = "".join(
            f'<td class="td-num charge-cell" {_cell_attrs(emp_charge_base, emp_charge_fte, emp_charge_prob, emp_charge_rate, emp, m)}>'
            f'{float(emp_charge_base.loc[emp, m]) * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
            for m in MONTH_COLS
        )
        t_attrs, row_base = _total_attrs(emp_charge_base, emp_charge_fte, emp_charge_prob, emp_charge_rate, emp)
        emp_html += (
            f'<tr>'
            f'<td class="td-emp">{emp}</td>'
            f'{cells}'
            f'<td class="td-total charge-total-cell" {t_attrs}>'
            f'{row_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
            f'</tr>'
        )

    emt = _month_totals_attrs(emp_charge_base, emp_charge_fte, emp_charge_prob, emp_charge_rate)
    eg_base = sum(r[0] for r in emt)
    eg_fte  = sum(r[1] for r in emt)
    eg_prob = sum(r[2] for r in emt)
    eg_rate = sum(r[3] for r in emt)
    emp_total_cells = "".join(
        f'<td data-base="{r[0]:.4f}" data-base-fte="{r[1]:.4f}" data-base-prob="{r[2]:.4f}" data-base-rate="{r[3]:.4f}">'
        f'{r[0] * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
        for r in emt
    )
    emp_cum_cells = "".join(
        f'<td class="fte-cell-cum charge-cell"'
        f' data-base="{sum(r[0] for r in emt[i:]):.4f}"'
        f' data-base-fte="{sum(r[1] for r in emt[i:]):.4f}"'
        f' data-base-prob="{sum(r[2] for r in emt[i:]):.4f}"'
        f' data-base-rate="{sum(r[3] for r in emt[i:]):.4f}">'
        f'{sum(r[0] for r in emt[i:]) * _DEFAULT_MONTHLY_HOURS:,.0f}</td>'
        for i in range(len(emt))
    )
    emp_html += (
        f'<tr class="proj-table-total"><td>Total</td>{emp_total_cells}'
        f'<td data-base="{eg_base:.4f}" data-base-fte="{eg_fte:.4f}" data-base-prob="{eg_prob:.4f}" data-base-rate="{eg_rate:.4f}">'
        f'{eg_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td></tr>'
        f'<tr class="cum-row"><td class="cum-label">REMAINING</td>{emp_cum_cells}'
        f'<td class="fte-cell-cum" data-base="{eg_base:.4f}" data-base-fte="{eg_fte:.4f}" data-base-prob="{eg_prob:.4f}" data-base-rate="{eg_rate:.4f}">'
        f'{eg_base * _DEFAULT_MONTHLY_HOURS:,.0f}</td></tr>'
    )
    emp_html += "</tbody></table>"

    return proj_html, grp_html, emp_html


def legend_html() -> str:
    items = [
        ("fte-zero", "0.00",      "No allocation"),
        ("fte-low",  "< 0.50",    "Under-utilised"),
        ("fte-mid",  "0.50–<1.0", "Partial"),
        ("fte-full", "= 1.0",     "Full-time"),
        ("fte-over", "> 1.0",     "Overallocated"),
    ]
    html = ""
    for cls, rng, label in items:
        html += (
            f'<span class="legend-item">'
            f'<span class="legend-swatch {cls}"></span>'
            f'{rng} — {label}'
            f'</span>'
        )
    return html

# ── Chart data (Project FTE Demand) ─────────────────────────────────────────
PROJECT_COLORS = [
    {"border": "#3b5998", "bg": "rgba(59,89,152,0.15)"},
    {"border": "#059669", "bg": "rgba(5,150,105,0.15)"},
    {"border": "#d97706", "bg": "rgba(217,119,6,0.15)"},
    {"border": "#7c3aed", "bg": "rgba(124,58,237,0.15)"},
    {"border": "#dc2626", "bg": "rgba(220,38,38,0.15)"},
]

chart_datasets = []
for i, proj in enumerate(PROJECTS):
    if proj not in proj_prob_month.index:
        continue
    c = PROJECT_COLORS[i % len(PROJECT_COLORS)]
    chart_datasets.append({
        "label": proj,
        "data": [round(float(v), 4) for v in proj_prob_month.loc[proj, MONTH_COLS]],
        "borderColor": c["border"],
        "backgroundColor": c["bg"],
        "borderWidth": 2,
        "pointRadius": 4,
        "pointHoverRadius": 6,
        "fill": True,
        "tension": 0.35,
    })

chart_data_json = json.dumps({
    "labels": MONTH_COLS,
    "datasets": chart_datasets,
})

def project_probability_rows() -> str:
    """Project × Month table of probability-weighted FTE — mirrors project_summary_rows style."""
    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    html = f"""
    <thead>
      <tr>
        <th class="th-emp">Project</th>
        {th_months}
        <th class="th-month">TOTAL</th>
      </tr>
    </thead><tbody>"""

    color_map = {
        "Project Alpha": "#3b5998",
        "Project Beta":  "#059669",
        "Project Gamma": "#d97706",
    }

    for proj in PROJECTS:
        if proj not in proj_prob_month.index:
            continue
        group  = _proj_group.get(proj, "")
        avg_p  = proj_avg_prob.get(proj, 1.0)
        vals   = proj_prob_month.loc[proj, MONTH_COLS]
        total  = float(vals.sum())
        color  = color_map.get(proj, "#5f6978")
        prob_cls = "prob-high" if avg_p >= 0.85 else ("prob-mid" if avg_p >= 0.60 else "prob-low")
        cells = "".join(f'<td class="td-num">{fmt(v)}</td>' for v in vals)
        html += (
            f'<tr>'
            f'<td class="td-proj">'
            f'<span class="proj-dot" style="background:{color}"></span>'
            f'{proj}'
            f'<span class="proj-meta">{group}&ensp;<span class="prob-badge {prob_cls}">{avg_p:.0%}</span></span>'
            f'</td>'
            f'{cells}'
            f'<td class="td-total" style="color:{color}">{fmt(total)}</td>'
            f'</tr>'
        )

    # Totals row
    month_totals = proj_prob_month[MONTH_COLS].sum()
    grand_total  = float(month_totals.sum())
    total_cells  = "".join(f'<td class="fte-cell-total">{fmt(v)}</td>' for v in month_totals)
    html += (
        f'<tr class="heatmap-total-row">'
        f'<td class="td-emp total-label">TOTAL</td>'
        f'{total_cells}'
        f'<td class="fte-cell-total">{fmt(grand_total)}</td>'
        f'</tr>'
    )
    html += _backward_cum_row(month_totals, '<td class="td-emp cum-label">REMAINING</td>', f'<td class="fte-cell-cum">{fmt(grand_total)}</td>')
    html += "</tbody>"
    return html


# ── Probability chart data (stacked area by confidence band) ─────────────────
# Industry-standard bands:
#   Backlog     ≥ 90 %  — contracted / committed work
#   Outlook    50–90 %  — probable, in active pursuit
#   Opportunity < 50 %  — early-stage pipeline
_band_high = raw[raw["Probability"] >= 0.90]
_band_mid  = raw[(raw["Probability"] >= 0.50) & (raw["Probability"] < 0.90)]
_band_low  = raw[raw["Probability"] < 0.50]

def _band_monthly(subset) -> list:
    return [
        round(float((subset[m] * subset["Probability"]).sum()), 4)
        for m in MONTH_COLS
    ]

_prob_datasets = []
if not _band_high.empty:
    _prob_datasets.append({
        "label": "Backlog  (≥90 %)",
        "data": _band_monthly(_band_high),
        "backgroundColor": "#34d399",  # mid green
        "borderColor": "#10b981",
        "borderWidth": 1,
        "fill": True,
        "tension": 0.35,
        "pointRadius": 3,
    })
if not _band_mid.empty:
    _prob_datasets.append({
        "label": "Outlook  (50–90 %)",
        "data": _band_monthly(_band_mid),
        "backgroundColor": "#1d4ed8",  # dark blue
        "borderColor": "#1e3a8a",
        "borderWidth": 1,
        "fill": True,
        "tension": 0.35,
        "pointRadius": 3,
    })
if not _band_low.empty:
    _prob_datasets.append({
        "label": "Opportunity  (<50 %)",
        "data": _band_monthly(_band_low),
        "backgroundColor": "#60a5fa",  # mid blue
        "borderColor": "#3b82f6",
        "borderWidth": 1,
        "fill": True,
        "tension": 0.35,
        "pointRadius": 3,
    })

prob_chart_data_json = json.dumps({
    "labels": MONTH_COLS,
    "datasets": _prob_datasets,
})

# Always-present band rows (even if a band has zero FTE)
_BAND_DEFS = [
    ("Backlog",      "≥90 %",   "#34d399", _band_monthly(_band_high) if not _band_high.empty else [0.0]*len(MONTH_COLS)),
    ("Outlook",      "50–90 %", "#1d4ed8", _band_monthly(_band_mid)  if not _band_mid.empty  else [0.0]*len(MONTH_COLS)),
    ("Opportunity",  "<50 %",   "#60a5fa", _band_monthly(_band_low)  if not _band_low.empty  else [0.0]*len(MONTH_COLS)),
]

def prob_bands_table_rows() -> str:
    """Backlog / Outlook / Opportunity table — always shows all three rows."""
    th_months = "".join(f'<th class="th-month">{m}</th>' for m in MONTH_COLS)
    html = f"""
    <thead>
      <tr>
        <th class="th-emp">Pipeline Stage</th>
        {th_months}
        <th class="th-month">Total</th>
      </tr>
    </thead><tbody>"""
    for label, band_range, color, vals in _BAND_DEFS:
        row_total = round(sum(vals), 2)
        cells = "".join(f'<td class="td-num">{fmt(v)}</td>' for v in vals)
        html += (
            f'<tr>'
            f'<td class="td-proj">'
            f'<span class="proj-dot" style="background:{color}"></span>'
            f'<strong>{label}</strong> <span style="color:#8fa3bc;font-size:.8rem">{band_range}</span>'
            f'</td>'
            f'{cells}'
            f'<td class="td-total" style="color:{color}">{fmt(row_total)}</td>'
            f'</tr>'
        )
    # Totals row
    all_vals = [sum(row[3][i] for row in _BAND_DEFS) for i in range(len(MONTH_COLS))]
    grand = round(sum(all_vals), 2)
    total_cells = "".join(f'<td>{fmt(v)}</td>' for v in all_vals)
    html += (
        f'<tr class="proj-table-total"><td>Total</td>{total_cells}<td>{fmt(grand)}</td></tr>'
    )
    html += _backward_cum_row(all_vals, '<td class="cum-label">REMAINING</td>', f'<td class="fte-cell-cum">{fmt(grand)}</td>')
    html += "</tbody>"
    return html

# ── CSV data (raw) ────────────────────────────────────────────────────────────
csv_rows = [["Employee", "Project", "Project Group", "Probability"] + MONTH_COLS + ["Total"]]
for _, row in raw.sort_values(["Employee", "Project"]).iterrows():
    vals  = [row[m] for m in MONTH_COLS]
    group = row.get("Project Group", "")
    prob  = float(row.get("Probability", 1.0))
    csv_rows.append(
        [row["Employee"], row["Project"], group, f"{prob:.2f}"]
        + [f"{v:.2f}" for v in vals]
        + [f"{sum(vals):.2f}"]
    )
csv_json = json.dumps(csv_rows)

# ── Sidebar nav items ────────────────────────────────────────────────────────
_PROJ_ICON = ('<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
              'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
              '<path d="M2 3h6a4 4 0 014 4v14a3 3 0 00-3-3H2z"/>'
              '<path d="M22 3h-6a4 4 0 00-4 4v14a3 3 0 013-3h7z"/></svg>')
_DATA_ICON = ('<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
              'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
              '<rect x="3" y="3" width="18" height="18" rx="2"/>'
              '<path d="M3 9h18M3 15h18M9 3v18"/></svg>')
_PROB_ICON = ('<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
              'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
              '<path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>')

def _proj_letter(name: str) -> str:
    """First two capital letters of the project's distinguishing word."""
    skip = {"project", "programme", "program"}
    for part in name.strip().split():
        if part.lower() not in skip:
            return part[:2].upper()
    return name[:2].upper()

nav_proj_items = "\n    ".join(
    f'<button class="nav-item" data-tip="{p}" aria-label="{p}" '
    f"onclick=\"document.getElementById('sec-proj-{_slug(p)}').scrollIntoView({{behavior:'smooth'}});\">"
    f'<span class="nav-proj-letter">{_proj_letter(p)}</span></button>'
    for p in PROJECTS
)

# Group heatmap nav — only groups with >1 project (same condition as the section)
_nav_group_groups = [g for g in PROJECT_GROUPS if len([p for p in PROJECTS if _proj_group.get(p) == g]) > 1]
nav_group_items = "\n    ".join(
    f'<button class="nav-item" data-tip="{g}" aria-label="{g}" '
    f"onclick=\"document.getElementById('sec-group-{_slug(g)}').scrollIntoView({{behavior:'smooth'}});\">"
    f'<span class="nav-proj-letter" style="font-size:.72rem">{_proj_letter(g)[:2]}</span></button>'
    for g in (_nav_group_groups if len(PROJECT_GROUPS) > 1 else [])
)

nav_data_item = (
    '<button class="nav-item" data-tip="Data Table" aria-label="Data Table" '
    "onclick=\"document.getElementById('sec-data').scrollIntoView({behavior:'smooth'});\">"
    + _DATA_ICON + '</button>'
)
nav_prob_item = (
    '<button class="nav-item" data-tip="Probability Forecast" aria-label="Probability Forecast" '
    "onclick=\"document.getElementById('sec-prob').scrollIntoView({behavior:'smooth'});\">"
    + _PROB_ICON + '</button>'
)
_GROUP_FTE_ICON = (
    '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M3 3h18v4H3z"/><path d="M3 10h18v4H3z"/><path d="M3 17h18v4H3z"/></svg>'
)
nav_group_fte_item = (
    '<button class="nav-item" data-tip="Group FTE Demand" aria-label="Group FTE Demand" '
    "onclick=\"document.getElementById('sec-group-fte').scrollIntoView({behavior:'smooth'});\">"
    + _GROUP_FTE_ICON + '</button>'
)
_HEATMAP_ICON = (
    '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<rect x="3" y="3" width="7" height="7" rx="1"/>'
    '<rect x="14" y="3" width="7" height="7" rx="1"/>'
    '<rect x="3" y="14" width="7" height="7" rx="1"/>'
    '<rect x="14" y="14" width="7" height="7" rx="1"/></svg>'
)
nav_heatmap_item = (
    '<button class="nav-item" data-tip="Employee Utilisation" aria-label="Employee Utilisation" '
    "onclick=\"document.getElementById('sec-heatmap').scrollIntoView({behavior:'smooth'});\">"
    + _HEATMAP_ICON + '</button>'
)
_PROJ_FTE_ICON = (
    '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>'
)
nav_proj_fte_item = (
    '<button class="nav-item" data-tip="Project FTE Demand" aria-label="Project FTE Demand" '
    "onclick=\"document.getElementById('sec-ftechart').scrollIntoView({behavior:'smooth'});\">"
    + _PROJ_FTE_ICON + '</button>'
)
_CHARGE_ICON = (
    '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<line x1="12" y1="1" x2="12" y2="23"/>'
    '<path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>'
)
nav_charge_item = (
    '<button class="nav-item" data-tip="Charge" aria-label="Charge" '
    "onclick=\"document.getElementById('sec-charge').scrollIntoView({behavior:'smooth'});\">"
    + _CHARGE_ICON + '</button>'
)

_emp_opts   = "".join(f'<option value="{e}">{e}</option>' for e in EMPLOYEES)
_proj_opts  = "".join(f'<option value="{p}">{p}</option>' for p in PROJECTS)
_group_opts = "".join(f'<option value="{g}">{g}</option>' for g in PROJECT_GROUPS)
filter_bar_html = (
    '<div class="filter-bar">'
    f'<select id="flt-emp"   class="flt-select"><option value="">All Employees</option>{_emp_opts}</select>'
    f'<select id="flt-proj"  class="flt-select"><option value="">All Projects</option>{_proj_opts}</select>'
    f'<select id="flt-group" class="flt-select"><option value="">All Groups</option>{_group_opts}</select>'
    '<input id="flt-search" class="flt-input" type="text" placeholder="Search\u2026" />'
    '<button class="flt-clear" id="flt-clear">Clear</button>'
    '</div>'
)

# ── Assemble HTML ─────────────────────────────────────────────────────────────
generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
_charge_proj_html, _charge_grp_html, _charge_emp_html = charge_tables_html()

HTML = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Tyche · Manloading Report</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
  <style>
    :root {{
      --bg:              #f4f5f7;
      --surface:         #ffffff;
      --surface-soft:    #f8f9fb;
      --text:            #1f2733;
      --muted:           #5f6978;
      --accent:          #3b5998;
      --accent-strong:   #2d4373;
      --danger:          #dc2626;
      --warn:            #d97706;
      --ok:              #059669;
      --border:          #d5dbe4;
      --shadow:          0 8px 22px rgba(28,36,53,0.09);
      --radius-lg:       12px;
      --radius-md:       12px;
      --radius-sm:       8px;
      --header-gradient: linear-gradient(135deg,#2c3e6b 0%,#3b5998 100%);
    }}

    body.dark {{
      --bg:              #0f1623;
      --surface:         #1a2438;
      --surface-soft:    #1e2c42;
      --text:            #e2e8f2;
      --muted:           #8fa3bc;
      --accent:          #5b7fd4;
      --danger:          #f87171;
      --warn:            #fbbf24;
      --ok:              #34d399;
      --border:          #2a3a52;
      --shadow:          0 8px 22px rgba(0,0,0,0.45);
      --header-gradient: linear-gradient(135deg,#1a2438 0%,#253558 100%);
    }}

    *   {{ box-sizing:border-box; }}
    body {{
      margin:0; padding-left:52px; min-height:100vh;
      font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
      color:var(--text); background:var(--bg);
    }}

    /* ── Sidebar ── */
    .sidebar {{
      position:fixed; left:0; top:0; bottom:0; width:52px;
      background:linear-gradient(180deg,#2c3e6b 0%,#1e2d50 100%);
      display:flex; flex-direction:column;
      justify-content:space-between; padding:8px 0 10px;
      z-index:100; box-shadow:2px 0 10px rgba(0,0,0,0.18);
    }}
    .sidebar-top {{
      display:flex; flex-direction:column; align-items:center; gap:3px;
      flex:1; overflow-y:auto; padding-bottom:4px;
    }}
    .sidebar-bottom {{
      display:flex; flex-direction:column; align-items:center; gap:3px;
      flex-shrink:0;
    }}
    .sidebar-logo {{
      padding:6px 0 10px; display:flex; align-items:center;
      justify-content:center; opacity:.85; color:#fff;
    }}
    .sidebar-logo:hover {{ opacity:1; background:transparent !important; }}
    .nav-item {{
      width:38px; height:38px; border:none; background:transparent;
      border-radius:10px; cursor:pointer;
      display:flex; align-items:center; justify-content:center;
      color:rgba(255,255,255,.55); transition:background .15s,color .15s;
      position:relative;
    }}
    .nav-item svg {{ pointer-events:none; }}
    .nav-item:hover,.nav-item.is-active {{
      background:rgba(255,255,255,.13); color:#ffffff;
    }}
    .nav-item::after {{
      content:attr(data-tip);
      position:absolute; left:calc(100% + 10px); top:50%;
      transform:translateY(-50%);
      background:#1a2438; color:#e2e8f2;
      font-size:.73rem; white-space:nowrap;
      padding:4px 9px; border-radius:6px;
      pointer-events:none; z-index:200;
      box-shadow:0 2px 8px rgba(0,0,0,0.28);
      opacity:0; transition:opacity .15s;
    }}
    .nav-item:hover::after {{ opacity:1; }}
    .nav-item[data-tip=""]:hover::after,
    .nav-item:not([data-tip]):hover::after {{ opacity:0; pointer-events:none; }}

    /* ── Layout ── */
    .layout {{
      max-width:calc(100vw - 72px); margin:0 auto;
      padding:20px 20px 48px; display:grid; gap:18px;
    }}
    .layout > * {{ min-width:0; max-width:100%; }}

    /* ── Hero ── */
    .hero {{
      background:var(--header-gradient); color:#ecf2ff;
      border-radius:var(--radius-lg); padding:15px 28px;
      box-shadow:var(--shadow);
      display:flex; justify-content:space-between; align-items:center; gap:16px;
    }}
    .hero-brand {{ display:inline-flex; align-items:center; gap:10px; }}
    .hero h1 {{
      margin:0; font-size:clamp(1.28rem,1.8vw,1.7rem);
      font-weight:650; letter-spacing:.01em;
    }}
    .hero-meta {{
      display:inline-flex; flex-direction:column;
      align-items:flex-end; gap:0; line-height:1.25;
      color:#fff; font-size:.85rem;
    }}

    /* ── Surface card ── */
    .surface {{
      background:var(--surface); border:1px solid var(--border);
      border-radius:var(--radius-lg); box-shadow:var(--shadow);
    }}
    .card-header {{
      padding:14px 18px 0; display:flex;
      align-items:center; justify-content:space-between; gap:10px;
    }}
    .section-title {{
      margin:0; font-size:1rem; font-weight:700; color:#26364f;
    }}
    body.dark .section-title {{ color:var(--text); }}
    .section-sub {{
      font-size:.78rem; color:var(--muted); margin:2px 0 0;
    }}

    /* ── Stat cards ── */
    .stats {{
      display:grid; grid-template-columns:repeat(auto-fit,minmax(155px,1fr));
      gap:10px; padding:18px;
    }}
    .stat-card {{
      background:var(--surface-soft); border:1px solid var(--border);
      border-radius:var(--radius-md); padding:12px; min-height:90px;
      display:flex; flex-direction:column; position:relative;
    }}
    .stat-card-header {{ display:contents; }}
    .stat-card h3 {{
      margin:0; font-size:.68rem; text-transform:uppercase;
      letter-spacing:.06em; color:var(--muted); font-weight:700;
      padding-right:26px; word-break:break-word;
    }}
    .stat-card .tip {{
      position:absolute; top:10px; right:10px;
    }}
    .stat-value {{ margin-top:8px; font-size:1.8rem; font-weight:700; }}

    /* ── Help tip badge (same as Caerus) ── */
    .tip {{
      display:inline-flex; align-items:center; justify-content:center;
      position:relative; width:16px; height:16px; border-radius:999px;
      border:1px solid #8fb3c7; background:#e8f4fb; color:#0f4c6a;
      font-size:.65rem; line-height:1; font-weight:700; cursor:help;
      text-transform:none; letter-spacing:0; flex-shrink:0;
    }}
    .tip:hover::after, .tip:focus-visible::after {{
      content:attr(data-tip);
      position:absolute; left:50%; transform:translateX(-50%);
      bottom:calc(100% + 9px); min-width:180px; max-width:260px;
      background:#0f172a; color:#f8fafc; border-radius:8px;
      padding:7px 9px; font-size:.73rem; font-weight:500;
      line-height:1.35; text-transform:none; letter-spacing:0;
      white-space:normal; text-align:left; z-index:300;
      box-shadow:0 10px 24px rgba(15,23,42,0.34); pointer-events:none;
    }}
    .tip:hover::before, .tip:focus-visible::before {{
      content:""; position:absolute; left:50%; transform:translateX(-50%);
      bottom:calc(100% + 2px);
      border-left:6px solid transparent; border-right:6px solid transparent;
      border-top:7px solid #0f172a; z-index:300; pointer-events:none;
    }}

    /* ── Table shared ── */
    .table-wrap {{ padding:14px 18px 18px; overflow:auto; min-width:0; }}
    .table-scroll {{
      overflow:auto; border:1px solid var(--border);
      border-radius:var(--radius-md); background:var(--surface); max-width:100%;
    }}
    table {{ width:100%; border-collapse:collapse; }}
    th,td {{ padding:9px 8px; text-align:center; border-bottom:1px solid var(--border); font-size:.82rem; }}
    th {{
      position:sticky; top:0; background:#edf1f8; color:#3f4d66;
      font-size:.72rem; text-transform:uppercase;
      letter-spacing:.08em; z-index:2; white-space:nowrap;
    }}
    body.dark th {{ background:#1a2a42; color:var(--text); }}

    .heatmap-total-row:hover .fte-cell-total,
    .heatmap-total-row:hover .total-label {{
      background:#2c3e6b !important;
    }}
    body.dark .heatmap-total-row:hover .fte-cell-total,
    body.dark .heatmap-total-row:hover .total-label {{
      background:#1a2a42 !important;
    }}
    .th-emp {{ text-align:left !important; min-width:130px; }}
    .td-emp {{ text-align:left; font-weight:600; white-space:nowrap; }}
    .td-proj {{ text-align:left; white-space:nowrap; }}
    .td-num {{ font-variant-numeric:tabular-nums; min-width:52px; }}
    .td-total {{ min-width:80px; font-weight:700; font-variant-numeric:tabular-nums; }}
    .th-month {{ min-width:58px; }}
    .tr-shade td {{ background:var(--surface-soft); }}

    /* ── Heatmap table overrides ── */
    .heatmap-table {{ border-collapse:separate; border-spacing:3px; }}
    .heatmap-table th,
    .heatmap-table td {{ border:none; border-radius:6px; }}
    .heatmap-table th {{ background:#edf1f8; border-radius:6px; }}
    body.dark .heatmap-table th {{ background:#1a2a42; }}
    .heatmap-table .td-emp {{
      background:var(--surface-soft); border-radius:6px;
      border:1px solid var(--border);
    }}

    /* ── Heatmap cells ── */
    .fte-cell {{
      font-weight:600; font-variant-numeric:tabular-nums;
      position:relative;
    }}
    .fte-over-marker {{
      position:absolute; top:2px; right:2px;
      width:9px; height:9px; pointer-events:none;
      color:#000; display:block; line-height:1;
    }}
    body.dark .fte-over-marker {{ color:#fff; }}
    /* Project heatmap cells - same colours, no hover effect */
    .fte-cell-proj {{
      font-weight:600; font-variant-numeric:tabular-nums;
    }}
    .fte-zero {{ background:#f1f3f5; color:#aab2be; }}
    .fte-low  {{ background:#fef9c3; color:#854d0e; }}
    .fte-mid  {{ background:#bfdbfe; color:#1e3a8a; }}
    .fte-full {{ background:#bbf7d0; color:#14532d; }}
    .fte-over {{ background:#fca5a5; color:#7f1d1d; }}
    .avg-cell {{ font-size:.85rem; }}
    body.dark .fte-zero {{ background:#253347; color:#6a7f96; }}
    body.dark .fte-low  {{ background:#3a2c07; color:#fcd34d; }}
    body.dark .fte-mid  {{ background:#132454; color:#7dd3fc; }}
    body.dark .fte-full {{ background:#0b3321; color:#4ade80; }}
    body.dark .fte-over {{ background:#3d0c0c; color:#f87171; }}
    .fte-cell-proj.fte-zero {{ background:#f1f3f5; color:#aab2be; }}
    .fte-cell-proj.fte-low  {{ background:#fef9c3; color:#854d0e; }}
    .fte-cell-proj.fte-mid  {{ background:#bfdbfe; color:#1e3a8a; }}
    .fte-cell-proj.fte-full {{ background:#bbf7d0; color:#14532d; }}
    .fte-cell-proj.fte-over {{ background:#fca5a5; color:#7f1d1d; }}
    body.dark .fte-cell-proj.fte-zero {{ background:#253347; color:#6a7f96; }}
    body.dark .fte-cell-proj.fte-low  {{ background:#3a2c07; color:#fcd34d; }}
    body.dark .fte-cell-proj.fte-mid  {{ background:#132454; color:#7dd3fc; }}
    body.dark .fte-cell-proj.fte-full {{ background:#0b3321; color:#4ade80; }}
    body.dark .fte-cell-proj.fte-over {{ background:#3d0c0c; color:#f87171; }}

    /* ── Heatmap total row ── */
    .heatmap-total-row .td-emp.total-label {{
      background:#2c3e6b; color:#ecf2ff; font-weight:700;
      border-radius:6px; border:none;
    }}
    .fte-cell-total {{
      background:#2c3e6b; color:#ecf2ff;
      font-weight:700; font-variant-numeric:tabular-nums;
      border-radius:6px;
    }}
    body.dark .heatmap-total-row .td-emp.total-label,
    body.dark .fte-cell-total {{
      background:#1a2a42; color:#bfdbfe;
    }}

    /* ── Backward-cumulative (remaining) row ── */
    .cum-row .td-emp.cum-label, .cum-row .cum-label {{
      background:#16355f; color:#93c5fd; font-weight:600; font-size:.72rem;
      letter-spacing:.04em; border-radius:6px; border:none;
    }}
    .fte-cell-cum {{
      background:#16355f; color:#93c5fd;
      font-weight:600; font-variant-numeric:tabular-nums;
      border-radius:6px; font-size:.82rem;
    }}
    body.dark .cum-row .td-emp.cum-label, body.dark .cum-row .cum-label,
    body.dark .fte-cell-cum {{
      background:#0a1e35; color:#7dd3fc;
    }}

    /* ── Legend ── */
    .legend {{
      display:flex; flex-wrap:wrap; gap:10px;
      padding:0 18px 14px; font-size:.75rem; color:var(--muted);
    }}
    .legend-item {{ display:flex; align-items:center; gap:5px; }}
    .legend-swatch {{
      display:inline-block; width:12px; height:12px;
      border-radius:3px; flex-shrink:0;
    }}

    /* ── Chart ── */
    .chart-wrap {{
      padding:4px 18px 18px;
    }}
    .chart-canvas-wrap {{
      position:relative; height:260px;
    }}

    /* ── Download button ── */
    .btn-dl {{
      display:inline-flex; align-items:center; gap:6px;
      padding:7px 14px; border-radius:8px; border:1px solid var(--border);
      background:var(--surface-soft); color:var(--muted);
      font:inherit; font-size:.78rem; font-weight:700;
      cursor:pointer; text-decoration:none;
      transition:background .15s,color .15s,border-color .15s;
    }}
    .btn-dl:hover {{
      background:var(--accent); color:#fff; border-color:var(--accent);
    }}

    /* ── Project dot ── */
    .proj-dot {{
      display:inline-block; width:8px; height:8px;
      border-radius:50%; flex-shrink:0; margin-right:6px;
      vertical-align:middle;
    }}
    /* ── Logo ── */
    .logo-triangle {{ width:22px; height:22px; flex-shrink:0; }}

    /* ── Floating cell tooltip ── */
    #cell-tip {{
      position:fixed; z-index:500;
      background:#1a2438; color:#e2e8f2;
      padding:7px 11px; border-radius:8px;
      font-size:.76rem; line-height:1.6;
      pointer-events:none; opacity:0;
      transition:opacity .12s;
      box-shadow:0 4px 14px rgba(0,0,0,.35);
      max-width:260px; white-space:pre-line;
    }}

    /* ── Filter bar ── */
    .filter-bar {{ display:flex; flex-wrap:wrap; gap:8px; padding:12px 18px 0; align-items:center; }}
    .flt-select, .flt-input {{
      height:32px; border-radius:7px; border:1px solid var(--border);
      background:var(--surface-soft); color:var(--text);
      padding:0 10px; font-size:.82rem; outline:none;
    }}
    .flt-select:focus, .flt-input:focus {{ border-color:var(--accent); }}
    .flt-input {{ min-width:160px; }}
    .flt-clear {{
      height:32px; border-radius:7px; border:1px solid var(--border);
      background:transparent; color:var(--muted); cursor:pointer;
      padding:0 12px; font-size:.82rem; transition:background .15s;
    }}
    .flt-clear:hover {{ background:var(--border); color:var(--text); }}

    /* ── Row-click filtering ── */
    .filterable-table tbody tr:not(.proj-table-total):not(.cum-row):not(.heatmap-total-row) {{
      cursor: pointer;
    }}
    .filterable-table tbody tr:not(.proj-table-total):not(.cum-row):not(.heatmap-total-row):hover {{
      filter: brightness(0.95);
    }}
    .row-excluded {{
      opacity: 0.28;
      text-decoration: line-through;
      text-decoration-color: var(--muted);
    }}

    /* ── Charge section ── */
    .charge-controls {{
      display:flex; align-items:center; gap:8px;
    }}
    .charge-hours-label {{
      font-size:.78rem; color:var(--muted); white-space:nowrap;
    }}
    .charge-hours-input {{
      width:80px; height:32px; border-radius:7px;
      border:1px solid var(--border); background:var(--surface-soft);
      color:var(--text); padding:0 10px; font-size:.82rem;
      text-align:right; outline:none;
    }}
    .charge-hours-input:focus {{ border-color:var(--accent); }}
    body.dark .charge-hours-input {{ background:#1a2438; }}
    .charge-section-sub {{
      font-size:.72rem; text-transform:uppercase; letter-spacing:.08em;
      color:var(--muted); padding:14px 18px 2px; font-weight:600;
    }}

    /* ── Charge factor toggle chips ── */
    .charge-equation {{
      display:flex; align-items:center; flex-wrap:wrap; gap:5px;
      margin:4px 0 0; font-size:.82rem; color:var(--muted);
    }}
    .charge-op {{ opacity:.5; }}
    .charge-factor {{
      display:inline-flex; align-items:center; gap:5px;
      padding:3px 10px; border-radius:999px;
      border:1.5px solid var(--accent); background:var(--accent);
      color:#fff; font:inherit; font-size:.78rem; font-weight:600;
      cursor:pointer; transition:background .15s,color .15s,border-color .15s,opacity .15s;
      user-select:none;
    }}
    .charge-factor[aria-pressed="false"] {{
      background:transparent; color:var(--muted);
      border-color:var(--border); text-decoration:line-through;
      opacity:.55;
    }}
    .charge-factor:hover {{ filter:brightness(1.08); }}

    /* ── Charge meaning label ── */
    .charge-meaning {{
      color:var(--accent); font-weight:600; font-size:.82rem;
    }}
    #sec-charge .proj-table-total td,
    #sec-charge .cum-row td {{
      /* inherit the refined style already declared below */
    }}

    /* ── Sortable column headers ── */
    thead th.sortable {{ cursor:pointer; user-select:none; }}
    thead th.sortable::after {{ content:''; margin-left:4px; opacity:.3; font-size:.75em; }}
    thead th.sortable[data-sorted="asc"]::after  {{ content:'▲'; opacity:1; }}
    thead th.sortable[data-sorted="desc"]::after {{ content:'▼'; opacity:1; }}

    /* ── Total row in regular (non-heatmap) tables ── */
    /* fte-cell-total has border-radius for heatmap; strip it in collapse tables */
    table:not(.heatmap-table) .fte-cell-total {{
      border-radius: 0;
    }}
    /* proj-table-total: same dark tone as heatmap total rows, no border-radius */
    .proj-table-total td {{
      background: #2c3e6b; color: #ecf2ff;
      font-weight: 700; font-variant-numeric: tabular-nums;
      font-size: .78rem; text-transform: uppercase;
      letter-spacing: .06em; border-radius: 0;
    }}
    body.dark .proj-table-total td {{
      background: #1a2a42; color: #bfdbfe;
    }}

    /* ── Probability Forecast + Group FTE: refined total / remaining rows ── */
    #sec-prob .proj-table-total td,
    #sec-group-fte .proj-table-total td,
    #sec-ftechart .proj-table-total td,
    #sec-data .proj-table-total td,
    #sec-charge .proj-table-total td {{
      background: #eef2fa; color: #1e3a6e;
      border-top: 2px solid #3b5998; border-radius: 0;
      font-weight: 800; font-size: .80rem;
      letter-spacing: .05em; text-transform: uppercase;
    }}
    body.dark #sec-prob .proj-table-total td,
    body.dark #sec-group-fte .proj-table-total td,
    body.dark #sec-ftechart .proj-table-total td,
    body.dark #sec-data .proj-table-total td,
    body.dark #sec-charge .proj-table-total td {{
      background: #1a2d52; color: #93c5fd;
      border-top-color: #4a7ec4;
    }}
    #sec-prob .cum-row td,
    #sec-group-fte .cum-row td,
    #sec-ftechart .cum-row td,
    #sec-data .cum-row td,
    #sec-charge .cum-row td {{
      background: #f4f7ff; color: #4a6daa;
      border-top: 1px dashed #9ab3d8; border-radius: 0;
      font-weight: 500; font-size: .77rem;
      font-variant-numeric: tabular-nums; letter-spacing: .03em;
    }}
    #sec-prob .cum-row .cum-label,
    #sec-group-fte .cum-row .cum-label,
    #sec-ftechart .cum-row .cum-label,
    #sec-data .cum-row .cum-label,
    #sec-charge .cum-row .cum-label {{
      font-style: italic; font-weight: 600; color: #3a5d9c; background: #eaf0fb;
      border-radius: 0;
    }}
    body.dark #sec-prob .cum-row td,
    body.dark #sec-group-fte .cum-row td,
    body.dark #sec-ftechart .cum-row td,
    body.dark #sec-data .cum-row td,
    body.dark #sec-charge .cum-row td {{
      background: #0f1d36; color: #60a5fa;
      border-top-color: #2a4d7a;
    }}
    body.dark #sec-prob .cum-row .cum-label,
    body.dark #sec-group-fte .cum-row .cum-label,
    body.dark #sec-ftechart .cum-row .cum-label,
    body.dark #sec-data .cum-row .cum-label,
    body.dark #sec-charge .cum-row .cum-label {{
      color: #7dd3fc; background: #0c1a30;
    }}

    /* ── Nav project letter badge ── */
    .nav-proj-letter {{
      font-size:.9rem; font-weight:800; line-height:1;
      color:inherit; letter-spacing:-.01em;
    }}

    /* ── Project meta (group + prob badge inline under project name) ── */
    .proj-meta {{
      display:block; font-size:.72rem; font-weight:400;
      color:var(--muted); margin-top:2px; line-height:1;
    }}

    /* ── Group column ── */
    .td-group {{ text-align:left; font-size:.8rem; color:var(--muted); white-space:nowrap; }}

    /* ── Probability badge ── */
    .td-prob {{ text-align:center; }}
    .prob-badge {{
      display:inline-block; padding:2px 8px; border-radius:999px;
      font-size:.74rem; font-weight:700; letter-spacing:.03em;
    }}
    .prob-high {{ background:#dcfce7; color:#166534; }}
    .prob-mid  {{ background:#fef9c3; color:#854d0e; }}
    .prob-low  {{ background:#fee2e2; color:#991b1b; }}
    body.dark .prob-high {{ background:#14532d; color:#86efac; }}
    body.dark .prob-mid  {{ background:#3a2c07; color:#fcd34d; }}
    body.dark .prob-low  {{ background:#450a0a; color:#fca5a5; }}
  </style>
</head>
<body>

<!-- Sidebar -->
<nav class="sidebar" aria-label="Navigation">
  <div class="sidebar-top">
    <button class="nav-item sidebar-logo" id="scrollTop" data-tip="Back to top" aria-label="Back to top"
            onclick="window.scrollTo({{top:0,behavior:'smooth'}})">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M16 16l3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1z"/>
        <path d="M2 16l3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1z"/>
        <path d="M7 21h10"/><line x1="12" y1="3" x2="12" y2="21"/>
        <path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>
      </svg>
    </button>
    <button class="nav-item" data-tip="Overview" aria-label="Overview"
            onclick="document.getElementById('sec-hero').scrollIntoView({{behavior:'smooth'}})">
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="12" width="4" height="9"/><rect x="10" y="7" width="4" height="14"/><rect x="17" y="3" width="4" height="18"/></svg>
    </button>
    {nav_heatmap_item}
    {nav_prob_item}
    {nav_group_fte_item}
    {nav_proj_fte_item}
    {nav_proj_items}
    {nav_data_item}
    {nav_charge_item}
  </div>
  <div class="sidebar-bottom">
    <button class="nav-item" id="themeToggle" data-tip="Night mode" aria-label="Toggle night mode">
      <svg class="icon-moon" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>
      <svg class="icon-sun" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:none"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>
    </button>
  </div>
</nav>

<div class="layout">

  <!-- Hero -->
  <header class="hero" id="sec-hero">
    <div class="hero-brand">
      <svg class="logo-triangle" viewBox="0 0 24 24" fill="none" stroke="#ecf2ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">
        <path d="M16 16l3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1z"/>
        <path d="M2 16l3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1z"/>
        <path d="M7 21h10"/><line x1="12" y1="3" x2="12" y2="21"/>
        <path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>
      </svg>
      <h1>Tyche · Manloading Report</h1>
    </div>
    <div class="hero-meta">
      <span class="stamp-value">v{VERSION} &middot; Generated on {generated_at}</span>
    </div>
  </header>

  <!-- KPI Cards -->
  <section class="surface">
    <div class="card-header">
      <div>
        <p class="section-title">Summary</p>
        <p class="section-sub">Key metrics across all employees and projects</p>
      </div>
    </div>
    <div class="stats">
      {stat_cards()}
    </div>
  </section>

  <!-- Utilisation Heatmap -->
  <section class="surface" id="sec-heatmap">
    <div class="card-header">
      <div>
        <p class="section-title" style="display:flex;align-items:center;gap:6px;">Employee Utilisation Heatmap
          <span class="tip" tabindex="0"
            data-tip="Each cell shows expected FTE load (raw FTE x project probability). Colours: grey = 0, yellow = under 50%, blue = 50–<100%, green = 100%, red = over 100%. The small clock icon marks cells where the raw unweighted FTE exceeds 1.05 — a potential overallocation before probability discounting." aria-label="Heatmap help">?</span>
        </p>
        <p class="section-sub">Probability-weighted FTE per employee per month (raw FTE × project probability)</p>
      </div>
    </div>
    <div class="table-wrap">
      <div class="table-scroll">
        {utilisation_heatmap()}
      </div>
    </div>
  </section>

  <!-- Probability Forecast -->
  <section class="surface" id="sec-prob">
    <div class="card-header">
      <div>
        <p class="section-title">Probability-Weighted Forecast</p>
        <p class="section-sub">Probability-weighted FTE stacked by pipeline stage — Backlog (≥90 %, green) · Outlook (50–90 %, dark blue) · Opportunity (&lt;50 %, light blue)</p>
      </div>
    </div>
    <div class="chart-wrap">
      <div class="chart-canvas-wrap" style="height:300px">
        <canvas id="probForecastChart"></canvas>
      </div>
    </div>
    <div class="table-wrap">
      <div class="table-scroll">
        <table class="filterable-table">
          {prob_bands_table_rows()}
        </table>
      </div>
    </div>
  </section>

  <!-- Project Group FTE Demand -->
  <section class="surface" id="sec-group-fte">
    <div class="card-header">
      <div>
        <p class="section-title">Project Group FTE Demand</p>
        <p class="section-sub">Probability-weighted FTE aggregated by project group per month</p>
      </div>
    </div>
    <div class="chart-wrap">
      <div class="chart-canvas-wrap">
        <canvas id="groupFteChart"></canvas>
      </div>
    </div>
    <div class="table-wrap">
      <div class="table-scroll">
        <table class="filterable-table">
          {group_fte_demand_rows()}
        </table>
      </div>
    </div>
  </section>

  <!-- Project FTE Summary -->
  <section class="surface" id="sec-ftechart">
    <div class="card-header">
      <div>
        <p class="section-title">Project FTE Demand</p>
        <p class="section-sub">Probability-weighted FTE per project per month (raw FTE × probability)</p>
      </div>
    </div>    <div class="chart-wrap">
      <div class="chart-canvas-wrap">
        <canvas id="projectFteChart"></canvas>
      </div>
    </div>    <div class="table-wrap">
      <div class="table-scroll">
        <table class="filterable-table">
          {project_summary_rows()}
        </table>
      </div>
    </div>
  </section>

  <!-- Per-Project Heatmaps -->
  {project_heatmaps()}

  <!-- Employee × Project Detail -->
  <section class="surface" id="sec-data">
    <div class="card-header">
      <div>
        <p class="section-title">Employee × Project Detail</p>
        <p class="section-sub">Raw allocation data — one row per employee / project combination</p>
      </div>
      <button class="btn-dl" id="btnCsv">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Download CSV
      </button>
    </div>
    {filter_bar_html}
    <div class="table-wrap">
      <div class="table-scroll">
        <table id="detail-table">
          {employee_detail_rows()}
        </table>
      </div>
    </div>
  </section>

  <!-- Charge -->
  <section class="surface" id="sec-charge">
    <div class="card-header">
      <div>
        <p class="section-title">Charge</p>
        <p class="section-sub charge-equation">
          FTE
          <span class="charge-op">&times;</span>
          <button class="charge-factor" id="cfac-prob" data-factor="prob" aria-pressed="true">Probability</button>
          <span class="charge-op">&times;</span>
          <button class="charge-factor" id="cfac-rate" data-factor="rate" aria-pressed="true">Hourly Rate</button>
          <span class="charge-op">&times;</span>
          <button class="charge-factor" id="cfac-hours" data-factor="hours" aria-pressed="true">Monthly Hours</button>
          <span class="charge-op">=</span>
          <span class="charge-meaning" id="charge-meaning">Expected cost</span>
        </p>
      </div>
      <div class="charge-controls">
        <label class="charge-hours-label" for="chargeHours">Monthly hours</label>
        <input id="chargeHours" class="charge-hours-input" type="number" min="1" step="1" value="{_DEFAULT_MONTHLY_HOURS}">
      </div>
    </div>
    <p class="charge-section-sub">By Project</p>
    <div class="table-wrap"><div class="table-scroll">{_charge_proj_html}</div></div>
    {'<p class="charge-section-sub">By Group</p><div class="table-wrap"><div class="table-scroll">' + _charge_grp_html + '</div></div>' if _charge_grp_html else ''}
    <p class="charge-section-sub">By Employee</p>
    <div class="table-wrap"><div class="table-scroll">{_charge_emp_html}</div></div>
  </section>

</div>

<script>
  // ─ Floating cell tooltip ──────────────────────────────────────────────────────────
  (function () {{
    const tip = document.createElement('div');
    tip.id = 'cell-tip';
    document.body.appendChild(tip);
    function move(e) {{
      const x = e.clientX + 14, y = e.clientY + 14;
      tip.style.left = (x + tip.offsetWidth  > window.innerWidth  ? e.clientX - tip.offsetWidth  - 6 : x) + 'px';
      tip.style.top  = (y + tip.offsetHeight > window.innerHeight ? e.clientY - tip.offsetHeight - 6 : y) + 'px';
    }}
    const FTE_COLORS = {{
      light: {{
        'fte-zero': {{bg:'#f1f3f5', color:'#aab2be'}},
        'fte-low':  {{bg:'#fef9c3', color:'#854d0e'}},
        'fte-mid':  {{bg:'#bfdbfe', color:'#1e3a8a'}},
        'fte-full': {{bg:'#bbf7d0', color:'#14532d'}},
        'fte-over': {{bg:'#fca5a5', color:'#7f1d1d'}},
      }},
      dark: {{
        'fte-zero': {{bg:'#253347', color:'#6a7f96'}},
        'fte-low':  {{bg:'#3a2c07', color:'#fcd34d'}},
        'fte-mid':  {{bg:'#132454', color:'#7dd3fc'}},
        'fte-full': {{bg:'#0b3321', color:'#4ade80'}},
        'fte-over': {{bg:'#3d0c0c', color:'#f87171'}},
      }},
    }};
    document.querySelectorAll('[data-cell-tip]').forEach(el => {{
      el.addEventListener('mouseenter', e => {{
        if (!el.dataset.cellTip) return;
        tip.textContent = el.dataset.cellTip;
        const palette = FTE_COLORS[document.body.classList.contains('dark') ? 'dark' : 'light'];
        const matched = Object.keys(palette).find(cls => el.classList.contains(cls));
        if (matched) {{
          tip.style.background = palette[matched].bg;
          tip.style.color = palette[matched].color;
          tip.style.boxShadow = '0 4px 14px rgba(0,0,0,.20)';
        }} else {{
          tip.style.background = '#1a2438';
          tip.style.color = '#e2e8f2';
          tip.style.boxShadow = '0 4px 14px rgba(0,0,0,.35)';
        }}
        tip.style.opacity = '1'; move(e);
      }});
      el.addEventListener('mousemove',  move);
      el.addEventListener('mouseleave', ()  => {{ tip.style.opacity = '0'; }});
    }});
  }})();

  // ─ Dark mode ──────────────────────────────────────────────────────────────
  const toggle = document.getElementById('themeToggle');
  const moon   = toggle.querySelector('.icon-moon');
  const sun    = toggle.querySelector('.icon-sun');

  function applyTheme(dark) {{
    document.body.classList.toggle('dark', dark);
    moon.style.display = dark ? 'none' : '';
    sun.style.display  = dark ? '' : 'none';
    toggle.dataset.tip = dark ? 'Light mode' : 'Night mode';
    localStorage.setItem('tyche-theme', dark ? 'dark' : 'light');
    if (window._projectFteChart) {{
      const textColor = dark ? '#8fa3bc' : '#3f4d66';
      const gridColor = dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
      window._projectFteChart.options.scales.x.ticks.color = textColor;
      window._projectFteChart.options.scales.y.ticks.color = textColor;
      window._projectFteChart.options.scales.x.grid.color  = gridColor;
      window._projectFteChart.options.scales.y.grid.color  = gridColor;
      window._projectFteChart.options.plugins.legend.labels.color = textColor;
      window._projectFteChart.update();
    }}
    if (window._probForecastChart) {{
      const textColor = dark ? '#8fa3bc' : '#3f4d66';
      const gridColor = dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
      window._probForecastChart.options.scales.x.ticks.color = textColor;
      window._probForecastChart.options.scales.y.ticks.color = textColor;
      window._probForecastChart.options.scales.x.grid.color  = gridColor;
      window._probForecastChart.options.scales.y.grid.color  = gridColor;
      window._probForecastChart.options.plugins.legend.labels.color = textColor;
      window._probForecastChart.update();
    }}
    if (window._groupFteChart) {{
      const textColor = dark ? '#8fa3bc' : '#3f4d66';
      const gridColor = dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
      window._groupFteChart.options.scales.x.ticks.color = textColor;
      window._groupFteChart.options.scales.y.ticks.color = textColor;
      window._groupFteChart.options.scales.x.grid.color  = gridColor;
      window._groupFteChart.options.scales.y.grid.color  = gridColor;
      window._groupFteChart.options.plugins.legend.labels.color = textColor;
      window._groupFteChart.update();
    }}
  }}

  applyTheme(localStorage.getItem('tyche-theme') === 'dark');
  toggle.addEventListener('click', () => applyTheme(!document.body.classList.contains('dark')));

  // ─ Project FTE Chart ────────────────────────────────────────────────
  (function () {{
    const dark      = document.body.classList.contains('dark');
    const textColor = dark ? '#8fa3bc' : '#3f4d66';
    const gridColor = dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
    const data      = {chart_data_json};

    window._projectFteChart = new Chart(
      document.getElementById('projectFteChart'),
      {{
        type: 'line',
        data: data,
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          plugins: {{
            legend: {{
              position: 'top',
              labels: {{ color: textColor, usePointStyle: true, pointStyleWidth: 10, padding: 16, font: {{ size: 12 }} }}
            }},
            tooltip: {{
              callbacks: {{
                label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toFixed(2)}} FTE`
              }}
            }}
          }},
          scales: {{
            x: {{
              ticks: {{ color: textColor, font: {{ size: 11 }} }},
              grid:  {{ color: gridColor }}
            }},
            y: {{
              beginAtZero: true,
              ticks: {{ color: textColor, font: {{ size: 11 }}, callback: v => v.toFixed(1) + ' FTE' }},
              grid:  {{ color: gridColor }}
            }}
          }}
        }}
      }}
    );
  }})();
  // ─ Probability Forecast Chart (stacked area by confidence band) ─────────
  (function () {{
    const dark      = document.body.classList.contains('dark');
    const textColor = dark ? '#8fa3bc' : '#3f4d66';
    const gridColor = dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
    const data      = {prob_chart_data_json};

    window._probForecastChart = new Chart(
      document.getElementById('probForecastChart'),
      {{
        type: 'line',
        data: data,
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          plugins: {{
            legend: {{
              position: 'top',
              labels: {{ color: textColor, usePointStyle: true, pointStyleWidth: 10, padding: 16, font: {{ size: 12 }} }}
            }},
            tooltip: {{
              callbacks: {{
                label: ctx => ` ${{ctx.dataset.label}}: ${{(ctx.parsed.y||0).toFixed(2)}} FTE`
              }}
            }}
          }},
          scales: {{
            x: {{
              ticks: {{ color: textColor, font: {{ size: 11 }} }},
              grid:  {{ color: gridColor }}
            }},
            y: {{
              stacked: true,
              beginAtZero: true,
              ticks: {{ color: textColor, font: {{ size: 11 }}, callback: v => v.toFixed(1) + ' FTE' }},
              grid:  {{ color: gridColor }},
              title: {{ display: true, text: 'Expected FTE (prob-weighted)', color: textColor, font: {{ size: 11 }} }}
            }}
          }}
        }}
      }}
    );
  }})();
  // ─ Group FTE Chart ──────────────────────────────────────────────────
  (function () {{
    const dark      = document.body.classList.contains('dark');
    const textColor = dark ? '#8fa3bc' : '#3f4d66';
    const gridColor = dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)';
    const data      = {group_chart_data_json};

    window._groupFteChart = new Chart(
      document.getElementById('groupFteChart'),
      {{
        type: 'line',
        data: data,
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          plugins: {{
            legend: {{
              position: 'top',
              labels: {{ color: textColor, usePointStyle: true, pointStyleWidth: 10, padding: 16, font: {{ size: 12 }} }}
            }},
            tooltip: {{
              callbacks: {{
                label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.parsed.y.toFixed(2)}} FTE`
              }}
            }}
          }},
          scales: {{
            x: {{
              ticks: {{ color: textColor, font: {{ size: 11 }} }},
              grid:  {{ color: gridColor }}
            }},
            y: {{
              beginAtZero: true,
              ticks: {{ color: textColor, font: {{ size: 11 }}, callback: v => v.toFixed(1) + ' FTE' }},
              grid:  {{ color: gridColor }}
            }}
          }}
        }}
      }}
    );
  }})();
  // ─ Detail table: filter + sort ──────────────────────────────────────────────────────
  (function () {{
    const table = document.getElementById('detail-table');
    if (!table) return;
    const tbody    = table.querySelector('tbody');
    const selEmp   = document.getElementById('flt-emp');
    const selProj  = document.getElementById('flt-proj');
    const selGroup = document.getElementById('flt-group');
    const txtSrch  = document.getElementById('flt-search');
    const btnClr   = document.getElementById('flt-clear');

    let sortCol = -1, sortAsc = true;

    function updateTotals() {{
      const tfootRows = table.querySelectorAll('tfoot tr');
      const tfoot = tfootRows[0];
      const cumRow = tfootRows[1];
      if (!tfoot) return;
      const visRows = Array.from(tbody.rows).filter(r => r.style.display !== 'none');
      // tfoot cells: 0=label(colspan=4), 1..N=months, N+1=grand-total
      const nMonths = tfoot.cells.length - 2;
      const sums = Array(nMonths).fill(0);
      visRows.forEach(tr => {{
        for (let i = 0; i < nMonths; i++) {{
          // columns: 0=Emp, 1=Proj, 2=Group, 3=Prob, 4..=months
          sums[i] += parseFloat(tr.cells[i + 4].textContent) || 0;
        }}
      }});
      const grand = sums.reduce((a, b) => a + b, 0);
      for (let i = 0; i < nMonths; i++) {{
        tfoot.cells[i + 1].innerHTML = '<strong>' + sums[i].toFixed(2) + '</strong>';
      }}
      tfoot.cells[nMonths + 1].innerHTML = '<strong>' + grand.toFixed(2) + '</strong>';
      if (cumRow) {{
        for (let i = 0; i < nMonths; i++) {{
          const remaining = sums.slice(i).reduce((a, b) => a + b, 0);
          cumRow.cells[i + 1].textContent = remaining.toFixed(2);
        }}
        cumRow.cells[nMonths + 1].textContent = grand.toFixed(2);
      }}
    }}

    function filterRows() {{
      const emp   = selEmp.value;
      const proj  = selProj.value;
      const group = selGroup.value;
      const srch  = txtSrch.value.toLowerCase();
      Array.from(tbody.rows).forEach(tr => {{
        const eMatch = !emp   || tr.dataset.emp   === emp;
        const pMatch = !proj  || tr.dataset.proj  === proj;
        const gMatch = !group || tr.dataset.group === group;
        const sMatch = !srch  || tr.textContent.toLowerCase().includes(srch);
        tr.style.display = (eMatch && pMatch && gMatch && sMatch) ? '' : 'none';
      }});
      updateTotals();
    }}

    function sortRows(colIdx) {{
      if (sortCol === colIdx) {{ sortAsc = !sortAsc; }}
      else {{ sortCol = colIdx; sortAsc = true; }}
      Array.from(table.querySelectorAll('thead th')).forEach((th, i) => {{
        th.dataset.sorted = (i === colIdx) ? (sortAsc ? 'asc' : 'desc') : '';
      }});
      const rows = Array.from(tbody.rows);
      rows.sort((a, b) => {{
        const av = a.cells[colIdx] ? a.cells[colIdx].textContent.trim() : '';
        const bv = b.cells[colIdx] ? b.cells[colIdx].textContent.trim() : '';
        const an = parseFloat(av), bn = parseFloat(bv);
        const cmp = (isNaN(an) || isNaN(bn)) ? av.localeCompare(bv) : an - bn;
        return sortAsc ? cmp : -cmp;
      }});
      rows.forEach(r => tbody.appendChild(r));
    }}

    selEmp.addEventListener('change', filterRows);
    selProj.addEventListener('change', filterRows);
    selGroup.addEventListener('change', filterRows);
    txtSrch.addEventListener('input',  filterRows);
    btnClr.addEventListener('click', () => {{
      selEmp.value = ''; selProj.value = ''; selGroup.value = ''; txtSrch.value = '';
      filterRows();
    }});

    Array.from(table.querySelectorAll('thead th')).forEach((th, i) => {{
      th.classList.add('sortable');
      th.addEventListener('click', () => sortRows(i));
    }});
  }})();
  // ─ Row-click filtering on project/group tables ────────────────────────────
  (function () {{
    const EXCL = 'row-excluded';

    function recalc(table) {{
      const tbody = table.querySelector('tbody');
      if (!tbody) return;
      const isCharge = table.classList.contains('charge-table');
      const totalRow = tbody.querySelector('.proj-table-total') || tbody.querySelector('.heatmap-total-row');
      const cumRow   = tbody.querySelector('.cum-row');
      if (!totalRow) return;

      const isProjTotal = totalRow.classList.contains('proj-table-total');
      const nMonths = isProjTotal ? totalRow.cells.length - 2 : totalRow.cells.length - 1;

      // For charge tables, pick the right base key and hours multiplier from active factors
      let hours = 1;
      let chargeBase = cell => parseFloat(cell.dataset.base) || 0;
      let fmtVal = v => v.toFixed(2);
      if (isCharge) {{
        const cFactors = {{
          prob:  document.getElementById('cfac-prob')?.getAttribute('aria-pressed')  !== 'false',
          rate:  document.getElementById('cfac-rate')?.getAttribute('aria-pressed')  !== 'false',
          hours: document.getElementById('cfac-hours')?.getAttribute('aria-pressed') !== 'false',
        }};
        hours = cFactors.hours ? (parseFloat(document.getElementById('chargeHours')?.value) || 0) : 1;
        if (cFactors.prob && cFactors.rate)       chargeBase = c => parseFloat(c.dataset.base)     || 0;
        else if (cFactors.prob && !cFactors.rate) chargeBase = c => parseFloat(c.dataset.baseProb) || 0;
        else if (!cFactors.prob && cFactors.rate) chargeBase = c => parseFloat(c.dataset.baseRate) || 0;
        else                                      chargeBase = c => parseFloat(c.dataset.baseFte)  || 0;
        fmtVal = (cFactors.rate || cFactors.prob)
          ? v => Math.round(v).toLocaleString()
          : v => v.toFixed(2);
      }}

      const dataRows = Array.from(tbody.rows).filter(tr =>
        !tr.classList.contains('proj-table-total') &&
        !tr.classList.contains('cum-row') &&
        !tr.classList.contains('heatmap-total-row') &&
        !tr.classList.contains(EXCL)
      );

      const sums = Array(nMonths).fill(0);
      dataRows.forEach(tr => {{
        for (let i = 0; i < nMonths; i++) {{
          const cell = tr.cells[i + 1];
          if (!cell) continue;
          sums[i] += isCharge
            ? chargeBase(cell) * hours
            : parseFloat(cell.textContent.replace(/[^0-9.-]/g, '')) || 0;
        }}
      }});

      const grand = sums.reduce((a, b) => a + b, 0);
      const fmt   = isCharge ? fmtVal : v => v.toFixed(2);

      for (let i = 0; i < nMonths; i++) totalRow.cells[i + 1].textContent = fmt(sums[i]);
      if (isProjTotal) totalRow.cells[nMonths + 1].textContent = fmt(grand);

      if (cumRow) {{
        for (let i = 0; i < nMonths; i++) {{
          cumRow.cells[i + 1].textContent = fmt(sums.slice(i).reduce((a, b) => a + b, 0));
        }}
        if (isProjTotal) cumRow.cells[nMonths + 1].textContent = fmt(grand);
      }}
    }}

    document.querySelectorAll('.filterable-table').forEach(table => {{
      const tbody = table.querySelector('tbody');
      if (!tbody) return;
      Array.from(tbody.rows).forEach(tr => {{
        if (tr.classList.contains('proj-table-total') || tr.classList.contains('cum-row') ||
            tr.classList.contains('heatmap-total-row')) return;
        tr.addEventListener('click', () => {{ tr.classList.toggle(EXCL); recalc(table); }});
      }});
    }});

    // Re-expose recalc so the charge hours input can also trigger it
    window._recalcFilterable = () => {{
      document.querySelectorAll('.filterable-table.charge-table').forEach(recalc);
    }};
  }})();

  // ─ Charge: factor toggles + live recalculation ────────────────────────────
  (function () {{
    const input   = document.getElementById('chargeHours');
    if (!input) return;

    // Factor state
    const factors = {{ prob: true, rate: true, hours: true }};

    function getBase(el) {{
      // Pick the pre-aggregated base that matches the active prob/rate toggles
      const useProp = factors.prob;
      const useRate = factors.rate;
      if (useProp && useRate)  return parseFloat(el.dataset.base)      || 0;  // FTE × Prob × Rate
      if (useProp && !useRate) return parseFloat(el.dataset.baseProb)  || 0;  // FTE × Prob
      if (!useProp && useRate) return parseFloat(el.dataset.baseRate)  || 0;  // FTE × Rate
      return parseFloat(el.dataset.baseFte) || 0;                             // FTE only
    }}

    function fmt(v) {{
      return factors.rate || factors.prob
        ? Math.round(v).toLocaleString()
        : v.toFixed(2);
    }}

    function recalcCharge() {{
      const h = factors.hours ? (parseFloat(input.value) || 0) : 1;
      document.querySelectorAll('#sec-charge [data-base]').forEach(el => {{
        if (el.closest('tr')?.classList.contains('row-excluded')) return;
        el.textContent = fmt(getBase(el) * h);
      }});
      if (window._recalcFilterable) window._recalcFilterable();
    }}

    const MEANINGS = [
      [true,  true,  true,  'Expected cost'],
      [true,  true,  false, 'Cost per hour of capacity (rate base)'],
      [true,  false, true,  'Expected FTE\u00b7hours (prob-weighted)'],
      [true,  false, false, 'Expected FTE\u00b7months (prob-weighted)'],
      [false, true,  true,  'Budget ceiling \u2014 ignores probability'],
      [false, true,  false, 'Rate\u00b7FTE base (per hour, ceiling)'],
      [false, false, true,  'Raw FTE\u00b7hours (unweighted)'],
      [false, false, false, 'Raw FTE\u00b7months (unweighted)'],
    ];

    function updateMeaning() {{
      const el = document.getElementById('charge-meaning');
      if (!el) return;
      const m = MEANINGS.find(r => r[0]===factors.prob && r[1]===factors.rate && r[2]===factors.hours);
      el.textContent = m ? m[3] : '';
    }}

    // Wire factor buttons
    document.querySelectorAll('.charge-factor').forEach(btn => {{
      btn.addEventListener('click', () => {{
        const f = btn.dataset.factor;
        factors[f] = !factors[f];
        btn.setAttribute('aria-pressed', factors[f] ? 'true' : 'false');
        // Show/hide hours input based on hours factor
        if (f === 'hours') {{
          const ctrl = document.querySelector('.charge-controls');
          if (ctrl) ctrl.style.opacity = factors.hours ? '1' : '0.35';
        }}
        updateMeaning();
        recalcCharge();
      }});
    }});

    input.addEventListener('input', recalcCharge);

    // Expose for filterable recalc
    window._recalcCharge = recalcCharge;
  }})();

  // ─ CSV download ──────────────────────────────────────────────────
  document.getElementById('btnCsv').addEventListener('click', function () {{
    const rows = {csv_json};
    const NL   = String.fromCharCode(10);
    const csv  = rows.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join(NL);
    const blob = new Blob([csv], {{ type: 'text/csv' }});
    const url  = URL.createObjectURL(blob);
    const a    = Object.assign(document.createElement('a'), {{ href: url, download: 'manloading_data.csv' }});
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }});
</script>
</body>
</html>
"""

OUTPUT.write_text(HTML, encoding="utf-8")
print(f"Report saved → {OUTPUT}")
