# Tyche — Manloading Dashboard

A lightweight Python tool that reads employee allocation data from an Excel file and generates a self-contained, interactive HTML dashboard for manloading analysis.

---

## Features

- **Employee utilisation heatmap** — colour-coded by probability-weighted FTE (zero / low / mid / full / over)
- **Project summary table** — FTE per project per month, with backward-cumulative "remaining" rows
- **Project Group rollup** — portfolio-level aggregation of FTE and charge
- **Overallocation detection** — flags any employee × month where FTE exceeds 1.05, both raw and probability-weighted
- **Charge section with factor toggles** — click **Probability**, **Hourly Rate**, or **Monthly Hours** in the equation to switch each factor on/off; the table values and the meaning label update instantly
- **Probability pipeline chart** — stacked area chart and table breaking demand into Backlog (≥90 %), Outlook (50–90 %), and Opportunity (<50 %) bands
- **KPI stat cards** — total employees, projects, FTE·months, expected FTE·months, avg FTE per employee, overallocation counts
- **Dark mode** — toggle via the moon/sun icon in the sidebar
- **Filter & sort** — filter the detail table by employee, project, or group; click any row to exclude it from totals
- **CSV export** — download the raw allocation data as a CSV file
- **Single-file output** — all CSS and JS are inlined; the HTML report has no external dependencies

---

## Project Structure

```
tyche/
├── manloading_report.py   # Main script — reads Excel, generates HTML report
├── _make_sample.py        # Utility to regenerate sample_input.xlsx
├── sample_input.xlsx      # Sample allocation data (12 months, 13 employees, 4 projects)
└── output/
    └── manloading_report.html  # Generated report (git-ignored)
```

---

## Input Format

The script expects an Excel file (`.xlsx`) with the following structure:

| Row | Content |
|-----|---------|
| 1   | Optional "Months" label merged across month columns |
| 2   | Column headers (see below) |
| 3+  | One row per employee × project allocation |

**Required columns:**

| Column | Description |
|--------|-------------|
| `Employee` | Employee full name |
| `Project` | Project name |
| `Project Group` | Portfolio / group name (optional but recommended) |
| `Probability` | Likelihood the project proceeds (0.0 – 1.0; defaults to 1.0 if absent) |
| `Hourly Rate` | Hourly rate for charge calculations (optional; defaults to 0) |
| `YYYY-MM` ... | One column per month with FTE allocation values (e.g. `2026-04`, `2026-05`, …) |

Each row represents a single employee's allocation to a single project. An employee may appear on multiple rows (one per project).

---

## Installation

Python 3.9+ is required.

```bash
pip install pandas openpyxl
```

---

## Usage

### Generate the report

```bash
# Using the bundled sample data
python3 manloading_report.py

# Using your own Excel file
python3 manloading_report.py path/to/your_file.xlsx
```

The report is written to `output/manloading_report.html` (the directory is created automatically).

### Regenerate sample data

```bash
python3 _make_sample.py
```

This overwrites `sample_input.xlsx` with a fresh dataset covering 13 employees, 4 projects, 3 project groups, and 12 months (Apr 2026 – Mar 2027).

---

## Output

`output/manloading_report.html` — a fully self-contained single-file HTML page. Open it in any modern browser; no server or internet connection required.

### Dashboard sections

| Section | Description |
|---------|-------------|
| **KPI cards** | High-level metrics at a glance |
| **Utilisation heatmap** | Employee × month cells, probability-weighted FTE, with per-cell project breakdown tooltip |
| **Project FTE Demand** | Line/area chart + table of prob-weighted FTE per project per month |
| **Project Group FTE Demand** | Portfolio-level FTE rollup chart and table |
| **Per-project heatmaps** | One heatmap per project showing each employee's contribution |
| **Probability Forecast** | Pipeline breakdown by confidence band (Backlog / Outlook / Opportunity) |
| **Employee × Project Detail** | Raw allocation data, filterable and sortable; click rows to exclude from totals |
| **Charge** | Configurable cost table — toggle factors on/off via the equation in the header |

### Colour coding (FTE cells)

| Colour | Range | Meaning |
|--------|-------|---------|
| Grey | 0.00 | No allocation |
| Yellow | < 0.50 | Under-utilised |
| Blue | 0.50 – <1.0 | Partial allocation |
| Green | = 1.0 | Fully loaded |
| Red | > 1.0 | Overallocated |

---

## Key Concepts

**FTE (Full-Time Equivalent)** — a value of 1.0 means the employee is 100 % allocated to that project for that month.

**Probability-weighted FTE** — `FTE × Probability`. Used for planning when some projects are not yet confirmed. A project with `Probability = 0.55` contributes only 55 % of its raw FTE to expected totals.

**Charge** — the product of one or more factors: FTE, Probability, Hourly Rate, and Monthly Hours. Use the clickable equation in the Charge section header to toggle each factor independently. The meaning label after `=` updates to describe what the resulting numbers represent:

| Probability | Hourly Rate | Monthly Hours | Meaning |
|:-----------:|:-----------:|:-------------:|---------|
| ✓ | ✓ | ✓ | Expected cost |
| ✓ | ✓ | — | Cost per hour of capacity (rate base) |
| ✓ | — | ✓ | Expected FTE·hours (prob-weighted) |
| ✓ | — | — | Expected FTE·months (prob-weighted) |
| — | ✓ | ✓ | Budget ceiling — ignores probability |
| — | ✓ | — | Rate·FTE base (per hour, ceiling) |
| — | — | ✓ | Raw FTE·hours (unweighted) |
| — | — | — | Raw FTE·months (unweighted) |

**Overallocation** — reported in two flavours:
- *Raw*: total FTE across all projects for an employee exceeds 1.05 (ignores probability)
- *Expected*: probability-weighted total FTE exceeds 1.05


