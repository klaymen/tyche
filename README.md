# Tyche — Manloading Dashboard

A lightweight Python tool that reads employee allocation data from an Excel file and generates a self-contained, interactive HTML dashboard for manloading analysis.

---

## Features

- **Employee utilisation heatmap** — colour-coded by probability-weighted FTE (zero / low / mid / full / over)
- **Project summary table** — FTE per project per month, with backward-cumulative "remaining" rows
- **Project Group rollup** — portfolio-level aggregation of FTE and charge base
- **Overallocation detection** — flags any employee × month where FTE exceeds 1.05, both raw and probability-weighted
- **Charge base calculations** — FTE × Probability × Hourly Rate per project/group/employee
- **KPI stat cards** — total employees, projects, FTE·months, expected FTE·months, avg FTE per employee
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
| `Hourly Rate` | Hourly rate for charge base calculations (optional; defaults to 0) |
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
python manloading_report.py

# Using your own Excel file
python manloading_report.py path/to/your_file.xlsx
```

The report is written to `output/manloading_report.html` (the directory is created automatically).

### Regenerate sample data

```bash
python _make_sample.py
```

This overwrites `sample_input.xlsx` with a fresh 25-row dataset covering 13 employees, 4 projects, 3 project groups, and 12 months (Apr 2026 – Mar 2027).

---

## Output

`output/manloading_report.html` — a fully self-contained single-file HTML page. Open it in any modern browser; no server or internet connection required.

### Dashboard sections

| Section | Description |
|---------|-------------|
| **KPI cards** | High-level metrics at a glance |
| **Utilisation heatmap** | Employee rows × month columns; cells are probability-weighted FTE |
| **Project summary** | Project rows × month columns with a total column and cumulative-remaining row |
| **Project group summary** | Portfolio rollup of FTE and charge base |
| **Employee detail** | Per-employee breakdown across projects |

### Colour coding (FTE cells)

| Colour | Meaning |
|--------|---------|
| Grey | 0 FTE (no allocation) |
| Yellow | < 0.5 FTE (under-utilised) |
| Blue | 0.5 – <1.0 FTE |
| Green | Exactly 1.0 FTE (fully loaded) |
| Red | > 1.05 FTE (overallocated) |

---

## Key Concepts

**FTE (Full-Time Equivalent)** — a value of 1.0 means the employee is 100 % allocated to that project for that month.

**Probability-weighted FTE** — `FTE × Probability`. Used for planning when some projects are not yet confirmed. A project with `Probability = 0.55` contributes only 55 % of its raw FTE to expected totals.

**Charge base** — `FTE × Probability × Hourly Rate`. Multiply by the number of working hours in a month to get an estimated cost or revenue figure.

**Overallocation** — reported in two flavours:
- *Raw*: total FTE across all projects for an employee exceeds 1.05 (ignores probability)
- *Expected*: probability-weighted total FTE exceeds 1.05

---

## Versioning

Current version: **1.0.0**
