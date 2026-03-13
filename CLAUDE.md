# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup

```bash
pip install -r requirements.txt
```

## Running the tool

```bash
python main.py        # English version
python main_es.py     # Spanish version
```

Both scripts accept piped input for non-interactive testing:

```bash
printf "Acme Corp\n3\n12\n500000\nout.xlsx\n300000\n120000\n60000\n25000\n15000\n25\n0\n0\n..." | python main.py
```

## Git / GitHub workflow

**Always work on a feature branch — never commit directly to `master`.**

```bash
git checkout -b feature/<name>
# make changes, commit, then:
git push -u origin feature/<name>
gh pr create ...
```

## Architecture

The tool is split into three layers:

### 1. `calculator.py` — pure financial logic (no I/O, no language)
- `compute_financials(project_data)` is the single entry point; it drives everything and returns a `results` dict.
- `project_data` shape: `{ company_name, num_years, wacc (decimal), initial_investment, yearly_data: [{ revenue, cogs, opex, da, interest, tax_rate (decimal), capex, delta_wc }] }`
- `results` shape: `{ year_metrics: [...], npv, irr, simple_payback, discounted_payback, avg_gross_margin, avg_ebitda_margin, avg_net_margin }`
- `numpy_financial` is optional — IRR and NPV fall back to manual Newton-Raphson / sum if the package is missing.
- Tax is floored at 0 (no negative tax when EBT < 0).
- FCF formula: `net_income + da - capex - delta_wc`

### 2. `main.py` / `main_es.py` — CLI layer
- Collects metadata (company name, years, WACC, initial investment, filename), then loops year-by-year collecting income statement inputs.
- All monetary inputs are stored as raw dollar floats; percentages are divided by 100 before being passed to `calculator.py`.
- File save is wrapped in a `PermissionError` loop to handle the "file open in Excel" case.
- `main_es.py` additionally accepts comma as a decimal separator.

### 3. `excel_generator.py` / `excel_generator_es.py` — Excel output layer
- `generate_excel(project_data, results, filename)` is the only public function.
- Sheet 1 ("Cash Flow Projection" / "Proyección de Flujo de Caja"): income statement rows 5–18, blank separator at row 19, FCF rows 20–26. Panes frozen at B5.
- Sheet 2 ("Financial Summary" / "Resumen Financiero"): NPV/IRR cells get green/red fill based on sign (NPV) or vs. WACC (IRR). Payback displays "Not within projection" / "Fuera del período de proyección" when `None`.
- Styling is routed through `_apply_cell_style(cell, style_name)` — add new styles there rather than inline.
- `None` values in `year_metrics` render as "N/A" (English) or "N/D" (Spanish).

### Adding a new language
Create `main_<lang>.py` and `excel_generator_<lang>.py` — `calculator.py` is shared and requires no changes.
