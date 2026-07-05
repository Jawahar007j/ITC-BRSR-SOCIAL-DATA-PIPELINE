# ITC Social Performance Dashboard

*A Python ETL Pipeline and Interactive Excel Dashboard for ESG Analysis*

**Repository:** ITC-Social-Performance-Dashboard | **Author:** Jawahar

---

Python-based ETL pipeline and interactive Excel dashboard for analyzing ITC Limited's social ESG performance using annual report data. This project extracts, cleans, and structures social KPI data from ITC's Report and Accounts 2026 (BRSR / Sustainability disclosures), transforming unstructured PDF tables into a clean, analysis-ready dataset, then visualizes it through a single-page interactive Excel dashboard comparing FY2024-25 against FY2025-26 performance.

![Dashboard Preview](ITC_SOCIAL_DASHBOARD.png)
*Figure 1: Final ESG social performance dashboard (FY2024-25 vs FY2025-26)*

## 1. Project Objectives

- Convert unstructured PDF disclosures from a 300+ page corporate annual report into a structured, analysis-ready dataset.
- Isolate and validate social KPIs (Workforce, Wellbeing, Training, Safety, Human Rights, Sourcing, CSR) relevant to BRSR reporting.
- Programmatically reconstruct merged, multi-level PDF table headers without manual re-typing.
- Build a reproducible Python ETL pipeline covering parsing, classification, and transformation.
- Deliver a decision-ready, interactive Excel dashboard comparing year-over-year social performance.
- Demonstrate applied ESG analytics capability as part of a professional sustainability analytics portfolio.

## 2. Dashboard Highlights

- **Workforce Composition**: Permanent vs. Other-than-Permanent Employees and Workers, by gender.
- **Employee Wellbeing**: Health insurance coverage, maternity benefits, return-to-work and retention rates.
- **Training Coverage**: Health & Safety and Skill Upgradation training, by employee category and gender.
- **Performance & Career Development**: Review coverage across Employees and Workers.
- **Safety Metrics**: LTIFR, recordable injuries, and fatalities (Employees vs. Workers).
- **Human Rights & Pay Equity**: Gender pay gap and minimum wage compliance.
- **Sourcing**: MSME sourcing share, domestic sourcing share, and presence by location type.
- **CSR Impact**: Spend across aspirational districts and beneficiaries by program area.

## 3. Repository Structure

| File | Contents | Purpose |
|---|---|---|
| `ITC_SOCIAL_DASHBOARD.png`, `ITC_SOCIAL_DASHBOARD.pdf` | Dashboard exports | Final dashboard, for quick viewing |
| `ITC_Social_Dataset___Dashboard.xlsx` | Dashboard + 35 cleaned tables + chart-ready pivot tables | Full working workbook |
| `ITC_social_dataset.py` | ETL script | Parses raw BRSR tables into the refined, chart-ready workbook |
| `README.md` | — | Project documentation |

## 4. Data Source

Data is extracted from ITC Limited's publicly available Report and Accounts 2026, specifically the Business Responsibility and Sustainability Report (BRSR) Social section (pages 367–395), covering social disclosures for FY2024-25 and FY2025-26. The source PDF is not included in this repository due to file size; it can be downloaded from ITC Limited's official investor relations page.

## 5. ETL Pipeline Workflow

The pipeline is implemented as a single script, `ITC_social_dataset.py`, structured around the following stages:

| Stage | Function | Purpose |
|---|---|---|
| 1 | `load_tables()` | Loads all 35 raw BRSR Social tables extracted from the source PDF. |
| 2 | `classify_rows()` | Distinguishes header rows from data rows by detecting data-like cells (numbers, NA/dash markers, long free text) vs. pure label text. |
| 3 | `fill_headers_hierarchical()` | Reconstructs multi-level merged headers (e.g. FY → Gender → Metric), resetting the fill at each group boundary so labels don't bleed across unrelated columns. |
| 4 | `build_table_sheet()` | Renders a clean, formatted, human-readable version of each source table. |
| 5 | `build_metric_tables()` | Melts each table into per-KPI cross-tabs (Category × Financial Year). |
| 6 | `build_tables_sheet()` | Writes each metric's cross-tab as a self-contained, chart-ready table. |

## 6. Datasets

- **`ITC_Social_Dataset___Dashboard.xlsx`**: Contains the dashboard, 35 cleaned source tables (one per BRSR disclosure), and their corresponding `_Tables` sheets — one chart-ready cross-tab per KPI, with FY2024-25 and FY2025-26 values side by side.

Each `_Tables` sheet follows a consistent structure:

| Column | Description |
|---|---|
| Category | Row-level identity (e.g. Employees – Permanent – Male) |
| FY 2025-26 | Current financial year value |
| FY 2024-25 | Previous financial year value |
| Unit | Noted in the block title (Number or %) |

## 7. Key Year-over-Year Findings

- Total employees stood at 11,442, with female representation at 2,080 (~18%).
- Health insurance, maternity benefits, and return-to-work rates all held at 100% coverage.
- LTIFR improved to 0.02, with zero employee fatalities recorded.
- Worker union membership stood at 86%.
- Gender pay gap (gross wages paid to females as % of total) moved from 10% to 11%.
- CSR spend in aspirational districts totaled ~₹4.99 crore (in Rs. Lakhs) across 17+ states, with Women Empowerment and Health & Nutrition programs receiving the largest beneficiary allocations.

## 8. Tech Stack

- **Python**: `openpyxl`, `re`
- **Excel**: Pivot tables and charts — a native Excel dashboard with no external BI tool required.

## 9. How to Reproduce

```bash
pip install openpyxl
```

Update the `SRC` and `OUT` paths at the top of `ITC_social_dataset.py` to point to your local raw dataset and desired output location, then run:

```bash
python ITC_social_dataset.py
```

The refined workbook is written to the configured `OUT` path. The dashboard itself is built manually in Excel on top of the generated `_Tables` sheets.

## 10. About the Author

Built by **Jawahar**, an MBA and MSc graduate focused on ESG Analyst and Data/Business Analyst roles. This dashboard was built voluntarily using ITC Limited's publicly available report data, as part of a hands-on portfolio project demonstrating applied ESG analytics capability.

## 11. License

This project is for educational and portfolio purposes. Underlying data belongs to ITC Limited and is sourced from their publicly filed Report and Accounts 2026.
