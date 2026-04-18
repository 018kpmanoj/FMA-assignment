# FMA EC1 - Financial Statement Analysis

**Course:** Financial Management & Accounting (FMA)  
**Assignment:** EC1 - Financial Statement Analysis  
**Student:** K P Manoj | 2025MB26539  
**Programme:** MBA in AI for Business | BITS Pilani (WILP) | Semester 1, Jan 2026 Batch  

## Overview

Comparative financial statement analysis of **HCL Technologies Ltd.** vs **Wipro Ltd.** using 5 years (FY2021-FY2025) of consolidated financial data sourced from Moneycontrol.

## Files

| File | Description |
|------|-------------|
| `excel_report.xlsx` | Excel workbook with **complete** raw data (Balance Sheet, P&L, Cash Flow - all rows), computed ratios, DuPont Analysis, and Cash Conversion Cycle |
| `word_report.docx` | Analytical report (TNR 11, single-spaced) with Introduction, Intra-Company Analysis, Inter-Company Comparison, DuPont & CCC derivations, and Conclusion |
| `ratio_calculation_proofs.docx` | Enhanced calculation proofs for all 19 ratios + DuPont + CCC with formulas, **source derivation references** (where each value comes from), substitutions, and interpretations |
| `HCL Tech-1.xlsx` | Source data - HCL Technologies **consolidated only** financials (Balance Sheet, P&L, Cash Flow) from Moneycontrol |
| `Wipro-1.xlsx` | Source data - Wipro **consolidated only** financials (Balance Sheet, P&L, Cash Flow) from Moneycontrol |
| `_build_final.py` | Python script that cleans datasets, generates Excel and Word deliverables |

**Note:** The dataset files (`HCL Tech-1.xlsx`, `Wipro-1.xlsx`) contain only consolidated financial statements. Standalone data and the Moneycontrol pre-computed Ratios sheet have been removed since all ratios are independently computed in `excel_report.xlsx`.

## Ratios Computed

### Liquidity
- Current Ratio, Quick Ratio, Absolute Liquid Ratio

### Profitability
- Gross Profit Ratio, Operating Profit Ratio (EBIT Margin), Net Profit Ratio, ROE, ROCE, EPS

### Solvency
- Debt-Equity Ratio, Interest Coverage Ratio

### Turnover
- Debtors Turnover, Average Collection Period, Creditors Turnover, Average Payment Period, Inventory Turnover, Inventory Holding Days, Fixed Asset Turnover

### Valuation
- P/E Ratio

### Advanced Analysis
- DuPont Decomposition (NPM x Asset Turnover x Equity Multiplier)
- Cash Conversion Cycle (DIO + DSO - DPO)

## Key Finding

HCL Technologies outperforms Wipro across most financial parameters including revenue growth (CAGR 11.6% vs 9.5%), ROE (25.0% vs 15.9%), ROCE (30.8% vs 19.0%), and solvency (D/E 0.03x vs 0.20x). The DuPont analysis reveals asset turnover as the primary driver of the ROE gap.

## Tech Stack

- Python 3.x
- `openpyxl` - Excel generation
- `python-docx` - Word document generation
- Data Source: [Moneycontrol](https://www.moneycontrol.com/)
