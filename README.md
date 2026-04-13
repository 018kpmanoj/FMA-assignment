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
| `excel_report.xlsx` | Excel workbook with raw data (P&L, Balance Sheet) and computed ratios (Liquidity, Profitability, Solvency, Turnover, Valuation), DuPont Analysis, and Cash Conversion Cycle |
| `word_report.docx` | 6-page analytical report (TNR 11, single-spaced) covering Introduction, Intra-Company Analysis, Inter-Company Comparison, DuPont & CCC, and Conclusion |
| `ratio_calculation_proofs.docx` | Detailed calculation proofs for all 19 ratios + DuPont + CCC with formulas, substitutions, and interpretations |
| `HCL Tech-1.xlsx` | Source data - HCL Technologies consolidated financials (Moneycontrol) |
| `Wipro-1.xlsx` | Source data - Wipro consolidated financials (Moneycontrol) |
| `_build_final.py` | Python script used to generate the Excel and Word deliverables |

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
