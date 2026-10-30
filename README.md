# Northline Vendor Spend and Savings Model (Excel)

Twelve months of vendor spend by category and cost center, with a ranked list of cost savings from fee rate benchmarks, volume tiers and duplicate vendors.

> Northline Consumer Products is a fictional company. All figures are sample data, not from any real employer.

## Problem

Northline spent about $9.5M a year across 12 vendors (Jul 2025 to Jun 2026). Processing fees were the largest line, and some suppliers were set up twice under slightly different names. Finance needed a ranked, defensible savings list.

## What I built

`Northline_Vendor_Spend.xlsx`, built with `build_workbook.py` (openpyxl). Yellow cells are inputs and black cells are formulas.

| Sheet | Contents |
|---|---|
| Vendors | Vendor master: category, cost center, savings lever, current fee rate, target rate or discount |
| Monthly_Spend | Spend by vendor for 12 months with annual totals |
| Savings | Annual savings per vendor by formula, with rank |
| Summary | Savings by lever with SUMIF, and share of total spend |

Savings logic: processing fees save spend x (current rate minus target rate) / current rate. Other levers save spend x target discount.

## How to run

```
pip install openpyxl
python3 build_workbook.py
```

Open the workbook in Excel and change any yellow cell to test a new target.

## Key results (sample Northline figures)

| Lever | Annual savings |
|---|---|
| Rate benchmark (processing fees, staffing, telecom) | $317,038 |
| Volume tier (freight, hosting, office supply) | $88,118 |
| Duplicate vendor consolidation | $51,425 |
| **Total** | **$456,581 (4.8% of $9.54M spend)** |

Top opportunities:

1. Harbor Card Processing: fee rate 2.90% to 2.68%, saves $182,116
2. BrightPath Staffing: markup 18% to 17%, saves $61,261
3. Summit Payments: fee rate 3.10% to 2.90%, saves $57,980
4. Atlas Freight: 3% volume tier, saves $49,333
5. Cloudway Hosting: 3% volume tier, saves $23,245

## Conclusion

Renegotiating two card processors and consolidating duplicate suppliers delivers most of the roughly **$450K** of annual potential savings. The model gives procurement a ranked list and lets finance flex each target rate before negotiations.

## Stack

Excel, Python (openpyxl). No macros.
