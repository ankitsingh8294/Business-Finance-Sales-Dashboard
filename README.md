# Business Finance & Sales Performance Dashboard
**Fictional Company: Nexora Retail Solutions | Self-Initiated Portfolio Project by Ankit Anand**

> **Disclaimer:** This is a self-initiated portfolio project built with fictional company data.
> "Nexora Retail Solutions" is not a real company. No real employer, client, or business
> relationship is implied. All transactions, customers, and employees in the dataset are
> randomly generated for demonstration purposes only.

---

## Overview

A complete Excel-based finance and sales analysis system built from scratch to practice and
demonstrate practical business, accounting, and data-analysis skills for entry-level job
applications (Accounts Assistant, Junior Accountant, MIS Executive, Operations/Business
Support, and junior Data/Reporting roles).

The project models 12 months of order-level sales data for a fictional consumer-electronics
e-commerce retailer and turns it into a formula-driven KPI dashboard.

## Business Problem

A small retail business has order-level data sitting in a flat spreadsheet with no easy way to
see revenue, cost, profit, or performance trends across time, region, product category, or
salesperson. Decisions are being made without a clear, current view of business performance.

## Objectives

1. Build a clean, well-structured dataset of realistic sales transactions.
2. Calculate Revenue, COGS, Gross Profit, Operating Expense and Net Profit for every order
   using live formulas (not hardcoded numbers).
3. Summarise performance by month, region, category, product, salesperson, payment method,
   and order status.
4. Present the results on a single-page, recruiter-friendly dashboard with KPI cards and charts.
5. Draw business insights that are directly supported by the data.

## Dataset

| Attribute | Detail |
|---|---|
| Transactions | 560 individual orders |
| Time period | Aug 2025 - Jul 2026 (12 months) |
| Regions | 5 (North, South, East, West, Central) |
| Cities | 10 |
| Product categories | 6 (Audio, Mobile Accessories, Computer Peripherals, Smart Home, Wearables, Cameras) |
| Products | 22 |
| Salespersons | 8 |
| Payment methods | 5 (UPI, Credit Card, Debit Card, Net Banking, Cash on Delivery) |
| Order statuses | 4 (Delivered, Cancelled, Returned, Pending) |
| Duplicate Order IDs | 0 |
| Missing values | 0 |

**Columns:** Order ID, Order Date, Customer ID, Customer Name, Region, City, Salesperson,
Product, Category, Quantity, Unit Price, Discount %, Revenue, COGS, Gross Profit, Operating
Expense, Net Profit, Payment Method, Order Status.

## Tools & Technologies

- Microsoft Excel (built and verified for Excel 2016+)
- Excel Tables with AutoFilter
- Formulas: `SUMIFS`, `COUNTIFS`, `INDEX` / `MATCH`, `IF` / `IFERROR`, `LARGE`, `TEXT`, date
  functions
- Conditional formatting-ready structure
- SUMIFS-based pivot-style summary tables (with instructions included for building native
  Excel PivotTables & Slicers on the same data)
- Bar, Line and Pie charts

## Methodology

1. **Data generation** - a realistic transaction dataset was generated with logical
   relationships (e.g., Revenue = Quantity x Unit Price x (1 - Discount%); COGS is tied to a
   per-product unit cost; Operating Expense is a category-level % of Revenue).
2. **Data modelling** - a Product Master table (Product -> Category -> Unit Cost) and a
   Category Operating-Expense table were built as lookup references. Every financial column in
   Raw Data is a live formula that pulls from these reference tables using `INDEX`/`MATCH`,
   wrapped in `IFERROR` for safety.
3. **Revenue recognition** - Cancelled and Returned orders are excluded from every financial
   total (Revenue, COGS, Gross Profit, Operating Expense, Net Profit) and from the order/unit
   counts used for AOV and ASP. Only Delivered and Pending orders are treated as recognized
   business. Every `SUMIFS`/`COUNTIFS` aggregation in the workbook applies this same rule, so
   the KPI Summary, Pivot Analysis tables, Dashboard charts and Insights all agree with each
   other. The one deliberate exception is the Order Status Distribution table, which shows
   every status by design so the cancellation/return rate itself is visible.
4. **Aggregation** - `SUMIFS`/`COUNTIFS` summary tables were built for every dimension
   (region, category, product, salesperson, payment method, order status, month).
5. **KPI calculation** - headline KPIs were calculated once, centrally, and referenced
   everywhere else in the workbook so there is a single source of truth.
6. **Dashboard design** - KPI cards and 8 charts were laid out on one sheet, all pulling
   live from the summary tables.
7. **Insight generation** - insight sentences were written as formulas (`INDEX`/`MATCH`/`MAX`/
   `MIN` combined with `TEXT` and `&`) so they recalculate automatically and stay honestly tied
   to what the data shows.
8. **Quality audit** - checked for duplicate Order IDs, missing values, formula errors,
   cross-sheet total reconciliation, and unrealistic figures before finalising.

## KPIs Calculated

Total Revenue, Total COGS, Gross Profit, Gross Margin %, Operating Expenses, Net Profit, Net
Profit Margin %, Total Orders Placed, Recognized Orders, Cancelled/Returned Orders, Total
Units Sold, Average Order Value (AOV), Average Selling Price (ASP). All calculated live from
the Raw Data sheet - see 'KPI Summary'.

## Dashboard Features

- 8 KPI cards (Total Revenue, Total COGS, Gross Profit, Net Profit, Net Profit Margin, Total
  Orders, Total Units Sold, Average Order Value)
- Monthly Revenue Trend (line chart)
- Monthly Net Profit Trend (line chart)
- Region-wise Revenue (bar chart)
- Category-wise Revenue (bar chart)
- Top 10 Products by Gross Profit (bar chart)
- Salesperson Performance (bar chart)
- Expense Breakdown by Category (pie chart)
- Order Status Distribution (pie chart)
- Interactive filtering via the Raw Data Excel Table's built-in column filters, plus
  step-by-step instructions to add native PivotTable Slicers

## Key Insights

*(All calculated directly from the dataset - see the Insights sheet for the live formulas. All
revenue and profit figures below reflect recognized orders only - Cancelled and Returned orders
are excluded, per the Revenue Recognition Rule.)*

- Highest revenue month: **April 2026** (Rs. 3,84,240)
- Lowest revenue month: **November 2025** (Rs. 1,94,016)
- Highest-performing region by revenue: **Central**
- Highest gross-margin category: **Mobile Accessories** (56.1%)
- Top product by revenue and gross profit: **ViewClear 24-inch Monitor**
- Top-performing salesperson by revenue: **Suresh Kumar**
- Overall net profit margin: **32.7%** on recognized revenue of Rs. 32,82,715
- Cancelled + Returned orders make up **13.6%** of all 560 orders placed (76 orders) - these
  are excluded from every revenue and profit figure in the workbook

## Skills Demonstrated

- Microsoft Excel (formulas, tables, formatting)
- Data cleaning & preparation
- Data analysis
- Data visualization (charts)
- SUMIFS / COUNTIFS-based pivot-style summarisation (plus working knowledge of native
  PivotTables & PivotCharts)
- Excel formulas (INDEX/MATCH, IF/IFERROR, SUMIFS, COUNTIFS)
- KPI reporting
- Business reporting
- Financial analysis (revenue, cost, margin, profit)
- Sales analysis
- Dashboard development
- Drawing business insights that are supportable by data

## Limitations

- Unit costs and category operating-expense percentages are assumed reference figures set for
  this project, not real supplier or accounting data.
- Operating Expense is simplified to a flat percentage of Revenue per category rather than a
  full expense ledger (rent, salaries, ad spend tracked separately).
- The dataset is randomly generated, so while relationships between numbers are logical, it
  does not reflect a real company's actual seasonality or customer behaviour.
- Revenue recognition is simplified to a binary rule (Delivered/Pending = recognized,
  Cancelled/Returned = excluded). A real business might partially recognize revenue on
  Returned orders depending on when the return occurred, or track refund timing separately.
- Charts and summaries use SUMIFS-based "pivot-style" tables rather than native Excel
  PivotTables/Slicers by default, to keep the workbook simple, transparent, and robust across
  Excel versions. No slicer is connected to the Dashboard in this file; instructions to build
  native PivotTables/Slicers on the same data are included on the Pivot Analysis sheet for
  anyone who wants to add that interactivity themselves.

## Future Improvements

- Add native Excel PivotTables and Slicers directly in the workbook.
- Break Operating Expense into named sub-categories (logistics, marketing, packaging) with its
  own reference table.
- Add year-over-year comparison once a second year of data is available.
- Build a Power BI or Google Sheets version to demonstrate cross-tool reporting skills.
- Add a customer-level RFM (Recency, Frequency, Monetary) analysis for repeat-customer insight.

---
*Project by Ankit Anand - built independently as a portfolio piece.*
