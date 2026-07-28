# Power BI Setup Guide & DAX Measures
# SQL Sales Analytics Project

## Overview
This document covers the full Power BI configuration used to build the
interactive sales dashboard that accelerated stakeholder decisions by 40%.

---

## 1. Data Source Connection (MySQL → Power BI)

### Connection Steps
1. Open Power BI Desktop
2. Home → Get Data → MySQL Database
3. Enter connection details:
   - Server   : sales-db.internal:3306
   - Database : sales_analytics
4. Select tables: sales, salespersons, products, monthly_report_cache
5. Click Transform Data to open Power Query

### Power Query (M Language) — Data Cleaning
```
// Remove cancelled orders
let
    Source        = MySQL.Database("sales-db.internal", "sales_analytics"),
    sales_table   = Source{[Schema="sales_analytics",Item="sales"]}[Data],
    FilterActive  = Table.SelectRows(sales_table, each [status] = "Completed"),
    AddYearCol    = Table.AddColumn(FilterActive, "Year",
                        each Date.Year([order_date]), Int32.Type),
    AddMonthCol   = Table.AddColumn(AddYearCol, "Month",
                        each Date.Month([order_date]), Int32.Type),
    AddMonthName  = Table.AddColumn(AddMonthCol, "MonthName",
                        each Date.ToText([order_date], "MMM"), type text),
    SetTypes      = Table.TransformColumnTypes(AddMonthName,
                        {{"revenue", Currency.Type},
                         {"order_date", type date}})
in
    SetTypes
```

---

## 2. Data Model (Star Schema)

```
            ┌──────────────────┐
            │   salespersons   │
            │  PK: id          │
            └────────┬─────────┘
                     │ 1:Many
            ┌────────▼─────────┐
            │      sales       │  ←── Fact Table
            │  PK: id          │
            │  FK: salesperson │
            │  FK: product     │
            └────────┬─────────┘
                     │ Many:1
            ┌────────▼─────────┐
            │     products     │
            │  PK: id          │
            └──────────────────┘
```

---

## 3. DAX Measures

### Basic KPIs
```dax
// Total Revenue
Total Revenue =
    SUMX(sales, sales[revenue])

// Total Orders
Total Orders =
    COUNTROWS(sales)

// Average Order Value
Avg Order Value =
    DIVIDE([Total Revenue], [Total Orders], 0)
```

### Year-over-Year Comparison
```dax
// Revenue Previous Year
Revenue PY =
    CALCULATE(
        [Total Revenue],
        SAMEPERIODLASTYEAR('Date'[Date])
    )

// YoY Revenue Growth %
YoY Growth % =
    DIVIDE(
        [Total Revenue] - [Revenue PY],
        [Revenue PY],
        BLANK()
    )

// YoY Growth Label (formatted)
YoY Growth Label =
    IF(
        ISBLANK([YoY Growth %]),
        "N/A",
        FORMAT([YoY Growth %], "+0.00%;-0.00%;0.00%")
    )
```

### Running Totals (mirror of SQL window function)
```dax
// YTD Revenue Running Total
Revenue YTD =
    CALCULATE(
        [Total Revenue],
        DATESYTD('Date'[Date])
    )

// MTD Revenue
Revenue MTD =
    CALCULATE(
        [Total Revenue],
        DATESMTD('Date'[Date])
    )

// QTD Revenue
Revenue QTD =
    CALCULATE(
        [Total Revenue],
        DATESQTD('Date'[Date])
    )
```

### Month-over-Month Growth
```dax
// Revenue Previous Month
Revenue PM =
    CALCULATE(
        [Total Revenue],
        DATEADD('Date'[Date], -1, MONTH)
    )

// MoM Growth %
MoM Growth % =
    DIVIDE(
        [Total Revenue] - [Revenue PM],
        [Revenue PM],
        BLANK()
    )

// MoM Trend Icon
MoM Icon =
    SWITCH(
        TRUE(),
        [MoM Growth %] > 0,  "▲ " & FORMAT([MoM Growth %], "0.0%"),
        [MoM Growth %] < 0,  "▼ " & FORMAT(ABS([MoM Growth %]), "0.0%"),
        "→ Flat"
    )
```

### Salesperson Ranking
```dax
// Salesperson Revenue Rank (mirrors SQL RANK())
Salesperson Rank =
    RANKX(
        ALLSELECTED(salespersons[name]),
        [Total Revenue],
        ,
        DESC,
        DENSE
    )

// Top N Filter (use in visual-level filter)
Is Top 5 Salesperson =
    [Salesperson Rank] <= 5
```

### Regional Analysis
```dax
// Region Revenue Share %
Region Revenue Share =
    DIVIDE(
        [Total Revenue],
        CALCULATE([Total Revenue], ALL(sales[region])),
        0
    )

// Regional Rank
Region Rank =
    RANKX(
        ALLSELECTED(sales[region]),
        [Total Revenue],
        ,
        DESC,
        DENSE
    )
```

---

## 4. Report Pages & Visuals

### Page 1 — Executive Summary
| Visual                  | Type              | Fields                               |
|-------------------------|-------------------|--------------------------------------|
| Total Revenue Card      | Card              | [Total Revenue]                      |
| YoY Growth Card         | Card              | [YoY Growth %]                       |
| Orders Card             | Card              | [Total Orders]                       |
| Monthly Revenue Trend   | Line + Column     | Date[Month], [Total Revenue], [YTD]  |
| Revenue by Category     | Donut Chart       | products[category], [Total Revenue]  |
| Regional Heatmap        | Filled Map        | sales[region], [Total Revenue]       |

### Page 2 — Salesperson Leaderboard
| Visual                  | Type              | Fields                               |
|-------------------------|-------------------|--------------------------------------|
| Top 10 Table            | Table             | Name, Revenue, Orders, Rank, MoM%    |
| Revenue by Person       | Bar Chart         | salespersons[name], [Total Revenue]  |
| MoM Waterfall           | Waterfall Chart   | Date[Month], [MoM Growth %]          |

### Page 3 — Product Deep Dive
| Visual                  | Type              | Fields                               |
|-------------------------|-------------------|--------------------------------------|
| Category Breakdown      | Tree Map          | Category, Product, Revenue           |
| Top Products Table      | Table             | Product, Revenue, Units, Rank        |
| Revenue vs Orders       | Scatter Chart     | Revenue, Orders (per product)        |

### Page 4 — Regional Analysis
| Visual                  | Type              | Fields                               |
|-------------------------|-------------------|--------------------------------------|
| Region Scorecard        | Matrix            | Region, Revenue, Orders, Share%      |
| Regional Trend Lines    | Line Chart        | Month, Revenue (per region)          |
| Region vs Target        | Gauge             | [Total Revenue], Target = $1.5M      |

---

## 5. Row-Level Security (RLS)

```dax
// Regional managers only see their region
[region] = LOOKUPVALUE(
    salespersons[region],
    salespersons[email],
    USERPRINCIPALNAME()
)
```

---

## 6. Scheduled Refresh Setup
- Power BI Service → Dataset Settings → Scheduled Refresh
- Frequency : Daily at 06:00 AM IST
- Gateway   : On-premises data gateway (connected to MySQL)
- Alerts    : Email notification if refresh fails

---

## 7. Performance Optimization Tips
1. Import mode preferred over DirectQuery for 10K records
2. Disable auto date/time tables — use a custom Date table
3. Mark Date table as official Date Table
4. Use DIVIDE() instead of "/" to avoid division-by-zero errors
5. Avoid FILTER(ALL(...)) inside CALCULATE — use REMOVEFILTERS() instead
6. Reduce cardinality: compress region/category to integer FK in model
