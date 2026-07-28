-- ============================================================
--  SQL Sales Analytics — Window Function Queries
--  Database : sales_analytics (MySQL 8.0+)
--  Purpose  : Analyse 10,247+ sales records using window fns
-- ============================================================

USE sales_analytics;

-- ────────────────────────────────────────────────────────────
-- QUERY 1 : Running Revenue Total  (SUM OVER)
-- Shows cumulative revenue within each year, row by row
-- ────────────────────────────────────────────────────────────
SELECT
    s.order_id,
    sp.name                                          AS salesperson,
    p.product_name,
    s.order_date,
    s.revenue,
    SUM(s.revenue) OVER (
        PARTITION BY YEAR(s.order_date)
        ORDER BY     s.order_date
        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    )                                                AS ytd_running_total,
    SUM(s.revenue) OVER (
        PARTITION BY YEAR(s.order_date), MONTH(s.order_date)
    )                                                AS month_total
FROM  sales s
JOIN  salespersons sp ON s.salesperson_id = sp.id
JOIN  products     p  ON s.product_id     = p.id
WHERE s.status = 'Completed'
ORDER BY s.order_date;


-- ────────────────────────────────────────────────────────────
-- QUERY 2 : Salesperson Revenue Ranking  (RANK / DENSE_RANK)
-- Ranks every salesperson; ties share rank with RANK(),
-- no gaps with DENSE_RANK()
-- ────────────────────────────────────────────────────────────
SELECT
    sp.name                                          AS salesperson,
    sp.region,
    SUM(s.revenue)                                   AS total_revenue,
    COUNT(*)                                         AS total_orders,
    ROUND(AVG(s.revenue), 2)                         AS avg_order_value,
    RANK()       OVER (ORDER BY SUM(s.revenue) DESC) AS revenue_rank,
    DENSE_RANK() OVER (ORDER BY SUM(s.revenue) DESC) AS dense_rank,
    ROUND(
        SUM(s.revenue) * 100.0
        / SUM(SUM(s.revenue)) OVER (),
    2)                                               AS pct_of_total
FROM  sales s
JOIN  salespersons sp ON s.salesperson_id = sp.id
WHERE s.status = 'Completed'
  AND YEAR(s.order_date) = 2024
GROUP BY sp.id, sp.name, sp.region
ORDER BY revenue_rank;


-- ────────────────────────────────────────────────────────────
-- QUERY 3 : Month-over-Month Growth  (LAG)
-- Compares each month's revenue to the previous month
-- ────────────────────────────────────────────────────────────
WITH monthly AS (
    SELECT
        DATE_FORMAT(order_date, '%Y-%m')             AS report_month,
        SUM(revenue)                                 AS monthly_revenue,
        COUNT(*)                                     AS orders
    FROM  sales
    WHERE status = 'Completed'
    GROUP BY report_month
)
SELECT
    report_month,
    monthly_revenue,
    orders,
    LAG(monthly_revenue) OVER (ORDER BY report_month) AS prev_month_revenue,
    ROUND(
        (monthly_revenue - LAG(monthly_revenue) OVER (ORDER BY report_month))
        / NULLIF(LAG(monthly_revenue) OVER (ORDER BY report_month), 0) * 100,
    2)                                               AS mom_growth_pct,
    CASE
        WHEN monthly_revenue > LAG(monthly_revenue) OVER (ORDER BY report_month) THEN 'Growth'
        WHEN monthly_revenue < LAG(monthly_revenue) OVER (ORDER BY report_month) THEN 'Decline'
        ELSE 'Flat'
    END                                              AS trend
FROM monthly
ORDER BY report_month;


-- ────────────────────────────────────────────────────────────
-- QUERY 4 : Top Product per Category  (DENSE_RANK PARTITION)
-- Ranks products within each category independently
-- ────────────────────────────────────────────────────────────
WITH product_rank AS (
    SELECT
        p.category,
        p.product_name,
        SUM(s.revenue)                               AS total_revenue,
        COUNT(*)                                     AS units_sold,
        DENSE_RANK() OVER (
            PARTITION BY p.category
            ORDER BY     SUM(s.revenue) DESC
        )                                            AS rank_in_category
    FROM  sales s
    JOIN  products p ON s.product_id = p.id
    WHERE s.status = 'Completed'
    GROUP BY p.category, p.product_name
)
SELECT *
FROM  product_rank
WHERE rank_in_category <= 3          -- top 3 per category
ORDER BY category, rank_in_category;


-- ────────────────────────────────────────────────────────────
-- QUERY 5 : 3-Month Moving Average Revenue  (AVG OVER ROWS)
-- Smooths out monthly fluctuations for trend analysis
-- ────────────────────────────────────────────────────────────
WITH monthly AS (
    SELECT
        DATE_FORMAT(order_date, '%Y-%m') AS report_month,
        SUM(revenue)                     AS monthly_revenue
    FROM  sales
    WHERE status = 'Completed'
    GROUP BY report_month
)
SELECT
    report_month,
    monthly_revenue,
    ROUND(AVG(monthly_revenue) OVER (
        ORDER BY report_month
        ROWS BETWEEN 2 PRECEDING AND CURRENT ROW
    ), 2)                              AS moving_avg_3m,
    ROUND(AVG(monthly_revenue) OVER (
        ORDER BY report_month
        ROWS BETWEEN 5 PRECEDING AND CURRENT ROW
    ), 2)                              AS moving_avg_6m
FROM monthly
ORDER BY report_month;


-- ────────────────────────────────────────────────────────────
-- QUERY 6 : Regional Leaderboard with ROW_NUMBER
-- Unique sequential rank per region — no ties
-- ────────────────────────────────────────────────────────────
WITH regional AS (
    SELECT
        s.region,
        sp.name                                      AS salesperson,
        SUM(s.revenue)                               AS revenue,
        ROW_NUMBER() OVER (
            PARTITION BY s.region
            ORDER BY     SUM(s.revenue) DESC
        )                                            AS rank_in_region
    FROM  sales s
    JOIN  salespersons sp ON s.salesperson_id = sp.id
    WHERE s.status = 'Completed'
      AND YEAR(s.order_date) = 2024
    GROUP BY s.region, sp.id, sp.name
)
SELECT *
FROM  regional
WHERE rank_in_region = 1             -- #1 salesperson per region
ORDER BY revenue DESC;


-- ────────────────────────────────────────────────────────────
-- QUERY 7 : Next-Period Forecast using LEAD
-- Shows each month alongside the following month's revenue
-- ────────────────────────────────────────────────────────────
WITH monthly AS (
    SELECT
        DATE_FORMAT(order_date, '%Y-%m') AS report_month,
        SUM(revenue)                     AS monthly_revenue
    FROM  sales
    WHERE status = 'Completed'
    GROUP BY report_month
)
SELECT
    report_month,
    monthly_revenue,
    LEAD(monthly_revenue) OVER (ORDER BY report_month) AS next_month_revenue,
    ROUND(
        (LEAD(monthly_revenue) OVER (ORDER BY report_month) - monthly_revenue)
        / NULLIF(monthly_revenue, 0) * 100,
    2)                                                  AS projected_growth_pct
FROM monthly
ORDER BY report_month;


-- ────────────────────────────────────────────────────────────
-- QUERY 8 : Executive KPI Summary (used for Power BI)
-- Single query feeding the main dashboard cards
-- ────────────────────────────────────────────────────────────
SELECT
    -- Current year totals
    SUM(CASE WHEN YEAR(order_date) = 2024 THEN revenue ELSE 0 END)  AS revenue_2024,
    COUNT(CASE WHEN YEAR(order_date) = 2024 THEN 1 END)             AS orders_2024,
    AVG(CASE WHEN YEAR(order_date) = 2024 THEN revenue END)         AS aov_2024,

    -- Previous year totals
    SUM(CASE WHEN YEAR(order_date) = 2023 THEN revenue ELSE 0 END)  AS revenue_2023,
    COUNT(CASE WHEN YEAR(order_date) = 2023 THEN 1 END)             AS orders_2023,

    -- YoY growth
    ROUND(
        (SUM(CASE WHEN YEAR(order_date) = 2024 THEN revenue ELSE 0 END)
         - SUM(CASE WHEN YEAR(order_date) = 2023 THEN revenue ELSE 0 END))
        / NULLIF(SUM(CASE WHEN YEAR(order_date) = 2023 THEN revenue ELSE 0 END), 0) * 100,
    2)                                                               AS yoy_growth_pct
FROM sales
WHERE status = 'Completed';
