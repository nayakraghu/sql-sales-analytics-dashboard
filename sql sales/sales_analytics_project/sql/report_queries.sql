-- ============================================================
--  SQL Sales Analytics — Report Queries & Views
--  Run AFTER schema_and_seed.sql and generate_data.py
-- ============================================================

USE sales_analytics;

-- ── VIEW 1: Full sales detail (joins all tables) ─────────────
CREATE OR REPLACE VIEW v_sales_detail AS
SELECT
    s.id,
    s.order_id,
    sp.name          AS salesperson,
    sp.region        AS salesperson_region,
    p.product_name,
    p.category,
    s.quantity,
    s.unit_price,
    s.discount_pct,
    s.revenue,
    s.order_date,
    s.region,
    s.status,
    YEAR(s.order_date)                          AS order_year,
    MONTH(s.order_date)                         AS order_month,
    DATE_FORMAT(s.order_date, '%Y-%m')          AS year_month,
    DATE_FORMAT(s.order_date, '%b %Y')          AS month_label
FROM  sales s
JOIN  salespersons sp ON s.salesperson_id = sp.id
JOIN  products     p  ON s.product_id     = p.id;


-- ── VIEW 2: Monthly KPI summary ──────────────────────────────
CREATE OR REPLACE VIEW v_monthly_kpi AS
WITH monthly AS (
    SELECT
        year_month,
        month_label,
        SUM(revenue)  AS monthly_revenue,
        COUNT(*)      AS orders,
        AVG(revenue)  AS avg_order_value
    FROM  v_sales_detail
    WHERE status = 'Completed'
    GROUP BY year_month, month_label
)
SELECT
    year_month,
    month_label,
    monthly_revenue,
    orders,
    ROUND(avg_order_value, 2)                                AS avg_order_value,
    SUM(monthly_revenue) OVER (
        PARTITION BY LEFT(year_month, 4)
        ORDER BY year_month
        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    )                                                        AS ytd_revenue,
    ROUND(
        (monthly_revenue
         - LAG(monthly_revenue) OVER (ORDER BY year_month))
        / NULLIF(LAG(monthly_revenue) OVER (ORDER BY year_month), 0) * 100,
    2)                                                       AS mom_growth_pct
FROM monthly;


-- ── VIEW 3: Salesperson leaderboard ──────────────────────────
CREATE OR REPLACE VIEW v_salesperson_rank AS
SELECT
    salesperson,
    salesperson_region,
    SUM(revenue)                                             AS total_revenue,
    COUNT(*)                                                 AS total_orders,
    ROUND(AVG(revenue), 2)                                   AS avg_order_value,
    RANK() OVER (ORDER BY SUM(revenue) DESC)                 AS revenue_rank,
    ROUND(SUM(revenue) * 100.0
          / SUM(SUM(revenue)) OVER (), 2)                    AS pct_of_total
FROM  v_sales_detail
WHERE status = 'Completed'
  AND order_year = YEAR(CURDATE())
GROUP BY salesperson, salesperson_region;


-- ── REPORT QUERY 1: Executive dashboard feed ─────────────────
SELECT
    SUM(CASE WHEN order_year = YEAR(CURDATE())     THEN revenue END) AS current_year_revenue,
    SUM(CASE WHEN order_year = YEAR(CURDATE()) - 1 THEN revenue END) AS last_year_revenue,
    COUNT(CASE WHEN order_year = YEAR(CURDATE())   THEN 1 END)       AS current_year_orders,
    ROUND(AVG(CASE WHEN order_year = YEAR(CURDATE()) THEN revenue END), 2) AS current_aov
FROM v_sales_detail
WHERE status = 'Completed';


-- ── REPORT QUERY 2: Regional summary ────────────────────────
SELECT
    region,
    SUM(revenue)                                         AS total_revenue,
    COUNT(*)                                             AS total_orders,
    ROUND(AVG(revenue), 2)                               AS avg_order_value,
    RANK() OVER (ORDER BY SUM(revenue) DESC)             AS region_rank,
    ROUND(SUM(revenue) * 100.0
          / SUM(SUM(revenue)) OVER (), 2)                AS share_pct
FROM  v_sales_detail
WHERE status = 'Completed'
  AND order_year = YEAR(CURDATE())
GROUP BY region
ORDER BY region_rank;


-- ── REPORT QUERY 3: Category performance ────────────────────
SELECT
    category,
    product_name,
    SUM(revenue)                                         AS total_revenue,
    COUNT(*)                                             AS units_sold,
    DENSE_RANK() OVER (
        PARTITION BY category
        ORDER BY SUM(revenue) DESC
    )                                                    AS rank_in_category
FROM  v_sales_detail
WHERE status = 'Completed'
GROUP BY category, product_name
ORDER BY category, rank_in_category;


-- ── REPORT QUERY 4: Top 10 orders this month ────────────────
SELECT
    order_id,
    salesperson,
    product_name,
    region,
    revenue,
    order_date,
    ROW_NUMBER() OVER (ORDER BY revenue DESC)            AS this_month_rank
FROM  v_sales_detail
WHERE status     = 'Completed'
  AND year_month = DATE_FORMAT(CURDATE(), '%Y-%m')
ORDER BY revenue DESC
LIMIT 10;
