-- ============================================================
--  SQL Sales Analytics Project — Database Schema & Seed Data
--  Tool   : MySQL 8.0+
--  Author : Sales Analytics Team
--  Records: 10,247 (seed shows representative sample)
-- ============================================================

-- ── 1. CREATE DATABASE ───────────────────────────────────────
CREATE DATABASE IF NOT EXISTS sales_analytics;
USE sales_analytics;

-- ── 2. TABLES ────────────────────────────────────────────────

-- Salesperson master
CREATE TABLE IF NOT EXISTS salespersons (
    id            INT           AUTO_INCREMENT PRIMARY KEY,
    name          VARCHAR(100)  NOT NULL,
    email         VARCHAR(150)  UNIQUE NOT NULL,
    region        ENUM('North','South','East','West','Central') NOT NULL,
    hire_date     DATE          NOT NULL,
    is_active     BOOLEAN       DEFAULT TRUE,
    created_at    TIMESTAMP     DEFAULT CURRENT_TIMESTAMP
);

-- Product / category master
CREATE TABLE IF NOT EXISTS products (
    id            INT           AUTO_INCREMENT PRIMARY KEY,
    product_name  VARCHAR(150)  NOT NULL,
    category      ENUM('Software','Hardware','Services','Accessories') NOT NULL,
    unit_price    DECIMAL(10,2) NOT NULL,
    is_active     BOOLEAN       DEFAULT TRUE,
    created_at    TIMESTAMP     DEFAULT CURRENT_TIMESTAMP
);

-- Main sales fact table  (10,247 rows in production)
CREATE TABLE IF NOT EXISTS sales (
    id              INT             AUTO_INCREMENT PRIMARY KEY,
    order_id        VARCHAR(20)     UNIQUE NOT NULL,
    salesperson_id  INT             NOT NULL,
    product_id      INT             NOT NULL,
    quantity        INT             NOT NULL DEFAULT 1,
    unit_price      DECIMAL(10,2)   NOT NULL,
    discount_pct    DECIMAL(5,2)    DEFAULT 0.00,
    revenue         DECIMAL(12,2)   GENERATED ALWAYS AS
                        (quantity * unit_price * (1 - discount_pct/100)) STORED,
    order_date      DATE            NOT NULL,
    region          ENUM('North','South','East','West','Central') NOT NULL,
    status          ENUM('Completed','Pending','Cancelled') DEFAULT 'Completed',
    created_at      TIMESTAMP       DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (salesperson_id) REFERENCES salespersons(id),
    FOREIGN KEY (product_id)     REFERENCES products(id),
    INDEX idx_order_date   (order_date),
    INDEX idx_region       (region),
    INDEX idx_salesperson  (salesperson_id)
);

-- Monthly report cache  (populated by stored procedure)
CREATE TABLE IF NOT EXISTS monthly_report_cache (
    id              INT           AUTO_INCREMENT PRIMARY KEY,
    report_month    VARCHAR(7)    NOT NULL,   -- e.g. '2024-12'
    total_revenue   DECIMAL(14,2),
    total_orders    INT,
    avg_order_value DECIMAL(10,2),
    top_salesperson VARCHAR(100),
    top_product     VARCHAR(150),
    mom_growth_pct  DECIMAL(6,2),
    generated_at    TIMESTAMP     DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uq_month (report_month)
);

-- ── 3. SEED DATA ─────────────────────────────────────────────

INSERT INTO salespersons (name, email, region, hire_date) VALUES
('Alice Chen',  'alice.chen@company.com',  'North',   '2020-03-15'),
('Raj Verma',   'raj.verma@company.com',   'West',    '2019-07-01'),
('Carol Singh', 'carol.singh@company.com', 'East',    '2021-01-10'),
('Bob Patel',   'bob.patel@company.com',   'South',   '2020-11-22'),
('Diana Luo',   'diana.luo@company.com',   'North',   '2022-04-05'),
('Evan Kim',    'evan.kim@company.com',    'East',    '2021-09-14'),
('Fatima Nair', 'fatima.nair@company.com', 'Central', '2023-02-28'),
('George Tan',  'george.tan@company.com',  'West',    '2019-05-17');

INSERT INTO products (product_name, category, unit_price) VALUES
('Enterprise Suite',    'Software',     2400.00),
('Pro Laptop X1',       'Hardware',      850.00),
('Annual Support',      'Services',      230.00),
('Smart Monitor 4K',    'Hardware',      320.00),
('Cloud Storage Pro',   'Software',      140.00),
('Wireless Headset',    'Accessories',    75.00),
('Dev Toolkit License', 'Software',      580.00),
('On-site Training',    'Services',     1200.00);

-- Representative sample of sales records
INSERT INTO sales (order_id, salesperson_id, product_id, quantity, unit_price, discount_pct, order_date, region) VALUES
('ORD-4821', 1, 1, 1, 2400.00,  0.00, '2024-12-18', 'North'),
('ORD-4820', 2, 2, 1,  850.00,  5.00, '2024-12-17', 'West'),
('ORD-4819', 3, 3, 5,  230.00,  0.00, '2024-12-16', 'East'),
('ORD-4818', 4, 4, 2,  320.00, 10.00, '2024-12-15', 'South'),
('ORD-4817', 5, 1, 1, 2400.00,  0.00, '2024-12-14', 'North'),
('ORD-4816', 1, 5, 3,  140.00,  0.00, '2024-12-13', 'West'),
('ORD-4815', 2, 2, 1,  850.00,  0.00, '2024-12-12', 'Central'),
('ORD-4814', 3, 3, 4,  230.00,  0.00, '2024-12-11', 'North'),
('ORD-4813', 6, 1, 1, 2400.00,  0.00, '2024-12-10', 'East'),
('ORD-4812', 4, 4, 1,  320.00,  5.00, '2024-12-09', 'South'),
('ORD-4001', 1, 7, 2,  580.00,  0.00, '2024-01-05', 'North'),
('ORD-4002', 2, 8, 1, 1200.00,  0.00, '2024-01-12', 'West'),
('ORD-4003', 3, 3, 6,  230.00,  5.00, '2024-01-18', 'East'),
('ORD-4004', 5, 1, 1, 2400.00, 10.00, '2024-02-03', 'North'),
('ORD-4005', 6, 2, 2,  850.00,  0.00, '2024-02-14', 'East'),
('ORD-4006', 7, 5, 4,  140.00,  0.00, '2024-03-01', 'Central'),
('ORD-4007', 8, 4, 3,  320.00,  0.00, '2024-03-19', 'West'),
('ORD-4008', 1, 1, 1, 2400.00,  5.00, '2024-04-07', 'North'),
('ORD-4009', 2, 6, 10,  75.00,  0.00, '2024-04-22', 'West'),
('ORD-4010', 3, 8, 1, 1200.00,  0.00, '2024-05-11', 'East');
-- (10,227 more rows exist in production DB — generated via faker scripts)

-- ── 4. STORED PROCEDURE: Auto Monthly Report ─────────────────
DELIMITER $$

CREATE PROCEDURE IF NOT EXISTS GenerateMonthlyReport(IN p_month VARCHAR(7))
BEGIN
    DECLARE v_revenue     DECIMAL(14,2);
    DECLARE v_orders      INT;
    DECLARE v_aov         DECIMAL(10,2);
    DECLARE v_top_sp      VARCHAR(100);
    DECLARE v_top_prod    VARCHAR(150);
    DECLARE v_prev_rev    DECIMAL(14,2);
    DECLARE v_growth      DECIMAL(6,2);

    -- Aggregate for the given month
    SELECT SUM(revenue), COUNT(*), AVG(revenue)
    INTO   v_revenue, v_orders, v_aov
    FROM   sales
    WHERE  DATE_FORMAT(order_date, '%Y-%m') = p_month
      AND  status = 'Completed';

    -- Top salesperson
    SELECT sp.name INTO v_top_sp
    FROM   sales s JOIN salespersons sp ON s.salesperson_id = sp.id
    WHERE  DATE_FORMAT(s.order_date, '%Y-%m') = p_month
    GROUP  BY sp.name
    ORDER  BY SUM(s.revenue) DESC
    LIMIT  1;

    -- Top product
    SELECT p.product_name INTO v_top_prod
    FROM   sales s JOIN products p ON s.product_id = p.id
    WHERE  DATE_FORMAT(s.order_date, '%Y-%m') = p_month
    GROUP  BY p.product_name
    ORDER  BY SUM(s.revenue) DESC
    LIMIT  1;

    -- MoM growth
    SELECT SUM(revenue) INTO v_prev_rev
    FROM   sales
    WHERE  DATE_FORMAT(order_date, '%Y-%m') =
               DATE_FORMAT(DATE_SUB(STR_TO_DATE(CONCAT(p_month,'-01'),'%Y-%m-%d'), INTERVAL 1 MONTH), '%Y-%m')
      AND  status = 'Completed';

    SET v_growth = IF(v_prev_rev > 0,
                      ROUND((v_revenue - v_prev_rev) / v_prev_rev * 100, 2),
                      NULL);

    -- Upsert into cache
    INSERT INTO monthly_report_cache
        (report_month, total_revenue, total_orders, avg_order_value,
         top_salesperson, top_product, mom_growth_pct)
    VALUES
        (p_month, v_revenue, v_orders, v_aov, v_top_sp, v_top_prod, v_growth)
    ON DUPLICATE KEY UPDATE
        total_revenue   = v_revenue,
        total_orders    = v_orders,
        avg_order_value = v_aov,
        top_salesperson = v_top_sp,
        top_product     = v_top_prod,
        mom_growth_pct  = v_growth,
        generated_at    = CURRENT_TIMESTAMP;

    SELECT CONCAT('Report for ', p_month, ' generated successfully.') AS result;
END$$

DELIMITER ;

-- ── 5. EVENT: Auto-run report on 1st of each month ───────────
SET GLOBAL event_scheduler = ON;

CREATE EVENT IF NOT EXISTS evt_monthly_report
ON SCHEDULE EVERY 1 MONTH
STARTS '2024-02-01 00:05:00'
DO CALL GenerateMonthlyReport(DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 MONTH), '%Y-%m'));
