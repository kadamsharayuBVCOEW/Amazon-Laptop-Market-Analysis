-- ═══════════════════════════════════════════════════════════════════
--   E-COMMERCE SALES ANALYTICS — PHASE 3: SQL
--   Tool     : MySQL 8.0+ / MySQL Workbench
--   Dataset  : Amazon.in Laptop Scrape (1,226 rows → 1,152 cleaned)
--   Author   : [Your Name]
--   Date     : June 2026
--
--   Pipeline : Python (Scrape) → Excel (Store) → MySQL (Clean+Analyze) → Power BI (Dashboard)
--
--   SECTIONS:
--     PART A — Database & Table Setup
--     PART B — Data Import
--     PART C — Data Cleaning
--     PART D — Business Insight Queries (10 queries)
--     PART E — Views for Power BI (6 views)
-- ═══════════════════════════════════════════════════════════════════


-- ─────────────────────────────────────────────────────────────────
-- PART A: DATABASE & TABLE SETUP
-- ─────────────────────────────────────────────────────────────────

CREATE DATABASE IF NOT EXISTS ecommerce_analytics;
USE ecommerce_analytics;

DROP TABLE IF EXISTS laptops_raw;

CREATE TABLE laptops_raw (
    id                  INT AUTO_INCREMENT PRIMARY KEY,
    asin                VARCHAR(20),
    product_name        VARCHAR(500),
    brand               VARCHAR(100),
    price_inr           DECIMAL(10,2),
    original_price_inr  DECIMAL(10,2),
    discount_percent    DECIMAL(5,2),
    rating              DECIMAL(3,1),
    total_reviews       INT,
    listing_type        VARCHAR(20),
    search_keyword      VARCHAR(100),
    page_scraped        INT,
    scraped_at          DATETIME
);


-- ─────────────────────────────────────────────────────────────────
-- PART B: DATA IMPORT
-- ─────────────────────────────────────────────────────────────────

SET GLOBAL local_infile = 1;

LOAD DATA LOCAL INFILE 'D:/amazon_laptops_raw.csv'
INTO TABLE laptops_raw
FIELDS TERMINATED BY ','
ENCLOSED BY '"'
LINES TERMINATED BY '\r\n'
IGNORE 1 ROWS
(asin, product_name, brand, price_inr, original_price_inr,
 discount_percent, rating, total_reviews, listing_type,
 search_keyword, page_scraped, scraped_at);

-- Verify import
SELECT COUNT(*) AS total_rows FROM laptops_raw;


-- ─────────────────────────────────────────────────────────────────
-- PART C: DATA CLEANING
-- ─────────────────────────────────────────────────────────────────

-- Step 1: Check raw data quality before cleaning
SELECT
    COUNT(*)                        AS total_records,
    COUNT(DISTINCT asin)            AS unique_products,
    SUM(price_inr IS NULL)          AS missing_price,
    SUM(rating IS NULL)             AS missing_rating,
    SUM(total_reviews IS NULL)      AS missing_reviews,
    SUM(price_inr < 5000)           AS suspicious_low_price,
    SUM(price_inr > 500000)         AS suspicious_high_price
FROM laptops_raw;

-- Step 2: Check duplicate ASINs
SELECT asin, COUNT(*) AS occurrences
FROM laptops_raw
GROUP BY asin
HAVING COUNT(*) > 1
ORDER BY occurrences DESC
LIMIT 10;

-- Step 3: Create cleaned table with all transformations
DROP TABLE IF EXISTS laptops_cleaned;

CREATE TABLE laptops_cleaned AS
SELECT
    id, asin,
    TRIM(product_name) AS product_name,

    -- Standardize brand names
    CASE
        WHEN UPPER(product_name) LIKE '%MACBOOK%'    THEN 'Apple'
        WHEN UPPER(product_name) LIKE '%THINKPAD%'   THEN 'Lenovo'
        WHEN UPPER(product_name) LIKE '%IDEAPAD%'    THEN 'Lenovo'
        WHEN UPPER(product_name) LIKE '%INSPIRON%'   THEN 'Dell'
        WHEN UPPER(product_name) LIKE '%PAVILION%'   THEN 'HP'
        WHEN UPPER(product_name) LIKE '%ELITEBOOK%'  THEN 'HP'
        WHEN UPPER(product_name) LIKE '%VIVOBOOK%'   THEN 'ASUS'
        WHEN UPPER(product_name) LIKE '%ZENBOOK%'    THEN 'ASUS'
        WHEN UPPER(brand) LIKE '%LENOVO%'            THEN 'Lenovo'
        WHEN UPPER(brand) LIKE '%HP%'                THEN 'HP'
        WHEN UPPER(brand) LIKE '%DELL%'              THEN 'Dell'
        WHEN UPPER(brand) LIKE '%ASUS%'              THEN 'ASUS'
        WHEN UPPER(brand) LIKE '%ACER%'              THEN 'Acer'
        WHEN UPPER(brand) LIKE '%APPLE%'             THEN 'Apple'
        WHEN UPPER(brand) LIKE '%MSI%'               THEN 'MSI'
        WHEN UPPER(brand) LIKE '%SAMSUNG%'           THEN 'Samsung'
        WHEN UPPER(brand) LIKE '%MICROSOFT%'         THEN 'Microsoft'
        WHEN UPPER(brand) LIKE '%ALIENWARE%'         THEN 'Alienware'
        WHEN UPPER(brand) LIKE '%AVITA%'             THEN 'Avita'
        WHEN brand REGEXP '^[0-9]'                   THEN 'Other'
        ELSE TRIM(brand)
    END AS brand,

    price_inr,
    original_price_inr,

    -- Recalculate discount % from actual prices (more accurate than scraped value)
    CASE
        WHEN original_price_inr > price_inr
        THEN ROUND(((original_price_inr - price_inr) / original_price_inr) * 100, 1)
        ELSE NULL
    END AS discount_percent,

    -- Savings in INR
    CASE
        WHEN original_price_inr > price_inr
        THEN ROUND(original_price_inr - price_inr, 2)
        ELSE NULL
    END AS savings_inr,

    -- Treat rating 0 as NULL (means unrated, not zero stars)
    CASE
        WHEN rating > 0 AND rating <= 5 THEN rating
        ELSE NULL
    END AS rating,

    COALESCE(total_reviews, 0) AS total_reviews,

    -- Business categorization by price
    CASE
        WHEN price_inr < 30000                  THEN 'Budget (Under 30K)'
        WHEN price_inr BETWEEN 30000 AND 50000  THEN 'Mid-Range (30K-50K)'
        WHEN price_inr BETWEEN 50001 AND 80000  THEN 'Premium (50K-80K)'
        WHEN price_inr BETWEEN 80001 AND 120000 THEN 'High-End (80K-1.2L)'
        WHEN price_inr > 120000                 THEN 'Ultra-Premium (Above 1.2L)'
    END AS price_segment,

    -- Discount tier categorization
    CASE
        WHEN discount_percent IS NULL   THEN 'No Discount'
        WHEN discount_percent < 10      THEN 'Low (Under 10%)'
        WHEN discount_percent < 25      THEN 'Moderate (10-24%)'
        WHEN discount_percent < 40      THEN 'High (25-39%)'
        ELSE                                 'Mega Deal (40%+)'
    END AS discount_tier,

    -- Rating category
    CASE
        WHEN rating IS NULL  THEN 'Not Rated'
        WHEN rating >= 4.5   THEN 'Excellent (4.5+)'
        WHEN rating >= 4.0   THEN 'Good (4.0-4.4)'
        WHEN rating >= 3.5   THEN 'Average (3.5-3.9)'
        ELSE                      'Below Average'
    END AS rating_category,

    listing_type,
    search_keyword,
    scraped_at

FROM laptops_raw
WHERE
    price_inr IS NOT NULL
    AND price_inr > 5000
    AND price_inr < 500000
    AND product_name IS NOT NULL;

-- Step 4: Verify cleaning results
SELECT
    (SELECT COUNT(*) FROM laptops_raw)     AS raw_records,
    (SELECT COUNT(*) FROM laptops_cleaned) AS cleaned_records,
    (SELECT COUNT(*) FROM laptops_raw) -
    (SELECT COUNT(*) FROM laptops_cleaned) AS rows_removed;


-- ─────────────────────────────────────────────────────────────────
-- PART D: BUSINESS INSIGHT QUERIES
-- ─────────────────────────────────────────────────────────────────

-- Q1: Which brand has the most products?
SELECT brand, COUNT(*) AS total_products,
    ROUND(AVG(price_inr), 0) AS avg_price,
    ROUND(AVG(rating), 2) AS avg_rating
FROM laptops_cleaned
WHERE brand != 'Other'
GROUP BY brand
ORDER BY total_products DESC
LIMIT 10;

-- Q2: Which price segment is most popular?
SELECT price_segment, COUNT(*) AS total_products,
    ROUND(AVG(rating), 2) AS avg_rating,
    ROUND(AVG(discount_percent), 1) AS avg_discount
FROM laptops_cleaned
GROUP BY price_segment
ORDER BY total_products DESC;

-- Q3: Top 10 highest rated laptops
SELECT product_name, brand, price_inr, rating
FROM laptops_cleaned
WHERE rating IS NOT NULL
ORDER BY rating DESC, price_inr ASC
LIMIT 10;

-- Q4: Which brand gives highest average discount?
SELECT brand,
    ROUND(AVG(discount_percent), 1) AS avg_discount,
    ROUND(AVG(savings_inr), 0) AS avg_savings_inr,
    COUNT(*) AS products
FROM laptops_cleaned
WHERE discount_percent IS NOT NULL
AND brand != 'Other'
GROUP BY brand
HAVING COUNT(*) >= 5
ORDER BY avg_discount DESC
LIMIT 10;

-- Q5: Best value laptops (high rating + high discount)
SELECT product_name, brand, price_inr,
    discount_percent, rating
FROM laptops_cleaned
WHERE rating >= 4.0
    AND discount_percent >= 15
ORDER BY discount_percent DESC
LIMIT 10;

-- Q6: Rating distribution across all products
SELECT rating_category,
    COUNT(*) AS products,
    ROUND(COUNT(*) * 100.0 /
        (SELECT COUNT(*) FROM laptops_cleaned), 1) AS percentage
FROM laptops_cleaned
GROUP BY rating_category
ORDER BY products DESC;

-- Q7: Do expensive laptops get better ratings?
SELECT price_segment,
    COUNT(*) AS products,
    ROUND(AVG(rating), 2) AS avg_rating,
    ROUND(AVG(discount_percent), 1) AS avg_discount
FROM laptops_cleaned
WHERE rating IS NOT NULL
GROUP BY price_segment
ORDER BY AVG(price_inr);

-- Q8: Sponsored vs Organic listing performance
SELECT listing_type,
    COUNT(*) AS total_products,
    ROUND(AVG(price_inr), 0) AS avg_price,
    ROUND(AVG(rating), 2) AS avg_rating,
    ROUND(AVG(discount_percent), 1) AS avg_discount
FROM laptops_cleaned
GROUP BY listing_type;

-- Q9: Which search keyword returned best quality products?
SELECT search_keyword,
    COUNT(*) AS products,
    ROUND(AVG(price_inr), 0) AS avg_price,
    ROUND(AVG(rating), 2) AS avg_rating,
    ROUND(AVG(discount_percent), 1) AS avg_discount
FROM laptops_cleaned
GROUP BY search_keyword
ORDER BY avg_rating DESC;

-- Q10: Brand market positioning across price segments
SELECT brand, price_segment, COUNT(*) AS products
FROM laptops_cleaned
WHERE brand IN ('Lenovo','HP','ASUS','Acer','Dell','Apple','MSI')
GROUP BY brand, price_segment
ORDER BY brand, AVG(price_inr);


-- ─────────────────────────────────────────────────────────────────
-- PART E: VIEWS FOR POWER BI
-- These 6 views connect directly to Power BI dashboard
-- ─────────────────────────────────────────────────────────────────

-- View 1: Brand Performance Summary
CREATE OR REPLACE VIEW vw_brand_performance AS
SELECT
    brand,
    COUNT(*) AS total_products,
    ROUND(AVG(price_inr), 0) AS avg_price,
    ROUND(COALESCE(AVG(rating), 0), 2) AS avg_rating,
    ROUND(COALESCE(AVG(discount_percent), 0), 1) AS avg_discount,
    ROUND(COALESCE(AVG(savings_inr), 0), 0) AS avg_savings
FROM laptops_cleaned
WHERE brand != 'Other'
GROUP BY brand;

-- View 2: Price Segment Analysis
CREATE OR REPLACE VIEW vw_price_segments AS
SELECT
    price_segment,
    COUNT(*) AS total_products,
    ROUND(AVG(price_inr), 0) AS avg_price,
    ROUND(AVG(rating), 2) AS avg_rating,
    ROUND(AVG(discount_percent), 1) AS avg_discount,
    ROUND(AVG(savings_inr), 0) AS avg_savings
FROM laptops_cleaned
GROUP BY price_segment;

-- View 3: Discount Analysis by Brand and Tier
CREATE OR REPLACE VIEW vw_discount_analysis AS
SELECT
    brand,
    discount_tier,
    COUNT(*) AS product_count,
    ROUND(AVG(discount_percent), 1) AS avg_discount,
    ROUND(AVG(savings_inr), 0) AS avg_savings,
    MAX(discount_percent) AS max_discount
FROM laptops_cleaned
WHERE brand != 'Other'
GROUP BY brand, discount_tier;

-- View 4: Top Products (rated only)
CREATE OR REPLACE VIEW vw_top_products AS
SELECT
    product_name, brand, price_inr,
    original_price_inr,
    COALESCE(discount_percent, 0) AS discount_percent,
    COALESCE(savings_inr, 0) AS savings_inr,
    rating,
    rating_category,
    price_segment,
    discount_tier
FROM laptops_cleaned
WHERE rating IS NOT NULL
ORDER BY rating DESC, discount_percent DESC;

-- View 5: KPI Summary (single row of headline numbers)
CREATE OR REPLACE VIEW vw_kpi_summary AS
SELECT
    COUNT(*)                                AS total_products,
    COUNT(DISTINCT brand)                   AS total_brands,
    ROUND(AVG(price_inr), 0)               AS avg_price,
    MIN(price_inr)                          AS min_price,
    MAX(price_inr)                          AS max_price,
    ROUND(AVG(rating), 2)                   AS avg_rating,
    ROUND(AVG(discount_percent), 1)         AS avg_discount,
    SUM(CASE WHEN discount_percent >= 40
             THEN 1 ELSE 0 END)             AS mega_deals,
    SUM(CASE WHEN rating >= 4.5
             THEN 1 ELSE 0 END)             AS excellent_products
FROM laptops_cleaned;

-- View 6: Full clean dataset for Power BI slicers
CREATE OR REPLACE VIEW vw_all_products AS
SELECT
    id, asin, product_name, brand,
    price_inr, original_price_inr,
    COALESCE(discount_percent, 0) AS discount_percent,
    COALESCE(savings_inr, 0) AS savings_inr,
    COALESCE(rating, 0) AS rating,
    total_reviews,
    price_segment, discount_tier,
    rating_category, listing_type,
    search_keyword, scraped_at
FROM laptops_cleaned;

-- Verify all 6 views created
SHOW FULL TABLES WHERE Table_type = 'VIEW';
