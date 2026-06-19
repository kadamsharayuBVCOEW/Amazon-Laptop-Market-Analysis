# Amazon-Laptop-Market-Analysis
💻 Amazon.in Laptop Market Analysis

An end-to-end data analytics pipeline that scrapes real laptop listings from Amazon.in, cleans and analyzes them in MySQL, and visualizes business insights through an interactive Power BI dashboard.

Pipeline: Python (Scrape) → Excel (Store) → MySQL (Clean + Analyze) → Power BI (Visualize)


📊 Project Overview

MetricValueProducts scraped1,226Products after cleaning1,152Search keywords used7Brands identified39SQL business queries written10Power BI views created6Dashboard pages4


🎯 Why This Project

Instead of using a ready-made dataset, I wanted real-world experience handling messy, unpredictable data — from collection to final business insight. This project replicates how data actually moves inside a company: raw collection → storage → cleaning → analysis → visualization.


🛠️ Tools & Technologies

ToolPurposePython (requests, BeautifulSoup, pandas)Web scraping with anti-bot handlingExcelRaw data storage and initial inspectionMySQLData cleaning, business queries, viewsPower BIInteractive dashboard, DAX measures


🔄 Pipeline Walkthrough

Phase 1 — Python Web Scraping

Scraped Amazon.in across 7 keywords (laptops, gaming laptops, business laptops, ultrabook laptops, etc.) using requests and BeautifulSoup. Handled Amazon's anti-bot protection using:


Browser header simulation (fake Chrome identity)
Session-based cookie persistence
Randomized delays (4–8 sec) between requests
ASIN-based deduplication across keywords


Output: 1,226 unique product listings → amazon_laptops_raw.xlsx

Phase 2 — Excel

Performed initial data quality inspection — identified 52 suspiciously low-priced records (accessories, not laptops) and inconsistent brand naming. No cleaning was done here intentionally — raw data preserved, exported to CSV for SQL.

Phase 3 — MySQL


Created laptops_raw and laptops_cleaned tables
Removed 74 invalid records (price-based filtering)
Standardized 39 brand names using CASE statements and REGEXP
Recalculated discount % from actual price vs MRP
Created derived columns: price_segment, discount_tier, rating_category
Wrote 10 business insight queries
Built 6 SQL Views as a direct feed for Power BI


Phase 4 — Power BI

Connected live to MySQL via ODBC connector. Built a 4-page dashboard:


Executive Summary — KPI cards, market share donut, brand leaderboard
Brand Performance — price, rating, discount comparison across brands
Price & Discount Analysis — segment distribution, discount patterns
Top Products & Best Deals — product table, price-vs-discount scatter plot


Added 2 DAX measures: Premium Product % and High Rated %.


💡 Key Business Insights


Premium segment (₹50K–₹80K) dominates with 428 products — 37% of the market
72.74% of all laptops are priced above ₹50,000 — the market is premium-heavy
Dell offers the highest average discount (38.7%) while Apple offers the lowest (8.9%) — two completely opposite pricing strategies
Apple leads in customer satisfaction with a 4.64/5 average rating
42.7% of laptops are rated "Below Average" — a significant quality gap in the market
Higher-priced segments correlate with higher ratings (Ultra-Premium: 4.35 vs Budget: 3.61)



📁 Repository Contents

FileDescriptionamazon_scraper_final.pyPython scraper — multi-keyword, anti-block, checkpointingamazon_laptops_raw.xlsxRaw scraped dataset (1,226 rows × 12 columns)Amazon_Laptop_Analytics_FINAL.sqlComplete MySQL script — schema, cleaning, queries, viewsAmazon_Laptop_Market_Analysis.pbixPower BI dashboard file


🚀 How to Run This Project


Scraper: pip install requests beautifulsoup4 pandas openpyxl → python amazon_scraper_final.py
Database: Run Amazon_Laptop_Analytics_FINAL.sql in MySQL Workbench (creates schema, imports data, cleans, builds views)
Dashboard: Open Amazon_Laptop_Market_Analysis.pbix in Power BI Desktop → connect to your local MySQL instance → Refresh



📌 Data Source

All data was scraped directly from publicly available Amazon.in search result pages for educational and portfolio purposes.


📬 Contact

Feel free to connect with me on LinkedIn: www.linkedin.com/in/sharayu-kadam or reach out via email: sharayukadam95@gmail.com
