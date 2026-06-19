import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import re
import os
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════
#   AMAZON.IN MULTI-KEYWORD LAPTOP SCRAPER
#   Project  : E-Commerce Sales Analytics Dashboard
#   Phase    : 1 — Data Collection
#   Author   : [Your Name]
#   Target   : 5,000 – 8,000 unique laptop products
#
#   Keywords scraped:
#     1. laptops
#     2. gaming laptops
#     3. business laptops
#     4. laptops under 50000
#     5. laptops under 30000
#     6. student laptops
#     7. ultrabook laptops
#
#   Pipeline: Python → Excel → SQL → Power BI
# ═══════════════════════════════════════════════════════════════════


# ── CONFIGURATION ─────────────────────────────────────────────────
# Tweak these if you want more/fewer results

KEYWORDS = [
    "laptops",
    "gaming laptops",
    "business laptops",
    "laptops under 50000",
    "laptops under 30000",
    "student laptops",
    "ultrabook laptops",
]

MAX_PAGES_PER_KEYWORD = 20      # ~20 products/page × 20 pages × 7 keywords = ~2800 base
                                 # After dedup expect 5,000–8,000 unique products

DELAY_BETWEEN_PAGES  = (4, 8)   # Random wait in seconds between pages (avoids blocks)
DELAY_BETWEEN_KEYWORDS = (15, 25) # Longer pause between keywords (looks human)

OUTPUT_FILE = "amazon_laptops_raw.xlsx"
CHECKPOINT_FILE = "amazon_laptops_checkpoint.csv"  # Auto-saves progress


# ── HEADERS ───────────────────────────────────────────────────────
# Full Chrome browser headers — this is the #1 anti-block technique

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Cache-Control": "max-age=0",
    "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}

session = requests.Session()
session.headers.update(HEADERS)


# ── HELPER: WARM UP SESSION ───────────────────────────────────────
# Visit Amazon homepage first to collect cookies —
# makes subsequent requests look like real browsing

def warm_up_session():
    print("\n🔄 Warming up session (visiting Amazon homepage)...")
    try:
        r = session.get("https://www.amazon.in", timeout=15)
        print(f"   ✅ Homepage: {r.status_code} | Cookies collected: {len(session.cookies)}")
        time.sleep(random.uniform(3, 5))
    except Exception as e:
        print(f"   ⚠️  Warm-up warning: {e} — continuing anyway")


# ── HELPER: BUILD SEARCH URL ──────────────────────────────────────
# Converts a keyword + page number into Amazon search URL

def build_url(keyword, page):
    from urllib.parse import quote_plus
    encoded = quote_plus(keyword)
    return (
        f"https://www.amazon.in/s"
        f"?k={encoded}"
        f"&i=computers"
        f"&s=review-rank"       # Sort by most reviewed — better quality data
        f"&page={page}"
    )


# ── CORE: SCRAPE ONE PAGE ─────────────────────────────────────────
# Visits one search results page and extracts all product cards.
# Returns a list of product dictionaries.

def scrape_page(keyword, page_num):
    url = build_url(keyword, page_num)

    try:
        delay = random.uniform(*DELAY_BETWEEN_PAGES)
        print(f"    ⏳ Waiting {delay:.1f}s ...", end=" ", flush=True)
        time.sleep(delay)

        response = session.get(url, timeout=20)

        # Handle rate limiting
        if response.status_code == 503:
            print(f"\n    ⚠️  Rate limited (503). Sleeping 60s then retrying...")
            time.sleep(60)
            response = session.get(url, timeout=20)

        if response.status_code != 200:
            print(f"❌ Blocked (HTTP {response.status_code})")
            return [], False   # (products, should_continue)

        soup = BeautifulSoup(response.content, "html.parser")

        # CAPTCHA detection
        if "captcha" in response.text.lower() or "enter the characters" in response.text.lower():
            print("\n    🚨 CAPTCHA detected!")
            print("    👉 Open https://www.amazon.in in your browser, solve CAPTCHA,")
            print("       then press ENTER here to retry...")
            input()
            response = session.get(url, timeout=20)
            soup = BeautifulSoup(response.content, "html.parser")

        # Find all product cards
        cards = soup.find_all("div", {"data-component-type": "s-search-result"})
        print(f"Found {len(cards)} products")

        if len(cards) == 0:
            # Check if it's a "no results" page vs a block
            no_results = soup.find("span", string=re.compile("No results for", re.I))
            if no_results:
                return [], False  # Keyword exhausted
            return [], True       # Might be a block, keep trying

        products = []

        for card in cards:
            try:
                # ── ASIN (Amazon unique product ID) ──────────────
                asin = card.get("data-asin", "").strip()
                if not asin:
                    continue

                # ── Product Name ──────────────────────────────────
                name_tag = (
                    card.find("span", {"class": "a-size-medium"}) or
                    card.find("span", {"class": "a-size-base-plus"}) or
                    card.find("h2")
                )
                name = name_tag.get_text(strip=True) if name_tag else None
                if not name:
                    continue

                # ── Brand ─────────────────────────────────────────
                # Try dedicated brand tag first, else use first word of name
                brand_tag = card.find("span", {"class": "a-size-base-plus a-color-base"})
                brand = brand_tag.get_text(strip=True) if brand_tag else name.split()[0]

                # ── Current Selling Price ─────────────────────────
                price_tag = card.find("span", {"class": "a-price-whole"})
                price_text = (
                    price_tag.get_text(strip=True).replace(",", "").replace(".", "")
                    if price_tag else None
                )
                price = float(price_text) if price_text and price_text.isdigit() else None

                # ── MRP / Original Price ──────────────────────────
                original_price = None
                mrp_block = card.find("span", {"class": "a-price a-text-price"})
                if mrp_block:
                    offscreen = mrp_block.find("span", {"class": "a-offscreen"})
                    if offscreen:
                        raw = (
                            offscreen.get_text(strip=True)
                            .replace("₹", "")
                            .replace(",", "")
                            .strip()
                        )
                        try:
                            original_price = float(raw)
                        except:
                            pass

                # ── Discount % ────────────────────────────────────
                discount = None
                if price and original_price and original_price > price:
                    discount = round(((original_price - price) / original_price) * 100, 1)

                # ── Rating (out of 5) ─────────────────────────────
                rating = None
                rating_tag = card.find("span", {"class": "a-icon-alt"})
                if rating_tag:
                    match = re.search(r"(\d+\.?\d*)", rating_tag.get_text(strip=True))
                    if match:
                        rating = float(match.group(1))

                # ── Total Reviews Count ───────────────────────────
                reviews = None
                reviews_tag = card.find("span", {"class": "a-size-base", "dir": "auto"})
                if reviews_tag:
                    rev_text = reviews_tag.get_text(strip=True).replace(",", "").strip()
                    try:
                        reviews = int(rev_text)
                    except:
                        pass

                # ── Sponsored / Organic ───────────────────────────
                sponsored_tag = card.find("span", string=re.compile("Sponsored", re.I))
                listing_type = "Sponsored" if sponsored_tag else "Organic"

                # ── Search Keyword that found this product ────────
                # Useful for analysis: which keyword returns which products

                # Only save products that have at least a price
                if price:
                    products.append({
                        "asin":               asin,
                        "product_name":       name,
                        "brand":              brand,
                        "price_inr":          price,
                        "original_price_inr": original_price,
                        "discount_percent":   discount,
                        "rating":             rating,
                        "total_reviews":      reviews,
                        "listing_type":       listing_type,
                        "search_keyword":     keyword,
                        "page_scraped":       page_num,
                        "scraped_at":         datetime.now().strftime("%Y-%m-%d %H:%M"),
                    })

            except Exception:
                continue  # Skip broken cards silently

        return products, True

    except requests.exceptions.Timeout:
        print(f"⏱️  Timeout — skipping page {page_num}")
        return [], True
    except requests.exceptions.ConnectionError:
        print(f"🔌 Connection error — check your internet")
        return [], True
    except Exception as e:
        print(f"❌ Error: {e}")
        return [], True


# ── CORE: SCRAPE ALL KEYWORDS ─────────────────────────────────────
# Loops through every keyword, scrapes up to MAX_PAGES_PER_KEYWORD
# Saves a checkpoint CSV after each keyword (so you don't lose progress)

def scrape_all_keywords():
    all_products = []
    seen_asins = set()   # Track unique products across keywords

    print("\n" + "═" * 62)
    print("  🛒  Amazon.in Laptop Scraper — MULTI-KEYWORD MODE")
    print(f"  🎯  Target: 5,000–8,000 unique products")
    print(f"  🔑  Keywords: {len(KEYWORDS)}")
    print(f"  📄  Pages/keyword: {MAX_PAGES_PER_KEYWORD}")
    print("═" * 62)

    warm_up_session()

    for kw_index, keyword in enumerate(KEYWORDS, start=1):
        print(f"\n{'─'*62}")
        print(f"  🔑  Keyword {kw_index}/{len(KEYWORDS)}: '{keyword}'")
        print(f"  📦  Unique products so far: {len(all_products)}")
        print(f"{'─'*62}")

        kw_products = []
        empty_pages = 0

        for page in range(1, MAX_PAGES_PER_KEYWORD + 1):
            print(f"  📄 Page {page}/{MAX_PAGES_PER_KEYWORD} — ", end="", flush=True)

            page_products, should_continue = scrape_page(keyword, page)

            # Filter out already-seen ASINs (dedup across keywords)
            new_products = [
                p for p in page_products
                if p["asin"] not in seen_asins
            ]
            for p in new_products:
                seen_asins.add(p["asin"])

            kw_products.extend(new_products)
            all_products.extend(new_products)

            print(f"      ✨ {len(new_products)} new | 🔁 {len(page_products)-len(new_products)} dupes | 📦 Total: {len(all_products)}")

            if not should_continue:
                print(f"  ⚠️  Stopping keyword '{keyword}' early — no more results")
                break

            if len(page_products) == 0:
                empty_pages += 1
                if empty_pages >= 2:
                    print(f"  ⚠️  2 empty pages in a row — moving to next keyword")
                    break
            else:
                empty_pages = 0

        # ── CHECKPOINT SAVE after each keyword ───────────────────
        # If script crashes midway, you won't lose everything
        if all_products:
            checkpoint_df = pd.DataFrame(all_products)
            checkpoint_df.to_csv(CHECKPOINT_FILE, index=False)
            print(f"\n  💾 Checkpoint saved → {CHECKPOINT_FILE} ({len(all_products)} products)")

        # Longer pause between keywords — looks human to Amazon
        if kw_index < len(KEYWORDS):
            pause = random.uniform(*DELAY_BETWEEN_KEYWORDS)
            print(f"\n  😴 Resting {pause:.0f}s before next keyword...")
            time.sleep(pause)

    return all_products


# ── SAVE FINAL EXCEL ──────────────────────────────────────────────
# Produces the clean Excel file that goes to Phase 2

def save_to_excel(products):
    if not products:
        print("\n❌ No products collected. Check errors above.")
        return None

    df = pd.DataFrame(products)

    # Final dedup by ASIN (safety net)
    before = len(df)
    df.drop_duplicates(subset="asin", inplace=True)
    df.reset_index(drop=True, inplace=True)
    after = len(df)

    # Save raw data — NO cleaning yet (that's SQL's job in Phase 3)
    df.to_excel(OUTPUT_FILE, index=False)

    # ── FINAL REPORT ─────────────────────────────────────────────
    print("\n" + "═" * 62)
    print("  ✅  SCRAPING COMPLETE — FINAL REPORT")
    print("═" * 62)
    print(f"  📊 Total unique products   : {after:,}")
    print(f"  🔁 Duplicates removed      : {before - after:,}")
    print(f"  💾 Saved to                : {OUTPUT_FILE}")
    print(f"  🗂️  Columns                 : {list(df.columns)}")
    print("═" * 62)

    print("\n📋 PREVIEW — First 5 rows:")
    preview_cols = ["product_name", "brand", "price_inr", "discount_percent", "rating", "search_keyword"]
    print(df[preview_cols].head().to_string(index=False))

    print(f"\n📊 KEYWORD BREAKDOWN:")
    print(df.groupby("search_keyword")["asin"].count().rename("products").to_string())

    print(f"\n📈 QUICK STATS:")
    print(f"   Avg Selling Price   : ₹{df['price_inr'].mean():>10,.0f}")
    print(f"   Min Price           : ₹{df['price_inr'].min():>10,.0f}")
    print(f"   Max Price           : ₹{df['price_inr'].max():>10,.0f}")
    if df['rating'].notna().any():
        print(f"   Avg Rating          : {df['rating'].mean():>10.2f} / 5.0")
    print(f"   Unique Brands       : {df['brand'].nunique():>10,}")
    print(f"   % with Discount     : {df['discount_percent'].notna().mean()*100:>9.1f}%")
    print(f"   Avg Discount        : {df['discount_percent'].mean():>9.1f}%")
    print(f"   Sponsored listings  : {(df['listing_type']=='Sponsored').sum():>10,}")
    print(f"   Organic listings    : {(df['listing_type']=='Organic').sum():>10,}")

    print(f"\n✅ Phase 1 DONE! Open '{OUTPUT_FILE}' to verify your data.")
    print("   Next → Phase 2: Excel inspection")
    print("   Then → Phase 3: SQL (data cleaning + views)")
    print("   Then → Phase 4: Power BI dashboard\n")

    # Clean up checkpoint file
    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)
        print(f"   🧹 Checkpoint file removed (no longer needed)")

    return df


# ── RESUME FROM CHECKPOINT ────────────────────────────────────────
# If your script crashed midway, this loads saved progress
# and skips already-scraped ASINs

def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        print(f"\n⚡ Checkpoint found: {CHECKPOINT_FILE}")
        df = pd.read_csv(CHECKPOINT_FILE)
        print(f"   Loaded {len(df):,} previously scraped products")
        print("   Resuming from where we left off...\n")
        return df.to_dict("records"), set(df["asin"].tolist())
    return [], set()


# ── MAIN ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "═" * 62)
    print("  🚀  E-Commerce Analytics Project — Phase 1")
    print("  📌  Amazon.in Laptop Data Collection")
    print("═" * 62)

    products = scrape_all_keywords()
    df = save_to_excel(products)
