"""
sy_commercial_debug.py — Dump Square Yards commercial office page HTML.
Run: python sy_commercial_debug.py
"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

URLS = [
    "https://www.squareyards.com/rent/office-spaces-for-rent-in-hyderabad",
    "https://www.squareyards.com/sale/office-spaces-for-sale-in-hyderabad",
]

opts = Options()
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--window-size=1920,1080")
opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=opts)

for url in URLS:
    print(f"\n{'='*60}")
    print(f"URL: {url}")
    driver.get(url)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(8)

    for _ in range(5):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)

    soup = BeautifulSoup(driver.page_source, "lxml")

    # Save HTML
    fname = url.split("/")[-1] + ".html"
    with open(fname, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"Saved {fname}")

    # Try all card selectors
    for sel in ["article", "[class*='property']", "[class*='listing']",
                "[class*='card']", "[class*='tile']", "[class*='item']",
                "[class*='result']", "[class*='srp']", "li[class]", "div[class*='prop']"]:
        found = soup.select(sel)
        if len(found) >= 2:
            print(f"  Selector '{sel}': {len(found)} elements")
            if len(found) < 20:
                for el in found[:2]:
                    print(f"    classes: {el.get('class')}")
                    print(f"    text[:100]: {el.get_text(strip=True)[:100]}")
                    print()

    # Print all property-related links
    print("\n  Property links found:")
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href not in seen and any(x in href for x in ["hyderabad", "office", "commercial"]):
            if len(href) > 20 and "squareyards" in href:
                print(f"    {a.get_text(strip=True)[:50]:50s} -> {href[:80]}")
                seen.add(href)
                if len(seen) > 15:
                    break

driver.quit()
