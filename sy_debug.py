"""
sy_debug.py — Dump Square Yards page HTML and print all hrefs to find real property URLs.
Run: python sy_debug.py
"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

URL = "https://www.squareyards.com/hyderabad-real-estate"

opts = Options()
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--window-size=1920,1080")
opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=opts)
driver.get(URL)
WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
time.sleep(8)

# Scroll 5 times
for _ in range(5):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)

soup = BeautifulSoup(driver.page_source, "lxml")

# Save full HTML
with open("sy_dump.html", "w", encoding="utf-8") as f:
    f.write(driver.page_source)
print("Saved sy_dump.html")

# Print all unique hrefs containing project/property names
print("\n── All hrefs with 'hyderabad' or project names ──")
seen = set()
for a in soup.find_all("a", href=True):
    href = a["href"]
    text = a.get_text(strip=True)[:60]
    if href not in seen and len(href) > 5:
        if any(x in href.lower() for x in ["hyderabad", "project", "property", "prestige", "brigade", "godrej", "ramky"]):
            print(f"  TEXT: {text!r:50s}  HREF: {href}")
            seen.add(href)

# Print first 3 card HTML snippets
print("\n── First 3 card HTML snippets ──")
for sel in ["[class*='projectCard']","[class*='srpCard']","[class*='listingCard']",
            "[class*='PropertyCard']","[class*='projectTile']","article"]:
    cards = soup.select(sel)
    if len(cards) >= 2:
        print(f"\nSelector: {sel}  ({len(cards)} cards found)")
        for card in cards[:3]:
            print(card.prettify()[:800])
            print("---")
        break

driver.quit()
