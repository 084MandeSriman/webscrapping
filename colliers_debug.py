"""
colliers_debug.py — Dump Colliers India properties page HTML and find card selectors.
Run: python colliers_debug.py
"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

URL = "https://www.colliers.com/en-in/properties#q=Hyderabad&sort=relevancy&f:listingtype=[All%20Listings]&f:recenttransactions=[0]"
BASE_DOMAIN = "https://www.colliers.com"

opts = Options()
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--window-size=1920,1080")
opts.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36")


def _safe_text(tag, selector: str) -> str:
    if not tag:
        return ""
    el = tag.select_one(selector)
    return el.get_text(" ", strip=True) if el else ""


def parse_teaser_cards(soup: BeautifulSoup) -> list:
    cards = []
    for card in soup.select("div.teaser-card-container"):
        link_tag = card.select_one("a.CoveoResultLink[href]")
        href = link_tag["href"].strip() if link_tag else ""
        if href and href.startswith("/"):
            href = BASE_DOMAIN + href

        cards.append({
            "Listing ID": card.select_one("div.teaser-card").get("data-propertyid", "") if card.select_one("div.teaser-card") else "",
            "Status": _safe_text(card, ".teaser-card__type"),
            "Title": _safe_text(card, ".teaser-card__title"),
            "Address": _safe_text(card, ".teaser-card__address"),
            "Size": _safe_text(card, ".teaser-card__meta-detail--size"),
            "Property Type": _safe_text(card, ".teaser-card__propertytype-single"),
            "Details Link": href,
            "Image URL": card.select_one("img") and card.select_one("img").get("src", ""),
        })
    return cards


try:
    driver = webdriver.Chrome(options=opts)
    driver.get(URL)
    WebDriverWait(driver, 45).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.teaser-card-container"))
    )
    time.sleep(5)

    for _ in range(5):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    html = driver.page_source
    soup = BeautifulSoup(html, "lxml")

    with open("colliers_dump.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Saved colliers_dump.html")

    print("\n── Detected Colliers teaser cards ──")
    teaser_cards = parse_teaser_cards(soup)
    print(f"Found {len(teaser_cards)} teaser-card-container elements")
    for card in teaser_cards[:5]:
        print(f"- {card['Title']}")
        print(f"    Status: {card['Status']}")
        print(f"    Address: {card['Address']}")
        print(f"    Size: {card['Size']}")
        print(f"    Property Type: {card['Property Type']}")
        print(f"    Details Link: {card['Details Link']}")
        print()

    print("\n── Card selectors found ──")
    for sel in ["article", "div.teaser-card-container", "[class*='card']", "[class*='listing']", "[class*='property']",
                "[class*='tile']", "[class*='result']", "[class*='item']",
                "li[class]", "div[class*='prop']", "[class*='PropertyCard']",
                "[class*='SearchResult']", "[class*='ListingCard']"]:
        found = soup.select(sel)
        if 2 <= len(found) <= 200:
            print(f"  '{sel}': {len(found)} elements")
            for el in found[:2]:
                cls = el.get("class", [])
                txt = el.get_text(strip=True)[:80]
                print(f"    class={cls}")
                print(f"    text: {txt}")
                print()

    print("\n── Property links ──")
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href not in seen and "colliers" in href and len(href) > 30:
            if any(x in href.lower() for x in ["property", "listing", "office", "hyderabad", "india"]):
                print(f"  {a.get_text(strip=True)[:50]:50s} -> {href[:100]}")
                seen.add(href)
                if len(seen) > 15:
                    break
finally:
    driver.quit()
