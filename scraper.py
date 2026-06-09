"""
scraper.py — JLL India commercial office listings scraper.

Usage:
    python scraper.py                        # Hyderabad (default)
    python scraper.py --city Mumbai
    python scraper.py --no-headless
    python scraper.py --max-scrolls 30
"""

import os
import re
import time
import argparse
import traceback
import requests
from urllib.parse import urlencode

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils import setup_logger, ensure_output_dir, clean_dataframe, save_csv, save_excel_multi

# ── Constants ─────────────────────────────────────────────────────────────────

BASE_URL = "https://property.jll.co.in/search"

CITY_REGION_MAP = {
    "Hyderabad": "Telangana",
    "Mumbai":    "Maharashtra",
    "Bangalore": "Karnataka",
    "Chennai":   "Tamil Nadu",
    "Delhi":     "Delhi",
    "Pune":      "Maharashtra",
}

SCROLL_PAUSE = 3.0
MAX_RETRIES  = 3
OUTPUT_DIR   = "output"


# ── URL builder ───────────────────────────────────────────────────────────────

def build_url(city: str) -> str:
    region = CITY_REGION_MAP.get(city, "")
    params = {
        "tenureTypes":   "rent",
        "propertyTypes": "office",
        "cities":        city,
        "regions":       region,
        "sortBy":        "dateModified",
    }
    return f"{BASE_URL}?{urlencode(params)}"


# ── Driver setup ──────────────────────────────────────────────────────────────

def create_driver(headless: bool = True) -> webdriver.Chrome:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    # Method 1: Selenium 4.6+ built-in selenium-manager
    try:
        driver = webdriver.Chrome(options=opts)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception:
        pass

    # Method 2: webdriver-manager fallback
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        raise RuntimeError(f"Could not start ChromeDriver. Ensure Google Chrome is installed.\nDetails: {e}")


# ── Page loader ───────────────────────────────────────────────────────────────

def load_page(driver: webdriver.Chrome, url: str, logger, retries: int = MAX_RETRIES) -> bool:
    for attempt in range(1, retries + 1):
        try:
            logger.info(f"Loading URL (attempt {attempt}): {url}")
            driver.get(url)
            # Wait for the first property card to appear
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-cy='property-card']"))
            )
            time.sleep(3)
            return True
        except Exception as e:
            logger.warning(f"Attempt {attempt} failed: {e}")
            time.sleep(3)
    return False


# ── Infinite scroll ───────────────────────────────────────────────────────────

def _click_load_more(driver: webdriver.Chrome) -> None:
    """Click any visible Load More / Show More button."""
    for xpath in [
        "//button[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'LOAD MORE')]",
        "//button[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'SHOW MORE')]",
        "//a[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'LOAD MORE')]",
    ]:
        try:
            btn = driver.find_element(By.XPATH, xpath)
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(2)
                return
        except Exception:
            pass


def scroll_to_bottom(driver: webdriver.Chrome, logger, max_scrolls: int = 100) -> None:
    last_height = driver.execute_script("return document.body.scrollHeight")
    no_change_count = 0
    for i in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE)
        _click_load_more(driver)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            no_change_count += 1
            if no_change_count >= 3:
                logger.info(f"[JLL] Reached end of page after {i + 1} scrolls.")
                break
            time.sleep(3)
        else:
            no_change_count = 0
        last_height = new_height
        logger.debug(f"Scroll {i + 1}: height={new_height}px")


# ── Parsing ───────────────────────────────────────────────────────────────────

def _safe_text(tag, selector: str) -> str:
    el = tag.select_one(selector)
    return el.get_text(separator=" ", strip=True) if el else ""


def _extract_listing_id(url: str) -> str:
    # Extract slug from URL like /listings/building-name-123
    match = re.search(r"/listings/([^/?#]+)", url)
    return match.group(1) if match else ""


def parse_listings(html: str, city: str, base_domain: str = "https://property.jll.co.in") -> list:
    """
    Parse property cards using exact selectors found in JLL's HTML:
    - Card:    [data-cy="property-card"]
    - Details: [data-cy="property-card-details"]
    - Name:    .text-jll-red (red uppercase building name)
    - Address: two .block.text-base spans after the name
    - Rent:    p containing "Rent" span
    - Size:    p containing "Size" span
    - Type:    .bg-stone-200 badge
    - Link:    <a href="/listings/...">
    - Image:   <img> inside the card image div
    """
    soup = BeautifulSoup(html, "lxml")
    records = []

    cards = soup.select("[data-cy='property-card']")

    for card in cards:
        # ── Link ──────────────────────────────────────────────────────────────
        link_tag = card.select_one("a[href^='/listings/']")
        raw_href = link_tag["href"] if link_tag else ""
        prop_link = base_domain + raw_href if raw_href else ""

        # ── Image ─────────────────────────────────────────────────────────────
        img = card.select_one("img[srcset]")
        image_url = img.get("src", "") if img else ""

        # ── Details block ─────────────────────────────────────────────────────
        details_div = card.select_one("[data-cy='property-card-details']")

        # Building name — red uppercase span
        name = _safe_text(details_div or card, ".text-jll-red") if details_div else ""

        # Address — two consecutive .block.text-base spans
        addr_spans = (details_div or card).select("a .block.text-base") if details_div else []
        address = ", ".join(s.get_text(strip=True) for s in addr_spans if s.get_text(strip=True))

        # Property type badge
        prop_type = _safe_text(card, ".bg-stone-200") or "Office"

        # Rent — paragraph containing "Rent" label
        rent = ""
        size = ""
        if details_div:
            for p in details_div.select("p"):
                text = p.get_text(" ", strip=True)
                if "Rent" in text:
                    # Get the bold value span
                    bold = p.select_one(".font-bold")
                    rent = bold.get_text(strip=True) if bold else text
                elif "Size" in text:
                    bold = p.select_one(".font-bold")
                    size = bold.get_text(strip=True) if bold else text

        record = {
            "Building Name":    name,
            "Property Type":    prop_type,
            "Address":          address,
            "Area / Size":      size,
            "Rent":             rent,
            "City":             city,
            "Region":           CITY_REGION_MAP.get(city, ""),
            "Property Details": "",
            "Property Link":    prop_link,
            "Image URL":        image_url,
            "Listing ID":       _extract_listing_id(prop_link),
        }

        if record["Building Name"] or record["Property Link"]:
            records.append(record)

    return records


# ── Post-processing / data cleaning rules ────────────────────────────────────

def _parse_sqft(value) -> float:
    """Extract the highest numeric sqft value from a string like '25,000 - 50,000 sqft'."""
    if not isinstance(value, str) or not value.strip():
        return 0.0
    numbers = re.findall(r"[\d]+", value.replace(",", ""))
    floats = [float(n) for n in numbers if n]
    return max(floats) if floats else 0.0


def apply_combined_rules(df: pd.DataFrame) -> pd.DataFrame:
    # Rule 3: Clean and replace range with highest value in Area / Size
    def _keep_highest_sqft(val):
        if not isinstance(val, str):
            return val
        # Strip quotes, newlines, extra whitespace
        val = re.sub(r'["\n\r\t]', ' ', val)
        val = re.sub(r'\s+', ' ', val).strip()
        if not val:
            return val
        # Extract all numbers
        numbers = re.findall(r'[\d]+', val.replace(',', ''))
        if len(numbers) >= 2:
            highest = max(int(n) for n in numbers)
            unit = 'SF' if 'SF' in val else 'Sq. Ft' if 'Sq' in val else 'sqft'
            return f'{highest:,} {unit}'
        elif len(numbers) == 1:
            # Single value — just clean it up
            unit = 'SF' if 'SF' in val else 'Sq. Ft' if 'Sq' in val else 'sqft'
            return f'{int(numbers[0]):,} {unit}'
        return val

    df['Area / Size'] = df['Area / Size'].apply(_keep_highest_sqft)

    # Locality Filter: Keep only Kokapet, Kondapur, HITEC City, Financial District, Nanakramguda
    target_areas = ["kokapet", "kondapur", "hitec", "hitech", "hightech", "hi tech", "high tech", "hi-tech", "financial", "finaceal", "nanakramguda"]
    df = df[df.apply(
        lambda r: any(a in (str(r.get("Address", "")) + " " + str(r.get("Building Name", ""))).lower() for a in target_areas),
        axis=1
    )].copy()

    # Size & Generic Type Filter: Remove small spaces and generic listings
    df['_sqft_val'] = df['Area / Size'].apply(_parse_sqft)

    generic_keywords = [
        "office space for", "co-working space", "coworking space", 
        "shop for", "showroom for", "warehouse for", "work area",
        "independent building", "independent office"
    ]

    def _is_small_or_generic(row):
        name_lower = str(row.get("Building Name", "")).lower()
        addr_lower = str(row.get("Address", "")).lower()
        sqft = row["_sqft_val"]
        # Never drop Kokapet listings based on size/type
        if "kokapet" in addr_lower or "kokapet" in name_lower:
            return False
        if any(kw in name_lower for kw in generic_keywords):
            return True
        if 0 < sqft < 10000:
            return True
        return False

    df = df[~df.apply(_is_small_or_generic, axis=1)].copy()

    # Rule 1: For same building name (case-insensitive), keep only the row with highest sqft
    # BUT always keep residential rows — deduplicate commercial separately
    df_res = df[df["Source"] == "Residential"].copy()
    df_com = df[df["Source"] != "Residential"].copy()

    df_com['_name_lower'] = df_com['Building Name'].str.strip().str.lower()
    df_com = df_com.sort_values('_sqft_val', ascending=False)
    df_com = df_com.drop_duplicates(subset=['_name_lower'], keep='first')
    df_com = df_com.drop(columns=['_sqft_val', '_name_lower'])

    df_res['_name_lower'] = df_res['Building Name'].str.strip().str.lower()
    df_res = df_res.drop_duplicates(subset=['_name_lower'], keep='first')
    df_res = df_res.drop(columns=['_sqft_val', '_name_lower'], errors='ignore')

    df = pd.concat([df_com, df_res], ignore_index=True).reset_index(drop=True)
    return df


def scrape_jll(city: str = "Hyderabad", headless: bool = True, max_scrolls: int = 50) -> pd.DataFrame:
    logger = setup_logger()
    ensure_output_dir(OUTPUT_DIR)

    url = build_url(city)
    logger.info(f"[JLL] Target URL: {url}")

    driver = create_driver(headless=headless)
    all_records = []

    try:
        if not load_page(driver, url, logger):
            logger.error("[JLL] Failed to load page after retries.")
            return pd.DataFrame()

        logger.info("[JLL] Page loaded. Starting scroll...")
        scroll_to_bottom(driver, logger, max_scrolls=max_scrolls)

        logger.info("[JLL] Parsing HTML...")
        html = driver.page_source
        records = parse_listings(html, city)
        all_records.extend(records)
        logger.info(f"[JLL] Extracted {len(records)} raw listings.")

    except Exception:
        logger.error(f"[JLL] Unexpected error:\n{traceback.format_exc()}")
    finally:
        driver.quit()

    if not all_records:
        logger.warning("[JLL] No listings found.")
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    df = clean_dataframe(df)
    df["Source"] = "JLL"
    logger.info(f"[JLL] Clean listings: {len(df)}")
    return df


# ── CBRE scraper ──────────────────────────────────────────────────────────────

def scrape_cbre(city: str = "Hyderabad") -> pd.DataFrame:
    logger = setup_logger()
    if city.lower() != "hyderabad":
        logger.warning(f"[CBRE] Scraper is only configured for Hyderabad. Skipping CBRE for city: {city}")
        return pd.DataFrame()

    logger.info("[CBRE] Starting CBRE Hyderabad commercial office rentals scrape...")
    
    # We query the CBRE API directly to fetch all properties (setting PageSize=100 to get all in one call)
    api_url = (
        "https://www.cbre.co.in/property-api/propertylistings/query"
        "?Site=in-comm&RadiusType=Kilometers&CurrencyCode=INR&Unit=sqft"
        "&lon=78.47724389999999&Lat=17.406498&Lon=-0.12775829999998223"
        "&PolygonFilters=%5B%5B%2217.681159%2C80.917389%22%2C%2216.025792%2C80.917389%22%2C%2216.025792%2C77.615423%22%2C%2217.681159%2C77.615423%22%5D%5D"
        "&Common.Aspects=isLetting&Sort=asc(_distance)&Common.UsageType=Office"
        "&PageSize=100&Page=1"
        "&_select=Dynamic.PrimaryImage,Common.ActualAddress,Common.Charges,Common.NumberOfBedrooms,Common.PrimaryKey,Common.UsageType,Common.Coordinate,Common.Aspects,Common.ListingCount,Common.IsParent,Common.HomeSite,Common.Agents,Common.PropertySubType,Common.PropertyTypes,Common.ContactGroup,Common.Highlights,Common.Walkthrough,Common.MinimumSize,Common.MaximumSize,Common.TotalSize,Common.GeoLocation,Common.Sizes,Common.LeaseTypes"
    )
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://www.cbre.co.in/properties/office-space-for-rent-hyderabad",
    }
    
    try:
        response = requests.get(api_url, headers=headers, timeout=20)
        if response.status_code != 200:
            logger.error(f"[CBRE] API returned status code {response.status_code}")
            return pd.DataFrame()
            
        data = response.json()
        documents_list = data.get("Documents", [])
        if not documents_list or not isinstance(documents_list[0], list):
            logger.warning("[CBRE] No documents found in API response.")
            return pd.DataFrame()
            
        raw_listings = documents_list[0]
        records = []
        
        for item in raw_listings:
            addr_dict = item.get("Common.ActualAddress", {})
            building_name = addr_dict.get("Common.Line1", "").strip()
            
            # Form address
            line2 = addr_dict.get("Common.Line2", "").strip()
            locality = addr_dict.get("Common.Locallity", "").strip()
            address = line2 if line2 else locality
            # Ensure "Hyderabad" is in address
            if address and "hyderabad" not in address.lower():
                address = f"{address}, Hyderabad"
                
            # Form size
            size = ""
            sizes_list = item.get("Common.Sizes", [])
            if sizes_list:
                dims = sizes_list[0].get("Common.Dimensions", [])
                if dims:
                    size = f"{dims[0].get('Common.Amount'):,.0f} Sq. Ft"
            if not size:
                total_sizes = item.get("Common.TotalSize", [])
                if total_sizes:
                    min_area = total_sizes[0].get("Common.MinArea", 0)
                    if min_area:
                        size = f"{min_area:,.0f} Sq. Ft"
                        
            # Form rent
            rent = "Rent on Application"
            charges = item.get("Common.Charges", [])
            if charges:
                amount = charges[0].get("Common.Amount")
                per_unit = charges[0].get("Common.PerUnit", "sqft")
                rent_modifier = charges[0].get("Common.ChargeModifer", "").strip()
                
                modifier_str = ""
                if rent_modifier:
                    if rent_modifier.lower() == "from":
                        modifier_str = "From "
                    elif rent_modifier.lower() == "to":
                        modifier_str = "To "
                    else:
                        modifier_str = f"{rent_modifier} "
                        
                if amount:
                    rent = f"{modifier_str}₹{amount}/{per_unit} pm"
                    
            # Form details
            highlights_list = item.get("Common.Highlights", [])
            highlights = []
            for h in highlights_list:
                sub_h = h.get("Common.Highlight", [])
                if sub_h:
                    highlights.append(sub_h[0].get("Common.Text", ""))
            details = ", ".join(highlights)
            
            # Primary Key
            pk = item.get("Common.PrimaryKey", "")
            
            # Generate detail link slug
            slug = building_name.lower().strip()
            slug = re.sub(r'[^a-z0-9\s-]', '', slug)
            slug = re.sub(r'[\s-]+', '-', slug)
            detail_link = f"https://www.cbre.co.in/properties/office/details/{pk}/{slug}" if pk else ""
            
            # Image URL
            img_url = ""
            img_res = item.get("Dynamic.PrimaryImage", {}).get("Common.ImageResources", [])
            if img_res:
                img_url = img_res[0].get("Source.Uri", "")
                
            records.append({
                "Building Name":    building_name,
                "Property Type":    item.get("Common.UsageType", "Office"),
                "Address":          address,
                "Area / Size":      size,
                "Rent":             rent,
                "City":             city,
                "Region":           CITY_REGION_MAP.get(city, ""),
                "Property Details": details,
                "Property Link":    detail_link,
                "Image URL":        img_url,
                "Listing ID":       pk,
                "Source":           "CBRE"
            })
            
        df = pd.DataFrame(records)
        df = clean_dataframe(df)
        logger.info(f"[CBRE] {len(df)} listings successfully parsed.")
        return df
        
    except Exception as e:
        logger.error(f"[CBRE] Error querying API:\n{traceback.format_exc()}")
        return pd.DataFrame()


# ── Cushman & Wakefield scraper ───────────────────────────────────────────────

CW_URLS = [
    "https://www.cushmanwakefield.com/en/india/properties/for-lease/commercial/telangana/hyderabad/financial-district/eon-hyderabad-l",
]

# Hardcoded C&W listings — used as fallback if scraping fails or as static data
CW_STATIC_LISTINGS = [
    {
        "Building Name":       "EON Hyderabad",
        "Property Type":       "Office/Commercial",
        "Address":             "Financial District, Hyderabad, Telangana 500032, India",
        "Area / Size":         "2,211,560 SF",
        "Building Size":       "2,450,000 SF",
        "Rent":                "Contact us for pricing",
        "Year Built":          "2025",
        "Grade":               "A",
        "Construction Status": "The façade is currently under progress.",
        "City":                "Hyderabad",
        "Region":              "Telangana",
        "Property Details":    "Prime Location in Financial District. Excellent Connectivity to IT Corridor and Hyderabad International Airport (25-min drive). Modern Amenities planned. Tentatively six months from OC.",
        "Property Link":       "https://www.cushmanwakefield.com/en/india/properties/for-lease/commercial/telangana/hyderabad/financial-district/eon-hyderabad-l",
        "Image URL":           "",
        "Listing ID":          "eon-hyderabad-l",
        "Source":              "Cushman & Wakefield",
    },
]

def scrape_cw(headless: bool = True) -> pd.DataFrame:
    """Scrape Cushman & Wakefield pages, fall back to static data if scraping fails."""
    logger = setup_logger()
    logger.info(f"[C&W] Scraping {len(CW_URLS)} Cushman & Wakefield listings...")

    driver = create_driver(headless=headless)
    records = []

    try:
        for url in CW_URLS:
            record = _scrape_cw_page(driver, url, logger)
            if record:
                records.append(record)
    except Exception:
        logger.error(f"[C&W] Unexpected error:\n{traceback.format_exc()}")
    finally:
        driver.quit()

    # If scraping returned no data, use static hardcoded listings
    if not records:
        logger.warning("[C&W] Live scraping returned no data. Using static listings.")
        records = CW_STATIC_LISTINGS

    df = pd.DataFrame(records)
    df["Source"] = "Cushman & Wakefield"
    logger.info(f"[C&W] {len(df)} listings ready.")
    return df


def _scrape_cw_page(driver, url: str, logger) -> dict:
    """Scrape a single Cushman & Wakefield property detail page."""
    try:
        logger.info(f"[C&W] Loading: {url}")
        driver.get(url)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, "lxml")

        def _txt(sel):
            el = soup.select_one(sel)
            return el.get_text(" ", strip=True) if el else ""

        # Building name — h1 or main heading
        name = (_txt("h1") or _txt("h2") or "").strip()

        # Address — look for address-like text near the title
        address = ""
        for sel in ["[class*='address']", "[class*='location']", "[class*='subtitle']",
                    "[class*='Address']", "[class*='Location']"]:
            val = _txt(sel)
            if val and len(val) > 5:
                address = val
                break
        # Fallback: grab text after h1 that looks like an address
        if not address:
            h1 = soup.find("h1")
            if h1 and h1.find_next_sibling():
                address = h1.find_next_sibling().get_text(" ", strip=True)

        # Extract key details from the page text
        page_text = soup.get_text(" ", strip=True)

        def _extract_after(label: str, text: str) -> str:
            """Extract value that appears after a label in page text."""
            pattern = rf"{re.escape(label)}[:\s]*([^\n\r.]+)"
            m = re.search(pattern, text, re.IGNORECASE)
            return m.group(1).strip() if m else ""

        available_space  = _extract_after("Available Space", page_text)
        building_size    = _extract_after("Building Size", page_text)
        rental_price     = _extract_after("Rental Price", page_text)
        year_built       = _extract_after("Year Built", page_text)
        grade            = _extract_after(r"Grade[^A-Za-z]", page_text) or _extract_after("Grade", page_text)
        construction     = _extract_after("Construction Status", page_text)

        # Image
        img = soup.select_one("img[src*='cushmanwakefield'], img[src*='property'], img[class*='hero']")
        image_url = img["src"] if img else ""

        return {
            "Building Name":       name,
            "Property Type":       "Office/Commercial",
            "Address":             address,
            "Area / Size":         available_space,
            "Building Size":       building_size,
            "Rent":                rental_price,
            "Year Built":          year_built,
            "Grade":               grade,
            "Construction Status": construction,
            "City":                "Hyderabad",
            "Region":              "Telangana",
            "Property Details":    construction,
            "Property Link":       url,
            "Image URL":           image_url,
            "Listing ID":          url.rstrip("/").split("/")[-1],
        }
    except Exception as e:
        logger.warning(f"[C&W] Failed to scrape {url}: {e}")
        return {}


# ── Square Yards scraper ─────────────────────────────────────────────────────

# Square Yards — commercial office listings for Hyderabad only
SY_URLS = [
    "https://www.squareyards.com/sale/office-spaces-for-sale-in-hyderabad",
    "https://www.squareyards.com/rent/office-spaces-for-rent-in-hyderabad",
    "https://www.squareyards.com/rent/commercial-properties-for-rent-in-hyderabad",
    "https://www.squareyards.com/sale/commercial-properties-for-sale-in-hyderabad",
    "https://www.squareyards.com/kokapet-hyderabad-real-estate",
    "https://www.squareyards.com/sale/residential-properties-for-sale-in-kokapet-hyderabad",
    "https://www.squareyards.com/rent/residential-properties-for-rent-in-kokapet-hyderabad",
]

def scrape_squareyards(headless: bool = True) -> pd.DataFrame:
    logger = setup_logger()
    logger.info("[SY] Starting Square Yards Hyderabad commercial scrape...")

    driver = create_driver(headless=headless)
    all_records = []

    try:
        for sy_url in SY_URLS:
            logger.info(f"[SY] Loading: {sy_url}")
            driver.get(sy_url)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(8)

            # Scroll with retry
            last_height = driver.execute_script("return document.body.scrollHeight")
            no_change = 0
            for i in range(60):
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(3)
                for btn_text in ["LOAD MORE", "SHOW MORE", "VIEW MORE"]:
                    try:
                        btn = driver.find_element(
                            By.XPATH,
                            f"//button[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz',"
                            f"'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{btn_text}')]"
                        )
                        if btn.is_displayed():
                            driver.execute_script("arguments[0].click();", btn)
                            time.sleep(3)
                    except Exception:
                        pass
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    no_change += 1
                    if no_change >= 3:
                        logger.info(f"[SY] Scroll done after {i+1} scrolls for {sy_url}")
                        break
                    time.sleep(3)
                else:
                    no_change = 0
                last_height = new_height

            soup = BeautifulSoup(driver.page_source, "lxml")
            cards_data = _parse_sy_listings(soup)
            logger.info(f"[SY] Found {len(cards_data)} cards from {sy_url}")

            # Visit detail pages
            for i, card in enumerate(cards_data):
                url = card.get("Property Link", "")
                if url and url.startswith("http") and "squareyards.com" in url:
                    size, prop_type, details, detail_addr = _scrape_sy_detail(driver, url, logger)
                    card["Area / Size"]       = size
                    card["Property Type"]    = prop_type or card.get("Property Type", "Commercial")
                    card["Property Details"] = details
                    if not card.get("Address") or card["Address"] == "Hyderabad":
                        if detail_addr:
                            card["Address"] = detail_addr
                    logger.debug(f"[SY] {i+1}/{len(cards_data)} {card['Building Name']} → {size} | {card['Address']}")
                time.sleep(1)

            all_records.extend(cards_data)

    except Exception:
        logger.error(f"[SY] Error:\n{traceback.format_exc()}")
    finally:
        driver.quit()

    if not all_records:
        logger.warning("[SY] No listings found.")
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    df["Source"] = "Square Yards"
    df["City"]   = "Hyderabad"
    df["Region"] = "Telangana"

    # ── Filter: Commercial/Office OR Kokapet residential ──────────────────────
    PURELY_RESIDENTIAL = [
        "bhk flat", "bhk apartment", "bhk villa", "residential apartment",
        "residential project", "gated community", "township", "pg for rent",
        "independent house", "builder floor", "studio apartment",
    ]
    OFFICE_KEYWORDS = [
        "office space", "office spaces", "commercial space", "it park",
        "tech park", "business hub", "coworking", "co-working", "workspace",
        "sez", "cyber", "mindspace", "waverock", "salarpuria", "divyasree",
        "prestige sky", "rajapushpa", "avance business", "it hub",
        "commercial property", "commercial office", "office building",
    ]

    def _is_commercial_office(row):
        full_text = " ".join([
            str(row.get("Building Name", "")),
            str(row.get("Property Type", "")),
            str(row.get("Property Details", "")),
            str(row.get("Property Link", "")),
            str(row.get("Address", "")),
        ]).lower()

        # Always keep Kokapet residential listings
        if "kokapet" in full_text:
            return True
        # Reject purely residential outside Kokapet
        if any(kw in full_text for kw in PURELY_RESIDENTIAL):
            return False
        # Keep commercial/office
        if any(kw in full_text for kw in OFFICE_KEYWORDS):
            return True
        return True

    before = len(df)
    df = df[df.apply(_is_commercial_office, axis=1)].copy()
    logger.info(f"[SY] After commercial/Kokapet-residential filter: {len(df)}/{before} listings kept.")

    # Clean messy address (remove newlines, extra spaces)
    df["Address"] = df["Address"].apply(
        lambda v: re.sub(r"[\n\r\t]", " ", str(v)) if isinstance(v, str) else v
    )
    df["Address"] = df["Address"].apply(
        lambda v: re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v
    )

    # Remove duplicates within SY data
    df["_name_lower"] = df["Building Name"].str.strip().str.lower()
    df["_sqft"] = df["Area / Size"].apply(_parse_sqft)
    df = df.sort_values("_sqft", ascending=False)
    df = df.drop_duplicates(subset=["_name_lower"], keep="first")
    df = df.drop(columns=["_name_lower", "_sqft"]).reset_index(drop=True)

    logger.info(f"[SY] {len(df)} final listings ready.")
    return df


def _scrape_sy_detail(driver, url: str, logger) -> tuple:
    """Visit a Square Yards property detail page and extract size, type, details, address."""
    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, "lxml")
        text = soup.get_text(" ", strip=True)

        size = ""
        prop_type = ""
        details = ""
        address = ""

        # ── Address from detail page ─────────────────────────────────────────────
        for sel in ["[class*='project-address']", "[class*='locality']",
                    "[class*='location']", "[class*='address']",
                    "[class*='breadcrumb']", ".project-location",
                    "[class*='projectLocation']"]:
            el = soup.select_one(sel)
            if el:
                t = el.get_text(" ", strip=True)
                if t and "hyderabad" in t.lower() and len(t) < 150:
                    address = t
                    break
        # Fallback: regex from page text
        if not address:
            m = re.search(r"in\s+([A-Za-z\s]+),\s*Hyderabad", text)
            if m:
                address = f"{m.group(1).strip()}, Hyderabad"
        if not address:
            m = re.search(r"([A-Z][a-zA-Z\s]+),\s*Hyderabad", text)
            if m:
                address = m.group(0).strip()

        # ── Size ────────────────────────────────────────────────────────────────────────
        size_patterns = [
            r"Size[:\s]+([\d,\.]+\s*(?:to|-|–)\s*[\d,\.]+\s*Sq\.?\s*Ft\.?)",
            r"Size[:\s]+([\d,\.]+\s*Sq\.?\s*Ft\.?)",
            r"([\d,\.]+\s*(?:to|-|–)\s*[\d,\.]+\s*Sq\.?\s*Ft\.?)",
            r"([\d,\.]+\s*Sq\.?\s*Ft\.?)",
            r"Total\s+area[:\s]+([\d\.]+\s*Acres?)",
            r"spreads?\s+across\s+([\d\.]+\s*Acres?)",
        ]
        for pat in size_patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                raw = m.group(1).strip()
                nums = re.findall(r"[\d]+", raw.replace(",", ""))
                if nums:
                    try:
                        highest = max(float(n) for n in nums)
                        unit = "Acres" if "acre" in raw.lower() else "Sq. Ft"
                        size = f"{highest:,.0f} {unit}"
                    except Exception:
                        size = raw
                break

        # ── Property type ─────────────────────────────────────────────────────
        for sel in ["[class*='projectType']", "[class*='propertyType']",
                    "[class*='unitConfig']", "[class*='bhk']"]:
            el = soup.select_one(sel)
            if el:
                prop_type = el.get_text(" ", strip=True)[:60]
                break
        if not prop_type:
            m = re.search(r"(\d\s*BHK|Commercial|Office|Apartment|Villa|Plot)",
                          text, re.IGNORECASE)
            if m:
                prop_type = m.group(1).strip()

        # ── Short details summary ─────────────────────────────────────────────
        for sel in ["[class*='aboutProject']", "[class*='projectDesc']",
                    "[class*='description']", "[class*='overview']"]:
            el = soup.select_one(sel)
            if el:
                details = el.get_text(" ", strip=True)[:300]
                break

        return size, prop_type, details, address

    except Exception as e:
        logger.debug(f"[SY] Detail page failed {url}: {e}")
        return "", "", "", ""


def _parse_sy_listings(soup) -> list:
    """Parse Square Yards commercial listing cards using exact selectors from HTML."""
    records = []
    base = "https://www.squareyards.com"

    # Commercial pages use article.listing-card (confirmed from HTML dump)
    cards = soup.select("article.listing-card")
    if not cards:
        # fallback for project pages
        cards = soup.select("article.property-tile, article.focus-card")

    for card in cards:
        # ── URL: data-url on .listing-body ───────────────────────────────────────────────────────────────────
        prop_link = ""
        body = card.select_one(".listing-body")
        if body:
            prop_link = body.get("data-url", "")
        if not prop_link:
            item = card.select_one(".item[data-href]")
            if item:
                prop_link = item.get("data-href", "")
        if not prop_link:
            # project card fallback: onclick goToURL
            onclick = card.get("onclick", "")
            m = re.search(r"goToURL\(['\"]([^'\"]+)['\"]", onclick)
            if m:
                prop_link = m.group(1)
        if prop_link and not prop_link.startswith("http"):
            prop_link = base + "/" + prop_link.lstrip("/")

        # ── Name: h2.heading span (commercial) or h2/h3 (project) ────────────────────────────────────────────
        name = ""
        for sel in ["h2.heading span", "h2.heading a", "h2", "h3",
                    "[class*='project-name']", "[class*='projectName']"]:
            el = card.select_one(sel)
            if el:
                t = el.get_text(strip=True)
                if t and len(t) < 120 and "View All" not in t:
                    name = t
                    break
        if not name:
            fav = card.select_one(".favorite-btn")
            if fav:
                name = fav.get("data-propname", "") or fav.get("data-name", "")
        if not name:
            img = card.select_one("img")
            if img:
                name = img.get("alt", "")
        if not name or "View All" in name:
            continue

        # ── Address: p.location span (commercial) or locality selectors ────────────────────────────────────────────
        address = ""
        for sel in ["p.location span", "[class*='locality']",
                    "[class*='location']", "[class*='address']",
                    "figcaption p", ".project-address"]:
            el = card.select_one(sel)
            if el:
                t = el.get_text(" ", strip=True)
                if t and 3 < len(t) < 120:
                    address = t
                    break
        if not address:
            fav = card.select_one(".favorite-btn")
            if fav:
                address = fav.get("data-locality", "")
        if not address and prop_link:
            m = re.search(r"-([a-z][a-z-]+)-hyderabad", prop_link.lower())
            if m:
                address = m.group(1).replace("-", " ").title() + ", Hyderabad"
        if not address:
            address = "Hyderabad"
        address = re.sub(r"\s+", " ", address).strip()

        # ── Area: data-area on .avail-area (commercial) ───────────────────────────────────────────────────────────────────
        size = ""
        area_el = card.select_one(".avail-area")
        if area_el:
            area_val = area_el.get("data-area", "")
            unit_el  = area_el.select_one(".unit-label")
            unit     = unit_el.get_text(strip=True) if unit_el else "Sq.Ft."
            if area_val:
                try:
                    size = f"{int(float(area_val)):,} {unit}"
                except Exception:
                    size = f"{area_val} {unit}"
        if not size:
            fav = card.select_one(".favorite-btn")
            if fav:
                size = fav.get("data-area", "")

        # ── Price: p.listing-price strong (commercial) ───────────────────────────────────────────────────────────────────
        price = ""
        price_el = card.select_one("p.listing-price strong")
        if price_el:
            price = price_el.get_text(strip=True)
        if not price:
            fav = card.select_one(".favorite-btn")
            if fav:
                price = fav.get("data-totalprice", "")
        if not price:
            m = re.search(r"(₹[\s\d,\.]+(?:Cr|Lac|L|K)?|Price\s+On\s+Request)",
                          card.get_text(" "), re.IGNORECASE)
            if m:
                price = m.group(1).strip()

        # ── Property type ──────────────────────────────────────────────────────────────────────────────────────
        prop_type = "Office Space"
        fav = card.select_one(".favorite-btn")
        if fav:
            prop_type = (fav.get("data-unittype") or
                         fav.get("data-propertytype") or "Office Space")

        # ── Image ──────────────────────────────────────────────────────────────────────────────────────
        img = card.select_one("img.img-responsive")
        image_url = img.get("src", "") if img else ""

        # ── Description ────────────────────────────────────────────────────────────────────────────────
        desc_el = card.select_one(".description p")
        details = desc_el.get_text(" ", strip=True)[:300] if desc_el else ""

        records.append({
            "Building Name":    name,
            "Property Type":    prop_type,
            "Address":          address,
            "Area / Size":      size,
            "Rent":             price,
            "Property Details": details,
            "Property Link":    prop_link,
            "Image URL":        image_url,
            "Listing ID":       prop_link.rstrip("/").split("/")[-1] if prop_link else "",
        })

    return records

    for card in cards:
        # ── URL: extract from onclick="helperJS.goToURL('URL', ...)" ────────────
        prop_link = ""
        onclick = card.get("onclick", "")
        m = re.search(r"goToURL\(['\"]([^'\"]+)['\"]", onclick)
        if m:
            href = m.group(1)
            prop_link = href if href.startswith("http") else base + "/" + href.lstrip("/")
        # Also check figure onclick
        if not prop_link:
            fig = card.select_one("figure[onclick]")
            if fig:
                m = re.search(r"goToURL\(['\"]([^'\"]+)['\"]", fig.get("onclick", ""))
                if m:
                    href = m.group(1)
                    prop_link = href if href.startswith("http") else base + "/" + href.lstrip("/")
        # Fallback: <a href>
        if not prop_link:
            a = card.find("a", href=True)
            if a:
                href = a["href"]
                prop_link = href if href.startswith("http") else base + href

        # Skip non-Hyderabad
        if "hyderabad" not in prop_link.lower() and "hyderabad" not in card.get_text().lower():
            continue

        # ── Image ────────────────────────────────────────────────────────────────────
        img = card.find("img")
        image_url = ""
        if img:
            image_url = img.get("src") or img.get("data-src") or ""

        # ── Name ────────────────────────────────────────────────────────────────────────
        name = ""
        for sel in ["h2.project-name", "h2", "h3", "figcaption h2",
                    "[class*='project-name']", "[class*='projectName']",
                    "[class*='title']", "a[title]"]:
            el = card.select_one(sel)
            if el:
                t = el.get("title") or el.get_text(strip=True)
                if t and len(t) < 100 and "View All" not in t:
                    name = t
                    break
        if not name:
            # Try img alt
            if img and img.get("alt"):
                name = img["alt"]

        if not name or "View All" in name:
            continue

        # ── Address ─────────────────────────────────────────────────────────────────────
        address = ""
        for sel in ["[class*='locality']", "[class*='location']",
                    "[class*='address']", "[class*='subTitle']",
                    "[class*='project-location']", "span.locality",
                    "p.locality", ".content-left p", ".figure-body p",
                    "figcaption p", ".project-address"]:
            el = card.select_one(sel)
            if el:
                t = el.get_text(" ", strip=True)
                if t and 3 < len(t) < 120 and "hyderabad" in t.lower():
                    address = t
                    break

        # Fallback 1: extract locality from URL pattern
        # URLs like: /hyderabad-residential-property/brigade-gateway/324984/project
        # or: /prestige-golden-grove-tellapur-hyderabad-npd-343925
        if not address and prop_link:
            # Pattern 1: /project-name-locality-hyderabad-...
            m = re.search(r"-([a-z][a-z-]+)-hyderabad", prop_link.lower())
            if m:
                locality = m.group(1).replace("-", " ").title()
                address = f"{locality}, Hyderabad"

        # Fallback 2: extract from card text — look for "Locality, Hyderabad" pattern
        if not address:
            m = re.search(r"([A-Z][a-zA-Z\s]+),\s*Hyderabad", card.get_text(" "))
            if m:
                address = m.group(0).strip()

        # Fallback 3: check img alt or title attributes for location hints
        if not address and img:
            alt = img.get("alt", "") or img.get("title", "")
            m = re.search(r"([A-Z][a-zA-Z\s]+),?\s*Hyderabad", alt)
            if m:
                address = m.group(0).strip()

        if not address:
            address = "Hyderabad"

        # Clean address — remove newlines and extra spaces
        address = re.sub(r"[\n\r\t]", " ", address)
        address = re.sub(r"\s+", " ", address).strip()

        # ── Price ────────────────────────────────────────────────────────────────────────
        price = ""
        for sel in ["[class*='price']", "[class*='Price']",
                    "[class*='amount']", "[class*='cost']",
                    ".price-section", ".project-price"]:
            el = card.select_one(sel)
            if el:
                t = el.get_text(" ", strip=True)
                if t and ("₹" in t or "Rs" in t or "Cr" in t
                          or "Lac" in t or "request" in t.lower()):
                    price = t
                    break
        if not price:
            m = re.search(r"(₹[\s\d,\.]+(?:Cr|Lac|L|K)?|Price\s+On\s+Request)",
                          card.get_text(" "), re.IGNORECASE)
            if m:
                price = m.group(1).strip()

        records.append({
            "Building Name":    name,
            "Property Type":    "Commercial",
            "Address":          address,
            "Area / Size":      "",   # filled by detail page scraper
            "Rent":             price,
            "Property Details": "",
            "Property Link":    prop_link,
            "Image URL":        image_url,
            "Listing ID":       prop_link.rstrip("/").split("/")[-1],
        })

    return records


# ── Kokapet Residential scraper ─────────────────────────────────────────────

# Well-known Kokapet residential projects — static seed data
KOKAPET_RESIDENTIAL_STATIC = [
    {
        "Building Name":    "My Home Avatar",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Ultra-luxury residential towers at Kokapet.",
        "Property Link":    "https://www.squareyards.com/my-home-avatar-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "my-home-avatar-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Prestige Plots Kokapet",
        "Property Type":    "Residential Plot",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Gated residential plotted development at Kokapet.",
        "Property Link":    "https://www.squareyards.com/prestige-plots-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "prestige-plots-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Aparna Zenon",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Premium residential apartments at Kokapet.",
        "Property Link":    "https://www.squareyards.com/aparna-zenon-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "aparna-zenon-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Lodha Hyderabad",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Luxury residential project by Lodha at Kokapet.",
        "Property Link":    "https://www.lodhagroup.com/hyderabad",
        "Image URL":        "",
        "Listing ID":       "lodha-hyderabad-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Rajapushpa Atria",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "High-rise residential towers at Kokapet by Rajapushpa.",
        "Property Link":    "https://www.squareyards.com/rajapushpa-atria-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "rajapushpa-atria-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Incor One City",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Integrated township at Kokapet by Incor.",
        "Property Link":    "https://www.squareyards.com/incor-one-city-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "incor-one-city-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Bhavana Celestia",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Premium gated community at Kokapet.",
        "Property Link":    "https://www.squareyards.com/bhavana-celestia-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "bhavana-celestia-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "NSL Arena",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Residential apartments at Kokapet by NSL.",
        "Property Link":    "https://www.squareyards.com/nsl-arena-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "nsl-arena-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "Phoenix Kessaku",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Ultra-luxury high-rise residential at Kokapet by Phoenix.",
        "Property Link":    "https://www.squareyards.com/phoenix-kessaku-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "phoenix-kessaku-kokapet",
        "Source":           "Residential",
    },
    {
        "Building Name":    "My Home Tridasa",
        "Property Type":    "Residential Apartment",
        "Address":          "Kokapet, Hyderabad, Telangana",
        "Area / Size":      "",
        "Rent":             "Price on Request",
        "City":             "Hyderabad",
        "Region":           "Telangana",
        "Property Details": "Luxury residential towers at Kokapet by My Home Group.",
        "Property Link":    "https://www.squareyards.com/my-home-tridasa-kokapet-hyderabad",
        "Image URL":        "",
        "Listing ID":       "my-home-tridasa-kokapet",
        "Source":           "Residential",
    },
]


def scrape_kokapet_residential(headless: bool = True) -> pd.DataFrame:
    """
    Scrape Kokapet residential listings from Square Yards project pages.
    Falls back to static seed data if live scraping yields nothing.
    """
    logger = setup_logger()
    logger.info("[RES] Scraping Kokapet residential listings...")

    urls = [
        "https://www.squareyards.com/sale/residential-properties-for-sale-in-kokapet-hyderabad",
        "https://www.squareyards.com/rent/residential-properties-for-rent-in-kokapet-hyderabad",
        "https://www.squareyards.com/kokapet-hyderabad-real-estate",
    ]

    driver = create_driver(headless=headless)
    records = []

    try:
        for url in urls:
            logger.info(f"[RES] Loading: {url}")
            driver.get(url)
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
            except Exception:
                pass
            time.sleep(6)

            # scroll a few times to load cards
            last_h = driver.execute_script("return document.body.scrollHeight")
            for _ in range(15):
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2.5)
                nh = driver.execute_script("return document.body.scrollHeight")
                if nh == last_h:
                    break
                last_h = nh

            soup = BeautifulSoup(driver.page_source, "lxml")
            found = _parse_kokapet_res_cards(soup, url)
            logger.info(f"[RES] Found {len(found)} cards from {url}")
            records.extend(found)
    except Exception:
        logger.error(f"[RES] Error:\n{traceback.format_exc()}")
    finally:
        driver.quit()

    # Deduplicate by name
    seen = set()
    unique = []
    for r in records:
        key = r["Building Name"].strip().lower()
        if key and key not in seen:
            seen.add(key)
            unique.append(r)

    if not unique:
        logger.warning("[RES] Live scrape returned nothing. Using static Kokapet residential seed.")
        unique = KOKAPET_RESIDENTIAL_STATIC
    else:
        # Merge static entries not already found
        static_names = {r["Building Name"].strip().lower() for r in KOKAPET_RESIDENTIAL_STATIC}
        extra = [r for r in KOKAPET_RESIDENTIAL_STATIC if r["Building Name"].strip().lower() not in seen]
        unique.extend(extra)
        logger.info(f"[RES] Added {len(extra)} static seed entries not found live.")

    df = pd.DataFrame(unique)
    df["Source"] = "Residential"
    df["City"]   = "Hyderabad"
    df["Region"] = "Telangana"
    logger.info(f"[RES] {len(df)} Kokapet residential listings ready.")
    return df


def _parse_kokapet_res_cards(soup, source_url: str) -> list:
    """Parse project/listing cards from Kokapet residential pages."""
    records = []
    base = "https://www.squareyards.com"

    # Try listing cards first, then project tiles
    cards = soup.select("article.listing-card")
    if not cards:
        cards = soup.select("article.property-tile, article.focus-card, [class*='project-card'], [class*='projectCard']")

    for card in cards:
        # Name
        name = ""
        for sel in ["h2.heading span", "h2.heading a", "h2", "h3",
                    "[class*='project-name']", "[class*='projectName']"]:
            el = card.select_one(sel)
            if el:
                t = el.get_text(strip=True)
                if t and len(t) < 120 and "View All" not in t:
                    name = t
                    break
        if not name:
            fav = card.select_one(".favorite-btn")
            if fav:
                name = fav.get("data-propname", "") or fav.get("data-name", "")
        if not name:
            img = card.select_one("img")
            if img:
                name = img.get("alt", "")
        if not name or "View All" in name:
            continue

        # Link
        prop_link = ""
        body = card.select_one(".listing-body")
        if body:
            prop_link = body.get("data-url", "")
        if not prop_link:
            a = card.select_one("a[href]")
            if a:
                prop_link = a["href"]
        if prop_link and not prop_link.startswith("http"):
            prop_link = base + "/" + prop_link.lstrip("/")

        # Address
        address = "Kokapet, Hyderabad"
        for sel in ["p.location span", "[class*='locality']", "[class*='location']",
                    "[class*='address']", ".project-address"]:
            el = card.select_one(sel)
            if el:
                t = el.get_text(" ", strip=True)
                if t and 3 < len(t) < 120:
                    address = t
                    if "kokapet" not in address.lower():
                        address = "Kokapet, " + address
                    break

        # Price
        price = "Price on Request"
        price_el = card.select_one("p.listing-price strong")
        if price_el:
            price = price_el.get_text(strip=True)
        if not price or price == "Price on Request":
            fav = card.select_one(".favorite-btn")
            if fav and fav.get("data-totalprice"):
                price = fav["data-totalprice"]

        # Size
        size = ""
        area_el = card.select_one(".avail-area")
        if area_el:
            av = area_el.get("data-area", "")
            unit_el = area_el.select_one(".unit-label")
            unit = unit_el.get_text(strip=True) if unit_el else "Sq.Ft."
            if av:
                try:
                    size = f"{int(float(av)):,} {unit}"
                except Exception:
                    size = f"{av} {unit}"

        # Property type
        prop_type = "Residential"
        fav = card.select_one(".favorite-btn")
        if fav:
            prop_type = fav.get("data-unittype") or fav.get("data-propertytype") or "Residential"

        # Image
        img = card.select_one("img.img-responsive, img")
        image_url = img.get("src", "") if img else ""

        records.append({
            "Building Name":    name,
            "Property Type":    prop_type,
            "Address":          address,
            "Area / Size":      size,
            "Rent":             price,
            "Property Details": "",
            "Property Link":    prop_link,
            "Image URL":        image_url,
            "Listing ID":       prop_link.rstrip("/").split("/")[-1] if prop_link else "",
            "Source":           "Residential",
            "City":             "Hyderabad",
            "Region":           "Telangana",
        })

    return records


# ── Combined scraper ──────────────────────────────────────────────────────────

def scrape_all(city: str = "Hyderabad", headless: bool = True, max_scrolls: int = 50) -> pd.DataFrame:
    logger = setup_logger()
    ensure_output_dir(OUTPUT_DIR)

    df_jll  = scrape_jll(city=city, headless=headless, max_scrolls=max_scrolls)
    df_sy   = scrape_squareyards(headless=headless)
    df_cbre = scrape_cbre(city=city)
    df_res  = scrape_kokapet_residential(headless=headless)

    # Always use static C&W data
    logger.info("[C&W] Loading static Cushman & Wakefield listings...")
    df_cw = pd.DataFrame(CW_STATIC_LISTINGS)
    logger.info(f"[C&W] {len(df_cw)} listings ready.")

    # Align all columns
    all_cols = list(dict.fromkeys(
        list(df_jll.columns if not df_jll.empty else []) +
        list(df_cw.columns) +
        list(df_sy.columns if not df_sy.empty else []) +
        list(df_cbre.columns if not df_cbre.empty else []) +
        list(df_res.columns if not df_res.empty else [])
    ))
    for df in [df_jll, df_cw, df_sy, df_cbre, df_res]:
        for col in all_cols:
            if col not in df.columns:
                df[col] = ""

    df_combined = pd.concat(
        [d[all_cols] for d in [df_jll, df_cw, df_sy, df_cbre, df_res] if not d.empty],
        ignore_index=True
    )
    df_combined = apply_combined_rules(df_combined)

    logger.info(f"Combined: JLL={len(df_jll)}, C&W={len(df_cw)}, SY={len(df_sy)}, CBRE={len(df_cbre)}, Residential={len(df_res)}, Total={len(df_combined)}")

    sheets = {}
    if not df_jll.empty:   sheets["JLL"]                 = df_jll
    sheets["Cushman & Wakefield"]                        = df_cw
    if not df_sy.empty:    sheets["Square Yards"]         = df_sy
    if not df_cbre.empty:  sheets["CBRE"]                 = df_cbre
    if not df_res.empty:   sheets["Kokapet Residential"]  = df_res
    sheets["Combined"]                                   = df_combined

    excel_path = save_excel_multi(sheets, OUTPUT_DIR, city)
    csv_path   = save_csv(df_combined, OUTPUT_DIR, city, prefix="Combined")
    logger.info(f"Excel saved → {excel_path}")
    logger.info(f"CSV   saved → {csv_path}")
    return df_combined


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="JLL + C&W + Square Yards + CBRE Office Listings Scraper")
    parser.add_argument("--city",        default="Hyderabad", choices=list(CITY_REGION_MAP.keys()))
    parser.add_argument("--no-headless", action="store_true")
    parser.add_argument("--max-scrolls", type=int, default=50)
    parser.add_argument("--jll-only",    action="store_true")
    parser.add_argument("--cw-only",     action="store_true")
    parser.add_argument("--sy-only",     action="store_true", help="Scrape Square Yards only")
    parser.add_argument("--cbre-only",   action="store_true", help="Scrape CBRE only")
    args = parser.parse_args()

    headless = not args.no_headless

    if args.jll_only:
        df = scrape_jll(city=args.city, headless=headless, max_scrolls=args.max_scrolls)
    elif args.cw_only:
        df = scrape_cw(headless=headless)
    elif args.sy_only:
        df = scrape_squareyards(headless=headless)
    elif args.cbre_only:
        df = scrape_cbre(city=args.city)
    else:
        df = scrape_all(city=args.city, headless=headless, max_scrolls=args.max_scrolls)

    if not df.empty:
        print(f"\n✅ Done! {len(df)} total listings.")
        cols = [c for c in ["Building Name", "Address", "Area / Size", "Rent", "Source"] if c in df.columns]
        print(df[cols].head(15).to_string(index=False))
    else:
        print("\n⚠️  No data extracted. Check scraper.log for details.")


if __name__ == "__main__":
    main()
