# JLL India — Commercial Office Listings Scraper

Scrapes commercial office listings from [JLL India Property](https://property.jll.co.in) and exports them to Excel and CSV.

---

## Project Structure

```
web scrapping/
├── scraper.py          # Main scraper (Selenium + BeautifulSoup)
├── utils.py            # Save helpers, Excel formatting, logging
├── requirements.txt    # Python dependencies
├── README.md
├── scraper.log         # Auto-generated log file
└── output/             # Auto-created; Excel + CSV files saved here
```

---

## Installation

### 1. Prerequisites
- Python 3.10+
- Google Chrome browser installed

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

---

## Usage

### Default — Hyderabad (headless)
```bash
python scraper.py
```

### Other cities
```bash
python scraper.py --city Mumbai
python scraper.py --city Bangalore
python scraper.py --city Chennai
python scraper.py --city Delhi
python scraper.py --city Pune
```

### Visible browser (useful for debugging)
```bash
python scraper.py --no-headless
```

### Limit scroll depth
```bash
python scraper.py --max-scrolls 20
```

### Combine options
```bash
python scraper.py --city Mumbai --no-headless --max-scrolls 30
```

---

## Output

Files are saved in the `output/` folder with timestamps:

```
output/
├── JLL_Hyderabad_offices_20241201_143022.xlsx
└── JLL_Hyderabad_offices_20241201_143022.csv
```

### Excel features
- Frozen header row with dark-blue styling
- Alternating row shading
- Auto-adjusted column widths
- Metadata sheet (source, city, timestamp, count)

### Columns extracted

| Column           | Description                        |
|------------------|------------------------------------|
| Building Name    | Name of the property/building      |
| Property Type    | Office / Commercial                |
| Address          | Full address / locality            |
| Area / Size      | Carpet area or total area          |
| Rent             | Listed rent / price                |
| City             | Target city                        |
| Region           | State / region                     |
| Property Details | Additional description             |
| Property Link    | Direct URL to the listing          |
| Image URL        | Thumbnail image URL                |
| Listing ID       | Numeric ID extracted from URL      |

---

## Example Output (console)

```
2024-12-01 14:30:22 [INFO] Target URL: https://property.jll.co.in/search?...
2024-12-01 14:30:25 [INFO] Page loaded. Starting scroll...
2024-12-01 14:30:55 [INFO] Reached end of page after 18 scrolls.
2024-12-01 14:30:55 [INFO] Parsing HTML...
2024-12-01 14:30:56 [INFO] Extracted 87 raw listings.
2024-12-01 14:30:56 [INFO] Clean listings after dedup: 84
2024-12-01 14:30:56 [INFO] CSV   saved → output/JLL_Hyderabad_offices_20241201_143022.csv
2024-12-01 14:30:57 [INFO] Excel saved → output/JLL_Hyderabad_offices_20241201_143022.xlsx

✅ Done! 84 listings scraped.
Building Name          Address                  Area / Size    Rent
Raheja Mindspace       HITEC City, Hyderabad    50,000 sq ft   ₹65/sq ft
...
```

---

## Error Handling Guide

| Symptom | Likely Cause | Fix |
|---|---|---|
| `No listings found` | JLL changed CSS class names | Run with `--no-headless` and inspect the page; update selectors in `parse_listings()` |
| `Failed to load page` | Network issue or bot detection | Increase `SCROLL_PAUSE` in `scraper.py`; try `--no-headless` |
| `ChromeDriver error` | Chrome version mismatch | `webdriver-manager` auto-fixes this; ensure Chrome is installed |
| Empty `Building Name` | Selector mismatch | Check `scraper.log` and update `_parse_card()` selectors |
| Partial data | Page didn't fully load | Increase `--max-scrolls` or `SCROLL_PAUSE` constant |

---

## Scaling to Other Cities

The `CITY_REGION_MAP` dict in `scraper.py` maps cities to regions. To add a new city:

```python
CITY_REGION_MAP = {
    "Hyderabad": "Telangana",
    "NewCity":   "StateRegion",   # ← add here
}
```

Then run:
```bash
python scraper.py --city NewCity
```

---

## Notes

- The scraper uses a realistic Chrome user-agent and disables the `webdriver` flag to reduce bot detection.
- All logs are written to `scraper.log` in the project root.
- Duplicate listings (same `Property Link`) are automatically removed.
