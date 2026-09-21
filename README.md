# Manufacturer Product Scrapers

One script per manufacturer. Each script runs all of that manufacturer's scrapers **back to back**: when one section (or brand) finishes and its file is saved, the next one starts automatically, each in a fresh Chrome window.

| Script | Manufacturer | Runs | Saves |
|---|---|---|---|
| `acuity_all_scrapers.py` | Acuity Brands | 8 sections | 8 Excel files (`.xlsx`) |
| `cooper_lighting_all_scrapers.py` | Cooper Lighting | 8 brands | 8 CSV files |
| `milesight_scraper.py` | Milesight | 1 scraper (already a single script) | `milesight_products.csv` + `scraper_debug.log` |

## Requirements

- Python 3 and Google Chrome (the Acuity and Cooper scripts were tested with Python 3.13, Selenium 4.40 and Chrome 153)
- Packages:

```bash
pip install selenium openpyxl webdriver-manager
```

Acuity needs all three, Cooper needs only `selenium`, Milesight needs `selenium` and `webdriver-manager`.

## How to run

```bash
python3 acuity_all_scrapers.py
python3 cooper_lighting_all_scrapers.py
python3 milesight_scraper.py
```

Output files are saved in the folder you run the script from. These are long runs, so on a Mac you can stop it going to sleep with:

```bash
caffeinate -i python3 acuity_all_scrapers.py
```

## Acuity: `acuity_all_scrapers.py`

Collects five columns for every product: **Product Name, Brand, Product Series, Overview Text, Product URL**.

For each category it scrolls the page until everything has loaded, collects the product links, then opens every product page to read the name, the brand (from the URL), the Overview text (expanding "read more") and the Series (from the Specifications tab).

| # | Section | Categories | Output file |
|---|---|---|---|
| 1 | Confinement & Vandal | 10 | `acuity_confinement_vandal_products.xlsx` |
| 2 | Controls | 16 | `acuity_controls_products.xlsx` |
| 3 | DMX Networking | 9 | `acuity_dmx_networking_products.xlsx` |
| 4 | Industrial | 6 | `acuity_industrial_products.xlsx` |
| 5 | Life Safety | 8 | `acuity_life_safety_products.xlsx` |
| 6 | New Products | single page | `acuity_new_products.xlsx` |
| 7 | Residential Indoor | 15 | `acuity_residential_indoor_products.xlsx` |
| 8 | Residential | 14 | `acuity_residential_products.xlsx` |

- New Products is one long page, so it scrolls more slowly (4 seconds between scrolls) to let the whole list load.
- In Residential, the Downlights category clicks "View All" first.
- The old scripts mostly saved to the same file name (`acuity_products.xlsx`), so each section now has its own file and nothing overwrites anything.
- "Lighting Controls" was identical to "Controllers", and the standalone `acuity_scraper.py` was identical to "Residential Indoor", so each of those runs once.

## Cooper Lighting: `cooper_lighting_all_scrapers.py`

Collects four columns for every stock catalog item: **Name, Description, UPCCode, URL**.

For each brand it opens the brand's product list on cooperlighting.com, scrolls until all products have loaded, opens each product's configuration page, follows every link in its "Stock Catalog Number" table, and reads the name, description and UPC from each catalog page.

| # | Brand | Output file |
|---|---|---|
| 1 | Trellix Infrastructure | `trellix_infrastructure_catalog.csv` |
| 2 | WaveLinx | `wavelinx_catalog.csv` |
| 3 | MWS | `mws_catalog.csv` |
| 4 | Streetworks | `streetworks_catalog.csv` |
| 5 | Portfolio | `portfolio_catalog.csv` |
| 6 | RSA | `rsa_catalog.csv` |
| 7 | Metalux | `metalux_catalog.csv` |
| 8 | Halo | `halo_catalog.csv` |

- The order is quick brands first and the biggest (Halo, about 900 items last time) last.
- Every brand uses the same UPC lookup, which skips table rows that have no cells instead of giving up on the whole table. This is the version Halo and Metalux already used.
- If a brand's product list is empty on the site, its CSV contains only the header row and a warning is printed.
- Cooper folder names do not all match their code: the file in `Prentalux info` scrapes **Portfolio**, and the file in `Shaper` scrapes **RSA**. `Halo_scraper.py` was a duplicate of the WaveLinx scraper.

## Milesight: `milesight_scraper.py`

Already a single script, so this is an unchanged copy. It runs headless, walks the IoT Products menu (main categories, then subcategories, then products), and writes `milesight_products.csv` with **main_category, subcategory, model_name, product_url, description**, plus a `scraper_debug.log`.

## When something goes wrong

- A page that fails to load is skipped with a warning and the run carries on.
- If 5 pages fail in a row (usually a crashed browser), that section or brand stops, everything collected so far is saved, and the next one starts with a fresh browser.
- If a section cannot start at all, it is reported and the rest still run.
- Pressing Ctrl+C stops the whole run, and the file being written is saved first.
- A summary at the end lists anything that stopped early.

## Tips

- To skip a section or brand, put a `#` in front of every line of its entry in `SECTIONS` (Acuity) or `BRANDS` (Cooper). To re-run just one, comment out the others.
- UPCs are saved exactly as the site shows them, leading zero included. Excel drops leading zeros if you open and re-save the CSV, so import the UPC column as Text.

## Not included

- The `Kelvix Scrapper` and `Philips Lighting` folders were empty, so there was nothing to combine for them yet.
- The `NewRay info` and `SureLites` Cooper folders were also empty.
