# =====================================================
# ACUITY BRANDS - ALL-IN-ONE SCRAPER
# =====================================================
# Runs every Acuity scraper back to back. Each entry in SECTIONS (below) is one
# of the original scripts - same URLs, same scraping steps, same 5 columns.
# When a section finishes, its Excel file is saved, its browser is closed, and
# the next section starts on its own.
#
#   #  Section                Output file
#   1  Confinement & Vandal   acuity_confinement_vandal_products.xlsx
#   2  Controls               acuity_controls_products.xlsx
#   3  DMX Networking         acuity_dmx_networking_products.xlsx
#   4  Industrial             acuity_industrial_products.xlsx
#   5  Life Safety            acuity_life_safety_products.xlsx
#   6  New Products           acuity_new_products.xlsx
#   7  Residential Indoor     acuity_residential_indoor_products.xlsx
#   8  Residential            acuity_residential_products.xlsx
#
# Files are saved in the folder you run the script from.
# To skip a section, put a # in front of every line of its entry in SECTIONS.
# =====================================================
import time
import openpyxl
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urlparse

# =====================================================
# LOGGING
# =====================================================
def log(msg):
    print(f"[INFO] {msg}", flush=True)

def warn(msg):
    print(f"[WARN] {msg}", flush=True)

# =====================================================
# EXCEL-SAFE TEXT SANITIZATION
# =====================================================
ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

def sanitize_for_excel(text):
    if not text:
        return ""
    text = ILLEGAL_XML_CHARS.sub("", text)
    return text.strip()

# =====================================================
# SECTIONS (one per original script, run in this order)
# =====================================================
# base_url           page the section starts on (categories are base_url + path)
# categories         name -> path; None means the base_url page itself is the list
# output_file        Excel file this section saves
# sheet_title        worksheet name inside that file
# scroll_pause       seconds to wait between scrolls (default 2)
# view_all_categories  categories where the "View All" link is clicked first
SECTIONS = [
    {
        "name": "Confinement & Vandal",
        "base_url": "https://www.acuitybrands.com/products/confinement-vandal",
        "categories": {
            "Confinement": "/confinement",
            "Downlights": "/downlights",
            "Emergency Exit Sign": "/emergency-exit-sign",
            "Flat Panel": "/flat-panel",
            "Floodlight": "/floodlight",
            "Linear": "/linear",
            "Sconce": "/sconce",
            "Surface Light": "/surface-light",
            "Troffer": "/troffer",
            "Wallpack": "/wallpack",
        },
        "output_file": "acuity_confinement_vandal_products.xlsx",
        "sheet_title": "Acuity Products",
    },
    # Covers both "Conrollers Products" and "Lighting Controlls Products" - the
    # two original scripts were identical, so it only needs to run once.
    {
        "name": "Controls",
        "base_url": "https://www.acuitybrands.com/products/controls",
        "categories": {
            "Accessories": "/accessories",
            "Controllers": "/controllers",
            "Daylight Sensors": "/daylight-sensors",
            "Embedded Controls": "/embedded-controls",
            "Load Controllers": "/load-controllers",

            # Luminaires with Embedded Controls (SUBCATEGORIES)
            "Luminaires – Wired Networked (nLight)": "/luminaires-with-embedded-controls/wired-networked-nlight",
            "Luminaires – Wireless Networked (nLight)": "/luminaires-with-embedded-controls/wireless-networked-nlight",
            "Luminaires – Wired Standalone (nLight)": "/luminaires-with-embedded-controls/wired-standalone-nlight",
            "Luminaires – Wireless Standalone (SensorSwitch)": "/luminaires-with-embedded-controls/wireless-standalone-sensorswitch",

            "Occupancy Sensors": "/occupancy-sensors",
            "Outdoor Photocontrols and Sensors": "/outdoor-photocontrols-sensors",
            "Panels": "/panels",
            "Services": "/services",
            "Software and Mobile Apps": "/software-mobile-apps",
            "Wall Stations and Switches": "/wall-stations-switches",
            "Bringing Controls to Light": "/bringing-controls-to-light",
        },
        "output_file": "acuity_controls_products.xlsx",
        "sheet_title": "Acuity Products",
    },
    {
        "name": "DMX Networking",
        "base_url": "https://www.acuitybrands.com/products/dmx-networking-controls",
        "categories": {
            "Accessories": "/accessories",
            "Controls and Consoles": "/controls-consoles",
            "Data Receptacles": "/data-receptacles",
            "Ethernet Switches": "/ethernet-switches",
            "Gateways DMX RDM": "/gateways-dmx-rdm",
            "Interfaces DMX": "/interfaces-dmx",
            "Panels": "/panels",
            "Protocol Converter Router": "/protocol-converter-router",
            "Splitters DMX RDM": "/splitters-dmx-rdm",
        },
        "output_file": "acuity_dmx_networking_products.xlsx",
        "sheet_title": "Acuity Products",
    },
    # The original Industrial script saved to acuity_residential_products.xlsx
    # (the same name the Residential script uses), so it gets its own file here.
    {
        "name": "Industrial",
        "base_url": "https://www.acuitybrands.com/products/industrial",
        "categories": {
            "Linear High Bays": "/linear-high-bays",
            "Low Bays": "/low-bays",
            "Modular Wiring": "/modular-wiring",
            "Round High Bays": "/round-high-bays",
            "Special Applications": "/special-applications",
            "Strip Lights": "/strip-lights",
        },
        "output_file": "acuity_industrial_products.xlsx",
        "sheet_title": "Acuity Industrial Products",
    },
    {
        "name": "Life Safety",
        "base_url": "https://www.acuitybrands.com/products/life-safety",
        "categories": {
            "Canadian Life Safety": "/canadian-life-safety",
            "Emergency AC Inverters": "/emergency-ac-inverters",
            "Emergency Ballast & Drivers": "/emergency-ballast-drivers",
            "Emergency Control Devices": "/emergency-control-devices",
            "Emergency Lighting & Remotes": "/emergency-lighting-remotes",
            "Exit Signs & Combos": "/exit-signs-combos",
            "Self-Testing Automated Reporting": "/self-testing-automated-reporting",
            "Lithonia Life Safety Solutions": "/lithonia-life-safety-solutions",
        },
        "output_file": "acuity_life_safety_products.xlsx",
        "sheet_title": "Acuity Products",
    },
    # Single page, no categories, and a slower scroll (4 seconds) so the whole
    # list loads - exactly as in the original New Products script.
    {
        "name": "New Products",
        "base_url": "https://www.acuitybrands.com/products/our-new-products#sort=relevancy",
        "categories": None,
        "output_file": "acuity_new_products.xlsx",
        "sheet_title": "Acuity New Products",
        "scroll_pause": 4,
    },
    # Same as the original "Residential Indoor Products" script (and the
    # standalone acuity_scraper.py, which was identical to it).
    {
        "name": "Residential Indoor",
        "base_url": "https://www.acuitybrands.com/products/indoor",
        "categories": {
            "Architectural": "/architectural",
            "Decorative": "/decorative",
            "Downlights / Cylinders": "/downlights-cylinders",
            "Linear": "/linear",
            "Modular Wiring": "/modular-wiring",
            "Multiples": "/multiples",
            "Panels": "/panels",
            "Shop Lights": "/shop-lights",
            "Strip Lights": "/strip-lights",
            "Tapelight": "/tapelight",
            "Track Fixtures": "/track-fixtures",
            "Track Systems & Accessories": "/track-systems-accessories",
            "Troffers": "/troffers",
            "Wall Brackets": "/wall-brackets",
            "Wraps": "/wraps",
        },
        "output_file": "acuity_residential_indoor_products.xlsx",
        "sheet_title": "Acuity Products",
    },
    {
        "name": "Residential",
        "base_url": "https://www.acuitybrands.com/products/residential",
        "categories": {
            "Decorative Linear": "/decorative-linear",
            "Downlights": "/downlights",
            "Flush / Surface Mounts": "/flush-surface-mounts",
            "Linear Accent": "/linear-accent",
            "Pendants / Semi-Flush": "/pendants-semi-flush",
            "Sconces": "/sconces",
            "Smart Home": "/smart-home",
            "Step Lights": "/step-lights",
            "Tapelight": "/tapelight",
            "Track Fixtures": "/track-fixtures",
            "Track Systems": "/track-systems",
            "Undercabinet / Task": "/undercabinet-task",
            "Utility / Closet": "/utility-closet",
            "Vanity": "/vanity",
        },
        "output_file": "acuity_residential_products.xlsx",
        "sheet_title": "Acuity Residential Products",
        "view_all_categories": ["Downlights"],
    },
]

# =====================================================
# SHARED STATE (a fresh workbook + browser is set up for every section)
# =====================================================
workbook = None
sheet = None
driver = None
wait = None

def start_workbook(sheet_title):
    global workbook, sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_title
    sheet.append([
        "Product Name",
        "Brand",
        "Product Series",
        "Overview Text",
        "Product URL"
    ])

def start_browser():
    global driver, wait
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=options
    )

    wait = WebDriverWait(driver, 30)

def stop_browser():
    global driver, wait
    try:
        if driver is not None:
            driver.quit()
    except Exception:
        pass
    driver = None
    wait = None

# =====================================================
# UTILITIES
# =====================================================
def accept_cookies():
    try:
        btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Accept']"))
        )
        btn.click()
        time.sleep(1)
    except Exception:
        pass

def scroll_to_bottom(scroll_pause=2):
    last_height = 0
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(scroll_pause)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def get_product_links(scroll_pause=2):
    scroll_to_bottom(scroll_pause)
    links = set()
    for el in driver.find_elements(By.CSS_SELECTOR, "a[href*='/products/detail/']"):
        href = el.get_attribute("href")
        if href:
            links.add(href.split("#")[0])
    log(f"Found {len(links)} product links")
    return list(links)

def click_tab(tab_name):
    tab = wait.until(
        EC.element_to_be_clickable((By.XPATH, f"//li[@data-tabs-nav='{tab_name}']"))
    )
    driver.execute_script("arguments[0].click();", tab)
    time.sleep(1)

def expand_overview_read_more():
    try:
        btn = driver.find_element(
            By.XPATH,
            "//div[@data-tabs-section='overview']//div[contains(@class,'show-more-section-expand')]"
        )
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(1)
    except Exception:
        pass

def extract_brand_from_url(url):
    try:
        path_parts = urlparse(url).path.split("/")
        brand_slug = path_parts[4]
        brand = brand_slug.replace("-", " ").title().replace(" ", "")
        return f"AcuityLighting/{brand}"
    except Exception:
        return "AcuityLighting/Unknown"

def open_view_all(category):
    try:
        view_all = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[.//text()[contains(.,'View All')]]"
            ))
        )
        driver.execute_script("arguments[0].click();", view_all)
        time.sleep(2)
        log(f"Opened View All {category}")
    except Exception:
        warn(f"View All {category} not found")

# =====================================================
# PRODUCT SCRAPER
# =====================================================
def scrape_product_page(url):
    log(f"Opening product: {url}")
    driver.get(url)
    time.sleep(2)

    # Product Name
    try:
        name = wait.until(
            EC.presence_of_element_located((By.TAG_NAME, "h1"))
        ).text.strip()
    except Exception:
        warn("Product name not found")
        return

    # Brand
    brand = extract_brand_from_url(url)
    log(f"Brand detected: {brand}")

    # Overview
    overview_text = ""
    try:
        click_tab("overview")
        expand_overview_read_more()
        overview_text = driver.find_element(
            By.XPATH, "//div[@data-tabs-section='overview']"
        ).text
    except Exception:
        warn("Overview extraction failed")

    # Specifications (SERIES)
    series = ""
    try:
        click_tab("specifications")
        rows = driver.find_elements(
            By.CSS_SELECTOR,
            "table.info-table.specifications-table tr"
        )
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 2:
                label = cols[0].get_attribute("textContent").strip().lower()
                value = cols[1].get_attribute("textContent").strip()
                if label == "series":
                    series = value
                    log(f"Series found: {series}")
                    break
    except Exception as e:
        warn(f"Specifications extraction failed: {e}")

    sheet.append([
        sanitize_for_excel(name),
        sanitize_for_excel(brand),
        sanitize_for_excel(series),
        sanitize_for_excel(overview_text),
        url
    ])

# =====================================================
# ERROR HANDLING
# =====================================================
# A page that fails to load is logged and skipped, so one bad page can't stop a
# run that takes hours. If this many pages fail back-to-back the browser has
# most likely crashed: the section is stopped (its data is still saved) and the
# next section starts with a fresh browser.
MAX_FAILURES_IN_A_ROW = 5

def short_error(error):
    lines = str(error).strip().splitlines()
    return f"{type(error).__name__}: {lines[0] if lines else ''}"

class FailureStreak:
    def __init__(self):
        self.count = 0

    def ok(self):
        self.count = 0

    def failed(self, what, error):
        self.count += 1
        warn(f"{what} failed and was skipped ({short_error(error)})")
        if self.count >= MAX_FAILURES_IN_A_ROW:
            raise RuntimeError(
                f"{MAX_FAILURES_IN_A_ROW} pages failed in a row - the browser has probably crashed"
            ) from error

def scrape_products(product_urls, streak):
    for product_url in product_urls:
        try:
            scrape_product_page(product_url)
            streak.ok()
        except Exception as e:
            streak.failed(f"Product {product_url}", e)

# =====================================================
# SECTION RUNNER
# =====================================================
def run_section(section):
    name = section["name"]
    base_url = section["base_url"]
    categories = section.get("categories")
    output_file = section["output_file"]
    scroll_pause = section.get("scroll_pause", 2)
    view_all_categories = section.get("view_all_categories", [])

    start_workbook(section["sheet_title"])
    start_browser()
    streak = FailureStreak()

    try:
        try:
            driver.get(base_url)
            accept_cookies()

            if categories:
                for category, path in categories.items():
                    log(f"Opening category: {category}")
                    try:
                        driver.get(base_url + path)
                        time.sleep(2)

                        # SPECIAL CASE: DOWNLIGHTS -> VIEW ALL
                        if category in view_all_categories:
                            open_view_all(category)

                        product_urls = get_product_links(scroll_pause)
                    except Exception as e:
                        streak.failed(f"Category {category}", e)
                        continue

                    streak.ok()
                    scrape_products(product_urls, streak)
            else:
                # SINGLE PAGE (no categories)
                scrape_products(get_product_links(scroll_pause), streak)
        finally:
            # Always save whatever was collected, even if something went wrong
            workbook.save(output_file)
            log(f"Saved {sheet.max_row - 1} products to {output_file}")
    finally:
        stop_browser()

    log(f"COMPLETE - {name}")

# =====================================================
# MAIN FLOW
# =====================================================
def main():
    log(f"Starting Acuity all-in-one scraper ({len(SECTIONS)} sections)")
    stopped_early = []

    for number, section in enumerate(SECTIONS, start=1):
        log("=====================================================")
        log(f"SECTION {number} of {len(SECTIONS)}: {section['name']}")
        log("=====================================================")
        try:
            run_section(section)
        except Exception as e:
            warn(f"Section '{section['name']}' stopped early: {short_error(e)}")
            stopped_early.append(section["name"])

    log("=====================================================")
    if stopped_early:
        log("FINISHED - these sections stopped early: " + ", ".join(stopped_early))
    else:
        log("ALL SECTIONS COMPLETE")


if __name__ == "__main__":
    main()
