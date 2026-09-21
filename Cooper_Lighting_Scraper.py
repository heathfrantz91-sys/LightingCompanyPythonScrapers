# =====================================================
# COOPER LIGHTING - ALL-IN-ONE SCRAPER
# =====================================================
# Runs every Cooper Lighting brand scraper back to back. Each entry in BRANDS
# (below) is one of the original scripts - same start page, same scraping steps,
# same 4 columns (Name, Description, UPCCode, URL). When a brand finishes, its
# CSV is saved, its browser is closed, and the next brand starts on its own.
#
#   #  Brand                    Output file
#   1  Trellix Infrastructure   trellix_infrastructure_catalog.csv
#   2  WaveLinx                 wavelinx_catalog.csv
#   3  MWS                      mws_catalog.csv
#   4  Streetworks              streetworks_catalog.csv
#   5  Portfolio                portfolio_catalog.csv
#   6  RSA                      rsa_catalog.csv
#   7  Metalux                  metalux_catalog.csv
#   8  Halo                     halo_catalog.csv
#
# Files are saved in the folder you run the script from.
# To skip a brand, put a # in front of every line of its entry in BRANDS.
# =====================================================
import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# =========================
# BRANDS (one per original script, run in this order)
# =========================
# Ordered so the quick ones finish first and the biggest (Halo, about 900
# products last time) runs last.
BRANDS = [
    {
        "name": "Trellix Infrastructure",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/trellix-infrastructure",
        "output_file": "trellix_infrastructure_catalog.csv",
    },
    {
        "name": "WaveLinx",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/wavelinx",
        "output_file": "wavelinx_catalog.csv",
    },
    {
        "name": "MWS",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/mws",
        "output_file": "mws_catalog.csv",
    },
    {
        "name": "Streetworks",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/streetworks",
        "output_file": "streetworks_catalog.csv",
    },
    # This is the code that was saved in the "Prentalux info" folder.
    {
        "name": "Portfolio",
        "start_url": "https://www.cooperlighting.com/global/product-list#sort=alphanumeric&c=cooper-lighting%3Abrands%2Fportfolio&excludedMarketingStatus=Discontinued",
        "output_file": "portfolio_catalog.csv",
    },
    # This is the code that was saved in the "Shaper" folder.
    {
        "name": "RSA",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/rsa",
        "output_file": "rsa_catalog.csv",
    },
    {
        "name": "Metalux",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/metalux",
        "output_file": "metalux_catalog.csv",
    },
    {
        "name": "Halo",
        "start_url": "https://www.cooperlighting.com/global/product-list#c=cooper-lighting:brands/halo",
        "output_file": "halo_catalog.csv",
    },
]

# =========================
# ERROR HANDLING
# =========================
# A page that fails to load is logged and skipped, so one bad page can't stop a
# run that takes hours. If this many pages fail back-to-back the browser has
# most likely crashed: the brand is stopped (its rows so far are already in the
# CSV) and the next brand starts with a fresh browser.
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
        print(f"⚠️ {what} failed and was skipped ({short_error(error)})")
        if self.count >= MAX_FAILURES_IN_A_ROW:
            raise RuntimeError(
                f"{MAX_FAILURES_IN_A_ROW} pages failed in a row - the browser has probably crashed"
            ) from error

# =========================
# SELENIUM SETUP
# =========================
def start_browser():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    return driver, wait

# =========================
# PRODUCT PAGE -> CATALOG LINKS
# =========================
def get_catalog_urls(driver, wait, product_url):
    driver.get(product_url)
    time.sleep(3)

    # Scroll to Stock Catalog Number section
    try:
        stock_header = wait.until(
            EC.presence_of_element_located((By.XPATH, "//h2[contains(text(),'Stock Catalog Number')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", stock_header)
        time.sleep(2)
    except TimeoutException:
        print("⚠️ No Stock Catalog Number section found")
        return []

    # Collect catalog number links (ignore PDFs/spec sheets)
    catalog_links = driver.find_elements(By.CSS_SELECTOR, "a.product-table__link")
    catalog_urls = [
        link.get_attribute("href")
        for link in catalog_links
        if link.get_attribute("href") and "/global/brands/" in link.get_attribute("href")
    ]

    print(f"→ Found {len(catalog_urls)} catalog items")
    return catalog_urls

# =========================
# CATALOG PAGE -> NAME / DESCRIPTION / UPC
# =========================
def scrape_catalog_page(driver, catalog_url):
    driver.get(catalog_url)
    time.sleep(3)

    # Scroll to bottom to load full specs
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)

    # -------- Name
    try:
        name = driver.find_element(By.CSS_SELECTOR, "h1.product-detail__title").text.strip()
    except NoSuchElementException:
        name = "N/A"

    # -------- Description (from Product Description table)
    description = "N/A"
    try:
        desc_rows = driver.find_elements(
            By.XPATH,
            "//h2[contains(text(),'Product Description')]/following::table[1]//tr"
        )
        for row in desc_rows:
            label = row.find_element(By.XPATH, ".//td[1]").text.strip().lower()
            if label == "description":
                description = row.find_element(By.XPATH, ".//td[2]").text.strip()
                break
    except Exception:
        pass

    # -------- UPC (from Product Details table)
    upc = "N/A"
    try:
        rows = driver.find_elements(
            By.XPATH,
            "//h2[contains(text(),'Product Details')]/following::table[1]//tr"
        )
        for row in rows:
            try:
                label = row.find_element(By.XPATH, "./td[1]").text.strip()
                value = row.find_element(By.XPATH, "./td[2]").text.strip()
                if label.lower() == "upc":
                    upc = value
                    break
            except Exception:
                continue
    except Exception:
        pass

    return name, description, upc

# =========================
# ONE BRAND, START TO FINISH
# =========================
def scrape_brand(brand):
    brand_name = brand["name"]
    start_url = brand["start_url"]
    output_file = brand["output_file"]

    driver, wait = start_browser()
    streak = FailureStreak()
    rows_written = 0

    try:
        # =========================
        # OPEN MAIN PAGE
        # =========================
        driver.get(start_url)

        # Accept cookies if shown
        try:
            wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Agree')]"))
            ).click()
        except TimeoutException:
            pass

        time.sleep(4)

        # =========================
        # SCROLL TO LOAD ALL PRODUCTS
        # =========================
        print(f"🔄 Loading all {brand_name} products...")
        last_height = driver.execute_script("return document.body.scrollHeight")

        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(3)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        print("✅ All products loaded")

        # =========================
        # COLLECT "AVAILABLE STOCK / CONFIGURE" LINKS
        # =========================
        product_buttons = driver.find_elements(By.CSS_SELECTOR, "a.filter-product-card__cta.link")
        product_urls = list({btn.get_attribute("href") for btn in product_buttons if btn.get_attribute("href")})

        print(f"🔗 Found {len(product_urls)} product configuration pages")
        if not product_urls:
            print(f"⚠️ No products found for {brand_name} - the CSV will only contain the header row")

        # =========================
        # CSV SETUP
        # =========================
        with open(output_file, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Name", "Description", "UPCCode", "URL"])

            # =========================
            # LOOP EACH PRODUCT PAGE
            # =========================
            for p_index, product_url in enumerate(product_urls, start=1):
                print(f"\n📦 [{p_index}/{len(product_urls)}] Opening product page")

                try:
                    catalog_urls = get_catalog_urls(driver, wait, product_url)
                except Exception as e:
                    streak.failed(f"Product page {product_url}", e)
                    continue

                # A page that loaded but simply has nothing to list is not a failure.
                # (Otherwise only a successful catalog page resets the failure count.)
                if not catalog_urls:
                    streak.ok()

                # =========================
                # LOOP EACH CATALOG PAGE
                # =========================
                for catalog_url in catalog_urls:
                    try:
                        name, description, upc = scrape_catalog_page(driver, catalog_url)
                        streak.ok()
                    except Exception as e:
                        streak.failed(f"Catalog page {catalog_url}", e)
                        continue

                    print(f"✅ {name} | UPC: {upc}")
                    writer.writerow([name, description, upc, catalog_url])
                    rows_written += 1
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    print(f"\n🎉 DONE — {brand_name}: {rows_written} rows saved to {output_file}")
    return rows_written

# =========================
# MAIN FLOW
# =========================
def main():
    print(f"Starting Cooper Lighting all-in-one scraper ({len(BRANDS)} brands)")
    results = []

    for number, brand in enumerate(BRANDS, start=1):
        print("\n=====================================================")
        print(f"BRAND {number} of {len(BRANDS)}: {brand['name']}")
        print("=====================================================")
        try:
            rows = scrape_brand(brand)
            results.append(f"{brand['name']}: {rows} rows -> {brand['output_file']}")
        except Exception as e:
            print(f"⚠️ {brand['name']} stopped early: {short_error(e)}")
            print(f"   Rows collected before the problem are in {brand['output_file']} (if the file was created).")
            results.append(f"{brand['name']}: STOPPED EARLY -> {brand['output_file']}")

    print("\n=====================================================")
    print("ALL BRANDS FINISHED")
    print("=====================================================")
    for line in results:
        print(f"  {line}")


if __name__ == "__main__":
    main()
