import time
import re
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


# =========================================================
# SETTINGS
# =========================================================

START_URL = "https://www.kelvix.com/products/"
OUTPUT_FILE = "kelvix_products.xlsx"

BASE_URL = "https://www.kelvix.com"


# =========================================================
# CHROME SETUP
# =========================================================

options = webdriver.ChromeOptions()

# Leave Chrome visible
# options.add_argument("--headless=new")

options.add_argument("--window-size=1400,1000")

service = Service(ChromeDriverManager().install())

driver = webdriver.Chrome(
    service=service,
    options=options
)

wait = WebDriverWait(driver, 15)


# =========================================================
# EXCEL SETUP
# =========================================================

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Kelvix Products"

sheet.append([
    "Product Name",
    "Product Description",
    "Product Code",
    "Product URL"
])


# =========================================================
# HELPERS
# =========================================================

def scroll_to_bottom():
    last_height = driver.execute_script(
        "return document.body.scrollHeight"
    )

    while True:

        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);"
        )

        time.sleep(1.5)

        new_height = driver.execute_script(
            "return document.body.scrollHeight"
        )

        if new_height == last_height:
            break

        last_height = new_height


def clean_text(text):
    if not text:
        return ""

    return re.sub(r"\s+", " ", text).strip()


# =========================================================
# FIND PRODUCT LINKS
# =========================================================

def get_product_links():

    print("Opening Kelvix product catalog...")

    driver.get(START_URL)

    time.sleep(3)

    product_links = set()

    while True:

        scroll_to_bottom()

        links = driver.find_elements(By.TAG_NAME, "a")

        for link in links:

            try:

                href = link.get_attribute("href")

                if not href:
                    continue

                if "/product/" in href:
                    product_links.add(href.split("?")[0])

            except Exception:
                pass

        print(
            f"Products discovered so far: "
            f"{len(product_links)}"
        )

        # Try common pagination selectors
        next_button = None

        selectors = [
            "a.next",
            "a[rel='next']",
            ".pagination-next a",
            "a.next-page"
        ]

        for selector in selectors:

            try:

                elements = driver.find_elements(
                    By.CSS_SELECTOR,
                    selector
                )

                for element in elements:

                    if (
                        element.is_displayed()
                        and element.is_enabled()
                    ):
                        next_button = element
                        break

                if next_button:
                    break

            except Exception:
                pass

        if not next_button:
            break

        try:

            next_url = next_button.get_attribute("href")

            if next_url:
                driver.get(next_url)
            else:
                driver.execute_script(
                    "arguments[0].click();",
                    next_button
                )

            time.sleep(3)

        except Exception:
            break

    return sorted(product_links)


# =========================================================
# SCRAPE PRODUCT
# =========================================================

def scrape_product(url):

    print()
    print(f"Opening: {url}")

    driver.get(url)

    wait.until(
        EC.presence_of_element_located(
            (By.TAG_NAME, "body")
        )
    )

    time.sleep(1)

    # -----------------------------------------------------
    # PRODUCT NAME
    # -----------------------------------------------------

    product_name = ""

    try:

        product_name = clean_text(
            driver.find_element(
                By.TAG_NAME,
                "h1"
            ).text
        )

    except Exception:
        pass

    # -----------------------------------------------------
    # DESCRIPTION
    # -----------------------------------------------------

    description = ""

    try:

        h1 = driver.find_element(
            By.TAG_NAME,
            "h1"
        )

        paragraphs = h1.find_elements(
            By.XPATH,
            "following::p"
        )

        for paragraph in paragraphs:

            text = clean_text(paragraph.text)

            if text:
                description = text
                break

    except Exception:
        pass

    # -----------------------------------------------------
    # PRODUCT CODE
    # -----------------------------------------------------

    product_code = ""

    body_text = driver.find_element(
        By.TAG_NAME,
        "body"
    ).text

    patterns = [
        r"Product Code\s*:?\s*([A-Za-z0-9._\-]+)",
        r"YOUR PRODUCT CODE\s*:?\s*([A-Za-z0-9._\-]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            body_text,
            re.IGNORECASE
        )

        if match:
            product_code = match.group(1).strip()
            break

    if not product_code:

        print("    No product code found. Skipping.")

        return None

    print(f"    Name: {product_name}")
    print(f"    Code: {product_code}")

    return {
        "name": product_name,
        "description": description,
        "code": product_code,
        "url": url
    }


# =========================================================
# MAIN
# =========================================================

try:

    product_links = get_product_links()

    print()
    print("============================")
    print(f"TOTAL PRODUCTS: {len(product_links)}")
    print("============================")

    for index, url in enumerate(
        product_links,
        start=1
    ):

        print()
        print(
            f"[{index}/{len(product_links)}]"
        )

        try:

            product = scrape_product(url)

            if not product:
                continue

            sheet.append([
                product["name"],
                product["description"],
                product["code"],
                product["url"]
            ])

            # Save after every product
            workbook.save(OUTPUT_FILE)

        except Exception as error:

            print(f"    ERROR: {error}")

finally:

    driver.quit()


workbook.save(OUTPUT_FILE)

print()
print("============================")
print("KELVIX SCRAPE COMPLETE")
print("============================")
print(f"Saved to: {OUTPUT_FILE}")
