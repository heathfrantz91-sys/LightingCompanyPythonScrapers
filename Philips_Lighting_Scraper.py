import re
import time
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

START_URL = "https://www.lighting.philips.com/p"

OUTPUT_FILE = "philips_products.xlsx"

BASE_DOMAIN = "https://www.lighting.philips.com"


# =========================================================
# CHROME
# =========================================================

options = webdriver.ChromeOptions()

# Keep browser visible
# options.add_argument("--headless=new")

options.add_argument("--window-size=1500,1000")

service = Service(
    ChromeDriverManager().install()
)

driver = webdriver.Chrome(
    service=service,
    options=options
)

wait = WebDriverWait(driver, 15)


# =========================================================
# EXCEL
# =========================================================

workbook = openpyxl.Workbook()

sheet = workbook.active
sheet.title = "Products"

sheet.append([
    "Full Product Name",
    "Full Product Code",
    "Brand",
    "Product URL",
    "Product Description"
])


# =========================================================
# TEXT CLEANING
# =========================================================

def clean_text(text):

    if not text:
        return ""

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================================================
# COOKIE HANDLING
# =========================================================

def accept_cookies():

    button_texts = [
        "Accept",
        "Accept All",
        "Accept all",
        "Agree",
        "I agree"
    ]

    # -----------------------------------------------------
    # NORMAL DOM
    # -----------------------------------------------------

    for text in button_texts:

        try:

            buttons = driver.find_elements(
                By.XPATH,
                f"//button[contains(., '{text}')]"
            )

            for button in buttons:

                if button.is_displayed():

                    driver.execute_script(
                        "arguments[0].click();",
                        button
                    )

                    time.sleep(1)

                    return

        except Exception:
            pass

    # -----------------------------------------------------
    # IFRAMES
    # -----------------------------------------------------

    try:

        frames = driver.find_elements(
            By.TAG_NAME,
            "iframe"
        )

        for frame in frames:

            try:

                driver.switch_to.frame(frame)

                buttons = driver.find_elements(
                    By.XPATH,
                    "//button[contains("
                    "translate(., "
                    "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
                    "'abcdefghijklmnopqrstuvwxyz'), "
                    "'accept')]"
                )

                if buttons:

                    driver.execute_script(
                        "arguments[0].click();",
                        buttons[0]
                    )

                    driver.switch_to.default_content()

                    time.sleep(1)

                    return

                driver.switch_to.default_content()

            except Exception:

                driver.switch_to.default_content()

    except Exception:
        pass

    # -----------------------------------------------------
    # SHADOW DOM FALLBACK
    # -----------------------------------------------------

    try:

        driver.execute_script("""
            function findAccept(root) {

                const elements =
                    root.querySelectorAll('*');

                for (const element of elements) {

                    if (
                        element.shadowRoot
                    ) {

                        const result =
                            findAccept(
                                element.shadowRoot
                            );

                        if (result) {
                            return result;
                        }
                    }

                    if (
                        element.tagName === 'BUTTON'
                    ) {

                        const text =
                            element.innerText
                            .toLowerCase();

                        if (
                            text.includes('accept')
                            || text.includes('agree')
                        ) {
                            element.click();
                            return true;
                        }
                    }
                }

                return false;
            }

            return findAccept(document);
        """)

    except Exception:
        pass


# =========================================================
# SCROLL
# =========================================================

def scroll_full_page():

    last_height = 0

    while True:

        height = driver.execute_script(
            "return document.body.scrollHeight"
        )

        driver.execute_script(
            "window.scrollTo(0, "
            "document.body.scrollHeight);"
        )

        time.sleep(1.5)

        if height == last_height:
            break

        last_height = height


# =========================================================
# PRODUCT DATA TAB
# =========================================================

def click_product_data():

    possible_labels = [
        "Product Data",
        "Specifications",
        "Specification"
    ]

    for label in possible_labels:

        try:

            elements = driver.find_elements(
                By.XPATH,
                f"//*[contains("
                f"normalize-space(.), "
                f"'{label}')]"
            )

            for element in elements:

                try:

                    if (
                        element.is_displayed()
                        and element.is_enabled()
                    ):

                        driver.execute_script(
                            "arguments[0].click();",
                            element
                        )

                        time.sleep(1.5)

                        return True

                except Exception:
                    pass

        except Exception:
            pass

    return False


# =========================================================
# FIND SPECIFICATION
# =========================================================

def get_specification(spec_name):

    items = driver.find_elements(
        By.CSS_SELECTOR,
        "li.specification__list-item"
    )

    for item in items:

        try:

            header = item.find_element(
                By.CSS_SELECTOR,
                "span.specification__list-header"
            )

            header_text = clean_text(
                header.text
            )

            if header_text.lower() != spec_name.lower():
                continue

            # ---------------------------------------------
            # NORMAL DESCRIPTION SPAN
            # ---------------------------------------------

            try:

                description = item.find_element(
                    By.CSS_SELECTOR,
                    "span.specification__list-description"
                )

                value = clean_text(
                    description.text
                )

                if value:
                    return value

            except Exception:
                pass

            # ---------------------------------------------
            # DESCRIPTION WRAPPER
            # ---------------------------------------------

            try:

                wrapper = item.find_element(
                    By.CSS_SELECTOR,
                    ".specification__list-description-wrapper"
                )

                value = clean_text(
                    wrapper.text
                )

                if value:
                    return value

            except Exception:
                pass

            # ---------------------------------------------
            # RAW ITEM TEXT FALLBACK
            # ---------------------------------------------

            raw_text = clean_text(
                item.text
            )

            if raw_text:

                raw_text = raw_text.replace(
                    header_text,
                    "",
                    1
                ).strip()

                return raw_text

        except Exception:
            pass

    return ""


# =========================================================
# PRODUCT NAME
# =========================================================

def get_product_name():

    selectors = [
        "h1",
        ".pdp-product-title",
        "[data-testid='product-title']"
    ]

    for selector in selectors:

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector
            )

            for element in elements:

                value = clean_text(
                    element.text
                )

                if value:
                    return value

        except Exception:
            pass

    return ""


# =========================================================
# DESCRIPTION
# =========================================================

def get_description():

    selectors = [
        ".product-description",
        ".pdp-product-description",
        ".rich-text",
        "main p"
    ]

    for selector in selectors:

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector
            )

            for element in elements:

                text = clean_text(
                    element.text
                )

                if len(text) < 20:
                    continue

                # Remove order-code text that was leaking
                # into the description.
                text = re.sub(
                    r"Order code:\s*\d+",
                    "",
                    text,
                    flags=re.IGNORECASE
                )

                text = re.sub(
                    r"Full order code:\s*\d+",
                    "",
                    text,
                    flags=re.IGNORECASE
                )

                text = clean_text(text)

                if text:
                    return text

        except Exception:
            pass

    return ""


# =========================================================
# DISCOVER PRODUCT LINKS
# =========================================================

def get_product_links():

    print("Loading Philips Lighting catalog...")

    driver.get(START_URL)

    time.sleep(4)

    accept_cookies()

    scroll_full_page()

    product_links = set()

    anchors = driver.find_elements(
        By.TAG_NAME,
        "a"
    )

    for anchor in anchors:

        try:

            href = anchor.get_attribute(
                "href"
            )

            if not href:
                continue

            href = href.split("?")[0]

            if (
                "lighting.philips.com/p/"
                in href
            ):

                product_links.add(href)

        except Exception:
            pass

    return sorted(product_links)


# =========================================================
# SCRAPE PRODUCT PAGE
# =========================================================

def scrape_product(url):

    driver.get(url)

    wait.until(
        EC.presence_of_element_located(
            (By.TAG_NAME, "body")
        )
    )

    time.sleep(2)

    accept_cookies()

    scroll_full_page()

    product_name = get_product_name()

    click_product_data()

    full_product_code = ""

    possible_code_names = [
        "Full product code",
        "Full Product Code",
        "Order code",
        "Order Code"
    ]

    for code_name in possible_code_names:

        full_product_code = get_specification(
            code_name
        )

        if full_product_code:
            break

    description = get_description()

    # Philips product pages occasionally expose
    # order codes inside surrounding text.
    description = re.sub(
        r"Order code:\s*\d+",
        "",
        description,
        flags=re.IGNORECASE
    )

    description = re.sub(
        r"Full order code:\s*\d+",
        "",
        description,
        flags=re.IGNORECASE
    )

    description = clean_text(description)

    return {
        "name": product_name,
        "code": full_product_code,
        "brand": "Philips",
        "url": url,
        "description": description
    }


# =========================================================
# MAIN
# =========================================================

try:

    product_links = get_product_links()

    print()
    print("==============================")
    print(
        f"PRODUCT LINKS FOUND: "
        f"{len(product_links)}"
    )
    print("==============================")

    for index, url in enumerate(
        product_links,
        start=1
    ):

        print()
        print(
            f"[{index}/{len(product_links)}]"
        )

        print(url)

        try:

            product = scrape_product(url)

            print(
                f"    Name: "
                f"{product['name']}"
            )

            print(
                f"    Code: "
                f"{product['code']}"
            )

            sheet.append([
                product["name"],
                product["code"],
                product["brand"],
                product["url"],
                product["description"]
            ])

            # Incremental save so progress isn't lost
            workbook.save(
                OUTPUT_FILE
            )

        except Exception as error:

            print(
                f"    ERROR: {error}"
            )

finally:

    driver.quit()


workbook.save(
    OUTPUT_FILE
)

print()
print("==============================")
print("PHILIPS SCRAPE COMPLETE")
print("==============================")
print(
    f"Saved to: {OUTPUT_FILE}"
)
