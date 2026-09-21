import csv
import time

from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


BASE_URL = "https://www.milesight.com/"
OUTPUT_CSV = "milesight_products.csv"
LOG_FILE = "scraper_debug.log"


def log(msg):
    print(msg, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def start_driver():
    options = webdriver.ChromeOptions()

    # Headless mode
    options.add_argument("--headless=new")

    # Stability / speed
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    # Slight performance improvements
    prefs = {
        "profile.managed_default_content_settings.images": 2,
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    driver.set_page_load_timeout(60)
    return driver


def wait_for_page(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    time.sleep(0.25)
    try:
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)
    except Exception:
        driver.execute_script("arguments[0].click();", element)


def try_click_xpath(driver, xpath, timeout=8):
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        safe_click(driver, el)
        return True
    except Exception:
        return False


def accept_cookies(driver):
    xpaths = [
        "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
        "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'agree')]",
        "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'allow')]",
        "//*[@id='onetrust-accept-btn-handler']",
    ]
    for xpath in xpaths:
        if try_click_xpath(driver, xpath, timeout=3):
            log("[LOG] Cookie banner accepted.")
            time.sleep(0.5)
            return
    log("[LOG] No cookie banner found.")


def open_home_and_menu(driver):
    driver.get(BASE_URL)
    wait_for_page(driver)
    time.sleep(1)
    accept_cookies(driver)

    if not (
        try_click_xpath(driver, "//a[normalize-space()='IoT Products']", timeout=10)
        or try_click_xpath(driver, "//*[normalize-space()='IoT Products']", timeout=10)
    ):
        raise RuntimeError("Could not click IoT Products")

    time.sleep(0.75)

    if not (
        try_click_xpath(driver, "//a[normalize-space()='By Product']", timeout=10)
        or try_click_xpath(driver, "//*[normalize-space()='By Product']", timeout=10)
    ):
        raise RuntimeError("Could not click By Product")

    time.sleep(0.75)
    log("[LOG] Opened IoT Products > By Product menu.")


def get_anchor_pairs(driver):
    anchors = driver.find_elements(By.XPATH, "//a[@href]")
    pairs = []
    for a in anchors:
        try:
            text = " ".join(a.text.split()).strip()
            href = (a.get_attribute("href") or "").strip()
            if text and href:
                pairs.append((text, href))
        except StaleElementReferenceException:
            continue
    return pairs


def dedupe_pairs(pairs):
    seen = set()
    out = []
    for text, href in pairs:
        key = (text, href)
        if key not in seen:
            seen.add(key)
            out.append((text, href))
    return out


def discover_main_categories(driver):
    target_names = {
        "IoT LoRaWAN® Sensors",
        "LoRaWAN® Gateways",
        "Industrial Routers",
        "IoT Controllers",
        "Infinity",
        "Accessories",
        "LoRaWAN® Demo Kits",
        "Product Comparison Guide",
    }

    results = []
    for text, href in get_anchor_pairs(driver):
        if text in target_names and "milesight.com" in href:
            results.append((text, href))

    results = dedupe_pairs(results)
    log(f"[LOG] Found {len(results)} main categories.")
    for name, href in results:
        log(f"[MAIN CATEGORY] {name} | {href}")
    return results


def discover_subcategories_from_menu(driver, main_category_name):
    open_home_and_menu(driver)

    if not (
        try_click_xpath(driver, f"//a[normalize-space()='{main_category_name}']", timeout=8)
        or try_click_xpath(driver, f"//*[normalize-space()='{main_category_name}']", timeout=8)
    ):
        log(f"[LOG] Could not click main category in menu: {main_category_name}")
        return []

    time.sleep(1)

    subcategories = []
    for text, href in get_anchor_pairs(driver):
        looks_like_sub = (
            "Series" in text
            or text in {"IoT Display", "People Sensing Series", "CoWork Series"}
        )
        if looks_like_sub and "milesight.com" in href:
            subcategories.append((text, href))

    subcategories = dedupe_pairs(subcategories)

    filtered = []
    seen = set()
    for name, href in subcategories:
        if href in seen:
            continue
        seen.add(href)
        filtered.append((name, href))

    log(f"[LOG] Found {len(filtered)} subcategories under {main_category_name}.")
    for name, href in filtered:
        log(f"[SUBCATEGORY] {main_category_name} | {name} | {href}")
    return filtered


def is_valid_iot_product_url(href):
    if not href:
        return False

    href = href.split("#")[0].strip()

    allowed_prefixes = [
        "https://www.milesight.com/iot/product/",
    ]

    blocked_parts = [
        "/security/",
        "/solution/",
        "/solutions/",
        "/spotlight-highlight/",
        "/success-stories/",
        "/company/",
        "/support/",
        "/news/",
        "/blog/",
    ]

    if any(x in href for x in blocked_parts):
        return False

    return any(href.startswith(prefix) for prefix in allowed_prefixes)


def gather_product_links_from_page(driver):
    log(f"[LOG] Gathering product links from page: {driver.current_url}")

    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.7)

    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.4)

    anchors = driver.find_elements(By.XPATH, "//a[@href]")
    products = []
    seen = set()

    for a in anchors:
        try:
            href = (a.get_attribute("href") or "").split("#")[0].strip()
            text = " ".join(a.text.split()).strip()

            if is_valid_iot_product_url(href) and href not in seen:
                seen.add(href)
                products.append((text, href))
        except StaleElementReferenceException:
            continue

    log(f"[LOG] Found {len(products)} valid IoT product links on page.")
    for text, href in products:
        product_label = text if text else "NO_VISIBLE_PRODUCT_TEXT"
        log(f"[PRODUCT LINK] {product_label} | {href}")
    return products


def click_models_tab(driver):
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.5)

    possible_tabs = [
        "//a[normalize-space()='MODELS']",
        "//a[contains(normalize-space(), 'MODELS')]",
        "//div[normalize-space()='MODELS']",
        "//span[normalize-space()='MODELS']",
        "//*[contains(@class,'tab') and normalize-space()='MODELS']",
        "//*[contains(@class,'tab') and contains(normalize-space(), 'MODELS')]",
        "//a[normalize-space()='Models']",
        "//a[contains(normalize-space(), 'Models')]",
        "//div[normalize-space()='Models']",
        "//span[normalize-space()='Models']",
    ]

    tab_element = None
    for xpath in possible_tabs:
        try:
            tab_element = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            if tab_element:
                break
        except Exception:
            continue

    if not tab_element:
        log("[LOG] MODELS tab element not found.")
        return False

    try:
        safe_click(driver, tab_element)
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", tab_element)
        except Exception:
            log("[LOG] MODELS tab click failed.")
            return False

    time.sleep(1)

    active_checks = [
        "//a[contains(@class,'active') and contains(., 'MODELS')]",
        "//li[contains(@class,'active')]//*[contains(., 'MODELS')]",
        "//*[contains(@class,'cur') and contains(., 'MODELS')]",
        "//*[contains(@class,'on') and contains(., 'MODELS')]",
        "//*[contains(@class,'active') and normalize-space()='MODELS']",
        "//a[contains(@class,'active') and contains(., 'Models')]",
        "//li[contains(@class,'active')]//*[contains(., 'Models')]",
    ]

    for xpath in active_checks:
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            return True
        except Exception:
            continue

    spec_checks = [
        "//h1[contains(., 'Specifications')]",
        "//h2[contains(., 'Specifications')]",
        "//div[contains(., 'Specifications')]",
        "//h1[contains(., 'SPECIFICATIONS')]",
        "//h2[contains(., 'SPECIFICATIONS')]",
    ]
    for xpath in spec_checks:
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            return True
        except Exception:
            continue

    return False


def extract_description_from_features(driver):
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.5)

    clicked_features = (
        try_click_xpath(driver, "//a[normalize-space()='FEATURES']", timeout=4)
        or try_click_xpath(driver, "//a[contains(normalize-space(), 'FEATURES')]", timeout=4)
        or try_click_xpath(driver, "//div[normalize-space()='FEATURES']", timeout=4)
        or try_click_xpath(driver, "//span[normalize-space()='FEATURES']", timeout=4)
        or try_click_xpath(driver, "//a[normalize-space()='Features']", timeout=4)
        or try_click_xpath(driver, "//a[contains(normalize-space(), 'Features')]", timeout=4)
    )

    if clicked_features:
        time.sleep(0.8)

    selectors = [
        (By.CSS_SELECTOR, ".product-desc"),
        (By.CSS_SELECTOR, ".desc"),
        (By.CSS_SELECTOR, ".description"),
        (By.CSS_SELECTOR, ".features"),
        (By.CSS_SELECTOR, ".feature"),
        (By.XPATH, "//*[contains(@class, 'product-desc')]"),
        (By.XPATH, "//*[contains(@class, 'feature')]"),
    ]

    collected = []
    for by, value in selectors:
        try:
            els = driver.find_elements(by, value)
            for el in els:
                txt = " ".join(el.text.split())
                if txt and len(txt) > 20:
                    collected.append(txt)
        except Exception:
            continue

    final = []
    seen = set()
    for txt in collected:
        if txt not in seen:
            seen.add(txt)
            final.append(txt)

    return " | ".join(final)


def extract_model_names(driver):
    selectors = [
        (By.CSS_SELECTOR, ".model-name"),
        (By.XPATH, "//table//tr/td[1]"),
        (By.XPATH, "//tbody//tr/td[1]"),
        (By.XPATH, "//*[contains(@class, 'model')]"),
    ]

    names = []
    for by, value in selectors:
        try:
            els = driver.find_elements(by, value)
            temp = []
            for el in els:
                txt = " ".join(el.text.split()).strip()
                if txt and len(txt) < 120:
                    temp.append(txt)
            if temp:
                names = temp
                break
        except Exception:
            continue

    cleaned = []
    seen = set()
    junk = {
        "model",
        "models",
        "specification",
        "specifications",
        "hardware specifications",
        "software specifications"
    }

    for n in names:
        low = n.lower()
        if low in junk:
            continue
        if n not in seen:
            seen.add(n)
            cleaned.append(n)

    return cleaned


def scrape_one_product(driver, main_category, subcategory, product_url):
    log("-----------------------------------------------------")
    log(f"[PRODUCT URL REQUESTED] {product_url}")
    log(f"[MAIN CATEGORY] {main_category}")
    log(f"[SUBCATEGORY] {subcategory if subcategory else 'NONE'}")

    driver.get(product_url)
    wait_for_page(driver)
    time.sleep(1)

    actual_url = driver.current_url
    log(f"[PRODUCT URL VISITED] {actual_url}")

    models_clicked = click_models_tab(driver)
    log(f"[MODELS TAB CLICKED] {models_clicked}")

    if models_clicked:
        model_names = extract_model_names(driver)
    else:
        model_names = []

    description = extract_description_from_features(driver)
    description_found = bool(description.strip())

    log(f"[DESCRIPTION FOUND] {description_found}")
    if description_found:
        log(f"[DESCRIPTION VALUE] {description}")
    else:
        log("[DESCRIPTION VALUE] NONE")

    if model_names:
        for model in model_names:
            log(f"[MODEL FOUND] {model}")
    else:
        log("[MODEL FOUND] NONE")

    if not model_names:
        model_names = ["MODEL_NOT_FOUND"]

    rows = []
    for model in model_names:
        row = {
            "main_category": main_category,
            "subcategory": subcategory,
            "model_name": model,
            "product_url": actual_url,
            "description": description
        }
        rows.append(row)

    return rows


def scrape_main_category_direct(driver, main_category_name, main_category_url):
    log(f"[LOG] Opening main category directly: {main_category_name}")
    driver.get(main_category_url)
    wait_for_page(driver)
    time.sleep(1)

    products = gather_product_links_from_page(driver)
    rows = []
    for _, product_url in products:
        try:
            rows.extend(scrape_one_product(driver, main_category_name, "", product_url))
        except Exception as e:
            log(f"[ERROR] Product scrape failed | {product_url} | {e}")
    return rows


def scrape_subcategories(driver, main_category_name, subcategories):
    rows = []
    for sub_name, sub_url in subcategories:
        log(f"[LOG] Opening subcategory: {sub_name}")
        try:
            driver.get(sub_url)
            wait_for_page(driver)
            time.sleep(1)

            products = gather_product_links_from_page(driver)
            for _, product_url in products:
                try:
                    rows.extend(
                        scrape_one_product(driver, main_category_name, sub_name, product_url)
                    )
                except Exception as e:
                    log(f"[ERROR] Product scrape failed | {product_url} | {e}")
        except Exception as e:
            log(f"[ERROR] Subcategory open failed | {sub_name} | {sub_url} | {e}")
    return rows


def save_csv(rows, filename):
    fieldnames = [
        "main_category",
        "subcategory",
        "model_name",
        "product_url",
        "description",
    ]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main():
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("Milesight scraper log\n")

    driver = start_driver()
    all_rows = []

    try:
        open_home_and_menu(driver)
        main_categories = discover_main_categories(driver)

        if not main_categories:
            raise RuntimeError("No main categories were discovered from the By Product menu.")

        for main_name, main_url in main_categories:
            log("=====================================================")
            log(f"[PROCESSING MAIN CATEGORY] {main_name}")

            try:
                subcategories = discover_subcategories_from_menu(driver, main_name)

                if subcategories:
                    all_rows.extend(scrape_subcategories(driver, main_name, subcategories))
                else:
                    all_rows.extend(scrape_main_category_direct(driver, main_name, main_url))

            except Exception as e:
                log(f"[ERROR] Main category failure | {main_name} | {e}")

    finally:
        driver.quit()

    save_csv(all_rows, OUTPUT_CSV)
    log(f"[DONE] Saved {len(all_rows)} rows to {OUTPUT_CSV}")
    log(f"[DONE] Log written to {LOG_FILE}")


if __name__ == "__main__":
    main()