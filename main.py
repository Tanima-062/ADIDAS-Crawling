import os
import pytz
from datetime import datetime
import logging
import pytest
import unicodedata
from selenium.common.exceptions import TimeoutException
import traceback
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from difflib import SequenceMatcher
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
from bs4 import BeautifulSoup
import json
import re
import pandas as pd


base_path = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(base_path, 'adidas_products.xlsx')

# Function to generate the current timestamp
def get_japan_time():
    """Generate the current timestamp in Japan Standard Time (JST)."""
    jst = pytz.timezone('Asia/Tokyo')
    return datetime.now(jst).strftime('%Y%m%d_%H%M%S')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

class TestAdidas:
    def setup_method(self, method):
        timestamp = get_japan_time()

        options = uc.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("start-maximized")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")

        self.driver = uc.Chrome(options=options)
        self.driver.set_page_load_timeout(180)

        # Get script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_dir_name = f"{timestamp}"

        # Define directory paths relative to the script
        self.screenshot_dir = os.path.join(script_dir, f'screenshots_{base_dir_name}')
        self.error_dir = os.path.join(script_dir, f'error_{base_dir_name}')
        self.execution_dir = os.path.join(script_dir, f'execution_{base_dir_name}')

        # Create directories if they do not exist
        os.makedirs(self.screenshot_dir, exist_ok=True)
        os.makedirs(self.error_dir, exist_ok=True)
        os.makedirs(self.execution_dir, exist_ok=True)

        self.log_execution("Setup completed.")

    def teardown_method(self, method):
        if self.driver:
            self.driver.quit()

    def launch_driver(self):
        options = uc.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--blink-settings=imagesEnabled=false")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("start-maximized")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")

        self.driver = uc.Chrome(options=options)
        self.driver.set_page_load_timeout(180)
    def log_error(self, error_message):
        """Log error to a specific error log file in the error folder."""
        os.makedirs(self.error_dir, exist_ok=True)
        timestamp = get_japan_time()
        error_log_path = os.path.join(self.error_dir, f"error_log.txt")

        try:
            # Write the error message and traceback to the error log file
            with open(error_log_path, 'a', encoding="utf-8") as error_file:
                error_file.write(f"Error occurred at {timestamp}:\n")
                error_file.write(f"{error_message}\n")
                error_file.write(traceback.format_exc())
            logger.error(f"Error details saved in: {error_log_path}")
        except Exception as e:
            logger.error(f"Failed to log error: {e}")
            raise

    def log_execution(self, step_message):
        """Log execution message to the result log file."""
        os.makedirs(self.execution_dir, exist_ok=True)
        timestamp = get_japan_time()
        execution_log_path = os.path.join(self.execution_dir, f"execution_log.txt")

        try:
            # Write the execution message to the execution log file
            with open(execution_log_path, 'a', encoding="utf-8") as execution_file:
                execution_message = f"{timestamp}: {step_message}.\n"
                execution_file.write(execution_message)
            logger.info(f"Execution details saved in: {execution_log_path}")
        except Exception as e:
            logger.error(f"Failed to log execution: {e}")
            raise

    def take_screenshot(self, name, wait, type, loader_xpath="//*[@id='img-loader']"):
        """Take a screenshot after ensuring the page is loaded."""
        sub_dir = 'error' if type == 'error' else 'success'
        full_dir = os.path.join(self.screenshot_dir, sub_dir)
        os.makedirs(full_dir, exist_ok=True)
        timestamp = get_japan_time()
        screenshot_path = os.path.join(full_dir, f"{name}_{timestamp}.png")
        try:
            wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
            try:
                wait.until(EC.invisibility_of_element_located((By.XPATH, loader_xpath)))
            except TimeoutException:
                logger.warning("Loader (id='img-loader') is still visible after timeout, proceeding with screenshot.")
            self.driver.save_screenshot(screenshot_path)
            logger.info(f"Screenshot saved: {screenshot_path}")
        except Exception as e:
            logger.error(f"Failed to save screenshot: {str(e)}")

    def assert_expected_result(self, expected_text, wait):
        normalized_expected = unicodedata.normalize("NFKC", expected_text.strip())
        time.sleep(2)
        try:
            # Custom JavaScript: recursively extract all visible text
            page_text = self.driver.execute_script("""
                    function getTextContent(node) {
                        let text = "";
                        if (node.nodeType === Node.TEXT_NODE && node.nodeValue.trim() !== "") {
                            text += node.nodeValue;
                        }
                        for (let child of node.childNodes) {
                            text += getTextContent(child);
                        }
                        return text;
                    }
                    return getTextContent(document.body);
                """)
            page_text_cleaned = unicodedata.normalize("NFKC", " ".join(page_text.split()))

            # ✅ Exact match
            if normalized_expected in page_text_cleaned:
                self.log_execution(f"✅ Expected text found {expected_text}")
                self.take_screenshot(f"not_found_{expected_text}", wait, 'success')
                return

            # 🔍 Fuzzy match
            score = SequenceMatcher(None, normalized_expected, page_text_cleaned).ratio()
            if score > 0.75:
                self.log_execution(f"✅ Expected text match ({expected_text})")
                self.take_screenshot(f"予想テキスト_{expected_text}", wait, 'success')
            else:
                self.log_error(f"❌ Expected text not found ({expected_text})")
                self.take_screenshot(f"not_found_{expected_text}", wait, 'error')

        except Exception as e:
            self.log_error(f"❌ Error during assertion: {e}")
            self.take_screenshot(f"{expected_text}_error", wait, 'error')
    def test_adidas(self):
        wait = WebDriverWait(self.driver, 60)
        rows = []
        product_links = []
        try:
            self.driver.get("https://www.adidas.jp/men")
            mens_menu = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/men"]')))
            text = mens_menu.get_attribute('textContent').strip()
            self.assert_expected_result(text, wait)
        except Exception as e:
            self.log_error(f"Failed to open (https://www.adidas.jp/men): {e}")
        self.driver.set_window_size(1296, 775)
        self.log_execution(f"Navigate to URL (https://www.adidas.jp/men)")
        try:
            mens_menu = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/men"]')))
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'a[href="/men"]')))
            actions = ActionChains(self.driver)
            actions.move_to_element(mens_menu).perform()
            try:
                wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/メンズ-ウェア・服-tシャツ"]')))
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'a[href="/メンズ-ウェア・服-tシャツ"]')))
                tshirts_link = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/メンズ-ウェア・服-tシャツ"]')))
                try:
                    tshirts_link.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", tshirts_link)
                self.log_execution("Clicked on T-shirts link")
                while True:
                    try:
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                        wait.until(
                            EC.presence_of_all_elements_located(
                                (By.CSS_SELECTOR, 'article[data-testid="plp-product-card"]')))
                        articles = self.driver.find_elements(By.CSS_SELECTOR, 'article[data-testid="plp-product-card"]')
                        for product in articles:
                            try:
                                link_element = product.find_element(By.CSS_SELECTOR,
                                                                    'a[data-testid="product-card-image-link"]')
                                link = link_element.get_attribute('href')
                                product_links.append(link)
                            except Exception as e:
                                self.log_error(f"Error getting product link: {e}")
                    except Exception as e:
                        self.log_error(f"Error finding T-shirts list: {e}")
                        self.log_execution("Failed")

                    try:
                        next_button = wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-testid="pagination-next-button"]'))
                        )
                        href = next_button.get_attribute('href')
                        try:
                            self.driver.get(href)
                        except Exception as e:
                            self.log_error(f"Failed to open ({href}): {e}")
                    except TimeoutException:
                        break

                    try:
                        close_btn = self.driver.find_element(By.ID, "gl-modal__close-mf-account-portal")
                        if close_btn.is_displayed():
                            try:
                                close_btn.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", close_btn)
                    except Exception:
                        pass
            except Exception as e:
                self.log_error(f"Error finding T-shirts link: {e}")
                self.log_execution("Failed")
        except Exception as e:
            self.log_error(f"Error finding men's menu: {e}")
            self.log_execution("Failed")

        self.log_execution(f"Product Links: {len(product_links)}")
        for index, product_link in enumerate(product_links, start=1):
            try:
                self.log_execution(f"[{index}] Opened: {product_link}")
                self.driver.get(product_link)
                self.log_execution(f"Navigate to URL ({product_link})")
                time.sleep(2)

                breadcrumb = None
                image_url = None
                category = None
                productTitle = None
                price = None
                sizes = None
                sizeInfo = {}
                title = None
                description = None
                itemization = None
                reviews_data = []
                rating = None
                number_of_reviews = None
                coordinated_items_info = []

                try:
                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'ol[data-auto-id="breadcrumbs-desktop"] li')))

                    breadcrumb_items = self.driver.find_elements(By.CSS_SELECTOR, 'ol[data-auto-id="breadcrumbs-desktop"] li')

                    breadcrumb_texts = []
                    for item in breadcrumb_items[1:]:
                        try:
                            name = item.find_element(By.CSS_SELECTOR, '[property="name"]').text
                            breadcrumb_texts.append(name)
                        except:
                            continue

                    breadcrumb = ' / '.join(breadcrumb_texts)
                    self.log_execution(f"Breadcrumb: {breadcrumb}")

                except Exception as e:
                    self.log_error(f"Breadcrump not found: {e}")

                try:
                    category_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-auto-id="product-category"] span')))
                    category = category_element.get_attribute('textContent').strip()

                    self.log_execution(f"Category: {category}")
                except Exception as e:
                    self.log_error(f"Category not found: {e}")

                try:
                    img_element = wait.until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, 'picture[data-testid="pdp-gallery-picture"] img')
                                ))

                    image_url = img_element.get_attribute("src")
                    self.log_execution(f"Image URL: {image_url}")
                except Exception as e:
                    self.log_error(f"Image url not found: {e}")

                try:
                    product_element = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'h1[data-auto-id="product-title"] span')))
                    productTitle = product_element.get_attribute('textContent').strip()
                    if not productTitle:
                        productTitle = self.driver.execute_script("return arguments[0].textContent;",
                                                             product_element).strip()
                    self.log_execution(f"Product: {productTitle}")
                except Exception as e:
                    self.log_error(f"Product title not found: {e}")

                try:
                    price_element = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-testid="main-price"]')))
                    spans = price_element.find_elements(By.TAG_NAME, 'span')

                    if len(spans) >= 2:
                        price = spans[1].get_attribute('textContent').strip()
                        if not price:
                            price = self.driver.execute_script("return arguments[0].textContent;", spans[1]).strip()
                        self.log_execution(f"Price: {price}")
                except Exception as e:
                    self.log_error(f"Price not found: {e}")

                try:
                    wait.until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[data-auto-id="size-selector"] button')))
                    buttons = self.driver.find_elements(By.CSS_SELECTOR, 'div[data-auto-id="size-selector"] button')

                    available_sizes = []
                    for button in buttons:
                        class_attr = button.get_attribute("class")
                        if "unavailable" not in class_attr:
                            try:
                                size_text = button.find_element(By.TAG_NAME, "span").text
                                available_sizes.append(size_text)
                            except:
                                pass

                    sizes = ', '.join(available_sizes)
                    self.log_execution(f"Available Sizes: {sizes}")
                except Exception as e:
                    self.log_error(f"Available sizes not found: {e}")

                try:
                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button[data-auto-id="size-chart-link"]')))
                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[data-auto-id="size-chart-link"]')))
                    size_guide_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-auto-id="size-chart-link"]')))
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", size_guide_btn)
                    try:
                        size_guide_btn.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", size_guide_btn)
                    wait.until(EC.presence_of_element_located((By.ID, "gl-modal__size-chart-modal")))
                    wait.until(EC.visibility_of_element_located((By.ID, "gl-modal__size-chart-modal")))
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#gl-modal__size-chart-modal table")))

                    html = self.driver.page_source
                    soup = BeautifulSoup(html, "html.parser")
                    modal = soup.find("div", id="gl-modal__size-chart-modal")

                    tables = modal.find_all("table")
                    for table in tables:
                        thead = table.find("thead")
                        if not thead:
                            continue

                        headers = [th.get_text(strip=True) for th in thead.find_all("th")]
                        size_headers = headers[1:]

                        tbody = table.find("tbody")
                        for tr in tbody.find_all("tr"):
                            cells = [td.get_text(strip=True) for td in tr.find_all(["th", "td"])]
                            if len(cells) < 2:
                                continue

                            row_label = cells[0]
                            size_values = cells[1:]

                            size_dict = {}
                            for idx, size in enumerate(size_headers):
                                if idx < len(size_values):
                                    value = size_values[idx]
                                    if value:
                                        size_dict[size] = value

                            if row_label not in sizeInfo:
                                sizeInfo[row_label] = []

                            sizeInfo[row_label].append(size_dict)

                    self.log_execution(json.dumps(sizeInfo, ensure_ascii=False, indent=2))
                    close_button = wait.until(EC.element_to_be_clickable((By.ID, "gl-modal__close-size-chart-modal")))
                    try:
                        close_button.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", close_button)
                except Exception as e:
                    self.log_error(f"Error finding in size guide button: {e}")

                try:
                    review_container = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#navigation-target-reviews"))
                    )
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", review_container)
                    try:
                        review_container.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", review_container)
                    try:
                        wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class, 'ratings-label-container')]/span")))
                        rating_span = wait.until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//div[contains(@class, 'ratings-label-container')]/span"))
                        )
                        overall_rating = rating_span.get_attribute("textContent").strip()
                        self.log_execution(f"Overall rate: {overall_rating}")
                    except Exception as e:
                        self.log_error(f"Error finding overall rating: {e}")

                    try:
                        wait.until(EC.visibility_of_element_located((By.XPATH, "//div[contains(@class, 'reviews-header')]/h2")))
                        reviews_header = wait.until(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'reviews-header')]/h2")))
                        review_text = reviews_header.get_attribute("textContent").strip()
                        match = re.search(r'\((\d+)\)', review_text)
                        number_of_reviews = int(match.group(1)) if match else 0
                        self.log_execution(f"Number of reviews: {number_of_reviews}")
                    except Exception as e:
                        self.log_error(f"Error finding overall Number of reviews: {e}")

                    while True:
                        try:
                            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//button[@data-auto-id='reviews-load-more']")))
                            wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@data-auto-id='reviews-load-more']")))
                            load_more_btn = wait.until(EC.element_to_be_clickable(
                                (By.XPATH, "//button[@data-auto-id='reviews-load-more']")
                            ))
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",
                                                       load_more_btn)
                            before_count = len(
                                self.driver.find_elements(By.CSS_SELECTOR, '[data-auto-id="review"]'))
                            try:
                                load_more_btn.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", load_more_btn)

                            wait.until(
                                lambda d: len(
                                    d.find_elements(By.CSS_SELECTOR, '[data-auto-id="review"]')) > before_count
                            )
                            time.sleep(0.5)
                        except TimeoutException:
                            print("No more 'Read more reviews' button visible.")
                            break
                        except Exception as e:
                            print(f"Error clicking load more button: {e}")
                            break
                    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[data-auto-id="review"]')))
                    review_elements = self.driver.find_elements(By.CSS_SELECTOR, '[data-auto-id="review"]')
                    self.log_execution(f"Total reviews extracted: {len(review_elements)}")

                    for review in review_elements:
                        try:
                            reviewer_element = review.find_element(By.XPATH,
                                                                   ".//span[contains(@class, 'user-name')]")
                            date_element = review.find_element(By.XPATH, ".//span[contains(@class, 'date')]")
                            title_element = review.find_element(By.TAG_NAME, "h4")
                            desc_element = review.find_element(By.XPATH, ".//div[contains(@class, 'text')]")

                            reviewer_id = reviewer_element.get_attribute("textContent").strip()
                            date = date_element.get_attribute("textContent").strip()
                            review_title = title_element.get_attribute("textContent").strip()
                            review_description = desc_element.get_attribute("textContent").strip()

                            # Rating from masks
                            mask_elements = review.find_elements(By.CSS_SELECTOR, ".gl-star-rating__mask")
                            rating = 0
                            for mask in mask_elements:
                                style = mask.get_attribute("style")
                                match = re.search(r'width:\s*(\d+)', style)
                                if match:
                                    width_percent = int(match.group(1))
                                    if width_percent >= 50:
                                        rating += 1

                            reviews_data.append({
                                "date": date,
                                "rating": rating,
                                "review_title": review_title,
                                "review_description": review_description,
                                "reviewer_id": reviewer_id
                            })

                        except Exception as e:
                            print(f"Error processing a review: {e}")
                            continue

                    self.log_execution(json.dumps(reviews_data, indent=2, ensure_ascii=False))

                except Exception as e:
                    self.log_error(f"Error finding in review container: {e}")
                try:
                    desc_container = wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#navigation-target-description")
                    ))
                    title_element = desc_container.find_element(By.TAG_NAME, "h3")
                    desc_element = desc_container.find_element(By.CSS_SELECTOR, "p.gl-vspace")
                    title = title_element.get_attribute("textContent").strip()
                    description = desc_element.get_attribute("textContent").strip()
                    self.log_execution(f"Title of description: {title}")
                    self.log_execution(f"Description: {description}")
                except Exception as e:
                    self.log_error(f"Error finding in description container: {e}")
                try:
                    spec_container = wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#navigation-target-specifications")
                    ))
                    wait.until(EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, "#navigation-target-specifications li")
                    ))
                    bullet_elements = spec_container.find_elements(By.CSS_SELECTOR, "li")
                    bullets = ["• " + li.get_attribute("textContent").strip() for li in bullet_elements if li.get_attribute("textContent").strip()]
                    made_in_text = ""
                    table_rows = spec_container.find_elements(By.CSS_SELECTOR, ".gl-table__row--body")
                    for row in table_rows:
                        cells = row.find_elements(By.CSS_SELECTOR, ".gl-table__cell")
                        if len(cells) >= 2:
                            label_elem = cells[0].find_element(By.CSS_SELECTOR, ".gl-table__cell-inner")
                            value_elem = cells[1].find_element(By.CSS_SELECTOR, ".gl-table__cell-inner")

                            label = label_elem.get_attribute("textContent").strip()
                            value = value_elem.get_attribute("textContent").strip()

                            if "生産国" in label and value:
                                made_in_text = f"• {label}: {value}"
                                break

                    itemization = "\n".join(bullets + ([made_in_text] if made_in_text else []))

                    self.log_execution("General Description (itemization):\n" + itemization)

                except Exception as e:
                    self.log_error(f"Error extracting specifications: {e}")

                try:
                    carousel = wait.until(EC.presence_of_element_located((By.ID, "gl-carousel-system")))
                    wait.until(EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, "#gl-carousel-system a[data-testid='style-card']")))
                    link_index = 0
                    links = carousel.find_elements(By.CSS_SELECTOR, "a[data-testid='style-card']")
                    while True:
                        try:
                            if link_index >= len(links):
                                break

                            link = links[link_index]
                            href = link.get_attribute("href")
                            if not href:
                                link_index += 1
                                continue
                            try:
                                self.driver.get(href)
                            except Exception as e:
                                self.log_error(f"Failed ({href}): {e}")
                                continue
                            try:
                                product_cards = wait.until(
                                    EC.presence_of_all_elements_located(
                                        (By.CSS_SELECTOR, '[data-testid="product-card"]'))
                                )

                                for card in product_cards:
                                    try:
                                        product_link = card.find_element(By.CSS_SELECTOR, 'a')
                                        product_page_url = product_link.get_attribute("href")
                                        product_number = product_page_url.split('/')[-1].split('.')[0]

                                        image = card.find_element(By.CSS_SELECTOR, 'img')
                                        image_url = image.get_attribute("src")

                                        price_element = card.find_element(
                                            By.CSS_SELECTOR, '[data-testid="main-price"] span:nth-child(2)'
                                        )
                                        price = price_element.text.strip()

                                        coordinated_items_info.append({
                                            "product_page_url": product_page_url,
                                            "product_number": product_number,
                                            "image_url": image_url,
                                            "price": price
                                        })
                                    except Exception as e:
                                        self.log_error(f"Error processing coordinated product: {e}")
                            except Exception as e:
                                self.log_error(f"Error finding product cards: {e}")

                        except Exception as e:
                            self.log_error(f"Error finding coordinated item links: {e}")

                        link_index += 1

                except Exception as e:
                    self.log_error(f"Error handling coordinated products: {e}")

                json_data = json.dumps(coordinated_items_info, ensure_ascii=False, indent=2)
                self.log_execution(f"All coordinated items info JSON:\n{json_data}")

                if all([breadcrumb, category, productTitle, price, image_url]):
                    row = {
                        "Breadcrumb": breadcrumb,
                        "Image URL": image_url,
                        "Category": category,
                        "Product title": productTitle,
                        "Price": price,
                        "Sizes": sizes,
                        "Size info": sizeInfo,
                        "Title of description": title,
                        "Description": description,
                        "General description (itemization)": itemization,
                        "Rating": rating,
                        "Number of reviews": number_of_reviews,
                        "User Reviews": reviews_data,
                        "Coordinated product info": coordinated_items_info,
                    }
                    rows.append(row)
                if index % 3 == 0:
                    self.driver.quit()
                    time.sleep(5)
                    self.launch_driver()
                    wait = WebDriverWait(self.driver, 60)
            except Exception as e:
                self.log_error(f"Failed to open ({product_link}): {e}")
                continue

        df_new = pd.DataFrame(rows)

        if os.path.exists(excel_path):
            os.remove(excel_path)

        df_new.to_excel(excel_path, index=False)
        self.log_execution(f"New Excel file created with data: {excel_path}")

if __name__ == "__main__":
    pytest.main()