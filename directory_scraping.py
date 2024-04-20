import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd


URL = "https://www.whps.org/directory"
OUTPUT_FILE = "directory_output.xlsx"

ITEM_CONTAINER_CLASS = "fsConstituentItem"
FULL_NAME_CLASS = "fsFullName"
TITLE_CLASS = "fsTitles"
NEXTPAGE_BUTTON_CLASS = "fsNextPageLink"
FIRSTPAGE_BUTTON_CLASS = "fsFirstPageLink"

WAIT_TIME = 2  # seconds


def setup_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    service = webdriver.ChromeService()
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def go_to_first_page(driver):
    try:
        first_page_button = driver.find_element(By.CLASS_NAME, FIRSTPAGE_BUTTON_CLASS)
        first_page_button.click()
        print("Navigated to the first page.")
    except NoSuchElementException:
        print("First page link not found; assuming the page is already at the start.")


def scrape_directory(driver):
    items = driver.find_elements(By.CLASS_NAME, ITEM_CONTAINER_CLASS)
    data = []
    for item in items:
        name = item.find_element(By.CLASS_NAME, FULL_NAME_CLASS).text.strip()
        title_element = item.find_elements(By.CLASS_NAME, TITLE_CLASS)
        title = title_element[0].text.strip() if title_element else "N/A"
        email_elements = item.find_elements(
            By.XPATH, ".//a[contains(@href, 'mailto:')]"
        )
        email = (
            email_elements[0].get_attribute("href").replace("mailto:", "")
            if email_elements
            else "N/A"
        )
        data.append({"Name": name, "Title": title, "Email": email})
    return data


def loop_through_pages(driver, all_data):
    while True:
        time.sleep(WAIT_TIME)
        new_data = scrape_directory(driver)
        all_data.extend(new_data)
        print(f"Scraped {len(new_data)} entries on this page.")

        try:
            next_button = driver.find_element(By.CLASS_NAME, NEXTPAGE_BUTTON_CLASS)
            next_button.click()
        except NoSuchElementException:
            print("No more pages to scrape.")
            break


def main():
    driver = setup_driver()
    driver.get(URL)
    go_to_first_page(driver)
    all_data = []

    loop_through_pages(driver, all_data)

    driver.quit()

    # Exporting data to Excel
    df = pd.DataFrame(all_data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Data exported to '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    main()
