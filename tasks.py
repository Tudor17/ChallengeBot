import re
import os
import time
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl
from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class NewsScraper:
    def __init__(self, search_phrase, section, months_back):
        self.browser = Selenium()
        self.search_phrase = search_phrase
        self.section = section
        self.months_back = max(1, months_back)
        self.target_date = self._get_target_date()
        self.results = []
        self.excel_file = "news_data.xlsx"
        self.screenshot_directory = os.path.join("output", self.search_phrase)
        os.makedirs(self.screenshot_directory, exist_ok=True)

    def _get_target_date(self):
        target_date = datetime.now() - relativedelta(months=self.months_back - 1)
        return target_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

    def _extract_money(self, text):
        money_pattern = r'\$[\d,.]+|\d+ dollars|\d+ USD'
        matches = re.findall(money_pattern, text)
        return len(matches) > 0

    def open_site(self, url):
        self.browser.open_available_browser(url, browser_selection="chrome", maximized=True)

    def search_news(self):
        WebDriverWait(self.browser.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/ps-header/header/div[2]/button"))
        )
        self.browser.click_element("xpath:/html/body/ps-header/header/div[2]/button")

        WebDriverWait(self.browser.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/ps-header/header/div[2]/div[2]/form/label/input"))
        )
        search_input = self.browser.driver.find_element(By.XPATH, '/html/body/ps-header/header/div[2]/div[2]/form/label/input')
        search_input.send_keys(self.search_phrase)
        search_input.submit()

        WebDriverWait(self.browser.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select"))
        )
        self.browser.select_from_list_by_value(
            "xpath:/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select", "1"
        )

    def filter_by_section(self):
        WebDriverWait(self.browser.driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "checkbox-input-label"))
        )
        labels = self.browser.driver.find_elements(By.CLASS_NAME, 'checkbox-input-label')
        matching_element = next((label for label in labels if self.section.lower() in self.browser.get_text(label).lower()), None)
        if matching_element:
            self.browser.click_element(matching_element)
        time.sleep(10)

    def scrape_results(self):
        self.browser.set_screenshot_directory(self.screenshot_directory)
        last_page = False
        counter = 1

        while not last_page:
            WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "promo-wrapper")))
            news_elements = self.browser.driver.find_elements(By.CLASS_NAME, 'promo-wrapper')
            for index, news in enumerate(news_elements):
                WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[2]/p[2]')))
                date_element = self.browser.driver.find_element(By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[2]/p[2]')
                date = datetime.fromtimestamp(float(self.browser.get_element_attribute(date_element, 'data-timestamp')) / 1000)

                if date < self.target_date:
                    last_page = True
                    break

                WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[2]/div/h3')))
                title = self.browser.get_text(self.browser.driver.find_element(By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[2]/div/h3'))

                WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[2]/p[1]')))
                description = self.browser.get_text(self.browser.driver.find_element(By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[2]/p[1]'))

                WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[1]')))
                image_filename = self.browser.capture_element_screenshot(self.browser.driver.find_element(By.XPATH, f'/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li[{str(index+1)}]/ps-promo/div/div[1]'), str(counter) + '.png')

                title_count = title.lower().count(self.search_phrase.lower())
                description_count = description.lower().count(self.search_phrase.lower())
                contains_money = self._extract_money(title) or self._extract_money(description)

                self.results.append({
                    "title": title,
                    "date": date,
                    "description": description,
                    "image_filename": image_filename,
                    "title_count": title_count,
                    "description_count": description_count,
                    "contains_money": contains_money
                })
                counter += 1

            try:
                WebDriverWait(self.browser.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[2]/div[3]/a"))
                ).click()
            except:
                last_page = True

    def save_to_excel(self):
        if not os.path.exists(self.excel_file):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["Title", "Date", "Description", "Picture Filename", "Search Phrase Count (Title)", "Search Phrase Count (Description)", "Contains Money"])
        else:
            workbook = openpyxl.load_workbook(self.excel_file)
            sheet = workbook.active

        for result in self.results:
            sheet.append([
                result["title"], result["date"], result["description"], result["image_filename"],
                result["title_count"], result["description_count"], result["contains_money"]
            ])

        workbook.save(self.excel_file)
        shutil.move(self.excel_file, os.path.join(self.screenshot_directory, self.excel_file))

    def run(self):
        try:
            self.open_site("https://www.latimes.com/")
            self.search_news()
            #self.filter_by_section()
            self.scrape_results()
            self.save_to_excel()
        finally:
            self.browser.close_all_browsers()


scraper = NewsScraper(search_phrase='Donald Trump', section="business", months_back=1)
scraper.run()