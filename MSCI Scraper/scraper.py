import os

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import pickle
import datetime
import os

session = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
data = None

if len(os.listdir("CSV/")) >= 1:
    scraped_tickers = pd.read_csv("CSV/" + os.listdir("CSV/")[-1])
    scraped_tickers = scraped_tickers["temp_ticker"].to_list()
else:
    scraped_tickers = []


def is_element_exists_by_xpath(driver, xpath):
    try:
        element = driver.find_element(By.XPATH, xpath)
        return True
    except NoSuchElementException:
        return False


with open("tickers.pickle", 'rb') as f:
    tickers = pickle.load(f)
f.close()

chrome_options = Options()
chrome_options.add_argument("--headless")
article_driver = webdriver.Chrome(options=chrome_options)

url = 'https://www.msci.com/our-solutions/esg-investing/esg-ratings-climate-search-tool/issuer/microsoft-corporation/IID000000002143620'


while len(tickers) - len(scraped_tickers) > 0:
    for company in tickers:
        try:
            if company in scraped_tickers:
                continue

            driver = webdriver.Chrome()
            driver.get(url)

            company_data = dict()

            if is_element_exists_by_xpath(driver, xpath='//*[@id="_esgratingsprofile_keywords"]'):
                input_field = driver.find_element(By.XPATH, '//*[@id="_esgratingsprofile_keywords"]')
                input_field.send_keys(company)
                time.sleep(5)
            else:
                continue

            if is_element_exists_by_xpath(driver, '//*[@id="ui-id-1"]/li'):
                element = driver.find_element(By.XPATH, '//*[@id="ui-id-1"]/li')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(5)

                company_data["temp_ticker"] = company

                company_data["Company Name"] = driver.find_element(By.XPATH,
                                                                   '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/h1').text
                company_data["Ticker"] = driver.find_element(By.XPATH,
                                                             '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/div[1]').text
                company_data["Industry"] = driver.find_element(By.XPATH,
                                                               '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/div[2]/div[1]').text
                company_data["Country_Region"] = driver.find_element(By.XPATH,
                                                                     '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/div[2]/div[2]').text
            else:
                continue

            if is_element_exists_by_xpath(driver, '//*[@id="esg-commitment-toggle-link"]'):
                element = driver.find_element(By.XPATH,
                                              '//*[@id="esg-commitment-toggle-link"]')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(2)

                try:
                    company_data["decarbonization target"] = driver.find_element(By.XPATH,
                                                                                 '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[1]/p[2]').text
                except:
                    pass

                try:
                    company_data[
                        "decarbonization target that is considered in the calculation of Implied Temperature Rise?"] = driver.find_element(
                        By.XPATH,
                        '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[2]/p[2]').text
                except:
                    pass

                try:
                    company_data["Target Year"] = driver.find_element(By.XPATH,
                                                                      '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[3]/p[2]').text
                except:
                    pass

                try:
                    company_data["Comprehensiveness"] = driver.find_element(By.XPATH,
                                                                            '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[4]/p').text
                except:
                    pass

                try:
                    company_data["Ambition"] = driver.find_element(By.XPATH,
                                                                   '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[5]/p').text
                except:
                    pass

            if is_element_exists_by_xpath(driver, '//*[@id="esg-climate-toggle-link"]'):
                element = driver.find_element(By.XPATH,
                                              '//*[@id="esg-climate-toggle-link"]')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(2)

                try:
                    company_data["MSCI IMPLIED TEMPERATURE RISE"] = driver.find_element(By.XPATH,
                                                                                        '//*[@id="_esgratingsprofile_esg-company-transparency"]/div[1]/div[3]/span').text
                except:
                    pass

            if is_element_exists_by_xpath(driver, '//*[@id="esg-transparency-toggle-link"]'):
                element = driver.find_element(By.XPATH,
                                              '//*[@id="esg-transparency-toggle-link"]')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(5)

                if is_element_exists_by_xpath(driver,
                                              '//*[@id="_esgratingsprofile_esg-ratings-profile-header"]/div/div[1]/div[2]/div'):
                    company_data["MSCI ESG RATING"] = driver.find_element(By.XPATH,
                                                                          '//*[@id="_esgratingsprofile_esg-ratings-profile-header"]/div/div[1]/div[2]/div').get_attribute(
                        'class').split('-')[-1].upper()

            if is_element_exists_by_xpath(driver,
                                          '//*[@id="esg-controversies-toggle-link"]'):
                element = driver.find_element(By.XPATH,
                                              '//*[@id="esg-controversies-toggle-link"]')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(5)

                for i in range(1, 20):
                    try:
                        company_data[f'cont_col{i}'] = \
                        driver.find_element(By.XPATH, f'//*[@id="controversies-table"]/div[{i}]').get_attribute(
                            'class').split(
                            '-')[-1].title()
                    except:
                        continue

            if is_element_exists_by_xpath(driver, '//*[@id="esg-involvement-toggle-link"]'):
                element = driver.find_element(By.XPATH,
                                              '//*[@id="esg-involvement-toggle-link"]')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(2)

                try:
                    company_data["Banned Controversial Weapons"] = driver.find_element(By.XPATH,
                                                                                       '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[1]/div[1]/div[3]').text.title()
                except:
                    pass

                try:
                    company_data["Gambling"] = \
                        driver.find_element(By.XPATH,
                                            '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[2]/div[1]/div[3]').text.title()
                except:
                    pass
                try:
                    company_data["Tobacco"] = \
                        driver.find_element(By.XPATH,
                                            '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[3]/div[1]/div[3]').text.title()
                except:
                    pass

                try:
                    company_data["Alcohol"] = \
                        driver.find_element(By.XPATH,
                                            '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[4]/div[1]/div[3]').text.title()
                except:
                    pass

            if is_element_exists_by_xpath(driver, '//*[@id="esg-sdg-alignment-toggle-link"]'):
                element = driver.find_element(By.XPATH, '//*[@id="esg-sdg-alignment-toggle-link"]')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(2)

                goals_list = [
                    "No Poverty",
                    "No Hunger",
                    "Good Health and Well-Being",
                    "Quality Education",
                    "Gender Equality",
                    "Clean Water and Sanitation",
                    "Affordable and Clean Energy",
                    "Decent Work and Economic Growth",
                    "Industry, Innovation and Infrastructure",
                    "Reduced Inequalities",
                    "Sustainable Cities and Communities",
                    "Responsible Consumption and Production",
                    "Climate Action",
                    "Life under Water",
                    "Life on Land",
                    "Peace, Justice and Strong Institutions",
                    "Partnerships for the Goals"
                ]

                for i in range(2, 19):
                    if is_element_exists_by_xpath(driver,
                                                  f'//*[@id="esg-sdg-alignment-toggle"]/div/div[1]/div/div[{i}]'):
                        company_data[goals_list[i - 2]] = driver.find_element(By.XPATH,
                                                                              f'//*[@id="esg-sdg-alignment-toggle"]/div/div[1]/div/div[{i}]').get_attribute(
                            'class').title()

            for i in company_data.keys():
                company_data[i] = [company_data[i]]

            driver.quit()

            if data is None:
                data = pd.DataFrame.from_dict(company_data)
            else:
                data = pd.concat([data, pd.DataFrame.from_dict(company_data)], ignore_index=True)

            data.to_csv(f"CSV/{session}.csv")


        except Exception as e:
            if len(os.listdir("CSV/")) >= 1:
                scraped_tickers = pd.read_csv("CSV/" + os.listdir("CSV/")[-1])
                scraped_tickers = scraped_tickers["temp_ticker"].to_list()
            else:
                scraped_tickers = []

            print(e)
            driver.quit()
            continue
