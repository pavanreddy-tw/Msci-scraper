from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd
from bs4 import BeautifulSoup
import requests
import sys
import os

columns = ["Company Name", "Company Abbr", "Industry", "Country_Region", "decarbonization target",
           "decarbonization target that is considered in the calculation of Implied Temperature Rise?",
           "Target Year", "Comprehensiveness", "Ambition", "MSCI IMPLIED TEMPERATURE RISE",
           "MSCI ESG RATING", "Environment", "Social", "Customers", "Anticompetitive Practices", "Privacy & Data Security", "Human Rights & Community",
           "Labor Rights & Supply Chain", "Discrimination & Workforce Diversity", "Governance", "Banned Controversial Weapons", "Gambling", "Tobacco", "Alcohol"]

data = None



def is_element_exists_by_xpath(driver, xpath):
    try:
        element = driver.find_element(By.XPATH,xpath)
        return True
    except NoSuchElementException:
        return False



driver=webdriver.Chrome()

chrome_options = Options()
chrome_options.add_argument("--headless")
article_driver=webdriver.Chrome(options=chrome_options)

url = 'https://www.msci.com/our-solutions/esg-investing/esg-ratings-climate-search-tool/issuer/microsoft-corporation/IID000000002143620'
driver.get(url)




for company in ["AAPL", "MSFT", "NFLX", "META"]:
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

        company_data[columns[0]] = driver.find_element(By.XPATH, '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/h1').text
        company_data[columns[1]] = driver.find_element(By.XPATH,
                                                           '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/div[1]').text
        company_data[columns[2]] = driver.find_element(By.XPATH, '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/div[2]/div[1]').text
        company_data[columns[3]] = driver.find_element(By.XPATH, '//*[@id="_esgratingsprofile_esg-ratings-profile-container"]/div[1]/div[2]/div[2]').text
    else:
        continue

    dropdown_xpaths = ['//*[@id="esg-commitment-toggle-link"]',
                       '//*[@id="esg-climate-toggle-link"]',
                       '//*[@id="esg-transparency-toggle-link"]',
                       '//*[@id="esg-controversies-toggle-link"]',
                       '//*[@id="esg-involvement-toggle-link"]',
                       '//*[@id="esg-sdg-alignment-toggle-link"]']

    if is_element_exists_by_xpath(driver, '//*[@id="esg-commitment-toggle-link"]'):
        element = driver.find_element(By.XPATH, '//*[@id="esg-commitment-toggle-link"]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(2)

        company_data[columns[4]] = driver.find_element(By.XPATH,
                                                       '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[1]/p[2]').text
        company_data[columns[5]] = driver.find_element(By.XPATH,
                                                       '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[2]/p[2]').text
        company_data[columns[6]] = driver.find_element(By.XPATH,
                                                       '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[3]/p[2]').text
        company_data[columns[7]] = driver.find_element(By.XPATH,
                                                       '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[4]/p').text
        company_data[columns[8]] = driver.find_element(By.XPATH,
                                                      '//*[@id="esg-commitment-toggle"]/div/div[1]/div/div[2]/div[5]/p').text

    if is_element_exists_by_xpath(driver, '//*[@id="esg-climate-toggle-link"]'):
        element = driver.find_element(By.XPATH, '//*[@id="esg-climate-toggle-link"]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(2)

        company_data[columns[9]] = driver.find_element(By.XPATH,
                                                       '//*[@id="_esgratingsprofile_esg-company-transparency"]/div[1]/div[3]/span').text

    if is_element_exists_by_xpath(driver, '//*[@id="esg-transparency-toggle-link"]'):
        element = driver.find_element(By.XPATH, '//*[@id="esg-transparency-toggle-link"]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(5)

        company_data[columns[10]] = driver.find_element(By.XPATH, '//*[@id="_esgratingsprofile_esg-ratings-profile-header"]/div/div[1]/div[2]/div').get_attribute('class').split('-')[-1].upper()

    if is_element_exists_by_xpath(driver, '// *[ @ id = "esg-controversies-toggle-link"]'):
        element = driver.find_element(By.XPATH, '// *[ @ id = "esg-controversies-toggle-link"]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(2)

        company_data[columns[11]] = driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[1]').get_attribute('class').split('-')[-1].title()
        company_data[columns[12]] = \
        driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[2]').get_attribute('class').split('-')[
            -1].title()
        company_data[columns[13]] = \
        driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[3]').get_attribute('class').split('-')[
            -1].title()
        company_data[columns[14]] = \
        driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[4]').get_attribute('class').split('-')[
            -1].title()
        company_data[columns[15]] = \
        driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[7]').get_attribute('class').split('-')[
            -1].title()
        company_data[columns[16]] = \
            driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[10]').get_attribute('class').split('-')[
                -1].title()
        company_data[columns[17]] = \
            driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[11]').get_attribute('class').split('-')[
                -1].title()
        company_data[columns[18]] = \
            driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[14]').get_attribute('class').split('-')[
                -1].title()
        company_data[columns[18]] = \
            driver.find_element(By.XPATH, '//*[@id="controversies-table"]/div[19]').get_attribute('class').split('-')[
                -1].title()


    if is_element_exists_by_xpath(driver, '//*[@id="esg-involvement-toggle-link"]'):
        element = driver.find_element(By.XPATH, '//*[@id="esg-involvement-toggle-link"]')
        driver.execute_script("arguments[0].click();", element)
        time.sleep(2)

        company_data[columns[19]] = driver.find_element(By.XPATH, '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[1]/div[1]/div[3]').get_attribute('class').split('-')[-1].title()
        company_data[columns[20]] = \
        driver.find_element(By.XPATH, '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[2]/div[1]/div[3]').get_attribute('class').split('-')[
            -1].title()
        company_data[columns[21]] = \
        driver.find_element(By.XPATH, '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[3]/div[1]/div[3]').get_attribute('class').split('-')[
            -1].title()
        company_data[columns[22]] = \
        driver.find_element(By.XPATH, '//*[@id="esg-involvement-toggle"]/div/div[1]/div[3]/div[4]/div[1]/div[3]').get_attribute('class').split('-')[
            -1].title()

    if data is None:
        data = pd.DataFrame(company_data, index=[0])
    else:
        data = data.append(pd.DataFrame(company_data, index=[0]), ignore_index=True)


print(data)