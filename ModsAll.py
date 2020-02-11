import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.common.exceptions import TimeoutException
import time
import csv
from datetime import datetime
from selenium.common.exceptions import WebDriverException
from selenium.common import exceptions

# Use driver to locate information
driver = webdriver.Edge(executable_path="C://Windows//SysWOW64//MicrosoftWebDriver.exe")
driver.maximize_window()
# Using Edge to open the website
driver.get("https://www.ageofempires.com/mods")

# Do not comment out below line otherwise it will fail due to being slow on responding
driver.implicitly_wait(10)

table = driver.find_element_by_css_selector('#mods-listing > table')
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/ModsAll_%Y%m%d_%H%M.csv')
with open(filename, 'w', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)
    for row in table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])

driver.find_element_by_xpath('//*[@id="mods-paginav"]/ul/li[2]/button').click()
time.sleep(3)
table = driver.find_element_by_css_selector('#mods-listing > table')
with open(filename, 'a', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)
    for row in table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])

driver.find_element_by_xpath('//*[@id="mods-paginav"]/ul/li[3]/button').click()
time.sleep(3)
table = driver.find_element_by_css_selector('#mods-listing > table')
with open(filename, 'a', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)
    for row in table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])

i = 0
while i < 89:
    driver.find_element_by_xpath('//*[@id="mods-paginav"]/ul/li[5]/button').click()
    time.sleep(3)
    table = driver.find_element_by_css_selector('#mods-listing > table')
    with open(filename, 'a', newline='', encoding="utf-8") as csvfile:
        wr = csv.writer(csvfile)
        for row in table.find_elements_by_css_selector('tr'):
            wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])
    i += 1
else:
    print("This is the last page! ")
print("Finished... ")
driver.quit();


