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

# Locate Edge driver
driver = webdriver.Edge(executable_path="C://Windows//SysWOW64//MicrosoftWebDriver.exe")
driver.maximize_window()
# Using Edge to open the steam website
driver.get("https://partner.steampowered.com")

# Pause the driver for better performance
driver.implicitly_wait(10)

# Enter email address
login_un = driver.find_element_by_id('username').send_keys("")
# Enter password
login_pw = driver.find_element_by_id('password').send_keys("")
# Click sign in tp log in
driver.find_element_by_id('login_btn_signin').click()

# Find the desired link
driver.find_element_by_link_text('Age of Empires II: Definitive Edition').click()
time.sleep(3)
# Locate the option: 1 year
driver.find_element_by_css_selector('#ChartUnitsHistory_ranges > a:nth-child(1)').click()
time.sleep(3)

# Locate and click link to download csv. Need to enable the auto-download option with Edge to proceed auto downloading.
driver.find_element_by_link_text('view as .csv').click()
time.sleep(3)
print("AoE_II_Definitive Edition Units Sold csv is saved. ")

# Open local csv file and save data for Aoe_II_DE_Steam_details
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/AoE_II_DE_Steam_Details_%Y%m%d_%H%M.csv')
time.sleep(3)
with open(filename, 'w', newline='') as csvfile:
    wr = csv.writer(csvfile)
    a1 = "    "
    b1 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[1]/td[3]').text
    c1 = "Daily average during period"  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[1]/td[5]').text
    d1 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[1]/td[6]').text
    e1 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[1]/td[8]').text
    f1 = "Previous daily average"  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[1]/td[10]').text
    wr.writerow([a1, b1, c1, d1, e1, f1])

    a2 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[2]/td[1]').text
    b2 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[2]/td[3]').text
    c2 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[2]/td[5]').text
    d2 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[2]/td[6]').text
    e2 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[2]/td[8]').text
    f2 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[2]/td[10]').text
    wr.writerow([a2, b2, c2, d2, e2, f2])

    a3 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[3]/td[1]/em/a').text
    b3 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[3]/td[3]').text
    c3 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[3]/td[5]').text
    d3 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[3]/td[6]').text
    e3 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[3]/td[8]').text
    f3 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[3]/td[10]').text
    wr.writerow([a3, b3, c3, d3, e3, f3])

    a4 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[5]/td[1]').text
    b4 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[5]/td[3]').text
    c4 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[5]/td[5]').text
    d4 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[5]/td[6]').text
    e4 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[5]/td[8]').text
    f4 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[5]/td[10]').text
    wr.writerow([a4, b4, c4, d4, e4, f4])

    a5 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[6]/td[1]/em/a').text
    b5 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[6]/td[3]').text
    c5 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[6]/td[5]').text
    d5 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[6]/td[6]').text
    e5 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[6]/td[8]').text
    f5 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[6]/td[10]').text
    wr.writerow([a5, b5, c5, d5, e5, f5])

    a6 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[7]/td[1]/em/a').text
    b6 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[7]/td[3]').text
    c6 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[7]/td[5]').text
    d6 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[7]/td[6]').text
    e6 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[7]/td[8]').text
    f6 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[7]/td[10]').text
    wr.writerow([a6, b6, c6, d6, e6, f6])

    a7 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[8]/td[1]/em/a').text
    b7 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[8]/td[3]').text
    c7 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[8]/td[5]').text
    d7 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[8]/td[6]').text
    e7 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[8]/td[8]').text
    f7 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[8]/td[10]').text
    wr.writerow([a7, b7, c7, d7, e7, f7])

    a8 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[9]/td[1]/em/a').text
    b8 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[9]/td[3]').text
    c8 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[9]/td[5]').text
    d8 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[9]/td[6]').text
    e8 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[9]/td[8]').text
    f8 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[9]/td[10]').text
    wr.writerow([a8, b8, c8, d8, e8, f8])

    a9 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[11]/td[1]').text
    b9 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[11]/td[3]').text
    c9 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[11]/td[5]').text
    d9 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[11]/td[6]').text
    e9 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[11]/td[8]').text
    f9 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[11]/td[10]').text
    wr.writerow([a9, b9, c9, d9, e9, f9])

    a10 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[13]/td[1]').text
    b10 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[13]/td[3]').text
    c10 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[13]/td[5]').text
    d10 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[13]/td[6]').text
    e10 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[13]/td[8]').text
    f10 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[13]/td[10]').text
    wr.writerow([a10, b10, c10, d10, e10, f10])

    a11 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[14]/td[1]/em/a').text
    b11 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[14]/td[3]').text
    c11 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[14]/td[5]').text
    d11 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[14]/td[6]').text
    e11 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[14]/td[8]').text
    f11 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[14]/td[10]').text
    wr.writerow([a11, b11, c11, d11, e11, f11])

    a12 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[16]/td[1]').text
    b12 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[16]/td[3]').text
    c12 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[16]/td[5]').text
    d12 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[16]/td[6]').text
    e12 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[16]/td[8]').text
    f12 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[16]/td[10]').text
    wr.writerow([a12, b12, c12, d12, e12, f12])

    a13 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[18]/td').text
    b13 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[18]/td[3]').text
    c13 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[18]/td[5]').text
    d13 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[18]/td[6]').text
    e13 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[18]/td[8]').text
    f13 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[18]/td[10]').text
    wr.writerow([a13, b13, c13, d13, e13, f13])

    a14 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[19]/td[1]/a').text
    b14 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[19]/td[3]').text
    c14 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[19]/td[5]').text
    d14 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[19]/td[6]').text
    e14 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[19]/td[8]').text
    f14 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[19]/td[10]').text
    wr.writerow([a14, b14, c14, d14, e14, f14])

    a15 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[21]/td[1]').text
    b15 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[21]/td[3]').text
    c15 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[21]/td[5]').text
    d15 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[21]/td[6]').text
    e15 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[21]/td[8]').text
    f15 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[21]/td[10]').text
    wr.writerow([a15, b15, c15, d15, e15, f15])

    a16 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[22]/td[1]/em/a').text
    b16 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[22]/td[3]').text
    c16 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[22]/td[5]').text
    d16 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[22]/td[6]').text
    e16 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[22]/td[8]').text
    f16 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[22]/td[10]').text
    wr.writerow([a16, b16, c16, d16, e16, f16])

    a17 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[23]/td[1]/em/a').text
    b17 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[23]/td[3]').text
    c17 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[23]/td[5]').text
    d17 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[23]/td[6]').text
    e17 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[23]/td[8]').text
    f17 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[23]/td[10]').text
    wr.writerow([a17, b17, c17, d17, e17, f17])

    a18 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[24]/td[1]/em/a').text
    b18 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[24]/td[3]').text
    c18 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[24]/td[5]').text
    d18 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[24]/td[6]').text
    e18 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[24]/td[8]').text
    f18 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[24]/td[10]').text
    wr.writerow([a18, b18, c18, d18, e18, f18])

    a19 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[25]/td[1]/em/a').text
    b19 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[25]/td[3]').text
    c19 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[25]/td[5]').text
    d19 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[25]/td[6]').text
    e19 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[25]/td[8]').text
    f19 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[25]/td[10]').text
    wr.writerow([a19, b19, c19, d19, e19, f19])

    a20 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[26]/td[1]/em/a').text
    b20 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[26]/td[3]').text
    c20 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[26]/td[5]').text
    d20 = "    "  # driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[26]/td[6]').text
    e20 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[26]/td[8]').text
    f20 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[26]/td[10]').text
    wr.writerow([a20, b20, c20, d20, e20, f20])

    a21 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[28]/td[1]').text
    b21 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[28]/td[3]').text
    c21 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[28]/td[5]').text
    d21 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[28]/td[6]').text
    e21 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[28]/td[8]').text
    f21 = driver.find_element_by_xpath('//*[@id="gameDataLeft"]/div[4]/table/tbody/tr[28]/td[10]').text
    wr.writerow([a21, b21, c21, d21, e21, f21])

print("AoE_II_DE_Steam_Details data is saved. ")
time.sleep(3)
# # Locate the the detail table for AoE steams
# AoE_Steam_table = driver.find_element_by_css_selector('#gameDataLeft > div:nth-child(6) > table')
# # Open local csv file and save data
# filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/AoE_II_DE_Steam_details_%Y%m%d_%H%M.csv')
# with open(filename, 'w', newline='') as csvfile:
#     wr = csv.writer(csvfile)
#     for row in AoE_Steam_table.find_elements_by_css_selector('tr'):
#         wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])
# time.sleep(3)
# print("AoE_II_DE_Steam_Details data is saved. ")

# Locate the table element
AoE_Summary_table = driver.find_element_by_css_selector('#gameDataLeft > div:nth-child(1) > table')
# Open local csv file and save data
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/AoE_II_DE_Summary_%Y%m%d_%H%M.csv')
with open(filename, 'w', newline='') as csvfile:
    wr = csv.writer(csvfile)
    for row in AoE_Summary_table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])
time.sleep(3)
print("AoE_II_DE_Summary data is saved. ")

# Locate the link for Current Players
driver.find_element_by_css_selector(
    '#gameDataLeft > div:nth-child(1) > table > tbody > tr:nth-child(9) > td:nth-child(3) > a').click()
time.sleep(5)

# Locate 1 year for Current Players
driver.find_element_by_xpath('/html/body/center/div/div[3]/div[1]/em[1]').click()
# x.click()
time.sleep(3)

# Locate the summary statistics table for concurrent players
CCP_Summary_table = driver.find_element_by_css_selector("body > center > div > table")

# Open local csv and save data
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/Summary_Statistics_%Y%m%d_%H%M.csv')
with open(filename, 'w', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)
    for row in CCP_Summary_table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])
print("Concurrent player summary statistics data is saved. ")
time.sleep(3)

# # Locate the table element for CCP
# Concurrent_player_table = driver.find_element_by_css_selector('body > center > div > div:nth-child(13) > table')
#
# # Open local csv and save data
# filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/Concurrent_Players_%Y%m%d_%H%M.csv')
# with open(filename, 'w', newline='') as csvfile:
#     wr = csv.writer(csvfile)
#     for row in Concurrent_player_table .find_elements_by_xpath(".//tr"):
#         print([col.text for col in row.find_elements_by_xpath(".//td")])
# print("Concurrent_Player data is saved. ")

# Open local csv and save data for CCP
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/Concurrent_Play_Statistics_%Y%m%d_%H%M.csv')
with open(filename, 'w', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)

    i = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[1]/td[1]').text
    ii = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[1]/td[3]').text
    iii = "Daily average during period"  # driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[1]/td[5]').text
    iv = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[1]/td[6]').text
    v = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[1]/td[8]').text
    vi = "Previous daily average"  # driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[1]/td[10]').text
    wr.writerow([i, ii, iii, iv, v, vi])

    a = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[2]/td[1]').text
    b = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[2]/td[3]').text
    c = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[2]/td[5]').text
    d = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[2]/td[6]').text
    e = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[2]/td[8]').text
    f = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[2]/td[10]').text
    wr.writerow([a, b, c, d, e, f])

    g = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[3]/td[1]').text
    h = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[3]/td[3]').text
    i = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[3]/td[5]').text
    j = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[3]/td[6]').text
    k = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[3]/td[8]').text
    l = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[3]/td[10]').text
    wr.writerow([g, h, i, j, k, l])

    m = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[4]/td[1]').text
    n = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[4]/td[3]').text
    o = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[4]/td[5]').text
    p = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[4]/td[6]').text
    q = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[4]/td[8]').text
    r = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[4]/td[10]').text
    wr.writerow([m, n, o, p, q, r])

    s = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[5]/td[1]').text
    t = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[5]/td[3]').text
    u = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[5]/td[5]').text
    v = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[5]/td[6]').text
    w = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[5]/td[8]').text
    x = driver.find_element_by_xpath('/html/body/center/div/div[4]/table/tbody/tr[5]/td[10]').text
    wr.writerow([s, t, u, v, w, x])
print("Concurrent_Play_Statistics data is saved. ")
time.sleep(3)

# Go back to the main page
driver.find_element_by_css_selector('body > center > div > a').click()
time.sleep(3)

# Locate and click Median Time Played page
driver.find_element_by_css_selector(
    '#gameDataLeft > div:nth-child(1) > table > tbody > tr:nth-child(12) > td:nth-child(2) > a').click()
time.sleep(3)

# Locate the median time table
Median_Time_table = driver.find_element_by_css_selector('#gameDataLeft > table:nth-child(1)')

# Open local csv and save data
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/Median_Time_data_%Y%m%d_%H%M.csv')
with open(filename, 'w', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)
    for row in Median_Time_table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])
print("Median Time data is saved. ")
time.sleep(3)

# Locate the minimum time table
Minimun_Time_table = driver.find_element_by_css_selector('#gameDataLeft > table:nth-child(4)')

# Open local csv and save data
filename = datetime.now().strftime('C:/Users/v-huan/Desktop/Output/Minimun_Time_data_%Y%m%d_%H%M.csv')
with open(filename, 'w', newline='', encoding="utf-8") as csvfile:
    wr = csv.writer(csvfile)
    a = "Minimum time played"
    b = "Percentage of users"
    wr.writerow([a, b])
    for row in Minimun_Time_table.find_elements_by_css_selector('tr'):
        wr.writerow([d.text for d in row.find_elements_by_css_selector('td')])
print("Minimum Time data is saved. ")
time.sleep(3)

print("Finished. ")
time.sleep(3)
driver.quit()
