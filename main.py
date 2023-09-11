from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import openpyxl
# import pandas as pd


driver = webdriver.Chrome()

driver.get("https://a810-dobnow.nyc.gov/publish/Index.html#!/")
driver.delete_all_cookies()

print(driver.title)

#Click the Search By License button
licBtn = driver.find_element(By.XPATH, value="//*[@id='content']/div[3]/div[1]/div[4]/div[2]/button")
licBtn.click()
time.sleep(2)

#Click the Licensee Type drop down menu
licTypeBtn = driver.find_element(By.ID, value="LicSearchType")
licTypeBtn.click()
time.sleep(2)

#Select the type of license
select = Select(driver.find_element(By.ID, value="LicSearchType"))
select.select_by_value('5')

#Type Licensee Number into the search by license field
licSearchBox = driver.find_element(By.ID, value='LicLicenseNumbere')
# licSearchBox.send_keys('020056')
# licSearchBox.send_keys(Keys.RETURN)
# time.sleep(2)

#Get the results from search
# results = driver.find_element(By.ID, value='ngdialog1')
# print(results.text)

licTest = ['000001', '020056']
licList = []

wb = openpyxl.load_workbook('CS Associated DOBNOW Filings.xlsx')
ws = wb['Sheet2']
print(ws.max_row+1)

for x in range(2, ws.max_row+1):
    licList.append({
      "licensee": ws.cell(x, column=1).value
    })


for lic in licList:
    print(lic)
    licSearchBox.send_keys(lic["licensee"])
    licSearchBox.send_keys(Keys.RETURN)
    time.sleep(2)
    print(lic["licensee"])
    licList["lic"].append({
        "results": driver.find_element(By.XPATH, value="//*[contains(@id    , 'ngdialog')]")
    })
    # print(licList[lic])
    licSearchBox.clear()

print(licList[1])

# #for each licensee in excel list
# for item in list:
#     #Search the licensee number
#     licSearchBox.send_keys(item)
#     licSearchBox.send_keys(Keys.RETURN)
#     time.sleep(1)
#
#     #Get the results
#     results = driver.find_element(By.ID, value='ngdialog1')
#     print(results.text)
#
#     # If no record found, count = 0
#     # If there is record, count = total jobs, save each job into a dictionary -> 020056, job1: 'X000056121-I1' ...
#
#     #Close out of the results and clear the search box
#     closeBtn = driver.find_element(By.XPATH, value='//*[@id="ngdialog3"]/div[2]/div[4]')
#     closeBtn.click()

time.sleep(15)