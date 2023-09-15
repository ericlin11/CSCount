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

licTest = ['020056', '020150']
# licList = []
licList = {}

wb = openpyxl.load_workbook('CS Associated DOBNOW Filings.xlsx')
ws = wb['Sheet2']

for x in range(2, ws.max_row+1):
    licList[ws.cell(x,column=1).value] = []
    # licList.append({
    #     "licensee": ws.cell(x, column=1).value,
    #     "job1": ws.cell(x, column=2).value,
    #     "job2": ws.cell(x, column=3).value,
    #     "job3": ws.cell(x, column=4).value,
    #     "job4": ws.cell(x, column=5).value,
    #     "job5": ws.cell(x, column=6).value,
    #     "job6": ws.cell(x, column=7).value,
    #     "job7": ws.cell(x, column=8).value,
    #     "job8": ws.cell(x, column=9).value,
    #     "job9": ws.cell(x, column=10).value,
    #     "job10": ws.cell(x, column=11).value,
    #     "count": ws.cell(x, column=12).value
    # })

for i in licList:
    print(i)

#For each licensee in list
startingrow=2
for lic in licList:

    #Search the licensee number
    licSearchBox.send_keys(lic)
    licSearchBox.send_keys(Keys.RETURN)
    time.sleep(2)

    try:
        table = driver.find_element(By.CLASS_NAME, value='table')
    except:

    #Get the popup box results and filter out each job
    results = driver.find_element(By.CLASS_NAME, value='table').text
    results.replace('\\n', '|')
    licList[lic] += [results]
    ws.cell(startingrow, column=1).value = str(lic)
    ws.cell(startingrow, column=2).value = str(licList[lic])
    startingrow += 1
    print(licList[lic])

    wb.save('(new)CS Associated DOBNOW Filings.xlsx')
    print("saved")
    # try:
    #     licList[lic] += [results.text]
    #     ws.cell(startingrow, column=2).value = licList[lic]
    #     print(licList[lic])
    # except:
    #     print("error")
    # finally:
    #     driver.quit()
        # for row in results.find_element(By.XPATH, value='//*[contains(@id, "ngdialog")]/div[2]/div[2]/div[2]/table/tbody/tr[1]/td[1]'):


    # // *[ @ id = "ngdialog3"] / div[2] / div[2] / div[2] / table / tbody / tr[1] / td[1]

    # lic["job1"] = driver.find_element(By.XPATH, value="//*[contains(@id    , 'ngdialog')]").text
    # lic["job1"]= results
    # licList[lic] += [results]
    # print(lic + " : " + str(licList[lic]))


    licSearchBox.clear()




# #for each licensee in Excel list
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