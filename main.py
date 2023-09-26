from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import openpyxl

url = 'https://a810-dobnow.nyc.gov/publish/Index.html#!/'
# file_name = input("Please enter the filename: (Copy of Active CS licensees 9-11-23.xlsx) ")
# sheet_name = input("Please enter the sheet name of the excel: ")
# lic_col = int(input("What column number is the licensee number? "))
# starting_col = int(input("What column do you want to insert data, starting with Job Count? (22) "))
# starting_row = int(input("What row do you want to start with? "))


file_name = "Copy of Active CS licensees 9-11-23.xlsx"
sheet_name = "Sheet1"
lic_col = 1
starting_row = 2

driver = webdriver.Chrome()
driver.get(url)
driver.delete_all_cookies()

print(driver.title)

# Click the Search By License button
lic_btn = driver.find_element(By.XPATH, value="//*[@id='content']/div[3]/div[1]/div[4]/div[2]/button")
lic_btn.click()
time.sleep(2)

# Click the Licensee Type drop down menu
lic_type_btn = driver.find_element(By.ID, value="LicSearchType")
lic_type_btn.click()
time.sleep(2)

# Select the type of license
select = Select(driver.find_element(By.ID, value="LicSearchType"))
select.select_by_value('5')

# Type Licensee Number into the search by license field
lic_search_box = driver.find_element(By.ID, value='LicLicenseNumbere')

# Opens Excel file and go to specified worksheet
wb = openpyxl.load_workbook(file_name)
ws = wb[sheet_name]


def close_btn():
    try:
        close_btn = driver.find_element(By.XPATH, "//*[contains(@id, 'ngdialog')]/div[2]/div[3]/div/button")
        close_btn.click()
    except:
        try:
            close_btn = driver.find_element(By.XPATH, "//*[contains(@id, 'ngdialog')]/div[2]/div[1]/div[3]/button")
            close_btn.click()
        except:
            driver.quit()
            driver.get(url)
            driver.delete_all_cookies()
            lic_btn.click()
            time.sleep(2)
            lic_type_btn.click()
            time.sleep(2)
            select.select_by_value('5')


for x in range(starting_row, ws.max_row):
    starting_col = 22
    job_count = 0
    lic_num = ws.cell(starting_row, lic_col).value

    if ws.cell(starting_row, starting_col).value is None:

        # Fills License Search box with licensee number and searches
        lic_search_box.send_keys(lic_num)
        lic_search_box.send_keys(Keys.RETURN)
        time.sleep(2)

        # If there is no record, Total Job Count = 0
        # Example: License = 026836
        if "No records found for the given license number." in driver.page_source:
            print(str(starting_row) + " : No records")
            ws.cell(starting_row, starting_col).value = 0

        # Get the Total Job Count and Job Numbers and save it to Excel
        # Example: License = 026868
        elif "Associated Jobs with Active Permits" in driver.page_source:
            for i in range(1, 10):
                try:
                    job = driver.find_element(By.XPATH,
                                              "//*[contains(@id, 'ngdialog')]/div[2]/div[2]/div[2]/table/tbody/tr[" + str(
                                                  i) + "]/td[1]").text
                    starting_col += 1
                    ws.cell(starting_row, starting_col).value = job
                    i += 1
                    job_count += 1
                except:
                    print(str(starting_row) + " : " + str(job_count) + " records")
                    ws.cell(starting_row, column=22).value = job_count
                    break

        # Else Highlight Red

        wb.save(file_name)

        lic_search_box.clear()

        close_btn()

    starting_row += 1

time.sleep(15)
