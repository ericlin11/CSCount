from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl.styles.fills import PatternFill
from openpyxl.styles import Font, colors

url = 'https://a810-bisweb.nyc.gov/bisweb/JobsQueryByNumberServlet?passjobnumber=520465416&passdocnumber=&go10=+GO+&requestid=0'

driver = webdriver.Chrome()
driver.get(url)

time.sleep(5)
#
# driver = webdriver.Chrome()
# driver.get('https://a810-bisweb.nyc.gov/bisweb/bsqpm01.jsp')
# driver.find_element(By.XPATH, '/html/body/div/table/tbody/tr[4]/td/table/tbody/tr/td/div/table/tbody/tr[3]/td[3]/a').click()
# driver.find_element(By.ID, 'passjobnumber1').send_keys('520465416')
# driver.find_element(By.XPATH, '/html/body/div/table[2]/tbody/tr[17]/td/table/tbody/tr/td[3]/input').click()
# driver.delete_all_cookies()

time.sleep(5)
driver.get("https://a810-dobnow.nyc.gov/publish/Index.html#!/")
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

file_name = input("Please enter the filename: ")
sheet_name = input("Please enter the sheet name of the excel: ")
lic_col = int(input("What column number is the licensee number? "))
job_col = int(input("What column is the job number? "))
com_col = int(input("What column is the Comment? "))
starting_row = int(input("What row do you want to start with? "))

#Opens Excel file and go to specified worksheet
wb = openpyxl.load_workbook(file_name)
ws = wb[sheet_name]

def close_btn():
    try:
        close_btn = driver.find_element(By.XPATH, "//*[contains(@id, 'ngdialog')]/div[2]/div[3]/div/button")
        close_btn.click()
    except:
        close_btn = driver.find_element(By.XPATH, "//*[contains(@id, 'ngdialog')]/div[2]/div[1]/div[3]/button")
        close_btn.click()

def close_tab():
    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL + 'w')


for x in range(2, ws.max_row + 1):
    lic_num = ws.cell(starting_row, lic_col).value
    job_num = ws.cell(starting_row, job_col).value

    #Fills License Search box with licensee number and searches
    lic_search_box.send_keys(lic_num)
    lic_search_box.send_keys(Keys.RETURN)
    time.sleep(3)

    #Search BIS or DOBNOW Job, if it's under Licensee, Comment 'Good'
    if job_num in driver.page_source:
        print(str(starting_row) + ": Yes: " + lic_num + ": " + job_num)
        ws.cell(starting_row, com_col).value = "Good"

    #If Job Number is a BIS Job (All numbers) and not under Licensee, Search the BIS Intranet
    if job_num.isdigit() and len(job_num) == 9:

        #Open a new tab to BIS Intranet
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + 't')
        driver.get('https://a810-bisweb.nyc.gov/bisweb/JobsQueryByNumberServlet?passjob_number=' + job_num +
                   '&passdocnumber=&go10=+GO+&requestid=0')
        driver.delete_all_cookies()
        time.sleep(3)

        #Check if job is signed off, Comment 'Signed Off'
        if "SIGNED OFF" in driver.page_source:
            print(str(starting_row) + ": Signed Off: " + lic_num + ": " + job_num)
            ws.cell(starting_row, com_col).value = "Signed Off"

        #Check if job is released, Comment 'Released'
        elif "RELEASE CS-SSM-SSC" in driver.page_source:
            print(str(starting_row) + ": Released: " + lic_num + ": " + job_num)
            ws.cell(starting_row, com_col).value = "Released"

        #Go to All Permit, select the latest permit
        else:
            all_permit_btn = driver.find_element(By.XPATH, "/html/body/center/table[4]/tbody/tr[2]/td[5]/a")
            all_permit_btn.click()
            permit_btn = driver.find_element(By.XPATH, "/html/body/center/table[4]/tbody/tr[3]/td[1]/a")
            permit_btn.click()

            #Compare license number, if it is not the same, Comment 'Superseded'
            if job_num not in driver.page_source:
                print(str(starting_row) + ": Superseded: " + lic_num + ": " + job_num)
                ws.cell(starting_row, com_col).value = "Superseded"

            #If it's the same, highlight the Comment, to check later manaully
            if job_num not in driver.page_source:
                print(str(starting_row) + ": No: " + lic_num + ": " + job_num)
                cellFill = PatternFill(patternType='solid', fgColor=colors.Color(rgb='00FF0000'))
                ws.cell(starting_row, com_col).fill = cellFill

        #Close the BIS Intranet tab
        close_tab()

    #Search DOBNOW Job without the Filing# (I1, S1, ...), if it's under Licensee, Comment 'Another Filing# Counted'
    elif job_num[:9] in driver.page_source:
        print(str(starting_row) + ": Another job: " + lic_num + ": " + job_num)
        ws.cell(starting_row, com_col).value = "Another job counted"

    #Highlight the Comment cell to search manually
    else:
        print(str(starting_row) + ": No: " + lic_num + ": " + job_num)
        cellFill = PatternFill(patternType='solid', fgColor=colors.Color(rgb='00FF0000'))
        ws.cell(starting_row, com_col).fill = cellFill

    wb.save(file_name)

    close_btn()

    lic_search_box.clear()
    starting_row += 1
    time.sleep(1)

time.sleep(15)
