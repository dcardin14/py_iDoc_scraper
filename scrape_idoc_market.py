from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import lxml
from openpyxl import Workbook

login_url = 'https://www.idocmarket.com/Security/Register'
#url = 'https://www.idocmarket.com/RIOCO1/Document/Search' #Rio Blanco County, Colorado
url = 'https://www.idocmarket.com/CONWY1/Document/Search' #Converse County, Wyoming
driver = webdriver.Chrome()

# CREATE SHEET
workbook = Workbook()
sheet = workbook.active

# ADD HEADERS
header_values = ['Rec#', 'Bk#', 'RecDate', 'DocDate', 'DocType', 'Grantor', 'Grantee', 'Notes']
sheet.append(header_values)

# WEB TO EXCEL COLUMN MAPPING
mapping = {'Recorded': 2, 'Doc Date': 3, 'Grantors': 5, 'Grantees': 6,'Notes': 7}

'''
# LOGIN
driver.get(login_url)
#driver.find_element(By.ID, 'Login_Username').send_keys('dcardin14@gmail.com')
#driver.find_element(By.ID, 'Login_Password').send_keys('sumtoor99A+')
driver.find_element(By.ID, 'Login_Username').send_keys('jray@caminoresources.com')
driver.find_element(By.ID, 'Login_Password').send_keys('WRC2022New')

time.sleep(1)
driver.find_element(By.ID, 'loginForm').submit()
'''
# SEARCH
driver.get(url)
driver.find_element(By.ID, 'StartRecordDate').send_keys('4/18/2017')
driver.find_element(By.ID, 'EndRecordDate').send_keys('7/19/2023')
driver.find_element(By.ID, 'Section').send_keys('04')
driver.find_element(By.ID, 'Township').send_keys('T38N')
driver.find_element(By.ID, 'Range').send_keys('R74W')
time.sleep(1)
driver.find_element(By.ID, 'SearchForm').submit()


# ITERATE RESULTS
preview_modal = driver.find_element(By.ID, 'preview-modal')
results = driver.find_elements(By.CLASS_NAME, 'result-item')
for result in results:
    try:
        result.find_elements(By.CLASS_NAME, 'result-actions')[0].find_elements(By.TAG_NAME, 'a')[1].click()
        time.sleep(1)

        # CREATE EMPTY RECORD
        record = [None] * 8
    
        # ADD DOCUMENT TYPE
        docType = result.find_elements(By.CLASS_NAME, 'doc-title')[0].text
        record[4] = docType

      # ITERATE COLUMNS AND ADD RELEVANT DATA TO RECORD
        columns = preview_modal.find_elements(By.CLASS_NAME, 'row-fluid')
        for column in columns:
            header = column.find_elements(By.CLASS_NAME, 'span2')[0].find_element(By.TAG_NAME, 'strong').text.split(':')[0]
            data = column.find_elements(By.CLASS_NAME, 'span10')[0].text

            column_no = mapping.get(header, None)
            if (column_no):
                record[column_no] = data
    
        # ADD RECORD TO SHEET
        sheet.append(record)

        # CLOSE MODAL
        preview_modal.find_elements(By.CLASS_NAME, 'modal-footer')[0].find_elements(By.TAG_NAME, 'a')[1].click()
        time.sleep(1)
    except selenium.common.exceptions.ElementNotInteractableException:
        # Skip the current iteration and continue with the next result
        continue
workbook.save('wizdom_index.xlsx')
    