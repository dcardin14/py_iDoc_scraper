from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import lxml
from openpyxl import Workbook

login_url = 'https://www.idocmarket.com/Security/Register'
url = 'https://www.idocmarket.com/RIOCO1/Document/Search'
driver = webdriver.Chrome()

# CREATE SHEET
workbook = Workbook()
sheet = workbook.active

# ADD HEADERS
header_values = ['Rec#', 'Bk#', 'Pg#', 'RecDate', 'DocDate', 'DocType', 'Grantor', 'Grantee', 'Notes']
sheet.append(header_values)

# WEB TO EXCEL COLUMN MAPPING
#mapping = {'Rec': 0, 'Bk': 1, 'Pg': 2, 'Recorded': 3, 'Doc Date': 4, 'Grantors': 6, 'Grantees': 7,'Notes': 8}
mapping = {'Recorded': 3, 'Doc Date': 4, 'Grantors': 6, 'Grantees': 7,'Notes': 8}


# LOGIN
driver.get(login_url)
driver.find_element(By.ID, 'Login_Username').send_keys('dcardin14@gmail.com')
driver.find_element(By.ID, 'Login_Password').send_keys('sumtoor99A+')
time.sleep(1)
driver.find_element(By.ID, 'loginForm').submit()

# SEARCH
driver.get(url)
driver.find_element(By.ID, 'Section').send_keys('19')
driver.find_element(By.ID, 'Township').send_keys('2N')
driver.find_element(By.ID, 'Range').send_keys('97W')
time.sleep(1)
driver.find_element(By.ID, 'SearchForm').submit()


# ITERATE RESULTS
preview_modal = driver.find_element(By.ID, 'preview-modal')
results = driver.find_elements(By.CLASS_NAME, 'result-item')
for result in results:
    result.find_elements(By.CLASS_NAME, 'result-actions')[0].find_elements(By.TAG_NAME, 'a')[1].click()    
    time.sleep(1)

    # CREATE EMPTY RECORD
    record = [None] * 9
    
    # GET RECORDING INFO
    recordingInfo= result.find_elements(By.TAG_NAME, 'span')[0].text.split(' ')
    
    # ADD RECEPTION NUMBER AND DOCUMENT TYPE
    docType = result.find_elements(By.CLASS_NAME, 'doc-title')[0].text
    record[0] = recordingInfo[0][1:]
    record[5] = docType

    # ADD BOOK/PAGE NUMBERS
    if (len(recordingInfo) > 1):
        record[1] = recordingInfo[2]
        record[2] = recordingInfo[4]

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

workbook.save('./wizdom_index.xlsx')
    