from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import lxml
from openpyxl import Workbook
from bs4 import BeautifulSoup

login_url = 'https://idocmarket.com/Security/Register' 
url = 'https://idocmarket.com/WILND1/Document/Search' #Williams County, ND
driver = webdriver.Chrome()

# CREATE SHEET
workbook = Workbook()
sheet = workbook.active

# ADD HEADERS
header_values = ['Rec#', 'Bk#', 'RecDate', 'DocDate', 'DocType', 'Grantor', 'Grantee', 'Notes']
sheet.append(header_values)

# WEB TO EXCEL COLUMN MAPPING
mapping = {'Recorded': 2, 'Doc Date': 3, 'Grantors': 5, 'Grantees': 6,'Notes': 7}


# LOGIN
driver.get(login_url)
driver.find_element(By.ID, 'Login_Username').send_keys('info16@lonetreeenergy.com')
driver.find_element(By.ID, 'Login_Password').send_keys('LoneTree16$!')
#driver.find_element(By.ID, 'Login_Username').send_keys('jray@caminoresources.com')
#driver.find_element(By.ID, 'Login_Password').send_keys('WRC2022New')

time.sleep(1)
driver.find_element(By.ID, 'loginForm').submit()

# SEARCH
driver.get(url)
driver.find_element(By.ID, 'StartRecordDate').send_keys('3/6/2019')
time.sleep(2)
driver.find_element(By.ID, 'EndRecordDate').send_keys('9/8/2023')
time.sleep(2)
driver.find_element(By.ID, 'Section').send_keys('25')
time.sleep(2)
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
driver.find_element(By.ID, 'Township').send_keys('1')
time.sleep(2)
driver.find_element(By.ID, 'Range').send_keys('1')
driver.find_element(By.ID, 'Range').send_keys('1')
driver.find_element(By.ID, 'Range').send_keys('1')
driver.find_element(By.ID, 'Range').send_keys('1')

time.sleep(2)
driver.find_element(By.ID, 'vr_input').send_keys('SE4')
time.sleep(2)
driver.find_element(By.ID, 'SearchForm').submit()
#driver.find_element(By.ID, 'Search').click()
time.sleep(2)

# ITERATE RESULTS
preview_modal = driver.find_element(By.ID, 'preview-modal')
results = driver.find_elements(By.CLASS_NAME, 'result-item')
for result in results:

    result.find_elements(By.CLASS_NAME, 'result-actions')[0].find_elements(By.TAG_NAME, 'a')[0].click()
    time.sleep(3)

     # CREATE EMPTY RECORD
    record = [None] * 8
    
     # ADD DOCUMENT TYPE
     #docType = result.find_elements(By.CLASS_NAME, 'doc-title')[0].text
     #datapage = driver.find_element(By.ID, 'master-container')
    docType = driver.find_element(By.XPATH, '//*[@id="master-container"]/div[2]/h2').text #OK
    print(docType)
    record[4] = docType
    recNo = driver.find_element(By.XPATH, '//*[@id="indexing-child1"]/h4').text #OK
    print(recNo)
    time.sleep(2)
    RecDate = driver.find_element(By.CSS_SELECTOR, '#indexing-child1 > div:nth-child(4)').text #OK
    print(RecDate)
    time.sleep(2)
    DocDate = driver.find_element(By.XPATH, '//*[@id="indexing-child1"]/div[1]').text #OK
    print(DocDate)
    time.sleep(2)
    firstGrantor = driver.find_element(By.CSS_SELECTOR, '#parties > div:nth-child(3)').text #I thought I needed to select by selector for some reason
    print(firstGrantor)
    time.sleep(2)
    #firstGrantee = driver.find_element(By.CSS_SELECTOR, '#parties > div:nth-child(8)').text #Fails every time
    #print(firstGrantee)
    time.sleep(2)
    Notes = driver.find_element(By.XPATH,'//*[@id="indexing-child2"]/div[3]').text
    print(Notes)
    time.sleep(2)

      # ITERATE COLUMNS AND ADD RELEVANT DATA TO RECORD
       # columns = preview_modal.find_elements(By.ID, 'indexing_child1')
        #for column in columns:
            #header = column.find_elements(By.CLASS_NAME, 'span2')[0].find_element(By.TAG_NAME, 'strong').text.split(':')[0]
            #data = column.find_elements(By.CLASS_NAME, 'span10')[0].text
           # print("Looping iteration.")
            #column_no = mapping.get(header, None)
            #if (column_no):
    sheet.append(record)

    # CLOSE MODAL
    driver.execute_script("backToResults();")
    time.sleep(1)
        
    workbook.save('wizdom_index.xlsx')
    