import os
# Import Excel Library
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

#Import Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Open Excel File
wb = load_workbook('Amazon Order.xlsx')
wb2 = Workbook()
ws = wb.active
ws2 = wb2.active
ws2.title = "Data"

# Declare arrays
col_data = []
new_file_col = []
col_url = ""

# declare a drive

from selenium import webdriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
             "download.default_directory": 
                        os.path.dirname(__file__),
             "directory_upgrade": False}
options.add_experimental_option("prefs", prefs)
driver=webdriver.Chrome("./chromedriver/chromedriver.exe", options=options)


# Take Urls and append to col_data
for row in range(2, ws.max_row + 1):
    col_data.append(ws["AI" + str(row)].value)
    

# File Loop
for i in range (0, 2):
    driver.get(col_data[i])    