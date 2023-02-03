import os
# Import Excel Library
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

#Import Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

#Import zipfile
from zipfile import ZipFile

#Import Json
import json
import time

# Open Excel File
wb = load_workbook('Amazon Order.xlsx')
wb2 = Workbook()
ws = wb.active
ws2 = wb2.active
ws2.title = "Data"

# Declare variables
col_data = []
col_order = []
new_file_col = []
col_url = ""
counter = 0

# declare a drive
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
    col_order.append(ws["B"+ str(row)].value)
    col_data.append(ws["AI" + str(row)].value)
    

# File Loop
for i in range (0, 1):
    driver.get(col_data[i])
    time.sleep(5)
    for file in os.listdir(os.path.dirname(__file__)):
        print(file)
        if file.endswith(".zip"):
            
            with ZipFile(file, 'r') as zip:
                zip.extractall()
            
    for fileJson in os.listdir(os.path.dirname(__file__)):
        if fileJson.endswith(".json"):
            
            name = os.path.basename(fileJson)
            jsonName = os.path.splitext(name)[0]
            print(jsonName)
            f = open(jsonName+".json")
            data = json.load(f)
            for i in data:
                print(i)

