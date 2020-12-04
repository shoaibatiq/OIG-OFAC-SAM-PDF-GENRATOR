from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import pywinauto
import json
from pywinauto.application import Application
import pywinauto.mouse as mouse
import pywinauto.keyboard as keyboard
from time import sleep
import xlrd 
from settings import *
import os
from functions import *
import shutil


data = readData(filename=data_file)

def main(data,temp_dwnld,SAM_save):
    try:
        try:
            shutil.rmtree(temp_dwnld)
        except:
            pass
        sleep(1)
        os.mkdir(temp_dwnld)
    except Exception as e:
        print("Make sure there aren't any files in temp folder.",e)

    sam_url="https://www.sam.gov/SAM/pages/public/searchRecords/advancedPIRSearch.jsf"
    driver=createDriver()

    for emp in data:
        emp_id= int(emp['Employee ID'])
        emp_name = emp['Employee Name']

        driver.get(sam_url)

        driver.find_element('xpath','//input[@title="Single Search"]').click()

        name_field = driver.find_element('xpath','//input[@title="Name"]')
        name_field.clear()
        name_field.send_keys(emp_name)

        #Set exclusion status Active
        el=Select(driver.find_element('xpath','//select[contains(@id,"ExclStatus")]'))
        el.select_by_index(1)

        driver.find_element('xpath','//input[contains(@id,"SearchButton")]').click()

        driver.find_element('xpath','//input[@value="Save PDF"]').click()

        MoveFile(temp_dwnld, SAM_save, f"{emp_id}_{emp_name}_SAM.pdf")
            
    driver.quit()

if __name__ == '__main__':
    main(data,temp_dwnld,SAM_save)
    print("All Done")