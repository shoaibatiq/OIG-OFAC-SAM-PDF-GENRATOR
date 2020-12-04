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

data = readData(filename=data_file)

def main(data,OIG_save):
    driver=createDriver()
    ofac_url="https://sanctionssearch.ofac.treas.gov"
    for emp in data:
        emp_id= int(emp['Employee ID'])
        emp_name = emp['Employee Name']

        driver.get(ofac_url)

        el=Select(driver.find_element('xpath','//select[contains(@id,"MainContent_ddlType")]'))

        el.select_by_value('Individual')



        driver.find_element('xpath','//input[contains(@id,"MainContent_txtLastName")]').send_keys(emp_name)

        driver.find_element('xpath','//input[contains(@id,"btnSearch")]').click()


        try:
            driver.execute_script('window.print();')
        except:
            pass

        # sleep(2)

        # clickPrintBtn(driver)

        sleep(4)
        saveAshandle('Sanctions List Search - Google Chrome', OFAC_save, f"{emp_id}_{emp_name}_OFAC")
    driver.quit()
if __name__ == '__main__':
    main(data,OFAC_save)
    print("All Done")