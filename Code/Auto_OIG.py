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
    oig_url="https://exclusions.oig.hhs.gov/"
    for emp in data:

        emp_id= int(emp['Employee ID'])
        emp_name = emp['Employee Name']
        emp_name_split = emp_name.split(' ')
        fname, lname = emp_name_split[0], emp_name_split[-1]

        driver.get(oig_url)

        driver.find_element_by_xpath('//input[contains(@name,"Last")]').send_keys(fname)

        driver.find_element_by_xpath('//input[contains(@name,"First")]').send_keys(lname)

        driver.find_element_by_xpath('//input[@title="Search"]').click()

        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'mainContentInterior')))

        try:
            driver.execute_script('window.print();')
        except:
            pass

        # sleep(2)

        # clickPrintBtn(driver)

        sleep(4)
        saveAshandle('OIG Search Results - Google Chrome', OIG_save, f"{emp_id}_{emp_name}_OIG")
        
    driver.quit()

if __name__ == '__main__':
    main(data,OIG_save)
    print("All Done")