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

def readData(filename):
    data=[] 
    loc = (filename) 

    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 

    sheet.cell_value(0, 0) 

    headers= sheet.row_values(0)
    for row in range(1, sheet.nrows):
        data.append({k:v for k,v in zip(headers, sheet.row_values(row))})
    return data

def expand_shadow_element(driver,element):
    shadow_root = driver.execute_script('return arguments[0].shadowRoot', element)
    return shadow_root

def clickPrintBtn(driver):
    try:
        driver.switch_to.window(driver.window_handles[1])

        A = driver.find_element_by_tag_name('print-preview-app')

        shadow_rootA = expand_shadow_element(driver,A)

        B = shadow_rootA.find_element_by_id('sidebar')

        shadow_rootB = expand_shadow_element(driver,B)

        C = shadow_rootB.find_element_by_tag_name('print-preview-button-strip')

        shadow_rootC = expand_shadow_element(driver,C)

        printBtn = WebDriverWait(shadow_rootC, 100).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[class='action-button'][aria-disabled='false']")))
        
        printBtn.click()

        driver.switch_to.window(driver.window_handles[0])
    except:
        sleep(3)
        clickPrintBtn(driver)
        
def saveAshandle(win_title, save_loc, pdf_name, saveAs_titile = 'Save Print Output As'):
    handle=pywinauto.findwindows.find_window(best_match=win_title)
    app=Application().connect(handle=handle)

    dlg=app[saveAs_titile]
    dlg.Edit.set_edit_text(save_loc+pdf_name)
    for _ in range(5):
        try:
            dlg.Save.click()
            break
        except:
            dlg=app[saveAs_titile]
            sleep(2)
    try:
        dlg.Edit.type_keys("{ENTER}")
    except:
        pass
    
def GetFileName(dwnld_path):
    ls=os.listdir(dwnld_path)
    ls=[os.path.join(dwnld_path,i) for i in ls]
    files=[i for i in ls if not(os.path.isdir(i))]
    return files[0] if files != [] else ''

def MoveFile(dwnld_path, movePath,file_name):
    while GetFileName(dwnld_path) == '':
        pass
    sleep(1)
    dwnlded_file=GetFileName(dwnld_path)
    os.rename(dwnlded_file,movePath+file_name)

def createDriver():
    settings = {
        "appState": {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local"
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }  
    }
    prefs = {'printing.print_preview_sticky_settings': json.dumps(settings),
             "download.default_directory": temp_dwnld }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')

    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.implicitly_wait(10)
    
    return driver

    