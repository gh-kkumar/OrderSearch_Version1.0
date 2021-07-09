import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import os
import pyautogui
#import win32com.client as comclt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path
from selenium.webdriver.common import keys



FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet, 3, 3))

driver = webdriver.Chrome(executable_path = str(Path().resolve()) + r'\Browser\chromedriver_win32\chromedriver')
driver.maximize_window()
driver.get(Url)

#1. This is for HCP Login Page

FilePath = str(Path().resolve()) + r'\Excel Files\OrderSearch.xlsx'
Sheet = 'Login Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)
Seconds = 300 / 1000

for RowIndex in range(2, RowCount + 1):

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log in"]', 60)
    Element.click()

    if(RowIndex == 2):
        time.sleep(Seconds)
    elif(RowIndex > 2):
        time.sleep(7)

    print(driver.title)
    if (driver.title == 'Login'):
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/div/div/span/div/div', 60)
        if(RWDE.ReadData(FilePath, Sheet, RowIndex, 5) == Element.text):
            RWDE.WriteData(FilePath, Sheet, RowIndex, 6, Element.text)
            RWDE.WriteData(FilePath, Sheet, RowIndex, 7, 'Passed')
        else:
            RWDE.WriteData(FilePath, Sheet, RowIndex, 6, Element.text)
            RWDE.WriteData(FilePath, Sheet, RowIndex, 7, 'Failed')

        driver.execute_script('arguments[0].innerHTML = ""', Element)

        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Username"]', 60)
        Element.clear()

        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Password"]', 60)
        Element.clear()
    elif (driver.title == 'Guardant Health'):
        if (RWDE.ReadData(FilePath, Sheet, RowIndex, 5) == driver.title):
            RWDE.WriteData(FilePath, Sheet, RowIndex, 6, driver.title)
            RWDE.WriteData(FilePath, Sheet, RowIndex, 7, 'Passed')

        # 2. This is for Patient Search Flow
        Sheet = 'Order Search Page Data'
        RowCount = RWDE.RowCount(FilePath, Sheet)

        # Results Menu
        time.sleep(1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Results"]', 60)
        Element.click()

        time.sleep(2)
        for RowIndex1 in range(3, RowCount + 1):
            #OrderDateFrom
            time.sleep(1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[1]/div/lightning-input//input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 2)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 2)))
                Element.click()
            else:
                Element.clear()
                driver.execute_script('arguments[0].focus();', Element)

            #OrderDateTo
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div/lightning-input//input', 60)
            Element.clear()
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 3)) != 'None'):
                Element.send_keys(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 3)))
                Element.click()
            else:
                Element.clear()
                driver.execute_script('arguments[0].focus();', Element)

            # Status
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div/lightning-combobox//input[@placeholder = "Select an Option"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            # if(str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 4)) != 'None'):
            #    driver.execute_script('arguments[0].value = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 4)) + '";', Element)
            #    driver.execute_script('arguments[0].focus();', Element)
            #    #driver.execute_script('arguments[0].click()', Element)
            # else:
            #    driver.execute_script('arguments[0].value = ""', Element)
            #    driver.execute_script('arguments[0].focus();', Element)

            # Status Element
            time.sleep(Seconds)
            SpanElement = ''
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 5)) != 'None'):
                # SpanElement = '//span/span[. = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 4)) + '"]'
                SpanElement = '//lightning-base-combobox-item[' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 5)) + ']/span/span'
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, SpanElement, 60)
                driver.execute_script('return arguments[0].scrollIntoView(true);', Element)
                driver.execute_script('arguments[0].style.backgroundColor = "#FAEDEA";', Element)
                #time.sleep(1)
                Element.click()

                # Result
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[4]/div/lightning-combobox//input[@placeholder = "Select an Option"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                # Result Element
                time.sleep(Seconds)
                SpanElement = ''
                if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 7)) != 'None'):
                    SpanElement = '//div[4]/div/lightning-combobox/div/lightning-base-combobox/div/div[2]/lightning-base-combobox-item[' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 7)) + ']/span[2]'
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, SpanElement, 60)
                    #driver.execute_script('return arguments[0].scrollIntoView(true);', Element)
                    driver.execute_script('arguments[0].style.backgroundColor = "#FAEDEA";', Element)
                    #time.sleep(1)
                    Element.click()

            # Provider
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[5]//input', 60)
            driver.execute_script('arguments[0].click();', Element)

            # Provider Element
            time.sleep(1)
            SpanElement = ''
            if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 8)) != 'None'):
                SpanElement = '//span/span[. = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 8)) + '"]'#'//lightning-base-combobox-item[' + str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 9)) + ']/span[2]'
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, SpanElement, 60)
                #driver.execute_script('return arguments[0].scrollIntoView(true);', Element)
                driver.execute_script('arguments[0].style.backgroundColor = "#FAEDEA";', Element)
                #time.sleep(1)
                #driver.execute_script('arguments[0].click();', Element)
                Element.click()

            # Search Button
            time.sleep(1)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Search"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            time.sleep(1)
            # Table
            Element = '//div[2]/div/div/table'
            TableRows = Element + '/tbody/tr'
            row_count = len(driver.find_elements_by_xpath(TableRows))
            CellVal = ['', '', '', '', '', '']
            if (row_count > 0):
                for RIndex in range(1, row_count + 1):
                    TableColumns = driver.find_elements_by_xpath(Element + '/tbody/tr[' + str(RIndex) + ']/td')
                    TableCellValue1 = driver.find_element(By.XPATH, '//tbody/tr[' + str(RIndex) + ']/th')
                    CellVal[0] = TableCellValue1.text
                    for CIndex in range(1, len(TableColumns)):
                        TableCellValue = driver.find_element_by_xpath(Element + '/tbody/tr[' + str(RIndex) + ']/td[' + str(CIndex) + ']')
                        CellVal[CIndex] = TableCellValue.text
                        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex1, 15)) == 'None'):
                            CellValue = ''
                        else:
                            CellValue = RWDE.ReadData(FilePath, Sheet, RowIndex1, 15)

                    if (CellVal[0] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 10) and
                        CellVal[1] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 11) and
                        CellVal[2] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 12) and
                        CellVal[3] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 13) and
                        CellVal[4] == RWDE.ReadData(FilePath, Sheet, RowIndex1, 14) and
                        CellVal[5] == CellValue):
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 16, CellVal[0])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 17, CellVal[1])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 18, CellVal[2])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 19, CellVal[3])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 20, CellVal[4])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 21, CellVal[5])
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 22, 'Passed')
                    else:
                        RWDE.WriteData(FilePath, Sheet, RowIndex1, 22, 'Failed')
                    break
            else:
                RWDE.WriteData(FilePath, Sheet, RowIndex1, 22, 'Passed')


























