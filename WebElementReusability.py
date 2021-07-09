import time
import openpyxl
import xlsxwriter
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException


def PageElement(WebDriver, ElementType, ElementSelector, ElementString, ElementValue, Time):
    #print(ElementType + '--' + ElementSelector + '--' + ElementString + '=' + ElementValue)
    time.sleep(Time)
    if (ElementSelector == 'CSS Selector'):
        if(ElementValue != 'None'):
            Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.CSS_SELECTOR, ElementString, 60)
            WebDriver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if (ElementType == 'TextBox' or ElementType == 'TextArea'):
            if(ElementValue != 'None'):
                Element.send_keys(ElementValue)
                #Element.click()
            elif(ElementValue == 'None'):
                Element.click()
        elif (ElementType == 'Button'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'Image'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'Link'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'LookUp'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'Span'):
            if (ElementValue == 'Click'):
                Element.click()
                #WebDriver.execute_script("arguments[0].click();", Element)
        elif (ElementType == 'Tab'):
            if (ElementValue == 'Click'):
                #Element.click()
                WebDriver.execute_script("arguments[0].click();",Element)
        elif (ElementType == 'Div'):
            if (ElementValue == 'Click'):
                #Element.click()
                WebDriver.execute_script("arguments[0].click();",Element)
        elif (ElementType == 'CheckBox'):
            if (ElementValue == 'Yes'):
                Element.click()
        elif (ElementType == 'DropDownList'):
             ElementSelect = Select(Element)
             Element.click()
             ElementSelect.select_by_visible_text(ElementValue)

    if (ElementSelector == 'XPATH'):
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, ElementString, 60)
        WebDriver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if (ElementType == 'TextBox' or ElementType == 'TextArea'):
            if (ElementValue != 'None'):
                Element.send_keys(ElementValue)
                #Element.click()
            elif (ElementValue == 'None'):
                Element.click()
        elif (ElementType == 'Button'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'Image'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'Link'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'LookUp'):
            if (ElementValue == 'Click'):
                Element.click()
        elif (ElementType == 'Span'):
            if (ElementValue == 'Click'):
                #Element.click()
                WebDriver.execute_script("arguments[0].click();", Element)
        elif (ElementType == 'Tab'):
            if (ElementValue == 'Click'):
                # Element.click()
                WebDriver.execute_script("arguments[0].click();", Element)
        elif (ElementType == 'Div'):
            if (ElementValue == 'Click'):
                #Element.click()
                WebDriver.execute_script("arguments[0].click();",Element)
        elif (ElementType == 'CheckBox'):
            if (ElementValue == 'Yes'):
                Element.click()
        elif (ElementType == 'DropDownList'):
             ElementSelect = Select(Element)
             Element.click()
             ElementSelect.select_by_visible_text(ElementValue)


def PageElementErrorMessage(WebDriver, ElementType, ElementSelector, ElementString, ElementValue, Time):
    print(ElementType + '--' + ElementSelector + '--' + ElementString + '=' + ElementValue)
    time.sleep(Time)
    ErrorMessage = ''
    if (ElementSelector == 'CSS Selector'):
        if (ElementType == 'String'):
            if (ElementValue == 'None'):
                Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.CSS_SELECTOR, ElementString, 60)
                ErrorMessage = Element.text
            else:
                ErrorMessage = ' '

    if (ElementSelector == 'XPATH'):
        if (ElementType == 'String'):
            if (ElementValue == 'None'):
                Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, ElementString, 60)
                ErrorMessage = Element.text
            else:
                ErrorMessage = ' '

    return ErrorMessage

def check_exists_by_xpath(WebDriver,xpath):
    try:
        WebDriver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True