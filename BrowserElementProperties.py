import time
import openpyxl
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC

def WebElement(WebDriver, EC_Conditions, Method, Element_Locator, Time):
    Element = WebDriverWait(WebDriver, Time).until(EC_Conditions((Method, Element_Locator)))
    return Element