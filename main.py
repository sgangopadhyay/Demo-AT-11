# main testing file

# main.py

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By 
from webdriver_manager.firefox import GeckoDriverManager
from excel_functions import Suman_Excel_Functions
from locators import Test_Locators

excel_file = "C:\\Users\\sgang\\OneDrive\\Desktop\\DDTF\\test_data.xlsx"

sheet_number = "Sheet1"

s = Suman_Excel_Functions(excel_file, sheet_number)

url = "https://www.facebook.com"

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
driver.get(url)

rows = s.Row_Count()

for row in range(2, rows + 1):
    username = s.Read_Data(row, 6)
    password = s.Read_Data(row, 7)

    driver.find_element(by=By.NAME, value=Test_Locators().username_locator).send_keys(username)
    driver.find_element(by=By.NAME, value=Test_Locators().password_locator).send_keys(password)
    driver.find_element(by=By.NAME, value=Test_Locators().submitButton_locator).click()

    driver.implicitly_wait(10)

    if 'https://www.facebook.com/checkpoint/?next' in driver.current_url:
        print("SUCCESS : Login Success with Username {a}".format(a=username))
        s.Write_Data(row, 8, "TEST PASS")
        driver.back()
    elif("https://www.facebook.com" in driver.current_url):
        print("FAIL : Login Failure with Username {a}".format(a = username))
        s.Write_Data(row, 8, "TEST FAIL")
        driver.back()

driver.quit()

