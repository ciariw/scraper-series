from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import calendar
import time
import os
preferences = {
                "profile.default_content_settings.popups": 0,
                "download.default_directory": 'G:\My Drive\Worklog_Files\\',
                "directory_upgrade": True
            }
optionz = webdriver.ChromeOptions()
#optionz.add_argument("window-size=1920,1080")
#optionz.add_argument("--headless")

optionz.add_experimental_option('prefs', preferences)

driver = webdriver.Chrome(
    "C:/Program Files/Google/Chrome/Application/chromedriver.exe",chrome_options = optionz)


def download():
    a = driver.find_element_by_xpath("//a/span/span/span[text() = 'Excel']").get_attribute("id")
    driver.find_element_by_id(f"button-{a[7:11]}-btnIconEl").click()
def searchButton():

    a = driver.find_element_by_xpath("//a/span/span/span[text() = 'Search']").get_attribute("id")
    driver.find_element_by_id(f"button-{a[7:11]}-btnIconEl").click()

def loginAccess():
    driver.get("XXXXXXXXXXXXXXXXXXXXXXXXXXXX")
    driver.find_element_by_id("j_username").send_keys("XXXXXXX")
    driver.find_element_by_id("j_password").send_keys("XXXXXXXXXXX")
    driver.find_element_by_css_selector('input.btnWelcomeLogin').click()
    driver.get("XXXXXXXXXXXXXXXXXXXXXXXXXXX")

def setup():

    driver.find_element_by_xpath("//div[7]/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div/"
                                 "div/div/div[2]/div/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div").click()
    driver.find_element_by_xpath("/html/body/div[10]/div/div/table[1]/tbody/tr/td[1]/input").click()
    driver.find_element_by_xpath("//td[2]/div/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div/input").click()

def openDate(date):
    '''YYYY-MM-DD
    datefield-1454-inputEl
    datefield-1454-inputEl'''

    dTe = driver.find_element_by_xpath("//div[1]/div/div/table/tbody/tr/td[2]/table/tbody//td[1]/input")
    dTe.click()
    dTe.clear()
    dTe.send_keys(date)


def recursiveDl(dates):
    setup()
    for i in dates:
        openDate(i)
        searchButton()
        download()
    time.sleep(2)
    driver.quit()
