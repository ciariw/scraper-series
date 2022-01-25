from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import calendar
import time
import os

preferences = {
    "profile.default_content_settings.popups": 0,
    "download.default_directory": r'C:\Users\EmekaAriwodo\Desktop\mems\memstatus\\',
    "directory_upgrade": True
}
optionz = webdriver.ChromeOptions()

optionz.add_experimental_option('prefs', preferences)

driver = webdriver.Chrome(
    "C:/Program Files/Google/Chrome/Application/chromedriver.exe", chrome_options=optionz)
def download():
    names = ["Module Line 1", "Module Line 2", "Module Line 3"]
    driver.find_element_by_partial_link_text("Combined Status").click()
    for i in names:
        driver.find_element_by_name(":lineId").click()
        driver.find_element_by_xpath(f"//td[text() = '{i}']").click()
        a = driver.find_element_by_xpath("//a/span/span/span[text() = 'Search']").get_attribute("id")

        driver.find_element_by_id(f"button-{a[7:11]}-btnIconEl").click()
        a = driver.find_element_by_xpath("//a/span/span/span[text() = 'Excel']").get_attribute("id")
        driver.find_element_by_id(f"button-{a[7:11]}-btnIconEl").click()
def initialize():
    driver.get("http://10.60.110.42/emiplus/home")
    driver.find_element_by_id("j_username").send_keys("hqcell")
    driver.find_element_by_id("j_password").send_keys("1")
    driver.find_element_by_class_name("btnWelcomeLogin").click()
    driver.get("http://10.60.110.42/emiplus/home#EQP1020:viewReportId=EQP1020")
