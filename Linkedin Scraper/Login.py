import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.chrome.options import Options
def initiate(a=None):
    driver = webdriver.Chrome("C:\Program Files\Google\Chrome\Application\chromedriver")
    driver.get(a)
    return driver

def login(driver=None, id=None, pw=None):
    driver.find_element_by_id("session_key").send_keys(id)
    driver.find_element_by_id("session_password").send_keys(pw)
    driver.find_element_by_id("session_password").send_keys(Keys.ENTER)

