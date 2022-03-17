import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from selenium.webdriver.chrome.options import Options
import profiles
import json

def new_session(driver=None, ftr=None):
    # Holy grail of hacks. New Tab, call the function, close tab. Each tab retains all cookies, so there is
    # no need to log in again

    if ftr == "Job Listings":
        job_listings(driver)
    elif ftr == "Connection Info":
        connection_info(driver)
    elif ftr == "Company Employee list":
        company_employee_list(driver)

    driver.close()

def job_listings(driver=None):
    print("ASADSDADAS")

    sleep(5)
    return


def connection_info(driver=None):
    name_set = []
    temporary_search_depth = 5
    driver.get("https://www.linkedin.com/search/results/people/?network=%5B%22F%22%5D&sid=~x")
    # This is all the information that is available on the surface.
    # I can either use the URLs to grab more info on them, or I can just save it on the surface
    data_for_json = {}
    for xyz in range(0,temporary_search_depth):
        li = driver.find_elements_by_xpath("//div/span[1]/span/a/span/span[1]")
        lnk = driver.find_elements_by_xpath("//div[1]/div/span[1]/span/a")
        headers = driver.find_elements_by_xpath("//ul/li/div/div/div[2]/div[1]/div[2]/div/div[1]")
        locations = driver.find_elements_by_xpath("//div/div/div[1]/ul/li/div/div/div[2]/div[1]/div[2]/div/div[2]")

        for i, ids, hd, loc in zip(li, lnk, headers, locations):
            name_set.append(profiles.Person(i.text))
            name_set[-1].id = ids.get_attribute('href')
            name_set[-1].header = hd.text
            driver.execute_script(f"window.open('{ids.get_attribute('href')}', '_blank');")
            driver.switch_to.window(driver.window_handles[1])
            try:
                job = driver.find_element_by_xpath("//div[2]/ul/li[1]/a/h2/div").text
            except:
                print("problem")
                job = "N/A"
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            data_for_json[i.text] = [{"Description": hd.text, "Link": ids.get_attribute('href'), "Location": loc.text,
                                     "Current Job": job}]
        # This is kind of hacky but the next linkedin button wont load unless the button is within view
        driver.execute_script("window.scrollTo(0, 915)")
        while True:
            try:
                driver.find_element_by_xpath("//span[text()='Next']/parent::button").click()
            except:
                try:
                    driver.execute_script("window.scrollTo(0, 915)")
                except:
                    pass
            else:
                sleep(0.5)
                driver.execute_script("window.scrollTo(0, 915)")
                break

    with open('connection data.json', 'w', encoding='utf-8') as f:
        json.dump(data_for_json, f, ensure_ascii=False, indent=3)

    '''for i in name_set:
        print(f"Name:{i.name}\n desc: {i.header}\n URL: {i.id}\n List of Jobs: {i.jobs}\n--------------------------")'''
    return


def company_employee_list(driver=None):
    print("companylist")
    sleep(5)
    return
