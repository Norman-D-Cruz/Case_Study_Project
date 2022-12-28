import re
from datetime import datetime
from random import randint
from time import sleep
import time
import pandas as pd
import undetected_chromedriver.v2 as uc
import xlwings as xw
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import parameters as p  # Path: parameters.py
from parameters import capital_list, columns_list, country_list

script_start_time = datetime.now()


def main():
    print("Starting...")
    first_row_no = input("Enter first row number: ")
    last_row = input("Enter last row: ")
    
    # Start the session and use xlwings
    session = login(open_driver())
    wb = xw.Book.caller()
    sht_1 = wb.sheets["Sheet1"]
    sht_2 = wb.sheets["Sheet2"]
    
    profile_links = sht_1.range(f"A{first_row_no}:{last_row}").value

    column_index = 5
    row_no = int(first_row_no)
    
    for link in profile_links:
        start = time.time()
        print(f"Scraping Row {row_no}")
        try:
            name, headine, location = headline_scrape(session, link)
        except:
            print("wrong link, check if it links to profile")
        sht_2.range(f"A{row_no}").value = name
        sht_2.range(f"B{row_no}").value = headine
        sht_2.range(f"C{row_no}").value = location
        sleep(randint(2,5))
        try: 
            if link is not None:
                profile = get_profile_elements(session)
                employment_history, current_company = get_employment_history(profile)
                sht_2.range(f"D{row_no}").value = isolate_about(profile)
                sht_2.range(f"E{row_no}").value = current_company
                for employment in employment_history.values():
                    sht_2.range(f"{columns_list[column_index]}{row_no}").value = employment[0]
                    column_index += 1
                column_index = 5
            else:
                pass
        except:
            sht_2[f"A{row_no}"].color = (255,0,0)
        row_no += 1
        end = time.time() 
        final_time = round(end - start, 2)
        print(f"Execution Time: {final_time}\n")
    
    print("--------------Scraping Complete---------------")
    print(f"No. of Scraped Profile: {row_no - 2}")
    session.quit()

def open_driver(option="Yes"):
    """
    Opens chrome browser, sets visibility options (default is visible) and returns driver
    """
    print("Opening Chrome Driver...")

    if option == "Yes":
        driver = uc.Chrome()
        return driver
    elif option == "No":
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        driver = uc.Chrome(options=chrome_options)
        return driver

def login(driver):
    """
    Logs into LinkedIn, returns logged in session
    """
    print("Logging into LinkedIn...")

    driver.get("https://www.linkedin.com/")
    sleep(randint(2,5))
    driver.find_element(By.CSS_SELECTOR,"input[name='session_key']").send_keys(p.linkedin_username)
    sleep(randint(2,5))
    driver.find_element(By.CSS_SELECTOR,"input[name='session_password']").send_keys(p.linkedin_password)
    sleep(randint(2,5))
    driver.find_element(By.CSS_SELECTOR,"button[type='submit']").click()

    if driver.current_url == "https://www.linkedin.com/feed/":
        print("Login Successful!")
        return driver

    # Will trigger if Two Factor Authentication is enabled
    else:
        two_step_code = input("Enter 2FA code: ")
        driver.find_element(By.XPATH,"//form[@id='two-step-challenge']//input[@id='input__phone_verification_pin']").send_keys(two_step_code)
        sleep(randint(2,5))
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"label[for='recognizedDevice']")))
        element.click()
        sleep(randint(1,3))
        element.find_element(By.XPATH,"//form[@id='two-step-challenge']//button[@id='two-step-submit-button']").click()
        return driver

def get_profile_elements(driver):
    """
    Returns list of profile elements
    """
    print("Getting Profile Elements...")

    
    elements = driver.find_elements(By.XPATH,".//span[@class='visually-hidden']")

    # Gets the profile structure
    profile_elements = []
    for element in elements:
        profile_elements.append(element.get_attribute('innerHTML'))
    
    return profile_elements

def get_employment_history(elements):
    """
    Returns employment history as a list of dictionaries
    """
    print("Isolating Employment History...")

    if '<!---->Experience<!---->' in elements and '<!---->Education<!---->' in elements:
        experience_index = elements.index('<!---->Experience<!---->')
        education_index = elements.index('<!---->Education<!---->')
        dirty_employment = elements[experience_index+1:education_index]
    elif '<!---->Experience<!---->' in elements and '<!---->Interests<!---->' in elements:
        experience_index = elements.index('<!---->Experience<!---->')
        interests_index = elements.index('<!---->Interests<!---->')
        dirty_employment = elements[experience_index+1:interests_index]

    clean_employment = []

    for i in dirty_employment: 
        clean_employment.append(re.sub('<[^<]+?>', '', i))
    
    for i in clean_employment:
        if "\n" in i:
            clean_employment.remove(i)
    
    # The loops will drop title with countries or capital cities in it.
    for i in clean_employment:
        for word in i.split(" "):
            for country in country_list:
                try:
                    if word.lower() == country.lower():
                        clean_employment.remove(i)
                except:
                    pass

    for i in clean_employment:
        for word in i.split(" "):
            for capital in capital_list:
                try:
                    if word.lower().replace(",", "") == capital.lower():
                        clean_employment.remove(i)
                except:
                    pass

    # It will get the job title with the word present
    filter_object = filter(lambda a: 'Present' in a, clean_employment)
    filter_object_list = list(filter_object)
    present_index = []

    for i in filter_object_list:
        present_index.append((clean_employment.index(i)))

    emp_hist_dict = {}
    count = 0
    prev_index = 0

    for i in present_index:
        emp_hist_dict[count] = clean_employment[prev_index:i+1]
        prev_index = i+1
        count += 1
    
    current_company = ""

    for no, emp in emp_hist_dict.items():
        if " · " in emp[1] and "yrs" not in emp[1] and "mos" not in emp[1]:
            current_company = emp[1].split(' · ')[0]
            emp_hist_dict[no] = emp[0:1] + emp[2:]
            break
        elif "yrs" in emp[1] or "mos" in emp[1]:
            current_company = emp[0]
            emp_hist_dict[no] = emp[2:]
            break
        else:
            current_company = emp[1]

    # Temporary solution for when a key is empty caused by for loop filtering countries and capitals
    d = dict([(k,v) for k,v in emp_hist_dict.items() if len(v)>0])

    return d, current_company
    # return emp_hist_dict, current_company
    
def isolate_about(elements):
    """
    Returns about section
    """
    print("Isolating About Section...")

    if '<!---->About<!---->' in elements:
        about_index = elements.index('<!---->About<!---->')
        activity_index = elements.index('<!---->Activity<!---->')
        dirty_about = elements[about_index+1:activity_index]

        clean_about = []
        for i in dirty_about: 
            clean_about.append(re.sub('<[^<]+?>', '', i))

        return clean_about
    else:
        return None

def headline_scrape(driver, profile_link):
    """
    Scrapes headline from profile link

    returns 
    """
    print("Scraping Headline...")

    if profile_link is None:
        return None, None, None
    else:
        driver.get(profile_link)
        sleep(randint(2,5))
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        name_raw = driver.find_element(By.CSS_SELECTOR,".text-heading-xlarge.inline.t-24.v-align-middle.break-words").text
        print(name_raw)
        headline_raw = driver.find_element(By.CSS_SELECTOR,".text-body-medium.break-words").text

        location_raw = driver.find_element(By.CSS_SELECTOR,".text-body-small.inline.t-black--light.break-words").text
        
        return name_raw, headline_raw, location_raw


if __name__ == "__main__":
    xw.Book("linkedinscrape.xlsm").set_mock_caller()
    main()
