import getpass
import os
import sys
import time
import yaml 
import re
from pprint import pprint

import pandas as pd

# from bs4 import BeautifulSoup as bs

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

def data_config(yaml_file : str):
    try:
        with open(yaml_file, 'r') as stream:
            config = yaml.load(stream, Loader=yaml.FullLoader)
    except FileNotFoundError:
        print("Make sure your config.yaml file exists")
        print("You can just copy the example_data.yaml and fill it in.")
        sys.exit(0)

    if 'username' not in config or config['username'] == None:
        config['username'] = input("Enter KU username >")

    if 'password' not in config or config['password'] == None:
        config['password'] = getpass.getpass("Enter KU password >")

    return config

def main():

    # Config Stuff

    config = data_config('config.yaml')

    username = config['username']
    password = config['password']

    # Under the teaching schedule table, the class rosters start at 0 and goto n
    classRosterID = "CLASSROSTER$"
    classRosterNum = 0

    # Download Directory (Change for UNIX & Windows Machines)
    downloadPath = config['download_path']

    # Begin Script....
    if not os.path.exists(downloadPath):
        os.mkdir(downloadPath)


    # Create the web driver (Firefox)

    # Firefox only runs in headless mode for some reason and I can't be bothered to try and figure it out anymore
    options = Options()
    options.headless = True

    print(os.path.join(os.getcwd(), 'drivers/geckodriver'))
    driver = webdriver.Firefox(options=options,
            executable_path=os.path.join(os.getcwd(), 'drivers/geckodriver'))


    driver.get(config['base_uri'])


    # Wait for the username field to appear
    wait = WebDriverWait(driver, 5)
    userField = wait.until(EC.presence_of_element_located((By.ID, "userid")))
    passField = driver.find_element_by_id("pwd")

    print("User Text field found")

    # Clear the fields and enter the username and password
    userField.clear()
    passField.clear()
    userField.send_keys(username)
    passField.send_keys(password)
    driver.find_element_by_id("submitBtn").click()


    # Wait for faculty center to load and enter it
    fCtrImg = wait.until(EC.presence_of_element_located((By.ID, "PZFL_FACULTY_CENTER$1")))
    fCtrImg.click()

    # The content in "Faculty Center" is served via an iframe.
    # Switch the driver to that frame
    driver.switch_to.frame(driver.find_element_by_id("ptifrmtgtframe"))

    # Make sure in the same semester as configured in yaml script
    semesterText = driver.find_element_by_id("DERIVED_SSS_FCT_SSR_STDNTKEY_DESCR$9$").get_attribute("innerText")
    semesterText = semesterText.split("|")[0].rstrip()

    # Semester doesn't match, change
    if semesterText != config["semester"]:
        print("Selecting correct semester")
        driver.find_element_by_id("DERIVED_SSS_FCT_SSS_TERM_LINK").click()

        semesterTable = driver.find_element_by_class_name("PSLEVEL1GRID")
        rowInter = 0
        semesterFound = False

        # Loop through semester table and find matching semester name
        for row in semesterTable.find_elements_by_tag_name("tr"):
            # The first "row" in the table messes it up, just do a try catch
            try:
                row_text = row.find_element_by_id("TERM_VAL$" + str(rowInter)).get_attribute('innerText')
            except NoSuchElementException:
                continue 

            # print("ROW TEXT: ", row_text)
            if row_text == config["semester"]:
                # Semester found, click the radio button and break the loop
                row.find_element_by_id("SSR_DUMMY_RECV1$sels$" + str(rowInter) + "$$0").click()
                semesterFound = True
                break
            else:
                rowInter += 1
        
        if semesterFound == False:
            print("Semester not found in table, confirm spelling is correct")
    
        # Click "continue"
        driver.find_element_by_id("DERIVED_SSS_FCT_SSR_PB_GO$254$").click()

    # In the "My Teaching Schedule" loop through the classes & sections
    # As of 09/11/2019, the first TR containing a class is ID'd as "trINSTR_CLASS_VW$0_row1"

    print("Finding row")
    wait.until(EC.presence_of_element_located((By.ID, "trINSTR_CLASS_VW$0_row1")))

    while True:
        try: 
            # Go through the classes in the table
            classBtn = wait.until(EC.presence_of_element_located((By.ID, classRosterID + str(classRosterNum))))
            classBtn.click()

            # Grab the class LONG NAME to save the file as
            className = wait.until(EC.presence_of_element_located((By.ID, "DERIVED_SSR_FC_SSR_CLASSNAME_LONG")))
            className = className.text
            className = className.replace(' ', '_') + "_ROSTER.csv"


            classTable = driver.find_element_by_xpath("//table[@id='CLASS_ROSTER_VW$scroll$0']//table[@class='PSLEVEL1GRID']")

            # tableSoup = bs(classTable.get_attribute('innerHTML'), 'html.parser')

            # with open(os.path.join(downloadPath, className), 'w') as f:
            #     f.write(classTable.get_attribute('innerHTML'))

            tableData = pd.read_html(classTable.get_attribute('outerHTML'))

            emailElms = driver.find_elements_by_xpath("//a[starts-with(@id, 'EMAIL_LINK$')]")
            emailList = []

            for elm in emailElms:
                email = elm.get_attribute("href")
                email = email.replace("mailto:", "")
                emailList.append(email)

            emailSeries = pd.Series(emailList)
            
            tableData = tableData[0]
            tableData['Student Email'] = emailSeries

            # Drop columns
            tableData = tableData.drop(["Unnamed: 0", "Notify", "Photo", "Units", "Student Services Center"], axis=1)
            tableData = tableData.set_index('ID')

            tableData.to_csv(os.path.join(downloadPath, className))

            # os.rename(os.path.join(downloadPath, "ps.xls"),
            #             os.path.join(downloadPath, className))

            time.sleep(3)

            # Click the change class button and go to the next class
            classRosterNum += 1
            driver.find_element_by_id("DERIVED_SSR_FC_SSS_CHG_CLS_LINK").click()
            
        # TODO: I changed to wait until the class was found for some reason instea of just clicking on it.
        except TimeoutException as e:
            print("Error:", e)
            print("No more classes found, export complete")
            print("Goodbye")
            break

        except NoSuchElementException as e:
            print("Error:", e)
            print("No more classes found...")
            print("Script closing")
            break

    driver.close()
    

if __name__ == "__main__":
    main()