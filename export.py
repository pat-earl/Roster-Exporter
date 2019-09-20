import getpass
import os
import time
import yaml # pip install pyyaml
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

def data_config(yaml_file : str):
    with open(yaml_file, 'r') as stream:
        config = yaml.load(stream, Loader=yaml.FullLoader)

    if 'username' not in config or config['username'] == None:
        config['username'] = input("Enter KU username >")

    if 'password' not in config or config['password'] == None:
        config['password'] = getpass.getpass("Enter KU password >")

    return config

def main():

    # Config Stuff

    config = data_config('data.yaml')

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

    # Setup firefox profile to automatically save the "excel" files
    # Couldn't figure out how to set Firefox to not prompt on file download
    # Have to use chrome sadly

    # fp = webdriver.FirefoxProfile()
    # fp.set_preference("browser.download.folderList", 2)
    # fp.set_preference("browser.download.manger.showWhenStarting", False)
    # fp.set_preference("borwser.download.dir", downloadPath)
    # fp.set_preference("browser.helperApp.neverAsk.saveToDisk", "application/vnd.ms-excel")

    # Create the web driver (Firefox)
    # driver = webdriver.Firefox(firefox_profile=fp,executable_path='./drivers/geckodriver.exe')


    # Setup driver and chrome profile
    chromeOptions = webdriver.ChromeOptions()

    prefs = {"download.default_directory" : downloadPath}
    chromeOptions.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(executable_path='./drivers/geckodriver', chrome_options=chromeOptions)

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
            className = className.replace(' ', '_') + "_ROSTER.xls"


            # Download the "Excel" file
            driver.find_element_by_id("CLASS_ROSTER_VW$hexcel$0").click()

            time.sleep(3)

            os.rename(os.path.join(downloadPath, "ps.xls"),
                        os.path.join(downloadPath, className))

            # Click the change class button and go to the next class
            classRosterNum += 1
            driver.find_element_by_id("DERIVED_SSR_FC_SSS_CHG_CLS_LINK").click()
            

        except NoSuchElementException:
            print("No more classes found...")
            print("Script closing")
            break

    driver.close()
    

if __name__ == "__main__":
    main()