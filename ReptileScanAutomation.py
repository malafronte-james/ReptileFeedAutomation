# importing necessary classes
# from different modules
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from zipfile import ZipFile
from pathlib import Path
import creds
import time
import os
import ctypes  # An included library with Python install.

chrome_options = webdriver.ChromeOptions()

prefs = {"profile.default_content_setting_values.notifications": 2}
chrome_options.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome("chromedriver.exe")

print("Opening browser")

# open reptilescan.com using get() method
browser.get('https://login.reptilescan.com/')

# get username and password
username = creds.username
password = creds.password

print("Let's Begin")

element = browser.find_elements_by_xpath('//*[@id="Input_Email"]')
element[0].send_keys(username)

print("Username Entered")

element = browser.find_element_by_xpath('//*[@id="Input_Password"]')
element.send_keys(password)

print("Password Entered")

# logging in
# log_in = browser.find_elements_by_id('loginbutton')
log_in = browser.find_elements_by_xpath(
    '//*[@id="LoginForm"]/form/div[3]/button')
log_in[0].click()

print("Login Successful")

browser.get('https://login.reptilescan.com/ImportExport')

element = browser.find_element_by_xpath(
    '//*[@id="Content"]/main/div[2]/div[2]/a')
element.click()

time.sleep(3)  # seconds

# get the downloads folder path
downloads_path = str(Path.home() / "Downloads")

print(downloads_path)

# get the newest file (just downloaded)
files = [os.path.join(downloads_path, x)
         for x in os.listdir(downloads_path) if x.endswith(".zip")]
newest = max(files, key=os.path.getctime)

print("Recently modified Docs", newest)

# Close the browser
browser.close()

# opening the zip file in READ mode
with ZipFile(newest, 'r') as zip:
    # printing all the contents of the zip file
    zip.printdir()

    # extracting all the files
    print('Extracting all the files now...')
    zip.extract('reptiles.xls')
    print('Done!')

# open excel
# os.system('start excel.exe reptiles.xls')

# open box letting you know it's complete
# ctypes.windll.user32.MessageBoxW(0, "Action completed, opening folder", "Completed!", 1)

# get the current working directory
path = os.getcwd()

# open the file explorer
os.startfile(path)
