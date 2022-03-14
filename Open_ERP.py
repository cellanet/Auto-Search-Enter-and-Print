# Gia Minh Hoang
# Open a specific website (using Chrome) and login with a saved ID and Password.

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from subprocess import CREATE_NO_WINDOW
import time


# ======================================== OPEN ERP ================================================
chrome_service = ChromeService('C:/Users/***/OneDrive/Desktop/***/chromedriver.exe')
chrome_service.creationflags = CREATE_NO_WINDOW

driver = webdriver.Chrome(service=chrome_service)
driver.get("https://internal.*****erp.com/?ts=1642925866783#menu_id=108&action=101")
time.sleep(1.5)

# another tk for password
input_id = driver.find_element(By.NAME, "login")
input_id.send_keys("enter_id")
input_id = driver.find_element(By.NAME, "password")
input_id.send_keys("enter_password")
input_id.send_keys(Keys.RETURN)

# ======================================== DONE ERP ================================================