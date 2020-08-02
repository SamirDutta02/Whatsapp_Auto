from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
chrome_path=r'C:\web_drivers\chromedriver.exe'
driver=webdriver.Chrome(chrome_path)
driver.get("http://www.google.com")
try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "q"))
    )
finally:
    driver.quit()