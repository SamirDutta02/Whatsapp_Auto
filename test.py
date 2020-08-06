from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
chrome_path=r'C:\web_drivers\chromedriver.exe'
driver=webdriver.Chrome(chrome_path)
driver.get('https://web.whatsapp.com/')
time.sleep(15.0)
search = driver.find_element_by_xpath("//div[@class='_3FRCZ copyable-text selectable-text']")
#search.clear()
search.send_keys('yo')
time.sleep(1.0)
search.send_keys(Keys.RETURN)
text= driver.find_element_by_xpath("//div[@class='_3FRCZ copyable-text selectable-text' and @spellcheck]")
time.sleep(1.0)
text.send_keys('YO jhghjgj ')

send_btn=driver.find_element_by_xpath("//span[@data-testid='send']")
send_btn.click()