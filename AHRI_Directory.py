from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time

options = Options()
options.add_argument("--start-maximized")

PATH = 'C:\Program Files (x86)\chromedriver.exe'
driver = webdriver.Chrome(chrome_options=options, executable_path = PATH)
actions = ActionChains(driver)


driver.get('https://www.ahridirectory.org/NewSearch?programId=5&searchTypeId=3&productTypeId=1732')

link = driver.find_element(By.XPATH, '//*[@id="btnSearchQBottom"]')
link.click()

time.sleep(5)
element = driver.find_element(By.XPATH, '//*[@id="btnGenerateExcel"]')
element.click()

pages = range(2,5)
for number in pages:
    element = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/form/div/div/div[2]/div/div[2]/div/div[3]/div[4]/span/a[' + str(number) + ']')
    element.click()

    time.sleep(5)

    element = driver.find_element(By.XPATH, '//*[@id="btnGenerateExcel"]')
    element.click()

    time.sleep(5)



    
