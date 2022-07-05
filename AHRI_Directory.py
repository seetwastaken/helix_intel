from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import glob,os,os.path,time

#Downloading Excel Files from AHRI Directory
options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\Jason\Downloads\Excel",
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True,
  'profile.default_content_setting_values.automatic_downloads': 1
})
PATH = 'C:\Program Files (x86)\chromedriver.exe'
driver = webdriver.Chrome(chrome_options=options, executable_path = PATH)
actions = ActionChains(driver)

#Link of type of item to scrape
driver.get('https://www.ahridirectory.org/NewSearch?programId=30&searchTypeId=3')

search = driver.find_element(By.XPATH, '//*[@id="btnSearchQBottom"]')
search.click()

driver.implicitly_wait(10)

element = driver.find_element(By.XPATH, '//*[@id="btnGenerateExcel"]')
element.click()

driver.implicitly_wait(10)

max_pages = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div/div[2]/div/div[2]/div/div[3]/div[4]/span/a[6]').text
max_pages = int(max_pages) + 1


pages = range(2,max_pages)
for number in pages:
    element = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tblResults_next"]')))
    driver.execute_script("arguments[0].click();", element)

    driver.implicitly_wait(10)


    export = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnGenerateExcel"]')))
    driver.execute_script("arguments[0].click();", export)

file_name = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div/div[2]/div/div[1]/div[1]/h4').text

driver.implicitly_wait(10)
driver.close
#excel file compiler
#copy path of where files are located

time.sleep(5)
download = 'C:/Users/Jason/Downloads/Excel'
destination = 'C:/Users/Jason/Downloads/Excel/Completed'
file_list = glob.glob(download + "/*.xlsx")

excl_list = []

for file in file_list:
    if file.endswith('.xlsx'):
        excl_list.append(pd.read_excel(file))

excl_merged = pd.DataFrame()

for excl_file in excl_list:

    excl_merged = excl_merged.append(excl_file, ignore_index = True)

#transporting merged file into new folder 
excl_merged.to_excel('C:/Users/Jason/Downloads/Excel/Completed/ ' + file_name + '.xlsx' , index = False)

#Deleting the singular excel files
for done_file in file_list:
    os.remove(done_file)


    
