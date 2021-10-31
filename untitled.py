# +
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from RPA.Excel.Files import Files






# driver = webdriver.Chrome(executable_path=r"C:\Users\Programacion\Documents\Python\chromedriver.exe")
# driver.get("https://itdashboard.gov/")
# driver.maximize_window()
# textToSearch = "Department of Education"
# element = driver.find_element_by_xpath("//a[@aria-controls='home-dive-in']")
# element.click()

# agencieList = driver.find_elements_by_class_name("h4 w200")
# for agencie in agencieList:
#     print(agencie.text)
    
# print("Done.")

# element10 = WebDriverWait(driver, 10).until(
#     EC.element_to_be_clickable((By.XPATH, "(//span[@class='h4 w200'][normalize-space()='" + textToSearch + "'])[1]"))
# )
# element10.click()
# WebDriverWait(driver, 10).until(
#     EC.presence_of_element_located((By.XPATH, "//select[@aria-controls='investments-table-object']"))
# )

# comboAll = Select.(driver.find_element_by_xpath("//select[@aria-controls='investments-table-object']"))
# comboAll.select_by_visible_text('All')
