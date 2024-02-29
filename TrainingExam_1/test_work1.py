import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By

# Открываем Excel файл
workbook = openpyxl.load_workbook('task.xlsx')
sheet = workbook.active

# Читаем данные из ячеек
data = []
for row in sheet.iter_rows(values_only=True):
    data.append(row)

# print(data)


def filling_out_the_form(driver):
    element1 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelFirstName"]')
    element2 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelLastName"]')
    element3 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelAddress"]')
    element4 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelEmail"]')
    element5 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelPhone"]')
    element6 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelCompanyName"]')
    element7 = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelRole"]')
    
    element1.send_keys(data[1][0])
    element2.send_keys(data[1][1])
    element3.send_keys(data[1][2])
    element4.send_keys(data[1][3])
    element5.send_keys(data[1][4])
    element6.send_keys(data[1][5])
    element7.send_keys(data[1][6])
    

    login_button = driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input')
    login_button.click()


# Запускаем веб-драйвер (Chrome)
driver = webdriver.Chrome()
driver.get('https://rpachallenge.com/?lang=EN')

for _ in range(10):
    filling_out_the_form(driver=driver)

driver.quit()
