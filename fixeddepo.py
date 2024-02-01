import time

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait

import main

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.get('https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI'
           '-BSB001.html?classic=true#result')

driver.maximize_window()

file = "S:\\fixeddepo cal.xlsx"


rows = main.getRowCount(file,"Sheet1")


for r in range(2,rows+1):
    prin = main.readData(file,"Sheet1",r,1)
    rateofin = main.readData(file,"Sheet1",r,2)
    per1 = main.readData(file,"Sheet1",r,3)
    per2 = main.readData(file, "Sheet1", r, 4)
    freq = main.readData(file, "Sheet1", r, 5)
    exp_mvalu = main.readData(file, "Sheet1", r, 6)

    driver.find_element(By.XPATH,'//*[@id="principal"]').send_keys(prin)
    driver.find_element(By.XPATH, '//*[@id="interest"]').send_keys(rateofin)
    driver.find_element(By.XPATH, '//*[@id="tenure"]').send_keys(per1)
    perdrp = Select(driver.find_element(By.XPATH,'//*[@id="tenurePeriod"]'))
    perdrp.select_by_visible_text(per2)

    freqdrp = Select(driver.find_element(By.XPATH,'//*[@id="frequency"]'))
    freqdrp.select_by_visible_text(freq)

    driver.find_element(By.XPATH,'//*[@id="fdMatVal"]/div[2]/a[1]/img').click()
    act_mval = driver.find_element(By.XPATH,"(//span[@id='resp_matval'])/strong").text


    if float(exp_mvalu) == float(act_mval):
        print("test pass")
        main.writeData(file,"Sheet1",r,8,"Passed")
        main.fillGreenColor(file,"Sheet1",r,8)

    else:
        print("test failed")
        main.writeData(file, "Sheet1", r, 8, "Failed")
        main.fillRedColor(file, "Sheet1", r, 8)
    driver.find_element(By.XPATH,'//*[@id="fdMatVal"]/div[2]/a[2]/img').click()
    time.sleep(2)
