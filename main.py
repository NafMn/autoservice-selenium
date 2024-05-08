from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from openpyxl import load_workbook

import time

micro = load_workbook(filename="D:\\Aplikasi\\micronook.xlsx")
sheetrange = micro['Sheet1']

driver = webdriver.Chrome()

driver.get("https://www.micromentor.org/e-learning/seri-literasi-keuangan/pelatihan-cerdas-menabung/")
driver.maximize_window()
driver.implicitly_wait(10)


i = 5

while i <= 6:
    Email = sheetrange['A' + str(i)].value
    Password = sheetrange['B' + str(i)].value
    Nama = sheetrange['C' + str(i)].value


    # Close the cookie dialog if present
    try:
        cookie_dialog = driver.find_element(By.ID, "CybotCookiebotDialog")
        if cookie_dialog.is_displayed():
            close_button = cookie_dialog.find_element(By.ID, "CybotCookiebotDialogBodyLevelButtonLevelOptinDecline")
            close_button.click()
    except:
        pass
    button_masuk = driver.find_element(By.XPATH, '//a[@class="button primary hollow"]')
    button_masuk.click()

    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div/section[2]')))

        driver.find_element(By.ID,'id_username').send_keys(Email)
        driver.find_element(By.ID,'id_password').send_keys(Password)
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div/div/div/div/a')))

        next_1 = driver.find_element(By.XPATH,'//a[@class="button primary"]')
        next_1.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form/div[1]')))

        jawaban1 = driver.find_element(By.ID,'id_answer_1')
        jawaban1.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        driver.find_element_by_link_text('Lanjutkan').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))

        jawaban2 = driver.find_element(By.ID,'id_answer_0')
        jawaban2.click()
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        driver.find_element_by_link_text('Lanjutkan').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban3 = driver.find_element(By.ID,'id_answer_3')
        jawaban3.click()
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        driver.find_element_by_link_text('Lanjutkan').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban4 = driver.find_element(By.ID,'id_answer_3')
        jawaban4.click()
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        driver.find_element_by_link_text('Lanjutkan').click()

    except TimeoutException:
        print("Gagal Buos, mbaleni neh")
        pass
    
    time.sleep(5)
    i += 1

print("Oke")

driver.quit()
