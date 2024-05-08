from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By


from openpyxl import load_workbook

import time

micro = load_workbook(filename="D:\Aplikasi\micronook.xlsx")

sheetrange = micro['Sheet1']


driver = webdriver.Chrome()

driver.get("https://www.micromentor.org/e-learning/seri-literasi-keuangan/pelatihan-cerdas-menabung/")
driver.maximize_window()
driver.implicitly_wait(10)

# looping

i = 5

# len(sheetrange['A'])
while i <=6:
    Email = sheetrange['A'+str(i)].value
    Password = sheetrange['B'+str(i)].value
    Nama = sheetrange['C'+str(i)].value

    button_masuk = driver.find_element(By.XPATH, '//a[@class="button primary hollow"]')
    button_masuk.click()

    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="main-content"]/div/section[2]')))

        driver.find_element_by_id('id_username').send_keys(Email)
        driver.find_element_by_id('id_password').send_keys(Password)
        button_sub = driver.find_element_by_id('submit')
        button_sub.click()

        # Menunggu halaman selanjutnya dimuat
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div/div/div/div/a')))

        # Klik halaman selanjutnya
        next_1 = driver.find_element_by_css_selector('a.button.primary.hollow')
        # Menggunakan link text
        # driver.find_element_by_link_text('Mulai Pelatihan').click()
        next_1.click()


        # Menunggu halaman selanjutnya dimuat
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[2]/div/div/a')))

        # xpath
        driver.find_element(By.XPATH,'//a[@class="button primary training-nav__next"]').click()
        # Menggunakan link text
        # driver.find_element_by_link_text('Lanjutkan').click()

        # Menunggu halaman selanjutnya dimuat
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form/div[1]')))

        driver.find_element_by_id('id_answer_1')
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        # Menggunakan link text
        driver.find_element_by_link_text('Lanjutkan').click()
        # css
        # driver.find_element_by_css_selector('a.button.primary.training-nav__next').click()


        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))

        driver.find_element_by_id('id_answer_0')
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        # Menggunakan link text
        driver.find_element_by_link_text('Lanjutkan').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        driver.find_element_by_id('id_answer_3')
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        # Menggunakan link text
        driver.find_element_by_link_text('Lanjutkan').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        driver.find_element_by_id('id_answer_3')
        driver.find_element_by_id('submit').click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        # Menggunakan link text
        driver.find_element_by_link_text('Lanjutkan').click()

    except TimeoutException:
        print("Gagal Buos, mbaleni neh")
        pass
    
    time.sleep(5)
    i = i+1

print("oke")

# Tutup WebDriver
driver.quit()