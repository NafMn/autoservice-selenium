from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from openpyxl import load_workbook

import time

micro = load_workbook(filename = "new.xlsx")
sheetrange = micro['Sheet3']

driver = webdriver.Chrome()

driver.get("https://www.micromentor.org/e-learning/seri-literasi-keuangan/pelatihan-cerdas-menabung/")
driver.maximize_window()
driver.implicitly_wait(5)
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

i = 60
data_berhasil = 0
data_awal = i

while i <= 100:
    Email = sheetrange['A' + str(i)].value
    Password = sheetrange['B' + str(i)].value
    Nama = sheetrange['C' + str(i)].value


    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div/section[2]')))
        try:
            driver.find_element(By.ID,'id_username').send_keys(Email)
            driver.find_element(By.ID,'id_password').send_keys(Password)
            button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
            button_sub.click()
        except Exception as e:
            print(f"Error: Input salah pada indeks ke-{i} dalam perulangan.")
            print(e)

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div/div/div/div/a')))

        next_1 = driver.find_element(By.XPATH,'//a[@class="button primary"]')
        next_1.click()

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.button.primary.training-nav__next')))
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()


        # pretest no 1
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form/div[1]')))

        jawaban1 = driver.find_element(By.ID,'id_answer_1')
        jawaban1.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # pretest no 2
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))

        jawaban2 = driver.find_element(By.ID,'id_answer_0')
        jawaban2.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # pretest no 3
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban3 = driver.find_element(By.ID,'id_answer_3')
        jawaban3.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # pretest no 4
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban4 = driver.find_element(By.ID,'id_answer_3')
        jawaban4.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # pretest no 5
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_0')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # pretest no 6
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_0')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.button.primary.training-nav__next')))
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # learning
        # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[2]/div/div/a[2]')))
        # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.button.primary.training-nav__next')))

        # next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        # next_2.click()

        # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        # # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.button.primary.training-nav__next')))

        # next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        # next_2.click()

        # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[2]/div/div/a[2]')))
        # # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.button.primary.training-nav__next')))

        # next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        # next_2.click()

        # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//a[@class="button primary training-nav__next"]')))
        def lanjutkan():
            # next_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//a[@class="button primary training-nav__next"]')))
            # next_button.click()
            # WebDriverWait(driver, 10).until(EC.staleness_of(next_button))   


            # Temukan tombol "Lanjutkan" dengan mengecualikan tombol-tombol lainnya
            next_button_xpath = '//a[@class="button primary training-nav__next" and not(ancestor::div[contains(@style,"display: none")])]'
            # next_button_xpath = '//a[@class="button primary training-nav__next" and not(ancestor::div[contains(@style,"display: none")] and not(ancestor::div[contains(@class, "media-embed-block")]) and not(.//button[contains(@class, "ytp-large-play-button")])]'

            # Tunggu hingga tombol "Lanjutkan" muncul dan klik
            next_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, next_button_xpath)))
            next_button.click()

            # Tunggu hingga tombol "Lanjutkan" tidak lagi ada di DOM
            WebDriverWait(driver, 10).until(EC.staleness_of(next_button))


        lanjutkan()
        lanjutkan()
        lanjutkan()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))

        next_2 = driver.find_element(By.XPATH, '//a[@class="button primary training-nav__next"]')
        next_2.click()

        # post-test no 1
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form/div[1]')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_1')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # post-test no 2
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_0')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # post-test no 3
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_3')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # post-test no 4
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_3')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # post-test no 5
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_0')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # post-test no 6
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_1')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # post-test no 7
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]/div[1]/form')))
        
        jawaban5 = driver.find_element(By.ID,'id_answer_4')
        jawaban5.click()
        button_sub = driver.find_element(By.XPATH, '//button[@type="submit"]')
        button_sub.click()  

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        # selesai
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.button.primary.training-nav__next')))
        next_2 = driver.find_element(By.CSS_SELECTOR, 'a.button.primary.training-nav__next')
        next_2.click()

        dropdown = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//li[@class='is-dropdown-submenu-parent opens-right']")))
        # Klik elemen dropdown untuk membukanya
        dropdown.click()
        signout = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'sign_out_link')))
        signout.click()
        # signout = driver.find_element(By.ID, 'sign_out_link')
        # signout.click()

        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-content"]')))
        # kembali ke url masuk
        desired_url = "https://www.micromentor.org/accounts/login/?next=/e-learning/seri-literasi-keuangan/pelatihan-cerdas-menabung/"  
        driver.get(desired_url)
    
        data_berhasil += 1
    except Exception as e:
        print(f"Error: Input salah pada indeks ke-{i} dalam perulangan.")
        print(f"Jumlah data berhasil: {data_berhasil}")
        print(e)
        break
    
    time.sleep(5)
    i += 1
print("Jumlah data berhasil: ", data_berhasil)
print("gagal ndek nomer: ", data_awal + data_berhasil)

driver.quit()