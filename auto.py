import os.path
import random
import time

import bs4
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# Main method
custom_url = 'https://docs.google.com/forms/d/e/1FAIpQLSc-VAhp0EIB1b1Xo05LLPHdmQSk7kLeE2no2Jqo4GHY-X5LUw/viewform?usp=sf_link'

# Set the path to the ChromeDriver
# The ChromeDriver is a separate executable that Selenium WebDriver uses to control Chrome
chrome_options = Options()
# chrome_options.add_argument("--headless") # Hides the browser window
chrome_options.add_argument("--no-sandbox") # Disables the sandbox
chrome_options.add_argument("--disable-notifications") # Disables notifications
chrome_options.add_argument("--disable-web-security") # Disables web security
chrome_options.add_argument("--disable-translate") # Disables translation
chrome_options.add_argument("--disable-gpu") # Disables GPU acceleration
chrome_options.add_argument("--disable-blink-features") 
chrome_options.add_argument("--disable-blink-features=AutomationControlled") # Disables automation controlled
chrome_options.add_argument("--disable-extensions") # Disables extensions
chrome_options.add_argument("--disable-infobars") # Disables infobars
chrome_options.add_argument("--disable-dev-shm-usage") # Disables shared memory usage
chrome_options.add_argument("--ignore-certificate-errors") # Ignores certificate errors
chrome_options.add_argument("--allow-running-insecure-content") # Allows running insecure content
chrome_options.add_argument("window-size=1366x768") # Sets the window size
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3") # Sets the user agent
# chrome_options.add_argument("start-maximized") # Maximizes the window

# Initialize the ChromeDriver and open the browser
browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
browser.execute_script("document.body.style.zoom='80%'") # Zoom out the page to 80%

# Open the URL
browser.get(custom_url)

# Get data from Excel
data = pd.read_excel('data.xlsx', sheet_name='Sheet1')
for idx in range(1):
  # Get the email from Excel
  email = data.iloc[idx, 1]
  print('Email: ' + email + ' fill the form!\n')
  
  # Fill the form ((Page 1))
  try:
    email_field = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/form/div[2]/div/div[2]/div/div/div[1]/div[2]/div[1]/div/div[1]/input')))
  finally:
    time.sleep(3)
    email_field = browser.find_element(By.XPATH, '/html/body/div/div[2]/form/div[2]/div/div[2]/div/div/div[1]/div[2]/div[1]/div/div[1]/input')
    email_field.send_keys(email) # Loop through the data and fill the form

    next1 = browser.find_element(By.XPATH, '/html/body/div/div[2]/form/div[2]/div/div[3]/div/div[1]/div/span')
    browser.execute_script("arguments[0].click();", next1) # Click to the next page
  
  try:
    page2 = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div/span/div/div[1]/label/div')))
    time.sleep(1)
  finally:
    # Randomly select gender (Page 1)
    gender_options = ['male', 'female']
    selected_gender = random.choice(gender_options)

    # Selecting gender (Page 1)
    if selected_gender == 'male':
      gender_xpath = '//*[@id="i9"]/div[3]/div'
    else:
      gender_xpath = '//*[@id="i12"]/div[3]/div'
    
    gender_button = browser.find_element(By.XPATH, gender_xpath)
    browser.execute_script("arguments[0].click();", gender_button)
    
    # Selecting Độ tuổi (Page 1)
    age_options = ['18–27', '28-42', '42-53', '53-62', 'Trên 62']
    # selected_age = random.choice(age_options)
    # random_age_index = random.randint(1, 5)

    age_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div/span/div/div[{random.randint(1, 5)}]/label/div'
    age_button = browser.find_element(By.XPATH, age_xpath)
    browser.execute_script("arguments[0].click();", age_button)

    # Selecting Nghề Nghiệp (Page 1)
    job_options = ['Học sinh, sinh viên', 'Việc làm bán thời gian', 'Việc làm toàn thời gian', 'Tự làm chủ/Làm nghề tự do', 'Nghỉ hưu', 'Thất nghiệp']

    job_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div/span/div/div[{random.randint(1, 6)}]/label/div'
    job_button = browser.find_element(By.XPATH, job_xpath)
    browser.execute_script("arguments[0].click();", job_button)

    # Selecting Học Vấn (Page 1)
    education_options = ['Dưới trình độ trung học phổ thông', 'Bằng cấp trung học phổ thông hoặc tương đương', 'Cao đẳng/Đại học (Đã tốt nghiệp/Đang học)', 'Sau đại học']
    
    education_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div/span/div/div[{random.randint(1, 4)}]/label/div'
    education_button = browser.find_element(By.XPATH, education_xpath)
    browser.execute_script("arguments[0].click();", education_button)

    # Selecting Thu Nhập (Page 1)
    income_options = ['Dưới 5 triệu', '5 - 10 triệu', '10 - 20 triệu', 'Trên 20 triệu']

    income_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[8]/div/div/div[2]/div/div/span/div/div[{random.randint(1, 4)}]/label/div'
    income_button = browser.find_element(By.XPATH, income_xpath)
    browser.execute_script("arguments[0].click();", income_button)

    # Selecting Miền Sinh Sống (Page 1)
    region_options = ['Miền Bắc', 'Miền Trung', 'Miền Nam']
    
    region_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[9]/div/div/div[2]/div/div/span/div/div[{random.randint(1, 3)}]/label/div'
    region_button = browser.find_element(By.XPATH, region_xpath)
    browser.execute_script("arguments[0].click();", region_button)

    # Selecting Khu Vực (Page 1)
    area_options = ['Thành phố Lớn', 'Thành Thị', 'Nông Thôn']
    
    area_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[10]/div/div/div[2]/div/div/span/div/div[{random.randint(1,3)}]/label/div'
    area_button = browser.find_element(By.XPATH, area_xpath)
    browser.execute_script("arguments[0].click();", area_button)

    # Selecting Tình Trạng Hôn Nhân (Page 1)
    # marital_options = ['Đang hôn nhân', 'Đã hôn nhân']
    
    marital_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[11]/div/div/div[2]/div/div/span/div/div[{random.randint(1, 4)}]/label/div'
    marital_button = browser.find_element(By.XPATH, marital_xpath)
    browser.execute_script("arguments[0].click();", marital_button)

    next2_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[3]/div/div[1]/div[2]/div[2]'
    next2_button = browser.find_element(By.XPATH, next2_xpath)
    browser.execute_script("arguments[0].click();", next2_button)

    def form(num_form: int,
             num_answer: int,
             start_score: int,
             end_score: int,
             path_presence: str = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]',
             path_form: str = ' //*[@id="mG61Hd"]/div[2]/div/div[2]/div[{}]/div/div/div[2]/div/div[1]/div/div[{}]/span/div[{}]/div/div/div[3]/div'):
      '''
      num_form : number of form fields to fill
      num_answer : number of answers to fill
      start_score : starting score for filling forms
      end_score : ending score for filling forms
      path_presence : XPath for the element to check presence before proceeding
      path_form : XPath pattern for each form field
      '''
      try:
        presence = WebDriverWait(browser, 10).until(EC.presence_of_element_located(By.XPATH, path_presence))
        time.sleep(1)
      except:
        print("Element not found")
      finally:
        for form_section in range(4, num_form + 1):
          answer = random.randint(2, 6)
          # I want to click answer boxes in multiple forms and has multiple answers in this section
          for form_section in range(2, num_form + 1):
            for i in range(2, num_answer + 1):
              form_field_xpath = f'{path_form.format(form_section, i, answer)}'
              browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH, form_field_xpath))
          for i in range(2, num_answer + 1):
            form_field_xpath = f'{path_form.format(num_form, i, answer)}'
            browser.execute_script("arguments[0].click();", browser.find_element(By.XPATH, form_field_xpath))
    
    # Call the form() function with the desired parameters
    # //*[@id="mG61Hd"]/div[2]/div/div[2]/div[form_section]/div/div/div[2]/div/div[1]/div/div[4]/span/div[2]/div/div/div[3]/div
    form(num_form=5, num_answer=4, start_score=1, end_score=10)
    
    next3_xpath = f'//*[@id="mG61Hd"]/div[2]/div/div[3]/div/div[1]/div[2]/span'
    next3_button = browser.find_element(By.XPATH, next3_xpath)
    browser.execute_script("arguments[0].click();", next3_button)

    # For debugging purposes
    time.sleep(10)
    print('Done')