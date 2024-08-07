from msilib import text
from selenium import webdriver
from time import sleep 
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import parameters
from openpyxl import load_workbook

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
sleep(3)

driver.maximize_window()	
sleep(0.5)

driver.get('https://www.linkedin.com/')	
sleep(2)
driver.find_element("xpath", '//a[@class="nav__button-secondary btn-md btn-secondary-emphasis"]').click() #click on sign-in button
sleep(2)

# driver.find_element_by_xpath('//a[text()="Sign in"]').click()	
# sleep(2)

username_input = driver.find_element("name","session_key")

username_input.send_keys(parameters.username) 
sleep(0.5)

password_input = driver.find_element("name","session_password")	
password_input.send_keys(parameters.password)	
sleep(0.5)

driver.find_element("xpath", '//button[text()="Sign in"]').click()	
sleep(2)

book = load_workbook('C:/Users/HP/Desktop/yd/test.xlsx')
Sheet = book['Sheet1']	

for rows in Sheet.rows:

	arr=rows[2].value
	sleep(2)

	if str(arr)=="URL":
		continue

	driver.get(arr)
	sleep(2)	

	driver.find_element("link text","Message").click()
	sleep(2)
	
	driver.find_element("xpath", '//div[@aria-label="Write a message…"]/p').send_keys("Hey!! \nWelcome to Simulation World.\nFrom :- +Siddharth ")
	sleep(2)
    
	arr1=rows[0].value
	
	sleep(3)

	driver.find_element("xpath",'//button[@class="msg-form__send-button artdeco-button artdeco-button--1"]').click() 
	sleep(2)

	driver.find_element("xpath",'//button[@class="msg-overlay-bubble-header__control artdeco-button artdeco-button--circle artdeco-button--muted artdeco-button--1 artdeco-button--tertiary ember-view"]').click()
	
print("Automation Successful")
sleep(2)
driver.quit()


