from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from twilio.rest import Client
import requests
import time


path = '/usr/local/bin/chromedriver' 
service = Service(executable_path=path)
driver = webdriver.Chrome(service=service)
url = "http://reg.exam.dtu.ac.in/student/login"
driver.get(url)

TWILIO_ACCOUNT_SID = '<your_twilio_account_sid>'
TWILIO_AUTH_TOKEN = '<your_twilio_auth_token>'
TWILIO_PHONE_NUMBER = '<your_twilio_phone number>'
TO_PHONE_NUMBER = '<your_whatsapp_number>'
client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

def send(mssg):
    message = client.messages.create(
        body=mssg,
        from_=TWILIO_PHONE_NUMBER,
        to=TO_PHONE_NUMBER
    )



def check_site(url):
    try:
        response = requests.get(url)
        return response.status_code == 200
    except requests.RequestException:
        return False



search = driver.find_element(By.NAME, "roll_no")
search.send_keys("<your_rollnumber>")
search_pass = driver.find_element(By.NAME, "password")
search_pass.send_keys("<your_password>")
login = driver.find_element(By.XPATH, "//button[normalize-space()='Login']")
login.click()
course = WebDriverWait(driver, 0.8).until(
    EC.presence_of_element_located((By.XPATH, "//a[@class='text-decoration-none']"))
)
course.click()
re = 1
row = [] #this is the list of xpaths of the courses you want to register for

while True:
    if(check_site(url)==False):
       send("site is not reachable")
    registration_xpath = ""
    found_index = -1
    
    for i in range(len(row)):
        if int(driver.find_element(By.XPATH, f"//tbody/tr[{row[i]}]/td[5]").text) > 0:
            registration_xpath = f"//tbody/tr[{row[i]}]/td[6]/form[1]/button[1]"
            found_index = i
            break

    if registration_xpath:
        try:
            reg = WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.XPATH, registration_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", reg)
            driver.execute_script("arguments[0].click();", reg)
            send("Seat registered ✅")
            row.pop(found_index)
            
        except TimeoutException:
            send("❌ Failed to register seat. Not clickable.")

    driver.refresh()
    time.sleep(re)
