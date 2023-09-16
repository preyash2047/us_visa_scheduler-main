import time
import json
import random
import requests
import configparser
from datetime import datetime

import pandas as pd  # Added pandas library for reading Excel data
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as Wait
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

from embassy import Embassies

# Load user details from Excel file
excel_file = 'user_details.xlsx'
df = pd.read_excel(excel_file)

config_file = 'config.ini'
config = configparser.ConfigParser()
config.read(config_file)

LOCAL_USE = config['CHROMEDRIVER'].getboolean('LOCAL_USE')
HUB_ADDRESS = config['CHROMEDRIVER']['HUB_ADDRESS']
PUSHOVER_TOKEN = config['NOTIFICATION']['PUSHOVER_TOKEN']
PUSHOVER_USER = config['NOTIFICATION']['PUSHOVER_USER']
SENDGRID_API_KEY = config['NOTIFICATION']['SENDGRID_API_KEY']
PERSONAL_SITE_USER = config['NOTIFICATION']['PERSONAL_SITE_USER']
PERSONAL_SITE_PASS = config['NOTIFICATION']['PERSONAL_SITE_PASS']
PUSH_TARGET_EMAIL = config['NOTIFICATION']['PUSH_TARGET_EMAIL']
PERSONAL_PUSHER_URL = config['NOTIFICATION']['PERSONAL_PUSHER_URL']
RETRY_TIME_L_BOUND = config['TIME'].getfloat('RETRY_TIME_L_BOUND')
RETRY_TIME_U_BOUND = config['TIME'].getfloat('RETRY_TIME_U_BOUND')
WORK_LIMIT_TIME = config['TIME'].getfloat('WORK_LIMIT_TIME')
WORK_COOLDOWN_TIME = config['TIME'].getfloat('WORK_COOLDOWN_TIME')
BAN_COOLDOWN_TIME = config['TIME'].getfloat('BAN_COOLDOWN_TIME')

# Time Section:
minute = 60
hour = 60 * minute
# Time between steps (interactions with forms)
STEP_TIME = 0.5

JS_SCRIPT = ("var req = new XMLHttpRequest();"
             f"req.open('GET', '%s', false);"
             "req.setRequestHeader('Accept', 'application/json, text/javascript, */*; q=0.01');"
             "req.setRequestHeader('X-Requested-With', 'XMLHttpRequest');"
             f"req.setRequestHeader('Cookie', '_yatri_session=%s');"
             "req.send(null);"
             "return req.responseText;")

def initialize_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Enable headless mode
    chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration in headless mode
    chrome_options.add_argument("--no-sandbox")  # Disable sandboxing in headless mode
    chrome_options.add_argument("--disable-dev-shm-usage")  # Disable /dev/shm usage in headless mode
    if LOCAL_USE:
        # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver_path = "C:\chromedriver.exe"
        driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
        # driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    else:
        driver = webdriver.Remote(command_executor=HUB_ADDRESS, options=chrome_options)
    return driver


class VisaScheduler:
    def __init__(self,USERNAME,PASSWORD,SCHEDULE_ID,PRIOD_START,PRIOD_END,YOUR_EMBASSY, Embassies):
        self.driver = initialize_driver()
        self.LOG_FILE_NAME = "log_" + USERNAME + "_" +str(datetime.now().date()) + ".txt"   

        self.EMBASSY_COUNTER = 0
        self.first_loop = True

        self.USERNAME = USERNAME
        self.PASSWORD = PASSWORD
        self.SCHEDULE_ID = SCHEDULE_ID
        self.PRIOD_START = PRIOD_START.strftime('%Y-%m-%d')
        self.PRIOD_END = PRIOD_END.strftime('%Y-%m-%d')
        self.YOUR_EMBASSY = YOUR_EMBASSY
        self.Embassies = Embassies
        
        self.EMBASSY = self.Embassies[self.YOUR_EMBASSY][0]
        self.FACILITY_ID = self.Embassies[self.YOUR_EMBASSY][1]
        self.REGEX_CONTINUE = self.Embassies[self.YOUR_EMBASSY][2]

        self.SIGN_IN_LINK = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/users/sign_in"
        self.APPOINTMENT_URL = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/schedule/{self.SCHEDULE_ID}/appointment"
        self.DATE_URL = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/schedule/{self.SCHEDULE_ID}/appointment/days/{self.FACILITY_ID}.json?appointments[expedite]=false"
        self.TIME_URL = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/schedule/{self.SCHEDULE_ID}/appointment/times/{self.FACILITY_ID}.json?date=%s&appointments[expedite]=false"
        self.SIGN_OUT_LINK = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/users/sign_out"

    def send_notification(self, title, msg):
        print(f"Sending notification!")
        if SENDGRID_API_KEY:
            message = Mail(from_email=self.USERNAME, to_emails=self.USERNAME, subject=msg, html_content=msg)
            try:
                sg = SendGridAPIClient(SENDGRID_API_KEY)
                response = sg.send(message)
                print(response.status_code)
                print(response.body)
                print(response.headers)
            except Exception as e:
                print(e.message)
        if PUSHOVER_TOKEN:
            url = "https://api.pushover.net/1/messages.json"
            data = {
                "token": PUSHOVER_TOKEN,
                "user": PUSHOVER_USER,
                "message": msg
            }
            requests.post(url, data)
        if PERSONAL_SITE_USER:
            url = PERSONAL_PUSHER_URL
            data = {
                "title": "VISA - " + str(title),
                "user": PERSONAL_SITE_USER,
                "pass": PERSONAL_SITE_PASS,
                "email": PUSH_TARGET_EMAIL,
                "msg": msg,
            }
            requests.post(url, data)

    def auto_action(self, label, find_by, el_type, action, value, sleep_time=0):
        print("\t"+ label +":", end="")
        # Find Element By
        if find_by.lower() == 'id':
            item = self.driver.find_element(By.ID, el_type)
        elif find_by.lower() == 'name':
            item = self.driver.find_element(By.NAME, el_type)
        elif find_by.lower() == 'class':
            item = self.driver.find_element(By.CLASS_NAME, el_type)
        elif find_by.lower() == 'xpath':
            item = self.driver.find_element(By.XPATH, el_type)
        else:
            return 0

        # Do Action:
        if action.lower() == 'send':
            item.send_keys(value)
        elif action.lower() == 'click':
            item.click()
        else:
            return 0

        print("\t\tCheck!")
        if sleep_time:
            time.sleep(sleep_time)

    def start_process(self):
        # Bypass reCAPTCHA
        self.driver.get(self.SIGN_IN_LINK)
        time.sleep(STEP_TIME)
        Wait(self.driver, 60).until(EC.presence_of_element_located((By.NAME, "commit")))
        self.auto_action("Click bounce", "xpath", '//a[@class="down-arrow bounce"]', "click", "", STEP_TIME)
        self.auto_action("Email", "id", "user_email", "send", self.USERNAME, STEP_TIME)
        self.auto_action("Password", "id", "user_password", "send", self.PASSWORD, STEP_TIME)
        self.auto_action("Privacy", "class", "icheckbox", "click", "", STEP_TIME)
        self.auto_action("Enter Panel", "name", "commit", "click", "", STEP_TIME)
        Wait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(), '" + self.REGEX_CONTINUE + "')]")))
        print("\n\tlogin successful!\n")

    def reschedule(self, date):
        time = self.get_time(date)
        self.driver.get(self.APPOINTMENT_URL)
        headers = {
            "User-Agent": self.driver.execute_script("return navigator.userAgent;"),
            "Referer": self.APPOINTMENT_URL,
            "Cookie": "_yatri_session=" + self.driver.get_cookie("_yatri_session")["value"]
        }
        data = {
            "utf8": self.driver.find_element(by=By.NAME, value='utf8').get_attribute('value'),
            "authenticity_token": self.driver.find_element(by=By.NAME, value='authenticity_token').get_attribute('value'),
            "confirmed_limit_message": self.driver.find_element(by=By.NAME, value='confirmed_limit_message').get_attribute('value'),
            "use_consulate_appointment_capacity": self.driver.find_element(by=By.NAME, value='use_consulate_appointment_capacity').get_attribute('value'),
            "appointments[consulate_appointment][facility_id]": self.FACILITY_ID,
            "appointments[consulate_appointment][date]": date,
            "appointments[consulate_appointment][time]": time,
        }
        r = requests.post(self.APPOINTMENT_URL, headers=headers, data=data)
        if(r.text.find('Successfully Scheduled') != -1):
            title = "SUCCESS"
            msg = f"Rescheduled Successfully! {date} {time}"
        else:
            title = "FAIL"
            msg = f"Reschedule Failed!!! {date} {time}"
        return [title, msg]

    def get_date(self):
        # Requesting to get the whole available dates
        session = self.driver.get_cookie("_yatri_session")["value"]
        script = JS_SCRIPT % (str(self.DATE_URL), session)
        content = self.driver.execute_script(script)
        return json.loads(content)

    def get_time(self, date):
        time_url = self.TIME_URL % date
        session = self.driver.get_cookie("_yatri_session")["value"]
        script = JS_SCRIPT % (str(time_url), session)
        content = self.driver.execute_script(script)
        data = json.loads(content)
        time = data.get("available_times")[-1]
        print(f"Got time successfully! {date} {time}")
        return time

    def is_logged_in(self):
        content = self.driver.page_source
        if(content.find("error") != -1):
            return False
        return True

    def get_available_date(self, dates):
        # Evaluation of different available dates
        def is_in_period(date, PSD, PED):
            new_date = datetime.strptime(date, "%Y-%m-%d")
            result = ( PED >= new_date and new_date >= PSD )
            # print(f'{new_date.date()} : {result}', end=", ")
            return result
        
        PED = datetime.strptime(self.PRIOD_END, "%Y-%m-%d")
        PSD = datetime.strptime(self.PRIOD_START, "%Y-%m-%d")
        for d in dates:
            date = d.get('date')
            if is_in_period(date, PSD, PED):
                return date
        print(f"\n\nNo available dates between ({PSD.date()}) and ({PED.date()})!")

    def update_embassy(self):
        if len(list(self.Embassies)) == 0:
            # Ban Situation
            msg = f"Embassies List is empty, Probabely banned!\n\tSleep for {self.BAN_COOLDOWN_TIME} hours!\n"
            print(msg)
            self.info_logger(msg)
            self.send_notification("BAN", msg)
            self.driver.get(self.SIGN_OUT_LINK)
            time.sleep(self.BAN_COOLDOWN_TIME * hour)
            self.first_loop = True    

        if self.EMBASSY_COUNTER >= len(list(self.Embassies)):
            self.EMBASSY_COUNTER = 0
        self.YOUR_EMBASSY = list(self.Embassies)[self.EMBASSY_COUNTER]
        self.EMBASSY = self.Embassies[self.YOUR_EMBASSY][0]
        self.FACILITY_ID = self.Embassies[self.YOUR_EMBASSY][1]
        self.REGEX_CONTINUE = self.Embassies[self.YOUR_EMBASSY][2]
        print("Now looking for", self.YOUR_EMBASSY) 

        self.SIGN_IN_LINK = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/users/sign_in"
        self.APPOINTMENT_URL = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/schedule/{self.SCHEDULE_ID}/appointment"
        self.DATE_URL = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/schedule/{self.SCHEDULE_ID}/appointment/days/{self.FACILITY_ID}.json?appointments[expedite]=false"
        self.TIME_URL = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/schedule/{self.SCHEDULE_ID}/appointment/times/{self.FACILITY_ID}.json?date=%s&appointments[expedite]=false"
        self.SIGN_OUT_LINK = f"https://ais.usvisa-info.com/{self.EMBASSY}/niv/users/sign_out"
        self.EMBASSY_COUNTER += 1

    def info_logger(self, log):
        # file_path: e.g. "log.txt"
        with open(self.LOG_FILE_NAME, "a") as file:
            file.write(str(datetime.now().time()) + ":\n" + log + "\n")

    def run(self):
        while 1:
            if self.first_loop:
                t0 = time.time()
                total_time = 0
                Req_count = 0
                self.start_process()
                self.first_loop = False
            Req_count += 1
            try:
                msg = "-" * 60 + f"\nRequest count: {Req_count}, Log time: {datetime.today()}\n"
                print(msg)
                self.info_logger(msg)
                dates = self.get_date()
                if not dates:
                    # No Dats remove the state
                    del self.Embassies[self.YOUR_EMBASSY]
                    if self.EMBASSY_COUNTER > 0:
                        self.EMBASSY_COUNTER = 1
                    else:
                        self.EMBASSY_COUNTER -= 1   
                    msg = f"List is empty, So removed {self.YOUR_EMBASSY}!"
                    print(msg)
                    self.info_logger(msg)
                    self.send_notification("EMBASSY REMOVED", msg)

                # Print Available dates:
                msg = ""
                for d in dates:
                    msg = msg + "%s" % (d.get('date')) + ", "
                msg = "Available dates:\n"+ msg
                print(msg)
                self.info_logger(msg)
                date = self.get_available_date(dates)
                if date:
                    # A good date to schedule for
                    END_MSG_TITLE, msg = self.reschedule(date)
                    break
                self.update_embassy()
                self.RETRY_WAIT_TIME = random.randint(RETRY_TIME_L_BOUND, RETRY_TIME_U_BOUND)
                t1 = time.time()
                total_time = t1 - t0
                msg = "\nWorking Time:  ~ {:.2f} minutes".format(total_time/minute)
                print(msg)
                self.info_logger(msg)
                if total_time > WORK_LIMIT_TIME * hour:
                    # Let program rest a little
                    self.send_notification("REST", f"Break-time after {WORK_LIMIT_TIME} hours | Repeated {Req_count} times")
                    self.driver.get(self.SIGN_OUT_LINK)
                    time.sleep(WORK_COOLDOWN_TIME * hour)
                    self.first_loop = True
                else:
                    msg = "Retry Wait Time: "+ str(self.RETRY_WAIT_TIME)+ " seconds"
                    print(msg)
                    self.info_logger(msg)
                    time.sleep(self.RETRY_WAIT_TIME)
            except Exception as e:
                # Exception Occurred
                msg = f"Exception occurred: {str(e)}\n"
                self.END_MSG_TITLE = "EXCEPTION"
                print(msg)
                self.info_logger(msg)
                break

        print(msg)
        self.info_logger(msg)
        self.send_notification(self.END_MSG_TITLE, msg)
        self.driver.get(self.SIGN_OUT_LINK)
        self.driver.stop_client()
        self.driver.quit()

# Loop through each row in the Excel file and run the code for each user
for index, row in df.iterrows():
    username = row['USERNAME']
    password = row['PASSWORD']
    schedule_id = row['SCHEDULE_ID']
    period_start = row['PRIOD_START']
    period_end = row['PRIOD_END']
    your_embassy = row['YOUR_EMBASSY']

    # Call the function to run the visa scheduler for the current user
    visa_scheduler = VisaScheduler(username,password,schedule_id,period_start,period_end,your_embassy, Embassies)
    visa_scheduler.run()
