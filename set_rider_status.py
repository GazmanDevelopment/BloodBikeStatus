import configparser
import datetime as dt
import os
import sys
import time
import zoneinfo

from O365 import Account
from O365 import MSGraphProtocol
from selenium import webdriver
from selenium.webdriver.common.by import By

# Python Webservice
# https://towardsdatascience.com/build-your-own-python-restful-web-service-840ed7766832

# Create O365 Event
# https://stackoverflow.com/questions/71813869/creating-a-calendar-event-o365-package-in-python
# https://o365.github.io/python-o365/latest/getting_started.html 

config_path = "C:\\Users\\Username\\Path\\to\\settings.cfg"

geo_params = {}
url = ""
username = ""
password = ""

set_reminder = False
app_client_id = ""
client_secret = ""
event_subject = ""
resource_email = ""
time_zone = ""
reminder_hour = 20
reminder_minute = 30

def load_config():
    try:
        global url, username, password, geo_params, app_client_id, client_secret, event_subject, resource_email, time_zone, set_reminder, reminder_hour, reminder_minute

        if(os.path.exists(config_path)):
            config = configparser.ConfigParser()
            config.read(config_path)

            # Get the website details
            url = config.get('SITE', 'url')
            username = config.get('SITE', 'username')
            password = config.get('SITE', 'password')

            # Get the geolocation items from the config file
            geo_params = {
                "latitude": float(config.get('GEOLOCATION', 'latitude')),
                "longitude": float(config.get('GEOLOCATION', 'longitude')),
                "accuracy": int(config.get('GEOLOCATION', 'accuracy'))
            }

            #Load the MS365 details for the reminder details
            set_reminder = config.get('MS365', 'set_reminder')
            app_client_id = config.get('MS365', 'app_client_id')
            client_secret = config.get('MS365', 'client_secret')
            event_subject = config.get('MS365', 'event_subject')
            resource_email = config.get('MS365', 'account_email')
            time_zone = config.get('MS365', 'time_zone')
            reminder_hour = int(config.get('MS365', 'reminder_hour'))
            reminder_minute = int(config.get('MS365', 'reminder_minute'))
        else:
            raise Exception("Config file not found in specified location")
    except Exception as ex:
        print(f"load_config().  Error reading config: {ex}")
        exit()

def get_cli_args():
    try:
        status_value = sys.argv[1]
        print(f"Setting rider status to {status_value}")
        match status_value:
            case 'Available':
                return 'btnAvailable'
            case 'Unavailable':
                return 'btnUnavailable'
            case _:
                print("No valid argument entered, exiting program")
                exit()
    except Exception as ex:
        print(f"get_cli_args().  Error reading arguments: {ex}")

def login():
    try:
        browser = webdriver.Chrome()

        browser.execute_cdp_cmd("Page.setGeolocationOverride", geo_params)
        browser.get(url)

        assert 'Bloodbikes - Riders - Set your availability' in browser.title

        menu_button = browser.find_element(By.ID, 'dropdownMenu1')  # Find the search box
        menu_button.click()

        user_input = browser.find_element(By.ID, 'emailInput')
        password_input = browser.find_element(By.ID, 'passwordInput')

        user_input.send_keys(username)
        password_input.send_keys(password)

        submit_button = browser.find_element(By.XPATH, "//form[@id='frm-login']/div[3]/div")
        submit_button.click()

        return browser
    except Exception as ex:
        print(f"login().  Unable to open browser and log in: {ex}")
        return False

def set_status(browser, availability_status):
    try:
        time.sleep(3)
        assert 'As a rider' in browser.page_source

        browser.execute_script(f"SetAvailability({availability_status})")
        time.sleep(3)
        return True
    except Exception as ex:
        print(f"set_status().  Unable to set status to {availability_status}: {ex}")
        return False

def add_reminder():
    try:
        if (set_reminder == 'True'):
            TZI = zoneinfo.ZoneInfo(time_zone)
            now = dt.datetime.today().replace(tzinfo=TZI)
            reminder_datetime = now.replace(hour=reminder_hour, minute=reminder_minute)

            protocol_graph = MSGraphProtocol()
            account = Account(credentials=(app_client_id, client_secret), protocol=protocol_graph, scopes=['basic', 'Calendars.ReadWrite'])

            # Run one time only
            #  - Copy URL from terminal to browser
            #  - Login and authenticate
            #  - A blank page will appear, copy the URL back in to the terminal
            #  - Press the [ENTER] key
            #account.authenticate()

            schedule = account.schedule(resource=resource_email)
            calendar = schedule.get_default_calendar()
            new_event = calendar.new_event()
            new_event.subject = event_subject
            new_event.start = reminder_datetime
            new_event.save()
            print("add_reminder().  Reminder event set")
        else:
            print("add_reminder().  Reminder not configured to be set.")
    except Exception as ex:
        print(f"add_reminder().  Error setting reminder: {ex}")

if __name__ == "__main__":
    try:
        load_config()
        rider_status = get_cli_args()

        # Login and return the browser object
        browser_object = login()
        if (browser_object):
            #Once logged in, set the status to available
            if (set_status(browser_object, rider_status)):
                if (rider_status == "btnAvailable"):
                    add_reminder()
                print(f"__main__().  Status updated")
    except Exception as ex:
        print(f"__main__(). Error running program.  Error was: {ex}")