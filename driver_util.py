from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service  
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

from .config import LOGIN_URL, USERNAME, PASSWORD

def init_driver(download_path=None):
    options = Options()
    if download_path:
        prefs = {
            "download.default_directory": download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
    # return webdriver.Chrome(options=options)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


# def init_driver(download_path=None):
    # options = Options()
    # options.add_argument("--kiosk-printing")
    # options.add_argument("--disable-print-preview")
    # options.add_argument("--start-maximized")

    # if download_path:
    #     prefs = {
    #         "printing.print_preview_sticky_settings.appState": '''
    #         {
    #           "recentDestinations": [
    #             {
    #               "id": "Save as PDF",
    #               "origin": "local",
    #               "account": ""
    #             }
    #           ],
    #           "selectedDestinationId": "Save as PDF",
    #           "version": 2
    #         }
    #         ''',
    #         "savefile.default_directory": download_path,
    #         "download.default_directory": download_path,
    #         "download.prompt_for_download": False,
    #         "download.directory_upgrade": True,
    #         "plugins.always_open_pdf_externally": True
    #     }
    #     options.add_experimental_option("prefs", prefs)

    # service = Service(ChromeDriverManager().install())
    # return webdriver.Chrome(service=service, options=options)



def login(driver):
    driver.get(LOGIN_URL)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(USERNAME)
    driver.find_element(By.ID, "password").send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//button[contains(text(), '登入')]").click()
    time.sleep(1)