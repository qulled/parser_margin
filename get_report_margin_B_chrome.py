import datetime as dt
import logging
import os
import pickle
import time
import zipfile
from logging.handlers import RotatingFileHandler

from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


options = Options()

prefs = {'download.default_directory': r'C:\Users\ikaty\PycharmProjects\parser_margin\excel_docs'}

options.add_experimental_option('prefs', prefs)
options.add_argument("--disable-blink-features")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('--headless')

driver = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))

driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
driver.execute_cdp_cmd('Network.setUserAgentOverride', {
    "userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.53 Safari/537.36'})

stealth(driver,
        languages=["ru-Ru", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
        )


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, 'logs/')
log_file = os.path.join(BASE_DIR, 'logs/get_report_FBO.log')
console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.CRITICAL,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)


def auth(url,name):
    driver.get(url)
    cookies = pickle.load(open(f'cookies-{name}.py', 'rb'))
    for cookie in cookies:
        driver.add_cookie(cookie)
    time.sleep(25)
    # attempt = driver.find_element(By.CLASS_NAME,'WarningCookiesBannerCard__button__DSLFl2gcQr')
    # attempt.click()
    return time.sleep(1)


def get_margin(name,last_monday,last_sunday,month):
    pick_button = driver.find_element(By.CLASS_NAME,'Reports-table-row__menu-button__2dZVS0atXp')
    pick_button.click()
    time.sleep(2)
    detail_button = driver.find_element(By.CLASS_NAME, 'Navigation-item__button__1Ptxjizz-A')
    detail_button.click()
    time.sleep(5)
    download_button = driver.find_element(By.CLASS_NAME, 'DownloadButtons__download-button__9s6x08Q6Uh')
    download_button.click()
    time.sleep(2)
    detail_save_button = driver.find_element(By.CLASS_NAME,'Menu-block-item__button__1nvLrmLuUi')
    detail_save_button.click()
    time.sleep(15)
    for file_name in [f for f in os.listdir(dirparth)]:
        if file_name.startswith('Детализация'):
            zipname = file_name
    with zipfile.ZipFile(f'{dirparth}\\{zipname}', "r") as zip_ref:
        zip_ref.extractall(dirparth)
    try:
        file_oldname = os.path.join(dirparth, '0.xlsx')
        file_newname_newfile = os.path.join(dirparth, f'{name} {last_monday}-{last_sunday}.{month}.xlsx')
        os.rename(file_oldname, file_newname_newfile)
        path = os.path.join(dirparth, f'{zipname}')
        os.remove(path)
    except Exception as e:
        print(e)
    return time.sleep(2)


if __name__ == '__main__':
    date_from = dt.datetime.date(dt.datetime.now()) - dt.timedelta(days=14)
    date_to = dt.datetime.date(dt.datetime.now()) - dt.timedelta(days=8)
    last_monday = date_from.strftime("%d")
    last_sunday = date_to.strftime("%d")
    month = date_to.strftime("%m")
    dirparth = 'excel_docs/'
    try:
        name = 'Белотелов'
        auth('https://seller.wildberries.ru/suppliers-mutual-settlements/reports-implementations/reports-weekly',name)
        get_margin(name,last_monday,last_sunday,month)
    except Exception as e:
        print(e)
    finally:
        driver.close()
        exit()
