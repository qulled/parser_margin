from logging.handlers import RotatingFileHandler
from googleapiclient import discovery
from google.oauth2 import service_account
import logging
import os
import datetime as dt
import json
from dotenv import load_dotenv


CREDENTIALS_FILE = 'credentials_service.json'
credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
service = discovery.build('sheets', 'v4', credentials=credentials)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, 'logs/')
log_file = os.path.join(BASE_DIR, 'logs/creater.log')
console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')


if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
load_dotenv('.env ')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')


def copy_sheets(table_id, name, sheet_id, last_day, present_day, month):
    requests = {
        'duplicateSheet': {
            'sourceSheetId': sheet_id,
            'insertSheetIndex': 1,
            'newSheetName': f'{name} {last_day}-{present_day}.{month}'
        }
    }
    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(spreadsheetId=table_id,
                                                  body=body).execute()
    return response



if __name__ == '__main__':
    date_from = dt.datetime.date(dt.datetime.now()) - dt.timedelta(days=7)
    date_to = dt.datetime.date(dt.datetime.now()) - dt.timedelta(days=1)
    last_monday = date_from.strftime("%d")
    last_sunday = date_to.strftime("%d")
    month = date_to.strftime("%m")
    cred_file = os.path.join(BASE_DIR, 'credentials.json')
    table_id = SPREADSHEET_ID
    with open(cred_file, 'r', encoding="utf-8") as f:
        cred = json.load(f)
    for i in cred:
        if i == 'Белотелов':
            copy_sheets(table_id, i, 1, last_monday, last_sunday, month)
        elif i == 'Орлова':
            copy_sheets(table_id, i, 1, last_monday, last_sunday, month)
        elif i == 'Кулик':
            copy_sheets(table_id, i, 1, last_monday, last_sunday, month)
