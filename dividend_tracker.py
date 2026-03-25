from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
import gspread
import os
import json
from google.oauth2.service_account import Credentials

# auth google sheets

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

json_str = os.getenv("GOOGLE_CREDS")

if json_str: 
    print('Using GOOGLE_CREDS environment variable')
    print('Type of json_str:', type(json_str))  # should be <class 'str'>
    print('Length of json_str:', len(json_str))
    print('First 100 chars:', json_str[:100])
    try:
            creds_dict= json.loads(json_str) #running in GitHub Actions
            print('Type after json.loads:', type(creds_dict)) #must be <class 'dict'>
            creds = Credentials.from_service_account_info(creds_dict, scopes = scope)
    except json.JSONDecodeError as e:
            raise ValueError(f'Invalid JSON in Google_Creds: {e}')
    except Exception as e:
            raise RuntimeError(f'Failed to create credentials env var: {e}')

else: #local development fallback
    print('No Google_creds env var > using local credentials.json')
    try:
        creds = Credentials.from_service_account_file('credentials.json', scopes = scope)
    except FileNotFoundError:
        raise FileNotFoundError('credentials.json not found in current directory')
    except Exception as e:
        raise RuntimeError(f'Local credentials failed: {e}')

client = gspread.authorize(creds)
spreadsheet = client.open_by_key('1J1_80uGHwLL_kdiI88F37IRV0teKi5mM8nl5xkyhJUI')
sheet = spreadsheet.worksheet('PSE Dividend Tracker')

#scrape PSE Edge with JS rendering

url = 'https://edge.pse.com.ph/disclosureData/dividends_and_rights_info_form.do'

options = Options()
options.add_argument('--headless') #run without opening the browser window
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)

driver.get(url) #better wait: wait up to 15s for table to appear
try:
    WebDriverWait(driver, 15).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "table"))
    )
except Exception as e:
    print('Table not found after wait:', e)
    print(driver.page_source[:2000])
    driver.quit()
    exit(1)

soup = BeautifulSoup(driver.page_source, 'html.parser')
driver.quit()

table = soup.find('table') #hopefully now finds it; if not, inspect further

if not table:
    print('No table found even after wait')
    print(soup.prettify()[:2000]) #debug: see if table appears in full source 
    exit(1)

rows = table.find_all('tr')
data = []

for row in rows[1:]: #skip header
    cols = row.find_all('td')
    if len(cols)>=7:
        #Adjust indices based on actual columns(inspect in browser)
        company = cols[0].text.strip()
        classification = cols[1].text.strip()
        dividend_type = cols[2].text.strip()
        amount = cols[3].text.strip()
        ex_date = cols[4].text.strip()
        record_date = cols[5].text.strip()
        payment_date = cols[6].text.strip()

        data.append([company,classification, dividend_type,amount,ex_date,record_date, payment_date])

if not data:
    print('No Dividend data found')
    exit(0)

#convert to dataframe
df = pd.DataFrame(data, columns=[
    'Company',
    'Class',
    'Type',
    'Amount',
    'Ex-Date',
    'Record Date',
    'Payment Date'
]) 
if df.empty:
    print('N/A')
    exit()

#push to google sheets
try:
    sheet.clear()
    sheet.update('A1',[df.columns.tolist()])
    sheet.update('A2', df.values.tolist())
    print('Success! Data updated in Google Sheet.')
except Exception as e:
    print('Failed to update sheet: ',e)
    exit(1)