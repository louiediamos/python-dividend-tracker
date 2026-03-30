from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import gspread
import os
import json
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from gspread_formatting import format_cell_range, CellFormat

# auth google sheets
scope = [    'https://www.googleapis.com/auth/spreadsheets',    'https://www.googleapis.com/auth/drive']

json_str = os.getenv("GOOGLE_CREDS")

#running in GitHub Actions
if json_str: 
    print('Using GOOGLE_CREDS environment variable')
    creds_dict= json.loads(json_str) 
    creds = Credentials.from_service_account_info(creds_dict, scopes = scope)

#local development fallback
else: 
    print('No Google_creds env var > using local credentials.json')
    creds = Credentials.from_service_account_file('credentials.json', scopes = scope)
    
client = gspread.authorize(creds)
spreadsheet = client.open_by_key('1J1_80uGHwLL_kdiI88F37IRV0teKi5mM8nl5xkyhJUI')

#scrape PSE Edge with JS rendering
url = 'https://edge.pse.com.ph/disclosureData/dividends_and_rights_info_form.do'

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
driver.get(url) 
WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "table"))
    )

soup = BeautifulSoup(driver.page_source, 'html.parser')
driver.quit()
table = soup.find('table') 

rows = table.find_all('tr')
datacommon = []
datapref = []

for row in rows[1:]: 
    cols = row.find_all('td')
    if len(cols)>=7:
        company = cols[0].text.strip()
        shr_class = cols[1].text.strip()
        dividend_type = cols[2].text.strip()
        amount = cols[3].text.strip()
        ex_date = cols[4].text.strip()
        record_date = cols[5].text.strip()
        payment_date = cols[6].text.strip()

        shr_class = shr_class.upper()

        if shr_class == "COMMON":
            datacommon.append([company, shr_class, dividend_type,amount,ex_date,record_date, payment_date])            
        else:
            datapref.append([company, shr_class, dividend_type,amount,ex_date,record_date, payment_date])

#convert to dataframe
df_common = pd.DataFrame(datacommon, columns=['Company', 'Class', 'Type', 'Amount', 'Ex-Date', 'Record Date', 'Payment Date']) 
df_pref = pd.DataFrame(datapref, columns=[    'Company', 'Class', 'Type', 'Amount', 'Ex-Date', 'Record Date', 'Payment Date']) 
if df_common.empty and df_pref.empty:
    print('N/A')
    exit()

#Common Dividends worksheet
sheet_common = spreadsheet.worksheet('Common Dividends')

#Preferred Dividends worksheet
sheet_pref = spreadsheet.worksheet('Preferred Dividends')

#push to google sheets
try:
    #Common
    sheet_common.clear()
    sheet_common.merge_cells('A1:G1')
    title_format = CellFormat(textFormat=TextFormat(bold=True), horizontalAlignment='CENTER', verticalAlignment='MIDDLE', backgroundColor=Color(0.792, 0.929, 0.984))
    format_cell_range(sheet_common, 'A1:G1', title_format)

    header_format = CellFormat(    textFormat=TextFormat(bold=True), backgroundColor=Color(0.792, 0.929, 0.984)) #cornflower blue 3
    hearders = df_common.columns.tolist()
    sheet_common.update(values=[['Common Dividends']], range_name='A1')
    sheet_common.update(values=[hearders], range_name='A2')
    sheet_common.update('A3', df_common.values.tolist())

    #Preferred
    sheet_pref.clear()
    sheet_pref.merge_cells('A1:G1')
    sheet_pref.update(values=[['Preferred Dividends']],range_name='A1')
    sheet_pref.update(values=[hearders], range_name='A2')
    sheet_pref.update('A3', df_pref.values.tolist())

    format_cell_range(sheet_common and sheet_pref, 'A2:G2', header_format)
    format_cell_range(sheet_pref, 'A2:G2', header_format)
    
    set_frozen(sheet_common, rows=2)

    print('Success! Data updated in Google Sheet.')

except Exception as e:
    print('Failed to update sheet: ',e)
    exit(1)