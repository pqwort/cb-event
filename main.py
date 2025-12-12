import os
import sys
import json
import time
import datetime
import re
import base64
from bs4 import BeautifulSoup

# Google API
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ================= 設定區 =================
# 從環境變數讀取憑證 (GitHub Secrets)
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
# 從環境變數讀取日曆 ID，若無則報錯
CALENDAR_ID = os.getenv("CALENDAR_ID")

# 若沒有環境變數，嘗試讀取本地檔案 (測試用)
SERVICE_ACCOUNT_FILE = 'credentials.json'

# ================= Google Calendar 工具函式 =================
def get_calendar_service():
    scopes = ['https://www.googleapis.com/auth/calendar']
    creds = None

    if GOOGLE_CREDENTIALS_JSON:
        print("[Info] 使用環境變數中的憑證")
        try:
            # 嘗試直接解析 JSON
            info = json.loads(GOOGLE_CREDENTIALS_JSON)
            creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        except json.JSONDecodeError:
            # 若失敗，嘗試 Base64 解碼 (有時候 Secret 會存成 Base64)
            try:
                decoded = base64.b64decode(GOOGLE_CREDENTIALS_JSON).decode("utf-8")
                info = json.loads(decoded)
                creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
            except Exception as e:
                print(f"[Error] 憑證解析失敗: {e}")
                sys.exit(1)
    elif os.path.exists(SERVICE_ACCOUNT_FILE):
        print("[Info] 使用本地 credentials.json")
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    else:
        print("[Error] 找不到 Google 憑證，請設定 GOOGLE_CREDENTIALS_JSON 環境變數")
        sys.exit(1)

    return build('calendar', 'v3', credentials=creds)

def add_event_to_calendar(service, calendar_id, data):
    """
    寫入日曆事件
    data 結構: {'code': 'xxx', 'name': 'xxx', 'method': 'xxx', 'date': 'YYYY/MM/DD', 'subject': 'xxx', ...}
    """
    summary = f"📢 {data['code']} {data['name']} 代收價款公告 ({data['method']})"
    description = (
        f"發行方式：{data['method']}\n"
        f"主辦券商：{data.get('underwriter', '-')}\n"
        f"發行總額：{data.get('amount', '-')} 億\n"
        f"溢價率：{data.get('premium', '-')}\n"
        f"賣回條件：{data.get('put', '-')}\n"
        f"主旨：{data['subject']}\n"
        f"來源：MOPS 公開資訊觀測站"
    )
    
    # 設定時間：預設為當天全天事件，或設定在隔天早上 09:00 提醒
    # 這裡示範設定為「公告日期的隔天早上 09:00」
    announce_date = datetime.datetime.strptime(data['date'], "%Y/%m/%d").date()
    event_date = announce_date + datetime.timedelta(days=1)
    
    start_time = datetime.datetime.combine(event_date, datetime.time(9, 0)).isoformat()
    end_time = datetime.datetime.combine(event_date, datetime.time(9, 30)).isoformat()

    # 檢查重複 (利用 code 作為 unique key)
    # 使用 private extended property 來標記
    unique_key = f"mops_cb_{data['code']}_{data['date'].replace('/', '')}"
    
    print(f"   [Check] 檢查事件是否存在: {unique_key}")
    events_result = service.events().list(
        calendarId=calendar_id,
        privateExtendedProperty=f"uniqueID={unique_key}",
        singleEvents=True
    ).execute()
    
    if events_result.get('items'):
        print(f"   [Skip] 事件已存在，跳過。")
        return

    event_body = {
        'summary': summary,
        'location': '公開資訊觀測站',
        'description': description,
        'start': {
            'dateTime': start_time,
            'timeZone': 'Asia/Taipei',
        },
        'end': {
            'dateTime': end_time,
            'timeZone': 'Asia/Taipei',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 10}, # 10分鐘前通知
            ],
        },
        'extendedProperties': {
            'private': {
                'uniqueID': unique_key
            }
        }
    }

    try:
        event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        print(f"   [Success] 事件已建立: {event.get('htmlLink')}")
    except HttpError as e:
        print(f"   [Error] 寫入失敗: {e}")

# ================= 爬蟲邏輯 (從之前的代碼整合) =================

def parse_premium_value(text):
    try:
        clean_text = text.replace('%', '').strip()
        first_val = re.split(r'[~-]', clean_text)[0]
        match = re.search(r'\d+(\.\d+)?', first_val)
        if match:
            return float(match.group(0))
    except:
        pass
    return 0.0

def get_pscnet_db(driver):
    print("Step 1: 爬取統一證券資料庫...")
    url = "https://cbas16889.pscnet.com.tw/marketInfo/expectedRelease/"
    driver.get(url)
    time.sleep(5)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    tables = soup.find_all('table')
    psc_db = {}
    
    if tables:
        for table in tables:
            try:
                headers = [th.get_text(strip=True) for th in table.find_all('tr')[0].find_all('th')]
            except: continue

            col_idx = {'underwriter': -1, 'amount': -1, 'put': -1, 'duration': -1, 'premium': -1, 'tcri': -1}
            for i, h in enumerate(headers):
                if "主辦" in h: col_idx['underwriter'] = i
                elif "發行量" in h: col_idx['amount'] = i
                elif "賣回" in h: col_idx['put'] = i
                elif "年期" in h: col_idx['duration'] = i
                elif "溢價率" in h: col_idx['premium'] = i
                elif "TCRI" in h or "擔保" in h: col_idx['tcri'] = i

            rows = table.find_all('tr')[1:]
            for row in rows:
                cols = row.find_all('td')
                if not cols: continue
                row_text = row.get_text()
                col_texts = [c.get_text(strip=True) for c in cols]
                
                method = "未知"
                if "競拍" in row_text or "競價" in row_text: method = "💰 競價拍賣"
                elif "詢圈" in row_text or "詢價" in row_text: method = "⭕ 詢價圈購"
                
                premium_text = col_texts[col_idx['premium']] if col_idx['premium'] != -1 and len(cols) > col_idx['premium'] else "-"
                
                if method == "未知" and premium_text != "-":
                    if parse_premium_value(premium_text) > 105:
                        method = "⭕ 詢價圈購 (溢價率>105%)"

                code_match = re.search(r'\d{4}', row_text)
                if code_match:
                    possible = re.findall(r'\d{4}', row_text)
                    stock_code = None
                    for c in possible:
                        if not c.startswith("202"):
                            stock_code = c
                            break
                    
                    if stock_code and method != "未知":
                        psc_db[stock_code] = {
                            "method": method,
                            "premium": premium_text,
                            "amount": col_texts[col_idx['amount']] if col_idx['amount']!=-1 else "-",
                            "underwriter": col_texts[col_idx['underwriter']] if col_idx['underwriter']!=-1 else "-",
                            "put": col_texts[col_idx['put']] if col_idx['put']!=-1 else "-",
                        }
    return psc_db

def fetch_mops_data(driver, psc_db):
    print("Step 2: 爬取 MOPS 當日公告...")
    now = datetime.datetime.now()
    # 轉換為民國年
    year = str(now.year - 1911)
    month = str(now.month)
    day = str(now.day).zfill(2)
    
    # 測試用：強制指定有資料的日期 (正式使用請註解掉這三行)
    # year, month, day = "114", "12", "04"
    
    url = f"https://mopsplus.twse.com.tw/mops/#/web/t05st02?year={year}&month={month}&day={day}"
    print(f"Target: {year}/{month}/{day}")
    
    driver.get(url)
    time.sleep(5)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    rows = soup.find_all("tr")
    
    results = []
    
    for row in rows:
        row_text = row.get_text()
        if "代收價款" in row_text:
            cols = [c.text.strip() for c in row.find_all('td') if c.text.strip()]
            
            code = "N/A"
            name = "N/A"
            for col in cols:
                if col.isdigit() and len(col) == 4:
                    code = col
                    try: name = cols[cols.index(code) + 1]
                    except: pass
                    break
            
            if code != "N/A":
                info = psc_db.get(code, {})
                subject = row_text.split("公告")[1] if "公告" in row_text else row_text
                
                item = {
                    'code': code,
                    'name': name,
                    'subject': subject.strip(),
                    'date': f"{int(year)+1911}/{month}/{day}", # 存西元年方便 Calendar 處理
                    'method': info.get('method', "❓ 未知"),
                    'premium': info.get('premium', "-"),
                    'amount': info.get('amount', "-"),
                    'underwriter': info.get('underwriter', "-"),
                    'put': info.get('put', "-")
                }
                results.append(item)
                print(f"   Found: {code} {name} ({item['method']})")
                
    return results

# ================= 主程式 =================
def main():
    # 0. 初始化 Calendar Service
    if not CALENDAR_ID:
        print("[Error] 未設定 CALENDAR_ID 環境變數")
        sys.exit(1)
        
    calendar_service = get_calendar_service()
    
    # 1. 初始化 Selenium
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") # 無頭模式 (GitHub Actions 必須)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        # 2. 爬取資料
        db = get_pscnet_db(driver)
        data_list = fetch_mops_data(driver, db)
        
        # 3. 寫入日曆
        if data_list:
            print(f"Step 3: 寫入 {len(data_list)} 筆資料到 Google Calendar...")
            for data in data_list:
                add_event_to_calendar(calendar_service, CALENDAR_ID, data)
        else:
            print("今日無相關公告。")
            
    except Exception as e:
        print(f"[Error] {e}")
        sys.exit(1)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()