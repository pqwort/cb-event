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

# ================= 使用者設定區 (本地測試改這裡) =================

# 1. 填入您的日曆 ID (您的 Gmail)
LOCAL_CALENDAR_ID = "vip72@gmail.com"  # <--- 請改成您的 Gmail

# 2. 是否隱藏瀏覽器？ (本地測試建議 False，可以看到它在動)
HEADLESS_MODE = False 

# 3. 強制測試特定日期 (格式 "YYYY/MM/DD")
# 如果想抓「今天」，請把下面這行改成 None
TEST_DATE_OVERRIDE = "2025/12/04" 
# TEST_DATE_OVERRIDE = None

# ==============================================================

# 環境變數優先 (給 GitHub Actions 用)，若無則用上面的本地設定
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
CALENDAR_ID = os.getenv("CALENDAR_ID", LOCAL_CALENDAR_ID)
SERVICE_ACCOUNT_FILE = 'credentials.json'

# ================= Google Calendar 核心功能 =================
def get_calendar_service():
    scopes = ['https://www.googleapis.com/auth/calendar']
    creds = None

    if GOOGLE_CREDENTIALS_JSON:
        try:
            info = json.loads(GOOGLE_CREDENTIALS_JSON)
            creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        except:
            try:
                decoded = base64.b64decode(GOOGLE_CREDENTIALS_JSON).decode("utf-8")
                info = json.loads(decoded)
                creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
            except: pass
    
    if not creds and os.path.exists(SERVICE_ACCOUNT_FILE):
        print("[Info] 使用本地 credentials.json")
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    
    if not creds:
        print("[Error] 找不到憑證！請確認 credentials.json 是否在資料夾內。")
        sys.exit(1)

    return build('calendar', 'v3', credentials=creds)

def add_event_to_calendar(service, calendar_id, data):
    summary = f"💰 {data['code']} {data['name']} 代收價款 ({data['method'].split()[0]})"
    if "詢" in data['method']: summary = summary.replace("💰", "⭕")

    description = (
        f"【發行資訊】\n"
        f"• 發行方式：{data['method']}\n"
        f"• 轉換溢價：{data.get('premium', '-')}\n"
        f"• 發行總額：{data.get('amount', '-')} 億\n"
        f"• 主辦券商：{data.get('underwriter', '-')}\n"
        f"• 發行年期：{data.get('duration', '-')}\n"
        f"• 賣回條件：{data.get('put', '-')}\n"
        f"• 主旨：{data['subject']}\n"
        f"來源：MOPS 公開資訊觀測站"
    )
    
    event_date = data['date'].replace('/', '-')
    unique_key = f"mops_cb_{data['code']}_{event_date.replace('-', '')}"
    
    print(f"   [Check] 檢查事件: {unique_key}")
    try:
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
            'description': description,
            'start': {'date': event_date},
            'end': {'date': event_date},
            'reminders': {
                'useDefault': False,
                'overrides': [{'method': 'popup', 'minutes': 30}],
            },
            'extendedProperties': {'private': {'uniqueID': unique_key}}
        }

        event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        print(f"   [Success] 事件已建立: {event.get('htmlLink')}")
    except Exception as e:
        print(f"   [Error] 寫入失敗 (請確認日曆 ID 正確且已共用權限): {e}")

# ================= 爬蟲邏輯 =================

def parse_premium_value(text):
    try:
        clean_text = text.replace('%', '').strip()
        first_val = re.split(r'[~-]', clean_text)[0]
        match = re.search(r'\d+(\.\d+)?', first_val)
        if match: return float(match.group(0))
    except: pass
    return 0.0

def get_pscnet_detailed_database(driver):
    print("Step 1: 前往統一證券抓取詳細資料...")
    url = "https://cbas16889.pscnet.com.tw/marketInfo/expectedRelease/"
    driver.get(url)
    time.sleep(8) # 等待網頁載入
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    tables = soup.find_all('table')
    psc_db = {}
    
    if not tables:
        print("❌ 警告：找不到任何表格！")
        return {}

    print(f"   共發現 {len(tables)} 個表格，開始掃描...")
    
    for table in tables:
        headers = []
        try: headers = [th.get_text(strip=True) for th in table.find_all('tr')[0].find_all('th')]
        except: pass
        
        col_idx = {'underwriter': -1, 'amount': -1, 'put': -1, 'duration': -1, 'premium': -1, 'tcri': -1}
        for i, h in enumerate(headers):
            if "主辦" in h: col_idx['underwriter'] = i
            elif "發行量" in h: col_idx['amount'] = i
            elif "賣回" in h: col_idx['put'] = i
            elif "年期" in h: col_idx['duration'] = i
            elif "溢價率" in h: col_idx['premium'] = i
            elif "TCRI" in h or "擔保" in h: col_idx['tcri'] = i

        rows = table.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            if not cols: continue
            row_text = row.get_text()
            col_texts = [c.get_text(strip=True) for c in cols]
            
            def safe_get(idx): return col_texts[idx] if idx != -1 and idx < len(col_texts) else "-"

            method = "未知"
            if "競拍" in row_text or "競價" in row_text: method = "💰 競價拍賣"
            elif "詢圈" in row_text or "詢價" in row_text: method = "⭕ 詢價圈購"
            
            premium_text = safe_get(col_idx['premium'])
            if method == "未知" and premium_text != "-":
                if parse_premium_value(premium_text) > 105:
                    method = "⭕ 詢價圈購 (溢價率>105%)"

            possible_codes = re.findall(r'\d{4}', row_text)
            stock_code = None
            for c in possible_codes:
                if not c.startswith("202"): stock_code = c; break
            
            if stock_code:
                psc_db[stock_code] = {
                    "method": method,
                    "premium": premium_text,
                    "amount": safe_get(col_idx['amount']),
                    "underwriter": safe_get(col_idx['underwriter']),
                    "put": safe_get(col_idx['put']),
                    "duration": safe_get(col_idx['duration']),
                    "tcri": safe_get(col_idx['tcri'])
                }

    print(f"✅ 統一證券資料庫建立完成: {len(psc_db)} 筆")
    return psc_db

def fetch_and_process_mops(driver, psc_db):
    print("Step 2: 抓取 MOPS 公告...")
    
    if TEST_DATE_OVERRIDE:
        print(f"   [測試模式] 強制使用日期: {TEST_DATE_OVERRIDE}")
        dt = datetime.datetime.strptime(TEST_DATE_OVERRIDE, "%Y/%m/%d")
        year = str(dt.year - 1911)
        month = str(dt.month)
        day = str(dt.day).zfill(2)
    else:
        now = datetime.datetime.now()
        year = str(now.year - 1911)
        month = str(now.month)
        day = str(now.day).zfill(2)
    
    url = f"https://mopsplus.twse.com.tw/mops/#/web/t05st02?year={year}&month={month}&day={day}"
    print(f"   Target URL: {url}")
    
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
                
                data = {
                    'code': code,
                    'name': name,
                    'subject': subject.strip(),
                    'date': f"{int(year)+1911}/{month}/{day}",
                    'method': info.get('method', "❓ 未知"),
                    'premium': info.get('premium', "-"),
                    'amount': info.get('amount', "-"),
                    'underwriter': info.get('underwriter', "-"),
                    'put': info.get('put', "-"),
                    'duration': info.get('duration', "-"),
                    'tcri': info.get('tcri', "-")
                }
                results.append(data)
                print(f"   🎯 發現目標: {code} {name}")

    if not results:
        print("   ⚠️ 無相關公告。")
    
    return results

# ================= 主程式 =================
def main():
    if not CALENDAR_ID:
        print("[Error] 請在程式最上方填入您的 Gmail (LOCAL_CALENDAR_ID)")
        return

    # 初始化 Selenium
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1920,1080") 
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    if HEADLESS_MODE:
        options.add_argument("--headless")

    print(f"🚀 啟動爬蟲 (Headless: {HEADLESS_MODE})")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        db = get_pscnet_detailed_database(driver)
        final_data = fetch_and_process_mops(driver, db)
        
        if final_data:
            print(f"Step 3: 寫入 Google 日曆 ({len(final_data)} 筆)...")
            service = get_calendar_service()
            for item in final_data:
                add_event_to_calendar(service, CALENDAR_ID, item)
        else:
            print("無資料需寫入。")
            
    except Exception as e:
        print(f"❌ 發生錯誤: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("關閉瀏覽器...")
        driver.quit()

if __name__ == '__main__':
    main()