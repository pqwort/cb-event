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
# 從 GitHub Secrets 讀取
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
CALENDAR_ID = os.getenv("CALENDAR_ID")

# 本地測試用 (如果本地有檔案)
SERVICE_ACCOUNT_FILE = 'credentials.json'

# ================= Google Calendar 核心功能 =================
def get_calendar_service():
    scopes = ['https://www.googleapis.com/auth/calendar']
    creds = None

    if GOOGLE_CREDENTIALS_JSON:
        print("[Info] 使用環境變數中的憑證")
        try:
            info = json.loads(GOOGLE_CREDENTIALS_JSON)
            creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        except:
            # 嘗試 Base64 解碼 (防止 Secret 格式問題)
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
        print("[Error] 找不到 Google 憑證，請設定 Secrets")
        sys.exit(1)

    return build('calendar', 'v3', credentials=creds)

def add_event_to_calendar(service, calendar_id, data):
    """
    寫入日曆事件，並設定通知
    """
    # 標題範例：💰 6904 伯鑫 代收價款 (競價拍賣)
    icon = "💰" if "競" in data['method'] else "⭕"
    summary = f"{icon} {data['code']} {data['name']} 代收價款 ({data['method'].split()[0]})"
    
    description = (
        f"【發行資訊】\n"
        f"• 發行方式：{data['method']}\n"
        f"• 轉換溢價：{data.get('premium', '-')}\n"
        f"• 發行總額：{data.get('amount', '-')} 億\n"
        f"• 主辦券商：{data.get('underwriter', '-')}\n"
        f"• 發行年期：{data.get('duration', '-')}\n"
        f"• 賣回條件：{data.get('put', '-')}\n"
        f"• 擔保狀況：{data.get('tcri', '-')}\n\n"
        f"【公告內容】\n{data['subject']}\n\n"
        f"來源：公開資訊觀測站 & 統一證券"
    )
    
    # 設定時間：預設為「公告當日」的全天事件
    # 格式轉為 YYYY-MM-DD
    event_date = data['date'].replace('/', '-')
    
    # 唯一識別碼 (防止重複寫入)
    unique_key = f"mops_cb_{data['code']}_{event_date.replace('-', '')}"
    
    print(f"   [Check] 檢查事件: {unique_key}")
    
    # 檢查是否已存在
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
            'date': event_date, # 全天事件
        },
        'end': {
            'date': event_date, # 全天事件 (Google API 若 start=end 則為當天)
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'popup', 'minutes': 30}, # 30分鐘前通知 (對全天事件來說通常是前一天或當天9點)
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

# ================= 爬蟲邏輯 (MOPS + PSCNET) =================

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
    time.sleep(5)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    tables = soup.find_all('table')
    psc_db = {}
    
    if tables:
        for table in tables:
            try: headers = [th.get_text(strip=True) for th in table.find_all('tr')[0].find_all('th')]
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
                
                premium_text = col_texts[col_idx['premium']] if col_idx['premium']!=-1 and len(cols)>col_idx['premium'] else "-"
                
                # 溢價率判斷邏輯
                if method == "未知" and premium_text != "-":
                    if parse_premium_value(premium_text) > 105:
                        method = "⭕ 詢價圈購 (溢價率>105%)"

                # 抓取代號
                code_match = re.search(r'\d{4}', row_text)
                if code_match:
                    possible = re.findall(r'\d{4}', row_text)
                    stock_code = None
                    for c in possible:
                        if not c.startswith("202"): stock_code = c; break
                    
                    if stock_code and method != "未知":
                        psc_db[stock_code] = {
                            "method": method,
                            "premium": premium_text,
                            "amount": col_texts[col_idx['amount']] if col_idx['amount']!=-1 else "-",
                            "underwriter": col_texts[col_idx['underwriter']] if col_idx['underwriter']!=-1 else "-",
                            "put": col_texts[col_idx['put']] if col_idx['put']!=-1 else "-",
                            "duration": col_texts[col_idx['duration']] if col_idx['duration']!=-1 else "-",
                            "tcri": col_texts[col_idx['tcri']] if col_idx['tcri']!=-1 else "-"
                        }
    print(f"   統一證券資料庫建立完成: {len(psc_db)} 筆")
    return psc_db

def fetch_and_process_mops(driver, psc_db):
    print("Step 2: 抓取 MOPS 當日公告...")
    
    # 自動取得「今天」日期
    now = datetime.datetime.now()
    # 轉民國年
    target_year = str(now.year - 1911)
    target_month = str(now.month)
    target_day = str(now.day).zfill(2)
    
    # ★★★ 測試用：若要在今天(非交易日)測試，可暫時解開下面這行 ★★★
    # target_year, target_month, target_day = "114", "12", "04"
    
    url = f"https://mopsplus.twse.com.tw/mops/#/web/t05st02?year={target_year}&month={target_month}&day={target_day}"
    print(f"   Target URL: {url}")
    
    driver.get(url)
    time.sleep(5)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    rows = soup.find_all("tr")
    
    results = []
    found_any = False
    
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
                found_any = True
                info = psc_db.get(code, {})
                subject = row_text.split("公告")[1] if "公告" in row_text else row_text
                
                data = {
                    'code': code,
                    'name': name,
                    'subject': subject.strip(),
                    'date': f"{int(target_year)+1911}/{target_month}/{target_day}",
                    'method': info.get('method', "❓ 未知"),
                    'premium': info.get('premium', "-"),
                    'amount': info.get('amount', "-"),
                    'underwriter': info.get('underwriter', "-"),
                    'put': info.get('put', "-"),
                    'duration': info.get('duration', "-"),
                    'tcri': info.get('tcri', "-")
                }
                results.append(data)
                print(f"   Found Target: {code} {name}")

    if not found_any:
        print("   ⚠️ 本日無代收價款公告。")
    
    return results

# ================= 主程式 =================
def main():
    if not CALENDAR_ID:
        print("[Error] 請先在 GitHub Secrets 設定 CALENDAR_ID")
        return

    # 1. 初始化 Selenium (Headless 模式)
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") # 關鍵：GitHub Actions 無法顯示視窗
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        # 2. 爬蟲流程
        db = get_pscnet_detailed_database(driver)
        final_data = fetch_and_process_mops(driver, db)
        
        # 3. 寫入日曆
        if final_data:
            print(f"Step 3: 寫入 Google 日曆 ({len(final_data)} 筆)...")
            service = get_calendar_service()
            for item in final_data:
                add_event_to_calendar(service, CALENDAR_ID, item)
        else:
            print("今日無資料需寫入。")
            
    except Exception as e:
        print(f"❌ 發生錯誤: {e}")
        sys.exit(1) # 回傳錯誤碼讓 GitHub Actions 知道失敗
    finally:
        driver.quit()

if __name__ == '__main__':
    main()