from DrissionPage import ChromiumPage
from DrissionPage.common import Keys
import pandas as pd
import time
import re
import os
from datetime import datetime

# ================= 使用者設定區 =================
USER_ID = 'oskens'          
USER_PW = 'AAaa4652897'        # <--- 密碼記得改
TARGET_CLASS = '導師 一年六班' 

# 核心資料庫檔名
MASTER_FILE = "學生出缺席_總表.xlsx"

# 設定「檢查截止日」
START_DATE = "2025-01-04"   
# ==============================================

def get_clean_time(ele):
    try:
        text = ele.text.strip()
        match = re.search(r'\d{2}:\d{2}', text)
        return match.group(0) if match else ""
    except:
        return ""

def get_visible_status(ele):
    try:
        visible_texts = [span.text for span in ele.eles('tag:span') if 'ng-hide' not in span.attr('class')]
        return " ".join(visible_texts)
    except:
        return ""

def click_by_js_safe(page, ele_text):
    try:
        btn = page.ele(f'text:{ele_text}', timeout=5)
        if btn:
            btn.click(by_js=True)
            time.sleep(1)
            return True
        return False
    except:
        return False

def auto_login_and_navigate(page):
    print("🤖 1. 啟動登入與導航...")
    page.get('https://esa.ntpc.edu.tw/central/theme/01/index.html')
    
    if page.ele('text:登入(Login)'):
        page.ele('text:登入(Login)').click()
        time.sleep(1)

    if page.ele('@name=username'):
        page.ele('@name=username').input(USER_ID)
        page.ele('@name=password').input(USER_PW)
        if page.ele('#btn-submit'):
            page.ele('#btn-submit').click()
            time.sleep(3) 

    page.actions.key_down(Keys.ESCAPE).key_up(Keys.ESCAPE)
    time.sleep(0.5)
    if page.ele('text:確定', timeout=2):
        page.ele('text:確定').click(by_js=True)
        time.sleep(1)

    click_by_js_safe(page, TARGET_CLASS)
    if page.ele(f'text:{TARGET_CLASS}', timeout=2):
        click_by_js_safe(page, TARGET_CLASS)
        time.sleep(3)

    if not click_by_js_safe(page, '【新】學生出缺席'):
        return False
    time.sleep(2)

    if not click_by_js_safe(page, '學生到離校管理'):
        return False
    
    print("   ⏳ 等待資料表載入...")
    if page.ele('xpath://tr[contains(@ng-repeat, "rfidvm.list")]', timeout=10):
        print("   ✅ 成功抵達資料頁！")
        return True
    return True

def get_current_page_date(page):
    try:
        val = page.run_js('return document.querySelector(".md-datepicker-input").value')
        return val.strip() 
    except:
        return ""

def scrape_current_page(page, date_str):
    rows = page.eles('xpath://tr[contains(@ng-repeat, "rfidvm.list")]')
    if not rows:
        print(f"   ⚠️ {date_str} 無資料。")
        return []

    print(f"   📥 {date_str} 抓取中... ({len(rows)} 筆)")
    
    daily_data = []
    for row in rows:
        try:
            cols = row.eles('tag:td') 
            student_data = {
                '日期': date_str,
                '年班座號': cols[1].text,
                '姓名': cols[2].text,
                '到校時間': get_clean_time(cols[3]),
                '離校時間': get_clean_time(cols[6]), 
                '狀態註記': f"{get_visible_status(cols[4])} {get_visible_status(cols[6])}".strip()
            }
            daily_data.append(student_data)
        except:
            continue
    return daily_data

def go_to_prev_day(page):
    try:
        datepicker = page.ele('@ng-model=rfidvm.sdate')
        if datepicker:
            prev_btn = datepicker.prev()
            if prev_btn.tag != 'button':
                btn_inside = prev_btn.ele('tag:button')
                if btn_inside: prev_btn = btn_inside
            prev_btn.click(by_js=True)
            time.sleep(2) 
            return True
        return False
    except:
        return False

def get_existing_sheets(filename):
    """讀取 Excel 看看裡面已經有哪些日期的分頁"""
    if not os.path.exists(filename):
        return [] 
    
    try:
        xls = pd.ExcelFile(filename)
        return xls.sheet_names 
    except Exception as e:
        print(f"⚠️ 讀取舊檔失敗: {e}")
        return []

def main():
    existing_sheets = get_existing_sheets(MASTER_FILE)
    print(f"\n📂 讀取總表: {MASTER_FILE}")
    print(f"   已存在日期 ({len(existing_sheets)} 天): {existing_sheets}")

    page = ChromiumPage()
    if not auto_login_and_navigate(page):
        return

    data_book = {} 
    
    start_dt = datetime.strptime(START_DATE, "%Y-%m-%d")
    target_minguo_start = f"{start_dt.year - 1911}-{start_dt.month:02d}-{start_dt.day:02d}"

    print(f"\n🚀 開始檢查與補漏！(截止日: {target_minguo_start})\n")

    max_days = 60 
    count = 0

    while count < max_days:
        current_date_str = get_current_page_date(page)
        if not current_date_str:
            print("❌ 日期讀取失敗，停止。")
            break
            
        if current_date_str < target_minguo_start:
            print(f"🛑 日期 {current_date_str} 已早於設定起點，任務完成！")
            break

        if current_date_str in existing_sheets:
            print(f"🛑 發現已存在日期 {current_date_str}，資料已成功銜接，停止抓取。")
            break
        else:
            daily_data = scrape_current_page(page, current_date_str)
            if daily_data:
                data_book[current_date_str] = daily_data 

        if not go_to_prev_day(page):
            print("❌ 切換失敗，停止。")
            break
            
        count += 1
        if current_date_str in existing_sheets:
            time.sleep(0.5) 

    # ================= 存檔邏輯 (已修復錯誤) =================
    if data_book:
        print(f"\n💾 正在寫入 {len(data_book)} 天的新資料到總表...")
        
        # 判斷檔案是否存在，決定使用哪種模式
        if os.path.exists(MASTER_FILE):
            # 舊檔案存在 -> 使用 append 模式，並設定 if_sheet_exists
            writer = pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace')
        else:
            # 新檔案 -> 使用 write 模式 (不可加 if_sheet_exists)
            writer = pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w')
            
        with writer:
            for date_key, data_list in data_book.items():
                df = pd.DataFrame(data_list)
                df = df[['日期', '年班座號', '姓名', '到校時間', '離校時間', '狀態註記']]
                df.to_excel(writer, sheet_name=date_key, index=False)
                print(f"   -> 已存入分頁: {date_key}")
                
        print("\n" + "="*50)
        print(f"🎉 更新完成！總表已更新: {MASTER_FILE}")
        print("="*50)
    else:
        print("\n🎉 沒有發現新資料，總表已經是最新的了。")

if __name__ == '__main__':
    main()