import os
import sys
import requests
import json
import traceback
from datetime import datetime, timedelta
from scrapers import attendance_scraper
from utils.point_calculator import calculate_attendance_points

# 強制設定輸出為 UTF-8 以避免亂碼
if sys.stdout.encoding != 'utf-8':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except:
        pass

# ================= 設定區 =================
# 您的 GAS Web App URL
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyM4ys5HiOSdgZjxbb2Oj_ScBR_hU2yCRpQfryoDsNvyRTx2pzNmvEe8tMvLr9XfEGZrQ/exec" 

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "data", "學生出缺席_總表.xlsx")
# =========================================

def get_last_week_range():
    """自動取得上週一到週五的民國日期字串"""
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    last_friday = last_monday + timedelta(days=4)
    
    def to_minguo(dt):
        return f"{dt.year - 1911}-{dt.month:02d}-{dt.day:02d}"
    
    return to_minguo(last_monday), to_minguo(last_friday)

def main():
    try:
        print("=== 學生積分管理系統 (一鍵發分) ===")
        
        # 1. 爬蟲抓取
        choice = input("是否抓取最新出缺席資料? (y/n): ")
        if choice.lower() == 'y':
            print("🚀 啟動爬蟲...")
            attendance_scraper.MASTER_FILE = EXCEL_PATH
            attendance_scraper.main()

        # 2. 計算點數
        start, end = get_last_week_range()
        print(f"\n📅 自動設定區間為上週：{start} ~ {end}")
        use_custom = input("是否手動輸入日期區間? (y/n): ")
        if use_custom.lower() == 'y':
            start = input("請輸入開始日期 (如 114-02-23): ")
            end = input("請輸入結束日期 (如 114-02-27): ")

        points_map = calculate_attendance_points(EXCEL_PATH, start_date=start, end_date=end)
        
        if not points_map:
            print("ℹ️ 該區間無打卡資料。")
            input("按 Enter 結束...")
            return

        # 3. 預覽發分
        reason = f"{start}~{end} 每日打卡獎勵"
        updates = []
        print(f"\n[ 發分預覽 - {reason} ]")
        print("-" * 30)
        for seat, count in sorted(points_map.items()):
            pts = count * 5
            updates.append({"seatNo": seat, "points": pts})
            print(f"座號 {seat:02d}: 打卡 {count} 次 -> 發放 {pts} 分")
        print("-" * 30)

        # 4. 同步至雲端
        confirm = input(f"\n確認要發分給這 {len(updates)} 位學生嗎? (y/n): ")
        if confirm.lower() == 'y':
            print("🚀 正在發送到雲端積分銀行...")
            try:
                payload = {
                    "updates": updates,
                    "reason": reason
                }
                response = requests.post(WEB_APP_URL, data=json.dumps(payload), timeout=30)
                res_data = response.json()
                if res_data.get("success"):
                    print(f"✅ 成功！已更新 {res_data.get('updated')} 位同學的積分。")
                else:
                    print(f"❌ 失敗：{res_data.get('error')}")
            except Exception as e:
                print(f"❌ 連線錯誤：{e}")
        
        input("\n任務完成，按 Enter 鍵結束...")

    except Exception as e:
        print("\n❌ 程式發生嚴重錯誤！")
        print("-" * 30)
        traceback.print_exc()
        print("-" * 30)
        input("請按 Enter 鍵結束程式...")

if __name__ == "__main__":
    main()
