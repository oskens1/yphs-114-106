import pandas as pd
import os
import re

def calculate_attendance_points(excel_path, start_date=None, end_date=None):
    """
    讀取學生出缺席總表，計算每個學生的到校次數。
    start_date, end_date 格式: "114-02-23" (民國年格式，同 Excel 分頁名)
    """
    if not os.path.exists(excel_path):
        print(f"❌ 找不到檔案: {excel_path}")
        return {}

    print(f"📊 正在從 {excel_path} 計算積分...")
    if start_date or end_date:
        print(f"   📅 篩選區區間: {start_date or '不限'} ~ {end_date or '不限'}")
    
    xls = pd.ExcelFile(excel_path)
    points_map = {}

    for sheet_name in xls.sheet_names:
        # 篩選日期區間 (簡單字串比較即可，因為是 114-xx-xx 格式)
        if start_date and sheet_name < start_date:
            continue
        if end_date and sheet_name > end_date:
            continue

        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
        except:
            continue
        
        # 檢查必要的欄位
        if '年班座號' not in df.columns or '到校時間' not in df.columns:
            continue
            
        for _, row in df.iterrows():
            raw_seat = str(row['年班座號'])
            
            # 嘗試提取數字
            match = re.search(r'(\d+)', raw_seat)
            if match:
                val = int(match.group(1))
                # 如果是 10601 這種格式，取後兩位
                if val > 1000:
                    seat_no = val % 100
                else:
                    seat_no = val
            else:
                continue
            
            # 判斷是否有打卡 (到校時間不為空且不是 NaN)
            arrival_time = row['到校時間']
            if pd.notna(arrival_time) and str(arrival_time).strip() != "":
                points_map[seat_no] = points_map.get(seat_no, 0) + 1
    
    print(f"✅ 計算完成，共 {len(points_map)} 位學生有打卡紀錄。")
    return points_map

if __name__ == "__main__":
    # 測試用
    pts = calculate_attendance_points("../data/學生出缺席_總表.xlsx")
    for s, p in sorted(pts.items()):
        print(f"座號 {s}: {p} 點")
