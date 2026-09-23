import os
import json
import re
import pandas as pd
from playwright.sync_api import sync_playwright

excel_file = 'Runoff_Station_2026.xls'

# 1. ฟังก์ชันทำความสะอาดรหัสสถานี
def clean_text(text):
    if not text or pd.isna(text):
        return ""
    return re.sub(r'[\s\.\-\_\(\)]', '', str(text)).upper().strip()

# 2. โหลดข้อมูล Excel และทำ Dictionary
if not os.path.exists(excel_file):
    print(f"❌ ERROR: หาไฟล์ '{excel_file}' ไม่พบ!")
    exit()

print("1. Loading station coordinates from Excel...")
df_stations = pd.read_excel(excel_file)

code_dict = {}
detail_dict = {}

for _, row in df_stations.iterrows():
    raw_code = str(row['Code']).strip() if pd.notnull(row['Code']) else ""
    raw_detail = str(row['Detail']).strip() if pd.notnull(row['Detail']) else ""
    
    clean_code_key = clean_text(raw_code)
    clean_detail_key = clean_text(raw_detail)
    
    try:
        lat = float(row['Lat']) if pd.notnull(row['Lat']) else None
        lng = float(row['Long']) if pd.notnull(row['Long']) else None
    except:
        lat, lng = None, None

    info = {
        "code": raw_code,
        "name": raw_detail,
        "river": str(row.get('River', '-')).strip(),
        "province": str(row.get('Province', '-')).strip(),
        "region": str(row.get('Region', '-')).strip(),
        "basin": str(row.get('Basin', '-')).strip(),
        "lat": lat,
        "lng": lng
    }

    if clean_code_key:
        code_dict[clean_code_key] = info
    if clean_detail_key:
        detail_dict[clean_detail_key] = info

urls = [f"https://hyd-app-db.rid.go.th/hydro{i}hd_admsl.html" for i in range(1, 9)]
merged_results = []

print("2. Fetching & Parsing Horizontal Hydro Tables...")

scraped_count = 0
matched_coords_count = 0

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    for i, url in enumerate(urls, 1):
        try:
            print(f"   กำลังเปิดหน้าเว็บ hydro{i}...")
            page.goto(url, timeout=30000, wait_until="networkidle")
            
            rows = page.locator("table tr").all()
            all_table_data = []

            for row in rows:
                cols = row.locator("td, th").all_text_contents()
                cols = [c.strip() for c in cols]
                if any(cols): # สนใจเฉพาะแถวที่มีข้อมูล
                    all_table_data.append(cols)

            # ค้นหาแถวที่มีรหัสสถานี (เช่น P.20, P.1, Sw.5A)
            station_row_idx = -1
            for idx, r in enumerate(all_table_data):
                # แถวที่มีรหัสสถานีมักจะพบรหัสสถานีอย่างน้อย 2 ตัวขึ้นไป
                matches = sum(1 for item in r if clean_text(item) in code_dict)
                if matches >= 2:
                    station_row_idx = idx
                    break

            center_scraped = 0

            # หากพบแถวรหัสสถานี ให้แกะข้อมูลแบบแนวนอน
            if station_row_idx != -1:
                st_codes_row = all_table_data[station_row_idx]
                
                # พยายามหาแถวระดับน้ำ และ อัตราการไหล
                last_data_row = all_table_data[-1] if len(all_table_data) > station_row_idx else []
                second_last_row = all_table_data[-2] if len(all_table_data) > station_row_idx + 1 else []

                for col_idx, raw_st_code in enumerate(st_codes_row):
                    clean_code = clean_text(raw_st_code)
                    if not clean_code or clean_code in ['เวลา', 'DATE', 'TIME']:
                        continue

                    # จับคู่พิกัดจาก Excel
                    st_info = code_dict.get(clean_code)
                    
                    # ลองสกัดค่าน้ำจากคอลัมน์ที่ตรงกัน
                    water_level = "-"
                    discharge = "-"
                    bank_level = "-"

                    if col_idx < len(last_data_row):
                        water_level = last_data_row[col_idx]
                    if col_idx < len(second_last_row):
                        discharge = second_last_row[col_idx]

                    lat, lng = None, None
                    if st_info:
                        lat = st_info.get("lat")
                        lng = st_info.get("lng")
                        if lat is not None and lng is not None:
                            matched_coords_count += 1

                    status = "normal"
                    try:
                        wl = float(water_level)
                        if wl > 3.0: # สถานะตัวอย่าง
                            status = "warning"
                    except:
                        pass

                    station_data = {
                        "code": st_info.get("code") if st_info else raw_st_code,
                        "name": st_info.get("name") if st_info else raw_st_code,
                        "river": st_info.get("river", "-") if st_info else "-",
                        "province": st_info.get("province", "-") if st_info else "-",
                        "region": st_info.get("region", "-") if st_info else "-",
                        "basin": st_info.get("basin", "-") if st_info else "-",
                        "lat": lat,
                        "lng": lng,
                        "bank_level": bank_level,
                        "water_level": water_level,
                        "discharge": discharge,
                        "status": status
                    }

                    merged_results.append(station_data)
                    scraped_count += 1
                    center_scraped += 1

            print(f"   ศูนย์ hydro{i}: ดึงและแมตช์ได้ {center_scraped} สถานี")

        except Exception as e:
            print(f"   ❌ hydro{i} Error: {e}")

    browser.close()

print(f"\n3. สรุปผลการทำงาน:")
print(f"   - ดึงข้อมูลได้ทั้งหมด: {scraped_count} สถานี")
print(f"   - แมตช์พิกัดเจอทั้งหมด: {matched_coords_count} สถานี")

# บันทึกเป็นไฟล์ JSON
with open('hyd_merged_data.json', 'w', encoding='utf-8') as f:
    json.dump(merged_results, f, ensure_ascii=False, indent=2)

print("✅ บันทึกไฟล์ hyd_merged_data.json เรียบร้อยแล้ว!")