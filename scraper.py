import pandas as pd
import requests
from bs4 import BeautifulSoup
import json

# 1. โหลดข้อมูลพิกัดจากไฟล์ Excel
print("Loading station coordinates...")
df_stations = pd.read_excel('Runoff_Station_2026.xls')

# ทำความสะอาดข้อมูลรหัสสถานีเพื่อใช้เป็น Key
df_stations['Code_Clean'] = df_stations['Code'].astype(str).str.strip()

# แปลง DataFrame พิกัดเป็น Dictionary
stations_dict = {}
for _, row in df_stations.iterrows():
    code = row['Code_Clean']
    stations_dict[code] = {
        "code": code,
        "name": str(row['Detail']).strip(),
        "river": str(row['River']).strip(),
        "amphoe": str(row['Amphoe']).strip(),
        "province": str(row['Province']).strip(),
        "region": str(row['Region']).strip(),
        "basin": str(row['Basin']).strip(),
        "lat": float(row['Lat']) if pd.notnull(row['Lat']) else None,
        "lng": float(row['Long']) if pd.notnull(row['Long']) else None
    }

# 2. รายชื่อ URL ทั้ง 8 ศูนย์
urls = [f"https://hyd-app-db.rid.go.th/hydro{i}hd_admsl.html" for i in range(1, 9)]

merged_results = []

# 3. วนลูปดึงข้อมูลจากแต่ละศูนย์
for url in urls:
    print(f"Fetching data from: {url}")
    try:
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15, verify=False)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # ดึงข้อมูลจากตาราง HTML
        table = soup.find('table')
        if not table:
            continue
            
        rows = table.find_all('tr')
        for row in rows[1:]: # ข้าม Header
            cols = [td.get_text(strip=True) for td in row.find_all(['td', 'th'])]
            
            # โครงสร้างตารางหน้าเว็บโดยประมาณ: [ลำดับ, รหัสสถานี, ชื่อสถานี, ระดับตลิ่ง, ระดับน้ำ, อัตราไหล...]
            if len(cols) >= 5:
                st_code = cols[1].strip()
                bank_level = cols[3].strip() # ระดับตลิ่ง
                water_level = cols[4].strip() # ระดับน้ำปัจจุบัน
                discharge = cols[5].strip() if len(cols) > 5 else "-" # อัตราการไหล (Q)
                
                # แมตช์กับข้อมูลพิกัดใน Excel
                st_info = stations_dict.get(st_code, {})
                
                # คำนวณสถานะความเสี่ยง (ปกติ / เฝ้าระวัง / ล้นตลิ่ง)
                status = "normal"
                try:
                    wl = float(water_level)
                    bl = float(bank_level)
                    if wl >= bl:
                        status = "danger" # ล้นตลิ่ง
                    elif wl >= (bl - 0.5):
                        status = "warning" # เฝ้าระวัง (น้อยกว่าตลิ่งไม่เกิน 0.5ม.)
                except:
                    pass

                station_data = {
                    "code": st_code,
                    "name": st_info.get("name", cols[2]),
                    "river": st_info.get("river", "-"),
                    "province": st_info.get("province", "-"),
                    "region": st_info.get("region", "-"),
                    "basin": st_info.get("basin", "-"),
                    "lat": st_info.get("lat"),
                    "lng": st_info.get("lng"),
                    "bank_level": bank_level,
                    "water_level": water_level,
                    "discharge": discharge,
                    "status": status
                }
                
                # เก็บเฉพาะสถานีที่มีพิกัดถูกต้อง
                if station_data["lat"] and station_data["lng"]:
                    merged_results.append(station_data)
                    
    except Exception as e:
        print(f"Error fetching {url}: {e}")

# 4. บันทึกผลลัพธ์เป็นไฟล์ JSON สำหรับหน้าเว็บ
with open('hyd_merged_data.json', 'w', encoding='utf-8') as f:
    json.dump(merged_results, f, ensure_ascii=False, indent=2)

print(f"Successfully processed {len(merged_results)} stations into hyd_merged_data.json!")