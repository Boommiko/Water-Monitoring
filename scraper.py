import pandas as pd
import requests
from bs4 import BeautifulSoup
import json
import urllib3

# ปิดการเตือนเรื่อง SSL Certificate
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

print("1. Loading station coordinates from Excel...")
df_stations = pd.read_excel('Runoff_Station_2026.xls')

# แปลงให้เป็น String และตัดช่องว่างออก
df_stations['Code_Clean'] = df_stations['Code'].astype(str).str.strip().str.upper()

# แปลง DataFrame พิกัดเป็น Dictionary
stations_dict = {}
for _, row in df_stations.iterrows():
    code = row['Code_Clean']
    try:
        lat = float(row['Lat']) if pd.notnull(row['Lat']) else None
        lng = float(row['Long']) if pd.notnull(row['Long']) else None
    except:
        lat, lng = None, None

    stations_dict[code] = {
        "code": str(row['Code']).strip(),
        "name": str(row['Detail']).strip(),
        "river": str(row['River']).strip(),
        "province": str(row['Province']).strip(),
        "region": str(row['Region']).strip(),
        "basin": str(row['Basin']).strip(),
        "lat": lat,
        "lng": lng
    }

urls = [f"https://hyd-app-db.rid.go.th/hydro{i}hd_admsl.html" for i in range(1, 9)]
merged_results = []

print("2. Fetching real-time water data from 8 centers...")
for url in urls:
    try:
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15, verify=False)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        
        table = soup.find('table')
        if not table:
            continue
            
        rows = table.find_all('tr')
        for row in rows[1:]:
            cols = [td.get_text(strip=True) for td in row.find_all(['td', 'th'])]
            
            if len(cols) >= 5:
                st_code_raw = cols[1].strip()
                st_code_clean = st_code_raw.upper()
                
                bank_level = cols[3].strip()
                water_level = cols[4].strip()
                discharge = cols[5].strip() if len(cols) > 5 else "-"
                
                # ค้นหาใน Dictionary แบบยืดหยุ่น
                st_info = stations_dict.get(st_code_clean, {})
                
                lat = st_info.get("lat")
                lng = st_info.get("lng")
                
                # คำนวณสถานะความเสี่ยง
                status = "normal"
                try:
                    wl = float(water_level)
                    bl = float(bank_level)
                    if wl >= bl:
                        status = "danger"
                    elif wl >= (bl - 0.5):
                        status = "warning"
                except:
                    pass

                station_data = {
                    "code": st_code_raw,
                    "name": st_info.get("name", cols[2]),
                    "river": st_info.get("river", "-"),
                    "province": st_info.get("province", "-"),
                    "region": st_info.get("region", "-"),
                    "basin": st_info.get("basin", "-"),
                    "lat": lat,
                    "lng": lng,
                    "bank_level": bank_level,
                    "water_level": water_level,
                    "discharge": discharge,
                    "status": status
                }
                
                # เก็บข้อมูลเฉพาะสถานีที่มีพิกัด Latitude/Longitude เท่านั้น
                if lat is not None and lng is not None:
                    merged_results.append(station_data)
                    
    except Exception as e:
        print(f"Error fetching {url}: {e}")

print(f"3. Total valid stations with map coordinates: {len(merged_results)}")

with open('hyd_merged_data.json', 'w', encoding='utf-8') as f:
    json.dump(merged_results, f, ensure_ascii=False, indent=2)

print("Done! Saved to hyd_merged_data.json")