import asyncio
import json
import os
import re
import pandas as pd
from playwright.async_api import async_playwright

EXCEL_FILE = 'Runoff_Station_2026.xls'
JSON_OUTPUT = 'hyd_merged_data.json'

def load_excel_stations():
    """1. โหลดข้อมูลพิกัดและรายละเอียดสถานีจากไฟล์ Excel"""
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ ERROR: ไม่พบไฟล์ '{EXCEL_FILE}' ในโฟลเดอร์นี้")
        return {}

    print("1. Loading station coordinates from Excel...")
    try:
        df = pd.read_excel(EXCEL_FILE)
        print(f"   พบข้อมูลใน Excel ทั้งหมด {len(df)} แถว")
    except Exception as e:
        print(f"❌ ERROR อ่านไฟล์ Excel ไม่สำเร็จ: {e}")
        return {}

    col_map = {str(col).strip().lower(): col for col in df.columns}
    code_col = col_map.get('code') or 'Code'
    lat_col = col_map.get('lat') or 'Lat'
    lng_col = col_map.get('long') or col_map.get('lng') or 'Long'

    stations_dict = {}
    for _, row in df.iterrows():
        raw_code = str(row[code_col]).strip() if pd.notnull(row[code_col]) else ""
        clean_code = re.sub(r'[^A-Z0-9]', '', raw_code.upper())

        try:
            lat = float(row[lat_col]) if pd.notnull(row[lat_col]) else None
            lng = float(row[lng_col]) if pd.notnull(row[lng_col]) else None
        except:
            lat, lng = None, None

        station_obj = {
            "code": raw_code,
            "name": str(row.get('Detail', row.get('Name', ''))).strip(),
            "river": str(row.get('River', '-')).strip(),
            "province": str(row.get('Province', '-')).strip(),
            "region": str(row.get('Region', '-')).strip(),
            "basin": str(row.get('Basin', '-')).strip(),
            "lat": lat,
            "lng": lng
        }

        if clean_code and clean_code != "NAN":
            stations_dict[clean_code] = station_obj
        
        clean_name = re.sub(r'[^A-Z0-9ก-๙]', '', station_obj["name"].upper())
        if clean_name and clean_name != "NAN":
            stations_dict[clean_name] = station_obj

    return stations_dict


async def scrape_daily_hydro_data():
    stations_dict = load_excel_stations()
    all_stations_data = []

    print("\n2. Fetching Daily Water Level Data via Playwright...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(ignore_https_errors=True)
        page = await context.new_page()

        for i in range(1, 9):
            url = f"https://hyd-app-db.rid.go.th/hydro{i}d_admsl.html"
            print(f"   🌐 เชื่อมต่อศูนย์ hydro{i}: {url}")
            
            center_stations = []
            try:
                await page.goto(url, timeout=35000, wait_until="domcontentloaded")
                await page.wait_for_timeout(4000)

                # สั่งรัน JavaScript เพื่อขยาย/แสดงทุกแถวในตารางที่ถูกพับอยู่
                for frame in page.frames:
                    try:
                        await frame.evaluate('''() => {
                            // สั่งแสดงแถวทั้งหมดที่โดน hide/display:none
                            document.querySelectorAll('tr').forEach(tr => tr.style.display = '');
                            // จำลองการคลิกปุ่มขยายกลุ่ม (ถ้ามี)
                            document.querySelectorAll('.treegrid-expander, .expander, [onclick*="expand"]').forEach(el => el.click());
                        }''')
                    except:
                        pass

                await page.wait_for_timeout(2000)

                # สกัดข้อมูลจากตาราง
                for frame in page.frames:
                    rows_data = await frame.evaluate('''() => {
                        const trs = Array.from(document.querySelectorAll('tr'));
                        return trs.map(tr => {
                            const tds = Array.from(tr.querySelectorAll('td, th'));
                            return tds.map(td => td.innerText.trim().replace(/\\n/g, ' '));
                        }).filter(row => row.length >= 3);
                    }''')

                    for cols in rows_data:
                        row_str = " ".join(cols)
                        if 'ลำดับ' in row_str or 'รหัส' in row_str or 'CODE' in row_str.upper() or 'เกณฑ์เฝ้าระวัง' in row_str:
                            continue

                        # ค้นหารหัสสถานี (เช่น P.1, N.1, M.7, X.10A)
                        code_raw = ""
                        name_raw = ""
                        bank_level = "-"
                        water_level = "-"

                        for idx, val in enumerate(cols):
                            # มองหารูปแบบรหัสสถานี (มีตัวอักษรภาษาอังกฤษ + จุด/ตัวเลข)
                            if not code_raw and re.search(r'^[A-Z]{1,3}\.?[0-9]+[A-Z]?$', val.upper()):
                                code_raw = val
                                if idx + 1 < len(cols):
                                    name_raw = cols[idx + 1]
                                break

                        # ถ้าหาตาม pattern ไม่เจอ ให้ใช้ตำแหน่งคอลัมน์มาตรฐาน
                        if not code_raw and len(cols) >= 3:
                            code_raw = cols[1] if len(cols) > 1 else ""
                            name_raw = cols[2] if len(cols) > 2 else ""

                        if not code_raw or code_raw == "-" or len(code_raw) > 15:
                            continue

                        # ดึงระดับตลิ่ง (ถ้ามี)
                        if len(cols) >= 4:
                            for val in cols[3:5]:
                                clean_v = val.replace(',', '')
                                if re.match(r'^-?\d+(\.\d+)?$', clean_v):
                                    bank_level = val
                                    break

                        # ดึงค่าระดับน้ำของวันปัจจุบัน (สแกนย้อนหลังจากคอลัมน์ขวาสุด)
                        for col_val in reversed(cols):
                            clean_val = col_val.replace(',', '').replace('-', '').strip()
                            if clean_val and re.match(r'^-?\d+(\.\d+)?$', clean_val):
                                water_level = col_val
                                break

                        code_clean = re.sub(r'[^A-Z0-9]', '', code_raw.upper())
                        name_clean = re.sub(r'[^A-Z0-9ก-๙]', '', name_raw.upper())

                        # จับคู่พิกัดจาก Excel
                        st_info = stations_dict.get(code_clean) or stations_dict.get(name_clean) or {}

                        status = "normal"
                        try:
                            wl = float(water_level.replace(',', ''))
                            bl = float(bank_level.replace(',', ''))
                            if wl >= bl:
                                status = "danger"
                            elif wl >= (bl - 0.5):
                                status = "warning"
                        except:
                            status = "normal"

                        center_stations.append({
                            "code": st_info.get("code") or code_raw,
                            "name": st_info.get("name") or name_raw,
                            "river": st_info.get("river", "-"),
                            "province": st_info.get("province", "-"),
                            "region": st_info.get("region") or f"ศูนย์อุทกวิทยาภาค {i}",
                            "basin": st_info.get("basin", "-"),
                            "lat": st_info.get("lat"),
                            "lng": st_info.get("lng"),
                            "bank_level": bank_level,
                            "water_level": water_level,
                            "discharge": "-",
                            "status": status
                        })

                print(f"   ✅ ศูนย์ hydro{i}: ดึงได้ {len(center_stations)} สถานี")
                all_stations_data.extend(center_stations)

            except Exception as e:
                print(f"   ⚠️ hydro{i} เกิดข้อผิดพลาด: {e}")

        await browser.close()

    # ตัดตัวซ้ำ
    final_merged = []
    seen = set()
    for st in all_stations_data:
        key = (st['code'], st['name'])
        if key not in seen:
            seen.add(key)
            final_merged.append(st)

    matched_coords = sum(1 for s in final_merged if s.get('lat') is not None and s.get('lng') is not None)

    print(f"\n3. สรุปผลการทำงาน (รายวัน):")
    print(f"   - ดึงข้อมูลสถานีทั้งหมด: {len(final_merged)} สถานี")
    print(f"   - แมตช์พิกัด (Lat/Long) สำเร็จ: {matched_coords} สถานี")

    # บันทึกไฟล์ JSON
    with open(JSON_OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(final_merged, f, ensure_ascii=False, indent=2)

    print(f"✅ บันทึกไฟล์ {JSON_OUTPUT} เรียบร้อยแล้ว!")

if __name__ == "__main__":
    asyncio.run(scrape_daily_hydro_data())