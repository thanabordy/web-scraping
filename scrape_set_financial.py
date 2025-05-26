from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from io import StringIO
import pandas as pd
import time

# รายชื่อหุ้น
stock_list = [
    "24CS", "ARIN", "ARROW", "BC", "BKA", "BLESS", "BSM", "BTW", "CAZ", "CHEWA", "CPANEL", "CRD", "DHOUSE",
    "DIMET", "DPAINT", "FLOYD", "HYDRO", "IND", "JAK", "K", "META", "PANEL", "PPS", "PRI", "PROS", "PSG",
    "QTCG", "SENX", "SK", "SMART", "STC", "STX", "SVR", "TAPAC", "THANA", "TIGER", "TITLE", "WELL", "YONG"
]

# ตั้งค่า Chrome เป็น headless
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")

# เปิดเบราว์เซอร์
driver = webdriver.Chrome(options=chrome_options)

# เตรียมตัวเก็บข้อมูล
all_data = {}

for symbol in stock_list:
    try:
        print(f"📥 กำลังดึงข้อมูล: {symbol}")
        url = f'https://www.set.or.th/th/market/product/stock/quote/{symbol}/financial-statement/company-highlights'
        driver.get(url)

        wait = WebDriverWait(driver, 20)
        time.sleep(2)  # รอให้ JS ทำงานสมบูรณ์

        html = driver.page_source
        tables = pd.read_html(StringIO(html))

        # บันทึกตารางแรกของแต่ละหุ้นเท่านั้น (หรือจะเปลี่ยนให้เก็บทั้งหมดก็ได้)
        if tables:
            all_data[symbol] = tables[0]
        else:
            print(f"⚠️ ไม่พบตารางสำหรับ {symbol}")

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดกับ {symbol}: {e}")

driver.quit()

# ---------------------------------------------
# ❌ เวอร์ชันเดิม: สร้างหลาย Sheet
# with pd.ExcelWriter("set_financial_highlights_all.xlsx") as writer:
#     for symbol, df in all_data.items():
#         df.to_excel(writer, sheet_name=symbol[:31], index=False)  # จำกัดชื่อ sheet ไม่เกิน 31 ตัวอักษร
# print("✅ ดึงข้อมูลและบันทึก Excel เสร็จเรียบร้อย: set_financial_highlights_all.xlsx")
# ---------------------------------------------

# ✅ เวอร์ชันใหม่: รวมทุกตารางไว้ใน Sheet เดียว พร้อมหัวตารางชื่อหุ้น
with pd.ExcelWriter("set_financial_highlights_all_single_sheet.xlsx") as writer:
    combined_data = pd.DataFrame()

    for symbol, df in all_data.items():
        # เพิ่มหัวตารางว่าเป็นของหุ้นอะไร
        header_row = pd.DataFrame([[f"ข้อมูลของหุ้น {symbol}"] + [""] * (df.shape[1] - 1)], columns=df.columns)

        # รวมหัว + ข้อมูล
        df_with_header = pd.concat([header_row, df], ignore_index=True)

        # เว้นบรรทัดว่างท้ายตาราง
        empty_row = pd.DataFrame([[""] * df.shape[1]], columns=df.columns)

        # รวมเข้าตารางหลัก
        combined_data = pd.concat([combined_data, df_with_header, empty_row], ignore_index=True)

    # เขียนลง Excel Sheet เดียว
    combined_data.to_excel(writer, sheet_name="AllStocks", index=False)

print("✅ ดึงข้อมูลและบันทึก Excel แบบรวม Sheet เดียวเรียบร้อย: set_financial_highlights_all_single_sheet.xlsx")
