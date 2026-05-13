import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
from urllib.parse import urljoin, urlparse
import time

# ========== การตั้งค่า (Configuration) ==========
EMAIL_SENDER = "your-email@gmail.com"
EMAIL_PASSWORD = "your-app-password" 
EMAIL_RECEIVER = "receiver@gmail.com"
OUTPUT_DIR = "output"

def fetch_page(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        return response.text
    except Exception as e:
        print(f"❌ ไม่สามารถเข้าถึงเว็บไซต์ได้: {e}")
        return None

def clean_text(text):
    """ล้างตัวอักษรเพี้ยนและช่องว่างส่วนเกิน"""
    if not text: return ""
    # กำจัดตัวอักษรที่มักแสดงผลผิดพลาดใน Excel
    text = text.replace("â", "").replace("Â", "").strip()
    return text

def smart_parse(html):
    """วิเคราะห์และดึงข้อมูลจากหน้าเว็บอัตโนมัติ"""
    soup = BeautifulSoup(html, "html.parser")
    data = []

    # 1. พยายามหาตาราง (Table)
    table = soup.find("table")
    if table:
        print("🔎 ตรวจพบตาราง (Table)... กำลังสกัดข้อมูล")
        
        # ดึงหัวข้อจาก <th> (ถ้ามี)
        headers = [clean_text(th.get_text()) for th in table.find_all("th")]
        
        # ดึงแถวทั้งหมด
        rows = table.find_all("tr")
        if not rows: return []

        # ตรวจสอบว่าแถวแรกเป็น Header หรือไม่ เพื่อป้องกันข้อมูลซ้ำ
        start_idx = 0
        first_row_cells = rows[0].find_all(["td", "th"])
        
        # ถ้าแถวแรกมีแท็ก <th> หรือข้อความตรงกับ Headers ให้เริ่มที่แถวที่ 2
        if rows[0].find("th") or (headers and clean_text(rows[0].find("td").get_text() if rows[0].find("td") else "") == headers[0]):
            start_idx = 1
            if not headers: # กรณีแถวแรกเป็น th แต่ยังไม่มี headers
                headers = [clean_text(c.get_text()) for c in first_row_cells]

        for row in rows[start_idx:]:
            cols = row.find_all(["td", "th"])
            if cols:
                entry = {}
                for i, col in enumerate(cols):
                    col_name = headers[i] if i < len(headers) and headers[i] else f"Column_{i+1}"
                    entry[col_name] = clean_text(col.get_text())
                
                if any(entry.values()): # เก็บเฉพาะแถวที่ไม่ว่าง
                    data.append(entry)
        
        if data: return data

    # 2. กรณีไม่มีตาราง (Fallback) ดึงหัวข้อและเนื้อหาบทความ
    print("🔎 ไม่พบตาราง พยายามดึงเนื้อหาบทความ...")
    tags = soup.select("h1, h2, h3, p")
    for tag in tags:
        text = clean_text(tag.get_text())
        if len(text) > 25:
            data.append({"ประเภท": tag.name, "เนื้อหา": text})
            
    return data

def save_files(data, site_name):
    """บันทึก Excel แบบจัดรูปแบบสวยงามและเส้นตารางไม่เกิน"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_path = os.path.join(OUTPUT_DIR, f"{site_name}_{timestamp}.xlsx")
    
    df = pd.DataFrame(data)
    num_rows, num_cols = df.shape
    
    # ใช้ engine xlsxwriter เพื่อจัดการ Format
    writer = pd.ExcelWriter(xlsx_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Data')
    
    workbook  = writer.book
    worksheet = writer.sheets['Data']
    
    # --- กำหนดรูปแบบเซลล์ ---
    header_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    content_fmt = workbook.add_format({
        'text_wrap': True, 'valign': 'top', 'border': 1
    })
    
    # 1. ปรับความกว้างคอลัมน์ (เฉยๆ ไม่ใส่เส้นขอบทิ้งไว้)
    worksheet.set_column(0, num_cols - 1, 40)
    
    # 2. เขียนหัวตารางซ้ำเพื่อใส่สีและเส้นขอบ
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
        
    # 3. เขียนข้อมูลและตีเส้นขอบ "เฉพาะแถวที่มีข้อมูล"
    for r in range(num_rows):
        for c in range(num_cols):
            val = df.iloc[r, c]
            worksheet.write(r + 1, c, val, content_fmt)
            
    writer.close()
    return xlsx_path

def send_summary_email(file_path, count):
    """ฟังก์ชันส่งอีเมลพร้อมแนบไฟล์"""
    msg = EmailMessage()
    msg['Subject'] = f"📊 รายงานสรุปข้อมูลเว็บ [{datetime.now().strftime('%Y-%m-%d %H:%M')}]"
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER
    msg.set_content(f"ระบบทำงานเสร็จสมบูรณ์\nพบข้อมูลทั้งหมด: {count} รายการ\nกรุณาตรวจสอบไฟล์แนบ")

    try:
        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', 
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                               filename=os.path.basename(file_path))

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("📧 ส่งอีเมลสรุปเรียบร้อย!")
    except Exception as e:
        print(f"❌ ส่งอีเมลล้มเหลว: {e}")

def main():
    print("\n" + "="*30)
    print("🤖 Universal Smart Scraper")
    print("="*30)

    url = input("🌐 ใส่ URL ที่ต้องการดึงข้อมูล: ").strip()
    if not url.startswith("http"): url = "https://" + url

    html = fetch_page(url)
    if html:
        data = smart_parse(html)
        
        if data:
            domain_name = urlparse(url).netloc.replace("www.", "").split(".")[0]
            file_path = save_files(data, domain_name)
            print(f"📊 ดึงข้อมูลสำเร็จ {len(data)} รายการ")
            
            confirm = input("📧 ต้องการส่งรายงานเข้าอีเมลหรือไม่? (y/n): ").lower()
            if confirm == 'y':
                send_summary_email(file_path, len(data))
        else:
            print("⚠️ ไม่พบข้อมูลที่ดึงได้ในหน้าเว็บนี้")

    print("\n✨ โปรแกรมทำงานเสร็จสมบูรณ์!")

if __name__ == "__main__":
    main()