import streamlit as st
import pandas as pd
import google.generativeai as genai
import json
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image
import io

# --- 1. การตั้งค่าหน้าจอ ---
st.set_page_config(page_title="Laundry Data System (Gemini)", layout="wide")
st.title("🏨 ระบบจัดการข้อมูลซักรีด (Gemini Version)")

# --- 2. ดึงข้อมูลจาก Secrets ---
try:
    # สำหรับ Gemini API
    gemini_api_key = st.secrets["gemini_api_key"]
    
    # สำหรับ Google Sheets
    gcp_info = st.secrets["gcp_service_account"]
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_info(gcp_info, scopes=scope)
    gc = gspread.authorize(credentials)
    
    # เชื่อมต่อ Google Sheets (ใส่ชื่อไฟล์ Sheets ของคุณ)
    sh = gc.open("Laundry_Data") # ** ตรวจสอบว่าชื่อไฟล์ใน Google Sheets ตรงกัน **
    worksheet = sh.get_worksheet(0)
    
    genai.configure(api_key=gemini_api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')

except Exception as e:
    st.error(f"❌ การเชื่อมต่อล้มเหลว: {e}")
    st.stop()

# --- 3. ส่วนการอัปโหลดและประมวลผล ---
uploaded_file = st.file_uploader("📤 อัปโหลดรูปภาพบิลซักรีด", type=['png', 'jpg', 'jpeg'])

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption='รูปภาพที่อัปโหลด', width=400)
    
    if st.button("🤖 ให้ AI วิเคราะห์ข้อมูล"):
        with st.spinner('Gemini กำลังอ่านข้อมูล...'):
            try:
                # ส่งรูปให้ Gemini วิเคราะห์
                response = model.generate_content([
                    "วิเคราะห์รูปภาพบิลซักรีดนี้ และสรุปข้อมูลออกมาเป็น JSON format โดยมี key ดังนี้: "
                    "date (ว/ด/ป), department (แผนกที่ส่ง), items (รายการสินค้าและจำนวน), total_amount (ราคารวม). "
                    "ตอบเฉพาะ JSON เท่านั้น", 
                    image
                ])
                
                # ทำความสะอาดข้อความ JSON ที่ได้จาก AI
                clean_json = response.text.replace('```json', '').replace('```', '').strip()
                data = json.loads(clean_json)
                
                # แสดงผลที่ AI อ่านได้
                st.success("✅ อ่านข้อมูลสำเร็จ!")
                st.json(data)
                
                # ปุ่มบันทึกข้อมูล
                if st.button("💾 บันทึกเข้า Google Sheets"):
                    worksheet.append_row([
                        data.get('date'),
                        data.get('department'),
                        str(data.get('items')),
                        data.get('total_amount')
                    ])
                    st.balloons()
                    st.success("บันทึกลง Google Sheets เรียบร้อยแล้ว!")
            
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

# --- 4. แสดงตารางข้อมูลล่าสุด ---
st.divider()
st.subheader("📊 ข้อมูลล่าสุดในระบบ")
existing_data = pd.DataFrame(worksheet.get_all_records())
st.dataframe(existing_data, use_container_width=True)
