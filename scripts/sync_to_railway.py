"""
Sync local use cases to Railway production API directly with UTF-8
"""
import json, urllib.request, os

# อ่านข้อมูล local seed data 
with open("docs/seed_data.sql", "r", encoding="utf-8") as f:
    sql_text = f.read()

# ดึง use cases จาก local server ที่รันอยู่ หรือสร้างสคริปต์ยิงเข้า Railway API
# ขอ URL ของ Railway app จาก user หรือยิงตรงผ่าน API
print("พร้อมยิงข้อมูล UTF-8 ภาษาไทยสมบูรณ์ขึ้น Railway")
