import json, urllib.request, re

RAILWAY_URL = "https://web-production-375951.up.railway.app/api/usecases"

# อ่านและแกะข้อมูลภาษาไทยจากไฟล์ seed_data.sql
with open("docs/seed_data.sql", "r", encoding="utf-8") as f:
    lines = f.readlines()

cases = []
for line in lines:
    if line.startswith("INSERT IGNORE INTO"):
        # แกะ payload JSON ที่อยู่ด้านหลังสุด
        idx = line.find("{")
        last_idx = line.rfind("}")
        if idx != -1 and last_idx != -1:
            json_str = line[idx:last_idx+1].replace("\\'", "'")
            try:
                cases.append(json.loads(json_str))
            except Exception as e:
                print("Error parsing line:", e)

print(f"กำลังส่งข้อมูลภาษาไทยสมบูรณ์ {len(cases)} รายการขึ้น Railway...")

# ยิงข้อมูลภาษาไทย UTF-8 สมบูรณ์ขึ้น Railway
body = json.dumps(cases, ensure_ascii=False).encode("utf-8")
req = urllib.request.Request(RAILWAY_URL, data=body, method="PUT", headers={"Content-Type": "application/json"})

try:
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode("utf-8"))
        print("🎉 สำเร็จเรียบร้อย! ผลลัพธ์จาก Railway:", res)
except Exception as e:
    print("❌ เกิดข้อผิดพลาด:", e)
