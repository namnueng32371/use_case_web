# AGENTS.md — คำแนะนำและกฎการทำงานสำหรับ AI Agent

เอกสารนี้จัดทำขึ้นเพื่อให้ AI Coding Agent ทุกตัว (เช่น Claude Code, Antigravity หรือ Agent อื่นๆ) เข้าใจภาพรวมของโปรเจกต์ โครงสร้างไฟล์ และ**กฎข้อบังคับที่ต้องปฏิบัติตามอย่างเคร่งครัด**

---

## 🚨 กฎเหล็กสำคัญที่สุด (CRITICAL RULES)

> [!CAUTION]
> 1. **ห้ามทำ `git commit` หรือ `git push` โดยเด็ดขาด ถ้าผู้ใช้ (User) ไม่ได้สั่งอย่างชัดเจน**
>    - เมื่อทำการแก้ไขโค้ด ตรวจสอบบั๊ก หรือสร้างไฟล์ใหม่ ให้คงสถานะไว้ใน Working Directory เพื่อให้ผู้ใช้ตรวจสอบก่อนเสมอ
>    - ห้ามรันคำสั่ง `git commit`, `git push`, หรือแก้ไข Git history อัตโนมัติเด็ดขาด เว้นแต่จะมีคำสั่งชัดเจน เช่น *"commit ให้หน่อย"*, *"push ขึ้น GitHub เลย"*
>
> 2. **ห้ามรันสคริปต์ทดสอบ (`test_webapp.py`) โดยเด็ดขาด ถ้าผู้ใช้ (User) ไม่ได้สั่ง**
>    - ห้ามรันเทสอัตโนมัติในแต่ละรอบของการแก้ไข (โดยเฉพาะงานปรับแต่ง UI / CSS) เพื่อไม่ให้เสียเวลาและทำงานล่าช้า
>    - ให้รันคำสั่ง `python scripts/test_webapp.py` เฉพาะเมื่อผู้ใช้สั่งอย่างชัดเจนเท่านั้น เช่น *"เทสให้หน่อย"*, *"รัน test ให้ดู"*

---

## 1. ภาพรวมของโปรเจกต์ (Project Overview)

โปรเจกต์นี้คือ **AI Use Case Prioritization Tool**
* **Frontend**: หน้าเว็บไฟล์เดี่ยว (`AI_Use_Case_Program12.html`) พัฒนาด้วย Vanilla HTML/CSS/JavaScript
* **Backend**: Flask API (`server.py`) ที่รองรับทั้งการรันในเครื่อง Local และการ Deploy ขึ้น **Railway** (Production)
* **Database**: MySQL จัดเก็บข้อมูล Use Cases ในรูปแบบ JSON Payload ในตาราง `use_cases`
* **API Contract**: การบันทึกข้อมูลจะใช้รูปแบบ "แทนที่ทั้งก้อน" (`PUT /api/usecases`) เพื่อความเรียบง่ายและตรงกับพฤติกรรมของ Frontend

---

## 2. โครงสร้างโฟลเดอร์ (Directory Structure)

```
use_case_web/
├── server.py                   # Flask backend & REST API
├── AI_Use_Case_Program12.html  # Frontend single-page app
├── schema.sql                  # MySQL database schema
├── Procfile                    # คำสั่งสำหรับ Production บน Railway (Gunicorn)
├── runtime.txt                 # เวอร์ชัน Python (3.12.x)
├── requirements.txt            # Python dependencies (Flask, PyMySQL, gunicorn)
├── .gitignore                  # กรองไฟล์ที่ไม่ต้องการเข้า Git
├── AGENTS.md                   # คู่มือสำหรับ AI Agent (ไฟล์นี้)
├── README.md                   # คู่มือสำหรับผู้ใช้ทั่วไป
├── CONTEXT.md                  # ประวัติและข้อควรระวังทางสถาปัตยกรรม
│
├── docs/                       # เอกสารอ้างอิงและชุดข้อมูล
│   ├── seed_data.sql           # ข้อมูล Use Cases ทั้งหมดที่ใช้ Sync/Backup
│   └── use_cases_source/       # ไฟล์ต้นฉบับ (Excel / PPTX)
│
└── scripts/                    # สคริปต์ช่วยเหลือสำหรับ Developer
    ├── start_all.bat           # ดับเบิลคลิกเพื่อรัน MySQL + Server บน Windows Local
    ├── setup_mysql.ps1         # สคริปต์ติดตั้ง MySQL Service บน Windows
    ├── test_webapp.py          # ชุดทดสอบระบบ API และหน้าเว็บ (25 Test Cases)
    ├── export_to_sql.py        # ดึงข้อมูลจาก Local DB ออกมาเป็น docs/seed_data.sql
    └── push_seed.py            # ดันข้อมูล docs/seed_data.sql ขึ้น Railway Production
```

---

## 3. สคริปต์และขั้นตอนการทำงาน (Workflows)

### 3.1 การรันระบบในเครื่อง (Local Development)
- รันผ่าน `scripts/start_all.bat` (จะเปิด MySQL และ Flask server ที่พอร์ต `8765`)
- เข้าใช้งานได้ที่ `http://localhost:8765` หรือ `http://127.0.0.1:8765`

### 3.2 การทดสอบระบบ (Testing)
- **ไม่ต้องรันเทสอัตโนมัติในแต่ละรอบของการแก้ไข**
- รันคำสั่ง `python scripts/test_webapp.py` เฉพาะเมื่อผู้ใช้สั่งให้ทดสอบเท่านั้น เพื่อยืนยันว่า Endpoint ต่างๆ ยังทำงานสมบูรณ์ (ต้องผ่าน 25/25 PASS)

### 3.3 การอัปเดตข้อมูล Use Cases ขึ้น Production (Railway)
1. Export ข้อมูลล่าสุดใน Local ลงไฟล์:
   ```powershell
   python scripts\export_to_sql.py
   ```
2. ดันข้อมูลขึ้น Railway โดยตรง (รองรับ UTF-8 ภาษาไทยสมบูรณ์):
   ```powershell
   python scripts\push_seed.py
   ```

---

## 4. ข้อควรระวังในการแก้ไขโค้ด (Caution & Gotchas)

1. **ระวังการแก้ไขไฟล์ `AI_Use_Case_Program12.html`**:
   - ในไฟล์นี้มี JavaScript strings หลายจุดที่จำลองเอกสาร เช่น Word Document (`.doc`) ซึ่งมีแท็ก `</body>` และ `</html>` อยู่ในเนื้อหาข้อความ String Literal
   - **ห้าม** ใช้การค้นหาแบบง่ายๆ เช่น `replace("</body>", ...)` เพราะจะไปแทรกโค้ดเข้ากลาง String Literal ทำให้เกิด JavaScript SyntaxError
2. **ตัวแปร Environment Variables สำหรับ Railway**:
   - Production ใช้ตัวแปร: `DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME` (ค่า default ชี้ไปที่ `127.0.0.1` เมื่อรัน local)
3. **Encoding ภาษาไทยบน Windows**:
   - ให้ระบุ `encoding='utf-8'` เสมอเมื่อเปิดอ่านหรือเขียนไฟล์ใน Python
