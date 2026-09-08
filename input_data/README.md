# input_data/

โฟลเดอร์นี้เก็บไฟล์ Excel ต้นทาง (เช่น AI Transformation Plan ของแต่ละแผนก/ไลน์ผลิต) ที่ใช้เป็นข้อมูล input
สำหรับวิเคราะห์และดึงมาเพิ่มเป็น use case ใหม่ในเว็บแอป `AI_Use_Case_Program12.html`

## วิธีเพิ่ม use case ใหม่จากไฟล์ Excel

1. วางไฟล์ Excel ต้นทางไว้ในโฟลเดอร์นี้ (ตั้งชื่อไฟล์ให้สื่อความหมาย เช่น `AI_Transformation_Plan_<แผนก/ไลน์>.xlsx`)
2. วิเคราะห์เนื้อหาในไฟล์ (คอลัมน์ "Description for Redesign" / "Existing Work" ในแต่ละชีตมักเป็นจุดที่ระบุโอกาสใช้ AI)
   แล้วสรุปเป็น use case ตาม schema ของฟอร์มในเว็บแอป — ดูตัวอย่างได้ที่ `new_usecases_prod_ee.json` ในโฟลเดอร์นี้
3. บันทึกผลลัพธ์เป็นไฟล์ JSON (array ของ use case object) ไว้ในโฟลเดอร์นี้
4. เพิ่มเข้าเว็บแอปได้ 2 ทาง แล้วแต่สถานการณ์:
   - **มีเครื่องที่รัน `server.py` ต่อกับ MySQL อยู่แล้ว**: รัน `python import_usecases.py <ไฟล์.json>` เพื่อเพิ่ม
     use case ใหม่เข้าฐานข้อมูลที่กำลังใช้งานอยู่ทันที โดยไม่ลบข้อมูลเดิม (ข้าม id ที่ซ้ำอัตโนมัติ)
   - **ต้องการให้ use case นี้ติดมากับการติดตั้งเครื่องใหม่ทุกครั้ง** (เช่นจะ push ขึ้น git ให้ทีมอื่น clone แล้วรัน
     `schema.sql` ได้ข้อมูลชุดนี้ทันที): แปลง object แต่ละอันเป็นคำสั่ง `INSERT IGNORE INTO use_cases (id, name, dept, payload)
     VALUES (...)` ต่อท้ายไฟล์ [../schema.sql](../schema.sql) (ใช้ `INSERT IGNORE` เสมอ เพื่อไม่ให้ทับข้อมูลที่มีคนแก้ไขผ่าน
     หน้าเว็บไปแล้วเมื่อรัน `schema.sql` ซ้ำ) — ทำทั้งสองทางพร้อมกันได้ ไม่ขัดแย้งกัน

## ไฟล์ในโฟลเดอร์นี้

| ไฟล์ | ที่มา | หมายเหตุ |
|---|---|---|
| `AI_Transformation_Plan_Prod_EE.xlsx` | อัปโหลดโดยผู้ใช้ (2026-09-08) | ชีต: AI-Transform-Plan (template ตัวอย่างทั่วไป, ไม่ใช่ข้อมูลจริง), Prod / Prod (2) / Prod (3) (ขั้นตอนผลิต Molding ของ PE NPD), QWP (เอกสาร QWP-A07 อ้างอิงคุณภาพ), Man Alocation1-3 (แผน AI จัดสรรกำลังคนตาม Skill Matrix แบบละเอียด 3 ระดับ) |
| `new_usecases_prod_ee.json` | วิเคราะห์จากไฟล์ข้างต้น | use case ใหม่ 6 รายการ พร้อมนำเข้าเว็บแอปด้วย `import_usecases.py` |
| `AI_Transformation_Plan_RD_CENTER.xlsx` | อัปโหลดโดยผู้ใช้ (2026-09-08) | ไฟล์มีหลายสิบชีต วิเคราะห์จากชีต `AI-Transform-Plan (RD CENTER)` เป็นหลัก ครอบคลุมกระบวนการ RD Center ตั้งแต่รับ RFQ, ประเมิน Feasibility/Out Source, ออกแบบกระบวนการผลิต+Cost1/Cost2+Formulation, ตรวจสอบ Safety/QFD2 ของ QA, ติดตาม NPD Master Plan, จนถึงการทดลองผลิตสินค้าตัวอย่าง |
| `new_usecases_rd_center.json` | วิเคราะห์จากไฟล์ข้างต้น | use case ใหม่ 6 รายการ พร้อมนำเข้าเว็บแอปด้วย `import_usecases.py` |
