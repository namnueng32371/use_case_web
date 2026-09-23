# AGENTS.md — คำแนะนำและกฎการทำงานสำหรับ AI Agent

เอกสารนี้จัดทำขึ้นเพื่อให้ AI Coding Agent ทุกตัว (เช่น Claude Code, Antigravity หรือ Agent อื่นๆ) เข้าใจภาพรวมของโปรเจกต์ โครงสร้างไฟล์ และ**กฎข้อบังคับที่ต้องปฏิบัติตามอย่างเคร่งครัด**

---

## 🚨 กฎเหล็กสำคัญที่สุด (CRITICAL RULES)

> [!CAUTION]
> 1. **ห้ามทำ `git commit` หรือ `git push` โดยเด็ดขาด ถ้าผู้ใช้ (User) ไม่ได้สั่งอย่างชัดเจน**
>    - เมื่อทำการแก้ไขโค้ด ตรวจสอบบั๊ก หรือสร้างไฟล์ใหม่ ให้คงสถานะไว้ใน Working Directory เพื่อให้ผู้ใช้ตรวจสอบก่อนเสมอ
>    - ห้ามรันคำสั่ง `git commit`, `git push`, หรือแก้ไข Git history อัตโนมัติเด็ดขาด เว้นแต่จะมีคำสั่งชัดเจน เช่น *"commit ให้หน่อย"*, *"push ขึ้น GitHub เลย"*
>
> 2. **เมื่อไหร่ที่ commit (และ/หรือ push) ได้รับอนุญาตแล้ว ให้ publish `artifact.html` ขึ้น Claude Artifact ตัวหลักด้วยเสมอ**
>    - Artifact ตัวหลักคือ **"AI Use Case Prioritization Tool"** — publish โดยระบุ `url` ของ artifact เดิมเสมอ (ห้าม publish แบบไม่ใส่ `url` เพราะจะไปสร้างอันใหม่แยกหรือทับของเดิมผิดตัว — เคยเกิดปัญหานี้มาแล้ว ดู [CONTEXT.md](CONTEXT.md) หัวข้อ "ข้อควรรู้เรื่อง Artifact")
>    - ลำดับที่ถูกต้องเมื่อผู้ใช้สั่ง "commit": (1) แก้ไฟล์ + ทดสอบให้เสร็จ (2) `git commit` + `git push` (3) publish `artifact.html` ขึ้น Artifact ตัวหลัก (4) รายงานผลทั้งลิงก์ commit และลิงก์ Artifact ให้ผู้ใช้
>    - ถ้าผู้ใช้สั่งแค่ "commit" เฉยๆ โดยไม่พูดถึง artifact ให้ถือว่ารวม publish artifact ด้วยเสมอ (ไม่ต้องถามซ้ำ)
>
> 3. **`artifact.html` และ `AI_Use_Case_Program12.html` ต้องแก้คู่กันเสมอ**
>    - ทั้งสองไฟล์เป็นแอปตัวเดียวกัน มีโครงสร้าง field/ฟังก์ชันเดียวกัน แต่ **ไม่ได้ sync กันอัตโนมัติและมีเนื้อหาต่างกันเล็กน้อยในบางจุด** — ห้ามสมมติว่าเลขบรรทัดหรือข้อความตรงกันเป๊ะ ให้ตรวจสอบเนื้อหาจริงของแต่ละไฟล์ก่อนแก้ทุกครั้ง (grep/read ก่อน edit)
>    - แก้ไฟล์หนึ่งแล้วลืมอีกไฟล์ = ฟีเจอร์ไม่ตรงกันระหว่างเวอร์ชัน Artifact กับเวอร์ชันไฟล์เปล่า

---

## 1. ภาพรวมของโปรเจกต์ (Project Overview)

โปรเจกต์นี้คือ **AI Use Case Prioritization Tool** — เครื่องมือช่วยเก็บ วิเคราะห์ และจัดลำดับความสำคัญของ use case AI ในองค์กร (โรงงานผลิตยาง/พลาสติก)

* **สถาปัตยกรรม**: **ไม่มี Backend/Database แยกต่างหากแล้ว** — เดิมเคยเป็น Flask + MySQL แต่ถูกย้ายมาเป็น **Claude Artifact เพียวๆ** ทั้งหมด (ไฟล์เดี่ยว HTML/CSS/JavaScript, ไม่มี server, ไม่มี MySQL, ไม่มี Railway)
* **เก็บข้อมูลที่ไหน**: ใช้ `window.claude.use('db')` (Artifact's shared database) ผ่าน `claudeDb.doc('usecases/all').get()/.set({items:useCases})` — ข้อมูลอยู่ในตัว Artifact เอง
* **สองไฟล์ HTML คู่กัน**:
  - `artifact.html` — เวอร์ชันสำหรับ publish ขึ้น Claude Artifact (ไม่มี `<!DOCTYPE>`/`<html>` wrapper)
  - `AI_Use_Case_Program12.html` — เวอร์ชันไฟล์เต็มหน้า เปิดตรงจากเครื่อง/เบราว์เซอร์ได้โดยไม่ต้องมี server
* **Excel Template**: `docs/AI_UseCase_Template_v2.xlsx` เป็น template คู่ขนานที่แจกให้แผนกกรอกแบบออฟไลน์ — สร้างจากสคริปต์ `docs/build_template.py` โครงสร้าง field ต้องตรงกับฟอร์มในเว็บแอปเสมอ

---

## 2. โครงสร้างโฟลเดอร์ (Directory Structure)

```
use_case_web/
├── artifact.html                    # เวอร์ชัน publish ขึ้น Claude Artifact (ตัวหลักที่ผู้ใช้ใช้งานจริง)
├── AI_Use_Case_Program12.html       # เวอร์ชันไฟล์เต็มหน้า เปิดตรงได้โดยไม่ต้องมี server
├── AGENTS.md                        # คู่มือสำหรับ AI Agent (ไฟล์นี้)
├── CONTEXT.md                       # ประวัติ ข้อตกลง และบรีฟงานแบบละเอียด (อ่านก่อนเริ่มงานใหญ่)
├── PRODUCT.md                       # มุมมองผลิตภัณฑ์/ธุรกิจ
├── README.md                        # คู่มือสำหรับผู้ใช้ทั่วไป
├── .gitignore
│
├── data/                            # ข้อมูล export/backup
├── backup_ai_usecase_db_*.sql       # backup โครงสร้าง+ข้อมูลจากยุค MySQL (เผื่อกู้คืน)
│
└── docs/                            # เอกสารอ้างอิงและ template
    ├── AI_UseCase_Template_v2.xlsx  # Template Excel ที่แจกให้แผนกกรอก (ไฟล์ที่ส่งมอบจริง)
    ├── build_template.py            # สคริปต์สร้าง Excel template ด้านบน (ห้ามแก้ไฟล์ .xlsx มือ)
    ├── TEMPLATE_GUIDE.md            # คู่มืออธิบายทุกคอลัมน์ใน template
    ├── SCORING.md                   # ที่มาของสูตรคะแนนและเหตุผลที่ตัดบางปัจจัยออก
    ├── dept_codes.json
    └── use_cases_source*/           # ไฟล์ Excel ต้นฉบับของแต่ละแผนก (ก่อนทำ template รวม)
```

> ไม่มี `server.py`, `schema.sql`, `requirements.txt`, `Procfile`, `runtime.txt`, หรือโฟลเดอร์ `scripts/` แล้ว — ถูกลบไปตอนย้ายจาก Flask+MySQL มาเป็น Artifact-only ถ้าเจอไฟล์เหล่านี้กลับมาอีกหรือเอกสารเก่าอ้างถึง ให้ถือว่าเอกสารนั้นล้าสมัย

---

## 3. ขั้นตอนการทำงาน (Workflows)

### 3.1 การแก้ไขฟอร์ม/ฟีเจอร์ในเว็บแอป
1. แก้ `artifact.html` ก่อน (เป็นตัวอ้างอิงหลัก)
2. mirror การแก้ไขแบบเดียวกันไปที่ `AI_Use_Case_Program12.html` — **ตรวจเนื้อหาจริงก่อนแก้ทุกครั้ง** เพราะสองไฟล์เนื้อหาไม่ตรงกันเป๊ะ
3. ทดสอบผ่านเบราว์เซอร์ (เปิดไฟล์ตรงๆ หรือใช้ preview tool) — เช็ค console error, ทดสอบ add/edit/save use case จริง
4. รายงานผลแล้วรอคำสั่งก่อน commit (ดูกฎเหล็กข้อ 1)

### 3.2 การแก้ไข Excel Template
**ห้ามแก้มือใน Excel** ให้แก้ที่ `docs/build_template.py` แล้วรันสคริปต์สร้างใหม่เสมอ เพราะไฟล์มีสูตรและ dropdown หลายสิบชุดที่ต้องตรงกันทุกจุด

```bash
python docs/build_template.py
```

หลังรันเสร็จ **ต้องเปิดด้วย Excel ตัวจริงยืนยัน** ไม่ใช่เชื่อการอ่านไฟล์ด้วย Python อย่างเดียว (ดูรายละเอียดเช็คลิสต์ใน [CONTEXT.md](CONTEXT.md) หัวข้อ "สิ่งที่ต้องตรวจทุกครั้งหลังแก้ template")

### 3.3 การ commit + publish ขึ้น Artifact
เมื่อผู้ใช้สั่ง commit แล้ว ให้ทำตามลำดับนี้เสมอ (ดูกฎเหล็กข้อ 2):
1. `git add` เฉพาะไฟล์ที่เกี่ยวข้อง
2. `git commit` พร้อมข้อความอธิบาย "ทำไม" ไม่ใช่แค่ "ทำอะไร"
3. `git push` (ถ้าผู้ใช้สั่งหรืออนุญาตไว้)
4. **Publish `artifact.html` ขึ้น Artifact ตัวหลัก** โดยระบุ `url` เดิมเสมอ
5. รายงานทั้งผล commit/push และลิงก์ Artifact ให้ผู้ใช้

---

## 4. ข้อควรระวังในการแก้ไขโค้ด (Caution & Gotchas)

1. **ระวัง JavaScript string literal ที่มี `</body>`/`</html>` อยู่ข้างใน**:
   - ทั้ง `artifact.html` และ `AI_Use_Case_Program12.html` มีฟังก์ชันสร้าง Word document (`.doc`) ที่ประกอบ string HTML เต็มรูปแบบ (มี `<head>`, `</body></html>` อยู่ใน string literal เอง)
   - **ห้าม** ใช้การค้นหา/แทนที่แบบง่ายๆ (เช่น `replace("</body>", ...)`) เพราะจะไปแทรกโค้ดกลาง string literal ทำให้เกิด JavaScript SyntaxError — ใช้ Edit tool แบบระบุ context ให้ชัดเจนแทน
2. **field ในฟอร์มเว็บแอป ต้องตรงกับคอลัมน์ใน Excel template เสมอ**:
   - ห้ามสร้างหัวข้อ/กล่องใหม่ในฟอร์มที่ซ้ำความหมายกับคอลัมน์ที่มีอยู่แล้วใน template — ถ้าจะเพิ่ม field ใหม่ ให้เพิ่มใน `docs/build_template.py` (sheet 1/2) คู่กับใน `DATA_LISTS`/ฟอร์มของทั้งสองไฟล์ HTML ด้วย
   - รายชื่อ dropdown (บริษัท, หน่วยงาน ฯลฯ) ต้องแก้พร้อมกันทั้ง 3 ที่: `build_template.py` (LISTS), `artifact.html` (DATA_LISTS), `AI_Use_Case_Program12.html` (DATA_LISTS)
3. **Encoding ภาษาไทยบน Windows**:
   - ให้ระบุ `encoding='utf-8'` เสมอเมื่อเปิดอ่านหรือเขียนไฟล์ด้วย Python — คอนโซล Windows (PowerShell/cmd) มักแสดงภาษาไทยเพี้ยน (cp874) ทั้งที่ไฟล์จริงเป็น UTF-8 ปกติ ให้ตรวจด้วยการ redirect output ไปไฟล์แล้วอ่านด้วย UTF-8 แทนการเชื่อสายตาจาก terminal ตรงๆ
4. **อย่าลบไฟล์ backup ยุค MySQL โดยไม่ถาม**:
   - `backup_ai_usecase_db_*.sql` และ `data/use_cases_export.json` เป็น backup ข้อมูลจริงก่อนย้ายมา Artifact เก็บไว้เผื่อกู้คืน ห้ามลบเว้นแต่ผู้ใช้สั่ง
