# CONTEXT.md — บรีฟงานสำหรับ AI agent ที่จะไปติดตั้งโปรเจกต์นี้บนเครื่องอื่น

ไฟล์นี้เขียนสำหรับ AI coding agent (เช่น Claude Code session อีกตัวหนึ่ง) ไม่ใช่สำหรับมนุษย์อ่าน
ถ้าคุณเป็นมนุษย์ ให้อ่าน [README.md](README.md) แทน — มีข้อมูลชุดเดียวกันแต่เขียนเป็นขั้นตอนทีละขั้น ภาษาง่ายกว่า

## โปรเจกต์นี้คืออะไร

เครื่องมือหน้าเดียว ([AI_Use_Case_Program12.html](AI_Use_Case_Program12.html)) สำหรับบันทึกและจัดลำดับความสำคัญ use case AI ในองค์กร
เดิมเก็บข้อมูลใน `localStorage` ของเบราว์เซอร์ (เห็นคนเดียว) ถูกย้ายมาใช้ **MySQL** เป็น backend กลาง ผ่าน API ที่เขียนด้วย **Flask** ([server.py](server.py)) เพื่อให้หลายคนในวง LAN เดียวกันเห็นข้อมูลชุดเดียวกัน

## งานที่ต้องทำ

ตั้งค่าโปรเจกต์นี้ให้รันได้บนเครื่อง Windows เครื่องใหม่ โดย reproduce การตั้งค่าที่ใช้งานได้จริงตามที่อธิบายไว้ด้านล่างนี้เป๊ะๆ **ห้ามคิดสถาปัตยกรรมใหม่เอง ให้ทำตามเอกสารนี้**

1. ติดตั้ง Python 3.10 ขึ้นไป (ต้องเพิ่มเข้า PATH ด้วย)
2. ติดตั้ง MySQL Server (เวอร์ชัน 8.x/9.x เวอร์ชันใดก็ได้ที่ค่อนข้างใหม่ — **เลขเวอร์ชันไม่แน่นอนในแต่ละเครื่อง ต้อง detect เอง ห้าม hardcode "8.4"**)
3. รัน [setup_mysql.ps1](setup_mysql.ps1) แบบ Administrator สคริปต์นี้จะหาโฟลเดอร์เวอร์ชัน MySQL ที่ติดตั้งจริงใต้ `C:\Program Files\MySQL\` เอง, initialize data directory ด้วย `--initialize-insecure` (root **ไม่มีรหัสผ่าน**), เขียนไฟล์ `my.ini` แบบขั้นต่ำ, และติดตั้ง MySQL เป็น Windows service ชื่อ `MySQL84` (ชื่อนี้เป็นชื่อเก่าที่ตั้งไว้ตอนแรก ไม่จำเป็นต้องตรงกับเวอร์ชันจริง) เพื่อให้รันอัตโนมัติทุกครั้งที่เปิดเครื่อง
4. สร้าง database และตาราง โดยรัน [schema.sql](schema.sql) กับ server (`mysql -u root < schema.sql` **ไม่ต้องใส่ `-p`** เพราะ root ไม่มีรหัสผ่านโดยตั้งใจ ดูเหตุผลด้านล่าง)
5. `pip install -r requirements.txt` (Flask + PyMySQL)
6. `python server.py` — คำสั่งนี้จะเสิร์ฟทั้งหน้าเว็บ HTML และ REST API (`/api/usecases` แบบ GET/PUT) ที่ port `8765` โดยต่อกับ MySQL
7. ตรวจสอบโดยเปิด `http://localhost:8765` แล้วต้องเห็นหน้าเว็บโหลดขึ้นมาพร้อมรายการ use case ว่างเปล่า (ฐานข้อมูลตอนนี้มี 0 แถวโดยตั้งใจ — ข้อมูล mock/ทดสอบถูกล้างออกไปแล้ว **ห้าม seed ข้อมูลกลับเข้าไปเอง** เว้นแต่จะถูกขอให้ทำ)

[start_all.bat](start_all.bat) ช่วย automate ขั้นตอนที่ 6 เป็นต้นไปสำหรับการใช้งานประจำวันหลัง setup เสร็จแล้ว (start MySQL แบบ background process + `server.py` ในดับเบิลคลิกเดียว)

## จุดออกแบบสถาปัตยกรรมที่ต้องรักษาไว้ ห้ามเปลี่ยนเอง

- **MySQL root ไม่มีรหัสผ่าน ในทุกเครื่อง โดยตั้งใจ** เป็นการลดความซับซ้อนแบบตั้งใจสำหรับเครื่องมือที่ใช้ในวง LAN ภายในเท่านั้น — ทุกเครื่องที่รัน `server.py` จะต่อด้วย credential ชุดเดียวกัน (`root` / รหัสผ่านว่าง / `127.0.0.1:3306` / database `ai_usecase_db`) จะได้ไม่ต้องตั้งค่าหรือแชร์รหัสผ่านให้ตรงกัน **ห้ามเพิ่มรหัสผ่านเอง** เว้นแต่ user จะขอชัดเจน — ถ้าขอ ต้องบอกให้ตั้ง environment variable `DB_PASSWORD` (และอาจต้องตั้ง `DB_USER` ด้วย) ก่อนรัน `server.py` เพราะค่า default ใน [server.py](server.py) เป็นค่าว่าง
- **การตั้งค่านี้ใช้ได้เฉพาะในวง LAN เท่านั้น ห้ามเปิดออกอินเทอร์เน็ตเด็ดขาด** ไม่มีระบบ auth ไม่มี TLS ไม่มีรหัสผ่านป้องกัน DB ห้ามแนะนำให้เปิด port 3306 หรือ 8765 ออกสู่อินเทอร์เน็ตสาธารณะ
- **โครงสร้างข้อมูลตั้งใจให้ยืดหยุ่น ไม่ผูก schema ตายตัว** ตาราง `use_cases` มีคอลัมน์ `id`, `name`, `dept` บวกคอลัมน์ `payload` แบบ JSON ที่เก็บ object ของ use case ทั้งหมดไว้ข้างใน ออกแบบให้ตรงกับพฤติกรรมเดิมของ frontend (array ของ JS object ที่ field ไม่ตายตัว บันทึกทั้งก้อน) **อย่า normalize เป็นหลายคอลัมน์เพิ่มเอง** เว้นแต่จะถูกขอ
- **Contract การ save/load ของ frontend เป็นแบบ "แทนที่ทั้งก้อน" ไม่ใช่ CRUD ทีละแถว**: `GET /api/usecases` คืน array ทั้งหมด, `PUT /api/usecases` จะ**แทนที่เนื้อหาทั้งตารางในทีเดียวในหนึ่ง transaction** ตรงกับพฤติกรรมเดิมของ `localStorage` เป๊ะๆ (ฟังก์ชัน `loadData()`/`saveData()` ในไฟล์ HTML อ่าน/เขียนทั้ง list ทีเดียวเสมอ) — ดูที่ [AI_Use_Case_Program12.html:903](AI_Use_Case_Program12.html:903) **ถ้าจะแก้ API ให้รักษา contract นี้ไว้** อย่าเปลี่ยนเป็น endpoint แบบทีละรายการโดยไม่ปรึกษา user ก่อน เพราะต้องแก้ frontend ด้วย
- Frontend ยังมี fallback ไป `localStorage` แบบเงียบๆ ถ้า fetch `/api/usecases` ล้มเหลว (เช่น เปิดไฟล์ HTML ตรงๆ โดยไม่ได้รัน `server.py`) เป็นพฤติกรรมที่ตั้งใจไว้เป็น soft failure ไม่ใช่บั๊ก

## ปัญหาที่เคยเจอมาแล้ว (อย่าทำซ้ำ)

- **ไฟล์ PowerShell script ที่มีตัวอักษรไทยอาจ parse ไม่ผ่าน** Windows PowerShell 5.1 ตรวจจับ UTF-8 ของไฟล์ `.ps1` ที่ไม่มี BOM ได้ไม่แน่นอน การตีความ byte ผิดอาจสร้างตัวอักษรแปลกปลอมที่ทำให้ string literal พังและเกิด syntax error ด้วยเหตุนี้ `setup_mysql.ps1` จึงใช้**ข้อความภาษาอังกฤษล้วน** ถ้าจะแก้ไฟล์นี้ ให้คงเป็น ASCII ล้วน หรือถ้าจำเป็นต้องใช้ตัวอักษรที่ไม่ใช่ ASCII ต้องเพิ่ม UTF-8 BOM เข้าไปด้วย
- **ชื่อโฟลเดอร์เวอร์ชัน MySQL ไม่คงที่** เครื่องต้นทางได้ `MySQL Server 8.4` แต่เครื่องอื่นจะต่างออกไป (เช่น `MySQL Server 8.0`, `9.x`) ต้อง detect เสมอด้วย `Get-ChildItem "C:\Program Files\MySQL" -Directory` ห้าม hardcode เวอร์ชัน
- **การติดตั้ง Windows Service ต้องใช้สิทธิ์ Administrator / UAC elevation** ซึ่งไม่สามารถขอสิทธิ์แบบเงียบๆ จาก shell ที่ไม่ interactive ได้ ถ้าคุณเป็น AI agent ที่ไม่มี shell แบบ elevated ก็จะติดตั้ง service เองไม่ได้ — ต้องส่งคำสั่ง PowerShell ที่ต้องรันให้ user แล้วขอให้เขารันเองในหน้าต่าง PowerShell ที่เปิดแบบ elevated
- **ตัวติดตั้ง MySQL (GUI) จะบังคับให้ตั้งรหัสผ่าน root เป็นค่า default** ให้ข้าม/ปิดหน้านั้นระหว่างติดตั้ง (หรือปล่อยให้ตั้งไปก่อนแล้วค่อย overwrite ทีหลังด้วยการรัน `mysqld --initialize-insecure` ใหม่ใส่ data directory ที่ยังไม่เคยใช้) — อย่าไปสู้กับ GUI ใช้วิธีสคริปต์ใน `setup_mysql.ps1` แทน

## รายการไฟล์และหน้าที่

| ไฟล์ | หน้าที่ |
|---|---|
| `AI_Use_Case_Program12.html` | ตัวแอป frontend (ไฟล์เดียว vanilla JS) |
| `server.py` | Flask backend: เสิร์ฟหน้าเว็บ + REST API `/api/usecases` ที่ต่อกับ MySQL |
| `schema.sql` | สร้าง database `ai_usecase_db` และตาราง `use_cases` |
| `setup_mysql.ps1` | สคริปต์ตั้งค่า/ติดตั้ง MySQL เป็น Windows service แบบไม่มีรหัสผ่านในคลิกเดียว |
| `requirements.txt` | รายชื่อ Python package ที่ต้องใช้ (`Flask`, `PyMySQL`) |
| `start_all.bat` | ตัว launcher สำหรับใช้งานประจำวัน (ดับเบิลคลิกเดียว) |
| `.gitignore` | กันไฟล์ lock ของ Office, `__pycache__`, และไฟล์ backup `*_backup_*.sql` ไม่ให้หลุดเข้า git |
| `README.md` | คู่มือสำหรับมนุษย์ (ใช้งานประจำวัน, ติดตั้งเครื่องใหม่, วิธีใช้ git, troubleshooting) |
| `.claude/launch.json` | บอกเครื่องมือ browser-preview ของ Claude Code ให้รัน `python server.py` |

## เช็คลิสต์ตรวจสอบก่อนถือว่าเสร็จ

หลัง setup เสร็จ ต้องตรวจครบทุกข้อนี้ก่อนประกาศว่าสำเร็จ:

1. `mysqladmin -u root ping` → ต้องขึ้น `mysqld is alive`
2. `SELECT COUNT(*) FROM ai_usecase_db.use_cases;` → รันได้โดยไม่ error (0 แถวถือว่าปกติ/ถูกต้องสำหรับเครื่องที่เพิ่งติดตั้งใหม่)
3. เปิด `http://localhost:8765` แล้วหน้าแอปโหลดขึ้นมาได้
4. ทดสอบ `PUT /api/usecases` ด้วยข้อมูลทดสอบ 1 รายการ แล้ว `GET /api/usecases` ต้องได้ข้อมูลกลับมาตรงกัน จากนั้น**ต้องล้างข้อมูลทดสอบทิ้งด้วย** ห้ามปล่อยข้อมูลทดสอบค้างไว้
5. `git status` ต้องสะอาดถ้ามีการแก้ไฟล์ใดๆ — อย่าปล่อยไฟล์ untracked ค้างไว้แบบไม่ได้ตั้งใจ
