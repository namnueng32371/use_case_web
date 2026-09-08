# AI Use Case Prioritization Tool — คู่มือการติดตั้งและใช้งาน

เครื่องมือบันทึกและจัดลำดับความสำคัญ use case AI ขององค์กร ข้อมูลเก็บอยู่ใน **MySQL database** (`ai_usecase_db`) ผ่าน backend **Flask** ([server.py](server.py)) เพื่อให้หลายคนเข้าใช้งานแล้วเห็นข้อมูลชุดเดียวกัน

คู่มือนี้เป็นไฟล์เดียวครบทุกเรื่อง ไม่ต้องเปิดไฟล์อื่นประกอบ — ใครที่คุ้นเคยกับ Python/MySQL อยู่แล้วข้ามไปหัวข้อ "วิธีรันประจำวัน" ได้เลย ส่วนใครที่เครื่องยังไม่มีอะไรเลยให้เริ่มที่หัวข้อ "ติดตั้งบนเครื่องใหม่ทั้งหมด"

---

## องค์ประกอบของระบบ

| ส่วน | รายละเอียด |
|---|---|
| Database | MySQL Server, database `ai_usecase_db`, ตาราง `use_cases` |
| Backend | Flask ([server.py](server.py)) รันที่ port `8765` |
| Frontend | [AI_Use_Case_Program12.html](AI_Use_Case_Program12.html) เสิร์ฟผ่าน Flask ตัวเดียวกัน |

**หลักการสำคัญ: ทุกเครื่องตั้ง MySQL root แบบไม่มีรหัสผ่านเหมือนกันหมด** (เหมาะกับใช้งานในวง LAN ภายในที่ไว้ใจกันเท่านั้น — **ห้ามเปิด port 3306/8765 ออกอินเทอร์เน็ตเด็ดขาด** เพราะไม่มีรหัสผ่านป้องกัน) `server.py` ใช้ค่านี้เป็น default อยู่แล้ว ไม่ต้องตั้งค่าอะไรเพิ่ม

---

## วิธีรันประจำวัน (เครื่องที่ setup ไว้แล้ว)

ดับเบิลคลิกไฟล์ **[start_all.bat](start_all.bat)** — จะ start MySQL และแอปให้อัตโนมัติ แล้วเปิดเบราว์เซอร์ไปที่:

```
http://localhost:8765
```

**ห้ามปิดหน้าต่าง command ที่เปิดขึ้นมา** ระหว่างใช้งาน — ถ้าปิด แอปและการเชื่อมต่อฐานข้อมูลจะหยุดทำงาน

ให้คนอื่นในวง LAN เดียวกันเข้าใช้งานได้ผ่าน `http://<IP เครื่องนี้>:8765` (ดู IP ด้วยคำสั่ง `ipconfig`; อาจต้องเปิด Windows Firewall ให้ port 8765 ด้วย)

### วิธี manual ทีละขั้น (ถ้าไม่อยากใช้ start_all.bat)

```powershell
Start-Process "C:\Program Files\MySQL\MySQL Server 8.4\bin\mysqld.exe" -ArgumentList '--defaults-file="C:\ProgramData\MySQL\MySQL Server 8.4\my.ini"'
```
รอ 3-4 วินาทีให้ MySQL พร้อม แล้ว (ในโฟลเดอร์โปรเจกต์):
```bash
python server.py
```

> 🔶 เลข `8.4` ในคำสั่งข้างบนคือเวอร์ชัน MySQL — ถ้าเครื่องนั้นลงเวอร์ชันอื่น ให้เปิดดูโฟลเดอร์ `C:\Program Files\MySQL\` จริงก่อน แล้วแก้เลขให้ตรง

### เช็คว่า MySQL รันอยู่ไหม

```bash
"C:\Program Files\MySQL\MySQL Server 8.4\bin\mysqladmin.exe" -u root ping
```
ขึ้น `mysqld is alive` = ใช้งานได้ปกติ, error `Can't connect... (10061)` = MySQL ยังไม่ได้ start

### วิธีหยุดการทำงาน

- ปิดหน้าต่าง command ของ `start_all.bat` (หรือกด `Ctrl+C` ใน terminal ที่รัน `server.py`)
- หยุด MySQL: `"C:\Program Files\MySQL\MySQL Server 8.4\bin\mysqladmin.exe" -u root shutdown`

---

## ติดตั้งบนเครื่องใหม่ทั้งหมด

ทำตามลำดับ 1 → 6 ได้แม้ไม่มีพื้นฐานเทคนิคมาก่อน ใช้เวลาประมาณ 15-20 นาที

> 🔶 **จุดที่ต้องระวัง — แต่ละคอมไม่เหมือนกัน:** จะมีกล่องแบบนี้คั่นไว้ทุกจุดที่ค่าอาจไม่เหมือนกันในแต่ละเครื่อง อ่านให้ดีตรงจุดนั้นเป็นพิเศษ ส่วนที่เหลือทำตามได้เหมือนกันทุกเครื่อง

### ขั้นที่ 1 — ติดตั้ง Python

1. ไปที่ https://www.python.org/downloads/ กดดาวน์โหลดเวอร์ชันล่าสุด (3.10 ขึ้นไป)
2. รันตัวติดตั้ง **สำคัญ: ต้องติ๊กช่อง "Add python.exe to PATH"** ที่หน้าแรกก่อนกด Install
3. เช็คว่าใช้ได้แล้วโดยเปิด Command Prompt/PowerShell ใหม่ พิมพ์:
   ```bash
   python --version
   ```

### ขั้นที่ 2 — ติดตั้ง MySQL Server

1. ไปที่ https://dev.mysql.com/downloads/installer/ ดาวน์โหลด **MySQL Installer for Windows**
2. รันตัวติดตั้ง เลือก Setup Type เป็น **"Server only"**
3. ถ้ามีหน้าต่างถามให้ตั้งรหัสผ่าน/ตั้งค่าเชื่อมต่อ **ปิดทิ้งได้เลย ไม่ต้องกรอกอะไร** — ขั้นตอนถัดไปจะช่วยตั้งให้อัตโนมัติแทน เพื่อให้ได้ root ไม่มีรหัสผ่านตรงกับเครื่องอื่นๆ ในระบบนี้

### ขั้นที่ 3 — เอาโค้ดโปรเจกต์มาไว้ที่เครื่อง

ดูวิธีละเอียดในหัวข้อ [วิธีใช้ git กับโปรเจกต์นี้](#วิธีใช้-git-กับโปรเจกต์นี้) ด้านล่าง (`git clone` แล้ว `git checkout jeng_branch`) หรือถ้าไม่สะดวกใช้ git ให้ขอคนอื่น copy ทั้งโฟลเดอร์มาให้ตรงๆ — ต้องมีไฟล์เหล่านี้ครบ: `AI_Use_Case_Program12.html`, `server.py`, `schema.sql`, `requirements.txt`, `setup_mysql.ps1`

### ขั้นที่ 4 — ให้คอมตั้งค่า MySQL ให้อัตโนมัติ (คลิกเดียว)

1. กด Start พิมพ์ **PowerShell** เห็นโปรแกรมขึ้นมา **คลิกขวา → "Run as administrator"** (มีถามยืนยันให้กด Yes)
2. พิมพ์คำสั่งนี้ (แก้ path ให้ตรงกับที่วางโฟลเดอร์โปรเจกต์ไว้จริง):
   ```powershell
   powershell -ExecutionPolicy Bypass -File "C:\path\to\use_case_web\setup_mysql.ps1"
   ```
3. รอจนขึ้นข้อความสีเขียวว่า "Done! MySQL is ready..."

สคริปต์ [setup_mysql.ps1](setup_mysql.ps1) จะหาเวอร์ชัน MySQL ที่ติดตั้งจริงในเครื่องนั้นเองอัตโนมัติ สร้างที่เก็บข้อมูล ตั้งค่า root ไม่มีรหัสผ่าน และติดตั้งเป็น Windows Service ให้รันอัตโนมัติทุกครั้งที่เปิดเครื่อง — ไม่ต้องรู้เรื่องเวอร์ชันหรือแก้ค่าเอง

### ขั้นที่ 5 — สร้างฐานข้อมูล

เปิด PowerShell ธรรมดา (ไม่ต้อง admin) ไปที่โฟลเดอร์โปรเจกต์แล้วรัน [schema.sql](schema.sql) เพื่อสร้าง database และตาราง (มาพร้อม use case ตัวอย่าง 12 รายการที่วิเคราะห์จากไฟล์ Excel ใน `input_data/` ให้แล้ว):

```powershell
cd "C:\path\to\use_case_web"
& "C:\Program Files\MySQL\MySQL Server 8.4\bin\mysql.exe" -u root < schema.sql
```

> 🔶 เลข `8.4` ต้องตรงกับที่เห็นจริงในโฟลเดอร์ `C:\Program Files\MySQL\` ของเครื่องนั้น (เช่นอาจเป็น `MySQL Server 8.0`)

**ถ้ามีไฟล์ backup ข้อมูล (`.sql`) ที่คนอื่นส่งมาให้แทน** ใช้คำสั่งนี้แทนขั้นตอนข้างบน (ไฟล์ backup จะสร้าง database พร้อมข้อมูลให้ในตัวอยู่แล้ว ไม่ต้องรัน `schema.sql` ซ้ำ):
```powershell
& "C:\Program Files\MySQL\MySQL Server 8.4\bin\mysql.exe" -u root < ชื่อไฟล์ที่ได้รับมา.sql
```

### ขั้นที่ 6 — ติดตั้ง Python package แล้วรันแอป

```powershell
pip install -r requirements.txt
python server.py
```

เปิดเบราว์เซอร์ไปที่ `http://localhost:8765` — ควรเห็นหน้าแอปพร้อม use case ตัวอย่าง 12 รายการที่มากับ `schema.sql` (หรือข้อมูลจาก backup ถ้า import ไว้แทนในขั้นที่ 5) เพิ่มรายการใหม่ผ่านปุ่ม "+ เพิ่ม use case ใหม่" ได้เลย

### สรุปจุดที่ต่างกันในแต่ละเครื่อง (เช็คลิสต์)

| จุดที่ต้องดู | ทำไมถึงต่างกัน | ต้องทำอะไร |
|---|---|---|
| เลขเวอร์ชัน MySQL (เช่น 8.4, 8.0, 9.x) ในชื่อโฟลเดอร์ | แล้วแต่ว่าดาวน์โหลดเวอร์ชันไหนตอนติดตั้ง | เปิดดูโฟลเดอร์ `C:\Program Files\MySQL\` จริงก่อนพิมพ์คำสั่งที่มีเลขเวอร์ชัน |
| path ที่วางโฟลเดอร์โปรเจกต์ไว้ | แต่ละคนวางไว้คนละที่ (Desktop, D:\, ฯลฯ) | แก้ path ในคำสั่งให้ตรงกับที่ตัวเองวางไว้จริง |
| ชื่อไฟล์ backup ข้อมูล (.sql) ถ้ามี | วันที่ export ต่างกันแต่ละครั้ง | ใช้ชื่อไฟล์ที่ได้รับมาจริง |
| IP เครื่อง ถ้าจะให้คนอื่นเข้าผ่านเครือข่ายในออฟฟิศ | แต่ละเครื่องมี IP ไม่เหมือนกัน | เช็คด้วยคำสั่ง `ipconfig` ดูเลขที่ขึ้นตรง "IPv4 Address" |

---

## วิธีใช้ git กับโปรเจกต์นี้

Repo อยู่ที่ https://github.com/namnueng32371/use_case_web.git

### Clone ครั้งแรก

```bash
git clone https://github.com/namnueng32371/use_case_web.git
cd use_case_web
git checkout jeng_branch
```

### Branch ที่มีอยู่

| Branch | สถานะ |
|---|---|
| `main` | branch หลัก |
| `jeng_branch` | branch ที่ใช้พัฒนาอยู่ตอนนี้ |
| `non_branch` | มีอยู่ใน repo |

ตอนนี้ทั้ง 3 branch ยังชี้ไปที่ commit เดียวกัน (ยังไม่มีการแยกงานที่ต่างกันชัดเจน) — ถ้าไม่แน่ใจว่าควรทำงานบน branch ไหน ให้ใช้ `jeng_branch` เป็นค่าเริ่มต้น

### คำสั่งที่ใช้บ่อย

**ดูว่าอยู่ branch ไหน / มีอะไรเปลี่ยนแปลงบ้าง:**
```bash
git status
git branch
```

**สลับไป branch อื่น:**
```bash
git checkout main
git checkout jeng_branch
```

**ดึงของล่าสุดจาก remote มาก่อนเริ่มทำงาน:**
```bash
git pull origin jeng_branch
```

**บันทึกการเปลี่ยนแปลงที่ทำไว้ (commit):**
```bash
git add <ไฟล์ที่แก้ หรือ .>
git commit -m "อธิบายสั้นๆ ว่าแก้อะไร"
```

**ส่งขึ้น remote (push):**
```bash
git push origin jeng_branch
```

### ข้อควรระวัง

- **อย่า push ไฟล์ backup ข้อมูล (`.sql` ที่มีคำว่า backup) หรือไฟล์ชั่วคราว** — มี [.gitignore](.gitignore) กันไว้ให้ส่วนหนึ่งแล้ว แต่เช็ค `git status` ก่อน commit ทุกครั้งว่าไม่มีไฟล์แปลกๆ หลุดเข้าไป
- ก่อน push ควร `git pull` ก่อนเสมอ เผื่อมีคนอื่น push ไปก่อนแล้ว จะได้ไม่ชนกัน (conflict)
- ถ้าเจอ conflict ตอน pull/merge และไม่แน่ใจว่าจะแก้ยังไง ให้หยุดถามก่อน อย่าเดาแก้เอง เพราะอาจทำให้ของคนอื่นหายได้

---

## ข้อมูลการเชื่อมต่อฐานข้อมูล

| รายการ | ค่า |
|---|---|
| Host | `127.0.0.1` |
| Port | `3306` |
| User | `root` |
| Password | *(ไม่มี — ตั้งค่าแบบ local dev ทุกเครื่องเหมือนกัน)* |
| Database | `ai_usecase_db` |
| ตาราง | `use_cases` (คอลัมน์ `id`, `name`, `dept`, `payload` เป็น JSON เก็บข้อมูล use case ทั้งหมด) |

ปรับค่าเชื่อมต่อได้ผ่าน environment variables `DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME` ถ้าจำเป็น (ดูใน [server.py](server.py)) — ปกติไม่ต้องแตะเพราะทุกเครื่องตั้งเหมือนกันหมด

---

## Troubleshooting

| ปัญหา | สาเหตุที่เป็นไปได้ | วิธีแก้ |
|---|---|---|
| `'python' is not recognized` | ไม่ได้ติ๊ก Add to PATH ตอนลง Python | ลง Python ใหม่แล้วติ๊กช่องนั้น หรือเพิ่ม PATH เอง |
| `'mysql' is not recognized` | MySQL bin ไม่ได้อยู่ใน PATH | ใช้ full path เช่น `"C:\Program Files\MySQL\MySQL Server 8.4\bin\mysql.exe"` แทน |
| สคริปต์ `setup_mysql.ps1` บอกว่าไม่พบ MySQL | ยังไม่ได้ติดตั้ง MySQL ตามขั้นที่ 2 | กลับไปทำขั้นที่ 2 ให้เสร็จก่อน |
| `ERROR 1045 (28000): Access denied for user 'root'` | ตัวติดตั้ง MySQL ตั้งรหัสผ่าน root ไว้ก่อนรันสคริปต์ (ทับไม่สำเร็จ) | ลบโฟลเดอร์ data ของ MySQL เวอร์ชันนั้นใน `C:\ProgramData\MySQL\` ทิ้ง แล้วรัน `setup_mysql.ps1` ใหม่ |
| เปิด `localhost:8765` ไม่ได้ | ยังไม่ได้รัน `server.py` | รัน `start_all.bat` หรือ `python server.py` |
| หน้าเว็บเปิดได้แต่ข้อมูลไม่ขึ้น/บันทึกไม่ได้ | MySQL ยังไม่ได้ start | ตรวจด้วย `mysqladmin ping` ตามด้านบน แล้ว start MySQL |
| `Can't connect to MySQL server on 'localhost:3306'` | MySQL ไม่ได้รันอยู่ | Start MySQL ตามขั้นตอนด้านบน หรือ `net start MySQL84` ถ้าติดตั้งเป็น service แล้ว |
| ต้องการให้เพื่อนร่วมงานเข้าใช้งานจากเครื่องอื่นไม่ได้ | Windows Firewall block port 8765 | เปิด inbound rule สำหรับ port 8765 ใน Windows Firewall |

หากยังแก้ไม่ได้ ส่งข้อความ error ที่ขึ้นมาให้ทีมช่วยดูได้เลย
