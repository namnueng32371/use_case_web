# -*- coding: utf-8 -*-
"""ตรวจสอบความถูกต้องของ AI_UseCase_Template_v2.xlsx ทุกมิติ"""
import io, sys, warnings, re
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

PATH = sys.argv[1]
FIRST, LAST = 7, 36
fails, warns = [], []
def bad(m):  fails.append(m); print("  [FAIL]", m)
def warn(m): warns.append(m); print("  [WARN]", m)
def ok(m):   print("  [ok]  ", m)

wb = load_workbook(PATH)
S1N, S2N, S3N = "1-UseCase ทั้งหมด", "2-รายละเอียดเทคนิค", "3-คะแนนและลำดับ"
w1, w2, w3, wd = wb[S1N], wb[S2N], wb[S3N], wb["DATA"]

print("\n=== 1. โครงสร้าง sheet ===")
want = ["คู่มือการใช้งาน", S1N, S2N, S3N, "DATA"]
if wb.sheetnames == want: ok("ลำดับ sheet ถูกต้อง: " + " | ".join(want))
else: bad(f"ลำดับ sheet ผิด: {wb.sheetnames}")

print("\n=== 2. รายการ dropdown ใน DATA (ตัวเลือก vs คะแนน ต้องตรงบรรทัด) ===")
SCORE_OFFSET = 30
lists = {}
for ci in range(1, 40):
    L = get_column_letter(ci)
    head = wd[f"{L}3"].value
    if not head or L in ("AH", "AI"): continue   # AH/AI = บล็อกตั้งค่า ไม่ใช่รายการ dropdown
    opts, scores = [], []
    r = 4
    while wd[f"{L}{r}"].value not in (None, ""):
        opts.append(wd[f"{L}{r}"].value); r += 1
    r = SCORE_OFFSET
    while wd[f"{L}{r}"].value not in (None, ""):
        scores.append(wd[f"{L}{r}"].value); r += 1
    lists[L] = opts
    if len(opts) != len(scores):
        bad(f"คอลัมน์ {L} ({head}): ตัวเลือก {len(opts)} บรรทัด แต่คะแนน {len(scores)} บรรทัด")
    for i, s in enumerate(scores):
        if s != "-" and not (isinstance(s, (int, float)) and 1 <= s <= 5):
            bad(f"คอลัมน์ {L} บรรทัดที่ {i+1}: คะแนน '{s}' ไม่ใช่ 1-5 หรือ '-'")
    if len(set(opts)) != len(opts):
        bad(f"คอลัมน์ {L} ({head}): มีตัวเลือกซ้ำกัน")
ok(f"ตรวจรายการ dropdown ทั้งหมด {len(lists)} ชุด")

print("\n=== 3. dropdown ที่ผูกกับ sheet 1 / 2 ชี้ไปที่ช่วงถูกต้องมั้ย ===")
for ws in (w1, w2):
    for dv in ws.data_validations.dataValidation:
        if dv.type is None:      # กล่องคำอธิบายอย่างเดียว ไม่ได้จำกัดค่า
            if not (dv.prompt and dv.showInputMessage):
                bad(f"{ws.title} {dv.sqref}: validation ว่างเปล่า ไม่มีทั้งรายการและคำอธิบาย")
            continue
        m = re.match(r"DATA!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$", dv.formula1 or "")
        if not m:
            bad(f"{ws.title}: formula1 ผิดรูปแบบ -> {dv.formula1}"); continue
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        col = str(dv.sqref).split(":")[0].rstrip("0123456789")
        head = (ws[f"{col}5"].value or "").split("\n")[0]
        if c1 != c2: bad(f"{ws.title} คอลัมน์ {col}: ช่วงข้ามคอลัมน์")
        n = len(lists.get(c1, []))
        if r1 != 4 or r2 != 3 + n:
            bad(f"{ws.title} คอลัมน์ {col} ({head}): ช่วง {dv.formula1} "
                f"ไม่ตรงกับจำนวนตัวเลือก {n} (ควรเป็น DATA!${c1}$4:${c1}${3+n})")
        if dv.formula1.startswith("="): bad(f"{ws.title} {col}: formula1 มี '=' นำหน้า")
        if dv.showDropDown: bad(f"{ws.title} {col}: showDropDown=True จะซ่อนลูกศร")
        got = str(dv.sqref)
        if got != f"{col}{FIRST}:{col}{LAST}":
            warn(f"{ws.title} {col}: ผูกช่วง {got} (คาด {col}{FIRST}:{col}{LAST})")
def n_list(ws): return sum(1 for d in ws.data_validations.dataValidation if d.type)
def n_tip(ws):  return sum(1 for d in ws.data_validations.dataValidation if not d.type)
ok(f"sheet1 dropdown {n_list(w1)} ชุด + กล่องคำอธิบาย {n_tip(w1)} ช่อง / "
   f"sheet2 dropdown {n_list(w2)} ชุด + กล่องคำอธิบาย {n_tip(w2)} ช่อง")

print("\n=== 4. แถวตัวอย่าง (แถว 6) ต้องเป็นค่าที่มีในรายการจริง ===")
for ws in (w1, w2):
    for dv in ws.data_validations.dataValidation:
        if dv.type is None: continue
        col = str(dv.sqref).split(":")[0].rstrip("0123456789")
        v = ws[f"{col}6"].value
        if v is None: continue
        if isinstance(v, str) and v.startswith("="): continue
        c1 = re.match(r"DATA!\$([A-Z]+)", dv.formula1).group(1)
        if v not in lists[c1]:
            bad(f"{ws.title} {col}6 = '{v}' ไม่มีในรายการ DATA!{c1}")
ok("ตรวจค่าตัวอย่างครบทุกช่องที่เป็น dropdown")

print("\n=== 4b. ขีดจำกัดของ Excel (เกินแล้วไฟล์เปิดไม่ขึ้น) ===")
LIM = [("promptTitle", 32), ("prompt", 255), ("errorTitle", 32), ("error", 255)]
n_chk = 0
for ws in wb.worksheets:
    for d in ws.data_validations.dataValidation:
        for attr, mx in LIM:
            v = getattr(d, attr, None)
            if v is None: continue
            n_chk += 1
            if len(v) > mx:
                bad(f"{ws.title} {d.sqref}: {attr} ยาว {len(v)} ตัว เกินขีดจำกัด {mx}")
ok(f"ตรวจความยาวข้อความใน validation {n_chk} รายการ (ขีดจำกัด 32 / 255)")
import zipfile
_z = zipfile.ZipFile(PATH)
(ok if _z.testzip() is None else bad)(f"โครงสร้างไฟล์ zip ปกติ ({len(_z.namelist())} ส่วน)")

print("\n=== 5. หัวคอลัมน์และจำนวนช่อง ===")
for ws, n in ((w1, 23), (w2, 23), (w3, 21)):
    got = sum(1 for c in range(1, 40) if ws.cell(row=5, column=c).value)
    (ok if got == n else bad)(f"{ws.title}: {got} คอลัมน์ (คาด {n})")
    for c in range(1, got + 1):
        if not ws.cell(row=5, column=c).value:
            bad(f"{ws.title}: คอลัมน์ {get_column_letter(c)} ไม่มีหัวตาราง")

print("\n=== 6. สูตรใน sheet 3 อ้างอิงคอลัมน์ sheet 1 ถูกตัวมั้ย ===")
# คอลัมน์ sheet1 ที่ควรถูกอ้างถึงในแต่ละช่องคะแนนย่อย
expect = {4: "I", 5: "J", 7: "L", 8: "M", 9: "O", 11: "P", 12: "Q", 14: "R", 15: "S", 16: "T"}
for c, src in expect.items():
    f = w3.cell(row=FIRST, column=c).value or ""
    head = (w3.cell(row=5, column=c).value or "").replace("\n", " ")
    refs = set(re.findall(r"'1-UseCase ทั้งหมด'!([A-Z]+)\d+", f))
    if refs != {"B", src}:
        bad(f"sheet3 คอลัมน์ {get_column_letter(c)} ({head}): อ้างถึง {sorted(refs)} "
            f"แต่ควรเป็น ['B','{src}']")
    src_head = (w1[f"{src}5"].value or "").split("\n")[0]
    ok(f"{get_column_letter(c)} ({head}) <- sheet1 {src} ({src_head})")

print("\n=== 7. สูตรรวม / น้ำหนัก / เกณฑ์ ===")
tot = w3.cell(row=FIRST, column=18).value
for name, ref in [("Impact", "AI$4"), ("Feasibility", "AI$5"),
                  ("Data", "AI$6"), ("Risk", "AI$7")]:
    (ok if ref in tot else bad)(f"สูตรคะแนนรวมอ้างน้ำหนัก {name} ที่ DATA!${ref}")
wsum = sum(wd[f"AI{r}"].value for r in range(4, 8))
(ok if abs(wsum - 1.0) < 1e-9 else bad)(f"น้ำหนักรวม = {wsum:.0%} (ต้องเป็น 100%)")
lv = w3.cell(row=FIRST, column=19).value
for ref in ("AI$10", "AI$11"):
    (ok if ref in lv else bad)(f"สูตรระดับอ้างเกณฑ์ที่ DATA!${ref}")
(ok if wd["AI10"].value > wd["AI11"].value else bad)(
    f"เกณฑ์ สูง({wd['AC10'].value}) > กลาง({wd['AC11'].value})")
(ok if "(6-Q" in tot.replace(" ", "") else bad)("Risk ถูกกลับด้านด้วย (6 - Risk)")
(ok if ")*20," in tot else bad)("มีการคูณ 20 เพื่อขยายเป็นสเกล 100")

print("\n=== 8. การตรวจจับข้อมูลไม่ครบ ===")
mf = w3.cell(row=FIRST, column=21).value or ""
got = re.findall(r"COUNTBLANK\('1-UseCase ทั้งหมด'!([A-Z]+)\d+\)", mf)
need = list("IJLMOPQRST")
(ok if got == need else bad)(f"ช่องบังคับที่นับ = {got} (ต้องเป็น {need})")
for c, key in ((18, "คะแนนรวม"), (20, "อันดับ")):
    f = w3.cell(row=FIRST, column=c).value or ""
    (ok if f"U{FIRST}" in f or "R" + str(FIRST) in f else bad)(f"{key} มีเงื่อนไขกันข้อมูลไม่ครบ")
(ok if "ข้อมูลไม่ครบ" in (w3.cell(row=FIRST, column=19).value or "") else bad)(
    "ช่องระดับขึ้นข้อความ 'ข้อมูลไม่ครบ'")

print("\n=== 9. ลิงก์ข้ามหน้า ===")
for r in (6, FIRST, LAST):
    for c, src in ((1, "A"), (2, "B"), (3, "D")):
        f = w2.cell(row=r, column=c).value or ""
        if f"'1-UseCase ทั้งหมด'!{src}{r}" not in f:
            bad(f"sheet2 {get_column_letter(c)}{r} ไม่ได้ลิงก์ไป sheet1 {src}{r}")
ok("sheet2 คอลัมน์ A/B/C ลิงก์จาก sheet1 ครบ (ตรวจแถว 6, 7, 36)")

print("\n=== 10. ช่องที่เป็นสูตร ต้องไม่ถูกทำเป็นช่องกรอก ===")
n_formula = sum(1 for r in range(FIRST, LAST + 1) for c in range(1, 22)
                if isinstance(w3.cell(row=r, column=c).value, str)
                and w3.cell(row=r, column=c).value.startswith("="))
exp = (LAST - FIRST + 1) * 21
(ok if n_formula == exp else bad)(f"sheet3 มีสูตร {n_formula} ช่อง (คาด {exp})")

print("\n=== 11. ข้อความในเซลล์ที่ขึ้นต้นด้วย '=' (Excel จะมองเป็นสูตร) ===")
hits = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("=") and not re.match(r"^=[A-Z(]", v):
                bad(f"{ws.title}!{c.coordinate} ขึ้นต้นด้วย '=' : {v[:40]}"); hits += 1
if not hits: ok("ไม่พบข้อความที่จะถูกตีความผิดเป็นสูตร")

print("\n" + "=" * 60)
print(f"สรุป: FAIL {len(fails)} รายการ / WARN {len(warns)} รายการ")
sys.exit(1 if fails else 0)
