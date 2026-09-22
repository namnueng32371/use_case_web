# -*- coding: utf-8 -*-
"""สร้าง Excel Template เก็บ AI Use Case (v2)
   sheet: คู่มือ / 1-UseCase ทั้งหมด / 2-รายละเอียดเทคนิค / 3-คะแนนและลำดับ / DATA
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

OUT = r"C:\Users\Jengza\Desktop\use_case_web\docs\AI_UseCase_Template_v2.xlsx"
FIRST, LAST = 7, 36          # แถวข้อมูลจริง
S1 = "'1-UseCase ทั้งหมด'"   # ชื่อ sheet มีเว้นวรรค ต้องครอบ '

FONT = "Tahoma"
F_TITLE = Font(name=FONT, size=14, bold=True, color="FFFFFF")
F_GROUP = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_HEAD  = Font(name=FONT, size=9,  bold=True)
F_BODY  = Font(name=FONT, size=9)
F_LINK  = Font(name=FONT, size=9, color="008000")   # เขียว = ลิงก์มาจาก sheet อื่น
F_CALC  = Font(name=FONT, size=9, bold=True)
F_EX    = Font(name=FONT, size=9, italic=True, color="808080")
F_NOTE  = Font(name=FONT, size=9, color="C00000", bold=True)
F_H2    = Font(name=FONT, size=11, bold=True)
F_BLUE  = Font(name=FONT, size=10, bold=True, color="0000FF")

FILL_TITLE = PatternFill("solid", fgColor="1F3864")
FILL_BASIC = PatternFill("solid", fgColor="2E75B6")
FILL_IMPACT= PatternFill("solid", fgColor="C55A11")
FILL_DATA  = PatternFill("solid", fgColor="538135")
FILL_FEAS  = PatternFill("solid", fgColor="7030A0")
FILL_RISK  = PatternFill("solid", fgColor="BF3131")
FILL_PLAN  = PatternFill("solid", fgColor="595959")
FILL_TECH  = PatternFill("solid", fgColor="0F6E6E")
FILL_TOTAL = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")   # เหลือง = กรอกเอง
FILL_AUTO  = PatternFill("solid", fgColor="E2EFDA")   # เขียวอ่อน = ระบบเติมให้
FILL_EX    = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_T = Alignment(wrap_text=True, vertical="top")
WRAP_C = Alignment(wrap_text=True, vertical="center", horizontal="center")

wb = Workbook()

# ============================================================ DATA (ตัวเลือก)
ws_d = wb.create_sheet("DATA")
# (ตัวเลือก, คะแนน)  คะแนน None = ไม่ใช้คำนวณ
LISTS = {
    "A": ("บริษัท", [("S.K.Polymer", None), ("Thairubbtech", None), ("Polymate", None),
                     ("GTK", None), ("อื่นๆ", None)]),
    "B": ("ตอนนี้ข้อมูลเก็บอยู่ในรูปแบบไหน", [
        ("กระดาษ / เอกสารเขียนมือ - ต้องพิมพ์เข้าคอมก่อน", 1),
        ("ไฟล์ Excel / CSV - เปิดใช้ได้เลย", 4),
        ("ฐานข้อมูลในระบบ (ERP/MES/SQL) - ดึงอัตโนมัติได้", 5),
        ("รูปภาพ / ภาพถ่าย", 3),
        ("เสียง / วิดีโอ", 2),
        ("ยังไม่มีข้อมูลเก็บไว้ - ต้องเริ่มเก็บใหม่", 1)]),
    "C": ("มีข้อมูลย้อนหลังนานแค่ไหน", [
        ("ยังไม่มี - ต้องเริ่มเก็บใหม่", 1),
        ("น้อยกว่า 1 ปี - อาจยังไม่พอเทรน", 2),
        ("1 - 3 ปี - พอใช้ได้", 4),
        ("มากกว่า 3 ปี - เพียงพอแน่นอน", 5)]),
    "D": ("สิทธิ์นำข้อมูลมาใช้", [
        ("ได้ทันที - ข้อมูลอยู่ในมือแผนกเราเอง", 5),
        ("ต้องขออนุมัติก่อน - เป็นของแผนกอื่น / IT", 3),
        ("ขอไม่ได้ / ยังไม่แน่ใจ - เป็นของลูกค้าหรือ vendor", 1)]),
    "E": ("ต้องเชื่อมต่อกับระบบไอทีเดิมมั้ย", [
        ("ไม่มี - AI ทำงานแยกเดี่ยวได้เลย ไม่ต้องยุ่งกับระบบไหน", 5),
        ("มี - แค่ export ไฟล์ออกมาให้ AI แล้วเอาผลกลับเข้าไปทีหลังก็พอ", 3),
        ("มี - AI ต้องอ่าน/เขียนข้อมูลในระบบโดยตรงแบบทันที", 1),
        ("ยังไม่แน่ใจ - ต้องถาม IT ก่อน", 2)]),
    "F": ("มีคนดูแลระบบต่อหลังส่งมอบมั้ย", [
        ("มี - มีคนในแผนกรับดูแลต่อได้ (ระบุชื่อในหมายเหตุ)", 5),
        ("ยังไม่แน่ใจ - ต้องหาคนก่อน", 3),
        ("ไม่มี - ต้องพึ่งคนนอกแผนกตลอด", 1)]),
    "G": ("ถ้า AI ตอบผิด จะเกิดอะไรขึ้น", [
        ("ต่ำ - คนเห็นแล้วแก้ได้ทันที ไม่มีใครเดือดร้อน", 1),
        ("กลาง - เสียเวลาทำใหม่ หรือเสียต้นทุนบางส่วน", 3),
        ("สูง - ของเสียหลุดถึงลูกค้า / อันตราย / ผิดกฎหมาย", 5)]),
    "H": ("ต้องให้คนตรวจงานที่ AI ทำซ้ำมั้ย", [
        ("ไม่ต้อง - ใช้ผล AI ได้เลย", 1),
        ("สุ่มตรวจบางส่วน - ตรวจเฉพาะที่ AI ไม่มั่นใจ", 2),
        ("ตรวจซ้ำทุกครั้ง - คนต้องดูทุกชิ้น/ทุกรายการ", 4)]),
    "I": ("ข้อมูลชุดนี้ลับแค่ไหน", [
        ("สาธารณะ - หลุดออกไปก็ไม่เป็นไร", 1),
        ("ข้อมูลภายใน - ห้ามออกนอกบริษัท", 2),
        ("ลับ - ต้องจำกัดสิทธิ์ว่าใครเปิดดูได้บ้าง", 4),
        ("ลับมาก - มีข้อมูลส่วนบุคคล / ผูกกับกฎหมาย PDPA", 5)]),
    "J": ("ความสำคัญที่แผนกอยากได้", [("สูง - ต้องการด่วน", None), ("กลาง - (Q2 - Q3)", None),
                                      ("ต่ำ - Nice to have", None)]),
    "K": ("ประเภทข้อมูล Input", [("ข้อความ / เอกสาร", None), ("รูปภาพ", None), ("เสียง", None),
                                 ("วิดีโอ", None), ("ตัวเลข / ตาราง", None), ("ผสมหลายแบบ", None)]),
    "L": ("ภาษาที่ต้องการ", [("ไทย", None), ("อังกฤษ", None), ("ไทย + อังกฤษ", None),
                             ("ไม่เกี่ยวกับภาษา", None)]),
    "M": ("Latency ที่รับได้", [("Real-time (0 - 2 วินาที)", None),
                                ("Interactive (2 - 10 วินาที)", None),
                                ("Batch (นาที - ชั่วโมง)", None)]),
    "N": ("Availability", [("Business Hour (8:00 - 18:00)", None),
                           ("Extended (6:00 - 22:00)", None), ("24 ชม. x 7 วัน", None)]),
    "O": ("ความละเอียดรูป", [("ไม่เกี่ยวข้อง", None), ("ต่ำกว่า 1 MP", None), ("1 - 5 MP", None),
                             ("5 - 12 MP", None), ("มากกว่า 12 MP", None)]),
    "P": ("ที่ตั้งระบบ", [("On-premise เท่านั้น (ห้ามออกนอก)", None),
                          ("ใช้ Cloud / API ภายนอกได้", None),
                          ("ผสม (ข้อมูลลับอยู่ใน, ที่เหลือออกได้)", None)]),
    "Q": ("ความต้องการโมเดล", [("ใช้โมเดลสำเร็จรูปได้", None),
                               ("ต้อง Fine-tune ด้วยข้อมูลบริษัท", None), ("ยังไม่แน่ใจ", None)]),
    "R": ("เวลาที่ใช้กับงานนี้ตอนนี้", [
        ("น้อยกว่า 4 ชม./สัปดาห์  (ไม่ถึงครึ่งวันทำงาน)", 1),
        ("4 - 10 ชม./สัปดาห์  (ประมาณ 1 วันทำงาน)", 2),
        ("10 - 20 ชม./สัปดาห์  (ประมาณ 2 วันทำงาน)", 3),
        ("20 - 40 ชม./สัปดาห์  (เท่ากับใช้คน 1 คนเต็มเวลา)", 4),
        ("มากกว่า 40 ชม./สัปดาห์  (มากกว่า 1 คนเต็มเวลา)", 5)]),
    # "-" = ไม่นำมาคิดคะแนน (Impact จะใช้เฉพาะชั่วโมงแทน) ไม่ใช่ได้ 0 คะแนน
    "S": ("มูลค่าความเสียหาย/ของเสียต่อปี", [
        ("ประเมินเป็นเงินไม่ได้ / ไม่มีความเสียหายโดยตรง", None),
        ("น้อยกว่า 1 แสนบาท/ปี", 2),
        ("1 - 5 แสนบาท/ปี", 3),
        ("5 แสน - 2 ล้านบาท/ปี", 4),
        ("มากกว่า 2 ล้านบาท/ปี", 5)]),
    "T": ("จำนวน Users/วัน", [
        ("1 - 10 คน", None), ("11 - 50 คน", None), ("51 - 200 คน", None),
        ("201 - 1,000 คน", None), ("มากกว่า 1,000 คน", None)]),
    "U": ("Peak concurrent (คนใช้พร้อมกันสูงสุด)", [
        ("1 - 2 คน", None), ("3 - 5 คน", None), ("6 - 20 คน", None),
        ("21 - 50 คน", None), ("มากกว่า 50 คน", None)]),
    "V": ("Requests/วัน", [
        ("น้อยกว่า 100 ครั้ง", None), ("100 - 1,000 ครั้ง", None),
        ("1,000 - 10,000 ครั้ง", None), ("มากกว่า 10,000 ครั้ง", None)]),
    "W": ("ขนาดข้อมูล Input/วัน", [
        ("น้อยกว่า 1 GB", None), ("1 - 10 GB", None),
        ("10 - 100 GB", None), ("มากกว่า 100 GB", None)]),
    "X": ("ขนาดคลังเอกสารทั้งหมด (RAG)", [
        ("ไม่เกี่ยวข้อง - ไม่ใช่งานค้นหาเอกสาร", None), ("น้อยกว่า 1 GB", None),
        ("1 - 10 GB", None), ("10 - 100 GB", None), ("มากกว่า 100 GB", None)]),
    "Y": ("จำนวนรูป/วัน", [
        ("ไม่เกี่ยวข้อง - ไม่ใช่งานภาพ", None), ("น้อยกว่า 100 รูป", None),
        ("100 - 1,000 รูป", None), ("1,000 - 10,000 รูป", None),
        ("มากกว่า 10,000 รูป", None)]),
}
SCORE_OFFSET = 30   # คะแนนของแต่ละตัวเลือกอยู่ที่แถว 30+ คอลัมน์เดียวกัน

ws_d.sheet_view.showGridLines = False
ws_d["A1"] = "รายการตัวเลือก (Dropdown) — เพิ่ม/แก้ตัวเลือกได้ที่นี่ แล้วทุก sheet จะเปลี่ยนตาม"
ws_d["A1"].font = Font(name=FONT, size=11, bold=True, color="C00000")
for col, (head, items) in LISTS.items():
    h = ws_d[f"{col}3"]; h.value = head; h.font = F_HEAD
    h.fill = PatternFill("solid", fgColor="D9D9D9"); h.alignment = WRAP_C
    for i, (v, _s) in enumerate(items, start=4):
        c = ws_d[f"{col}{i}"]; c.value = v; c.font = F_BODY
    ws_d.column_dimensions[col].width = 44 if col == "R" else 32

ws_d[f"A{SCORE_OFFSET-2}"] = "ตารางแปลงตัวเลือกเป็นคะแนน 1-5 (เรียงตรงกับรายการด้านบนบรรทัดต่อบรรทัด) — ปรับได้"
ws_d[f"A{SCORE_OFFSET-2}"].font = Font(name=FONT, size=11, bold=True, color="C00000")
for col, (head, items) in LISTS.items():
    h = ws_d.cell(row=SCORE_OFFSET - 1, column=ws_d[f"{col}3"].column, value=head)
    h.font = F_HEAD; h.fill = PatternFill("solid", fgColor="D9D9D9"); h.alignment = WRAP_C
    for i, (_v, s) in enumerate(items, start=SCORE_OFFSET):
        c = ws_d[f"{col}{i}"]
        c.value = s if s is not None else "-"
        c.font = F_BLUE if s is not None else F_BODY
        c.alignment = Alignment(horizontal="center")

# --- น้ำหนักคะแนน + เกณฑ์ + ช่วงตัวเลข ------------------------------------
def put(cell, val, font=None, fmt=None):
    c = ws_d[cell]; c.value = val
    c.font = font or F_BODY
    if fmt: c.number_format = fmt

ws_d["AB1"] = "ค่าตั้งต้นของการคำนวณคะแนน (แก้ได้ ทุกสูตรจะเปลี่ยนตาม)"
ws_d["AB1"].font = Font(name=FONT, size=11, bold=True, color="C00000")
put("AB3", "น้ำหนักแต่ละปัจจัย", F_HEAD)
for r, (lbl, val) in enumerate([("Impact (ผลกระทบ)", 0.35), ("Feasibility (ความเป็นไปได้)", 0.25),
                                ("Data (ความพร้อมข้อมูล)", 0.25), ("Risk (ความเสี่ยง)", 0.15)], start=4):
    put(f"AB{r}", lbl)
    put(f"AC{r}", val, F_BLUE, "0%")
put("AB9", "เกณฑ์แบ่งระดับ", F_HEAD)
put("AB10", "คะแนน >= นี้ = สูง");   put("AC10", 75, F_BLUE)
put("AB11", "คะแนน >= นี้ = กลาง"); put("AC11", 50, F_BLUE)

put("AB14", "หมายเหตุ", F_HEAD)
put("AB15", "ช่วงชั่วโมงและช่วงมูลค่าเป็น dropdown แล้ว")
put("AB16", "แก้ตัวเลือกและคะแนนได้ที่คอลัมน์ R และ S")
put("AB17", 'คะแนน "-" แปลว่าไม่นำมาคิดคะแนน ไม่ใช่ได้ 0')
ws_d.column_dimensions["AB"].width = 34
ws_d.column_dimensions["AC"].width = 12

W_IMPACT, W_FEAS = "DATA!$AC$4", "DATA!$AC$5"
W_DATA, W_RISK   = "DATA!$AC$6", "DATA!$AC$7"
TH_HIGH, TH_MID  = "DATA!$AC$10", "DATA!$AC$11"

def opt_ref(col):
    return f"DATA!${col}$4:${col}${3+len(LISTS[col][1])}"

def score_ref(col):
    n = len(LISTS[col][1])
    return f"DATA!${col}${SCORE_OFFSET}:${col}${SCORE_OFFSET+n-1}"

def tip(ws, col, title, text):
    """กล่องคำอธิบายที่เด้งขึ้นตอนคลิกช่อง (Input Message) — ไม่จำกัดค่าที่กรอก"""
    v = DataValidation(type=None, allow_blank=True, showErrorMessage=False)
    v.promptTitle = title
    v.prompt = text
    v.showInputMessage = True
    ws.add_data_validation(v)
    v.add(f"{col}{FIRST}:{col}{LAST}")


def dv(col):
    """dropdown: formula1 ห้ามมี '=' นำหน้า, showDropDown=False = ให้แสดงลูกศร"""
    v = DataValidation(type="list", formula1=opt_ref(col), allow_blank=True,
                       showDropDown=False, showErrorMessage=True)
    v.errorTitle = "ค่าไม่ถูกต้อง"
    v.error = "กรุณาเลือกจากรายการเท่านั้น (กดลูกศรด้านขวาของช่อง)"
    v.promptTitle = LISTS[col][0]
    v.prompt = "เลือกจากรายการ — กดลูกศรด้านขวาของช่อง"
    v.showInputMessage = True
    return v

# ------------------------------------------------- helper สร้าง sheet ตาราง
def build_sheet(ws, title, note, groups, cols, example, auto_cols=()):
    ws.sheet_view.showGridLines = False
    ncol = len(cols)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    t = ws.cell(row=1, column=1, value=title)
    t.font = F_TITLE; t.fill = FILL_TITLE; t.alignment = WRAP_C
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    n = ws.cell(row=2, column=1, value=note)
    n.font = F_NOTE; n.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 22

    c0 = 1
    for gname, gspan, gfill in groups:
        if gspan > 1:
            ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + gspan - 1)
        g = ws.cell(row=4, column=c0, value=gname)
        g.font = F_GROUP; g.alignment = WRAP_C
        for k in range(c0, c0 + gspan):
            ws.cell(row=4, column=k).fill = gfill
            ws.cell(row=4, column=k).border = BORDER
        c0 += gspan
    ws.row_dimensions[4].height = 20

    for i, (head, width, _dvcol) in enumerate(cols, start=1):
        c = ws.cell(row=5, column=i, value=head)
        c.font = F_HEAD; c.alignment = WRAP_C; c.border = BORDER
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[5].height = 48

    if example:
        for i, v in enumerate(example, start=1):
            c = ws.cell(row=6, column=i, value=v)
            c.font = F_EX; c.fill = FILL_EX; c.alignment = WRAP_T; c.border = BORDER
        ws.row_dimensions[6].height = 58

    for r in range(FIRST, LAST + 1):
        ws.row_dimensions[r].height = 40
        for i in range(1, ncol + 1):
            c = ws.cell(row=r, column=i)
            c.font = F_LINK if i in auto_cols else F_BODY
            c.fill = FILL_AUTO if i in auto_cols else FILL_INPUT
            c.alignment = WRAP_T; c.border = BORDER

    for i, (_h, _w, dvcol) in enumerate(cols, start=1):
        if dvcol:
            v = dv(dvcol); ws.add_data_validation(v)
            L = get_column_letter(i)
            v.add(f"{L}{FIRST}:{L}{LAST}")
    ws.freeze_panes = "C6"

# ===================================== SHEET 1 : ทุก Use Case (คนกรอก)
ws1 = wb.active
ws1.title = "1-UseCase ทั้งหมด"
groups1 = [("ข้อมูลพื้นฐาน", 6, FILL_BASIC), ("ผลกระทบ (Impact)", 4, FILL_IMPACT),
           ("ความพร้อมข้อมูล (Data)", 5, FILL_DATA), ("ความเป็นไปได้ (Feasibility)", 2, FILL_FEAS),
           ("ความเสี่ยง (Risk)", 3, FILL_RISK), ("แผน", 3, FILL_PLAN)]
cols1 = [
    ("No.", 6, None),
    ("ชื่อ Use Case", 30, None),
    ("บริษัท", 16, "A"),
    ("ฝ่าย / แผนก", 16, None),
    ("ผู้รับผิดชอบ\n(Owner)", 14, None),
    ("ปัญหาที่ต้องการแก้\n(ทำไมถึงอยากได้ AI)", 38, None),
    ("วิธีการทำงานปัจจุบัน\nExisting Work\n(ตอนนี้ทำยังไง ใครทำ)", 38, None),
    ("แนวทางหลังนำ AI มาใช้\nDescription for Redesign\n(อยากให้ AI ทำอะไร ผลลัพธ์ที่คาดหวัง)", 38, None),
    ("เวลาที่ใช้กับงานนี้ตอนนี้ (ยังไม่ใช้ AI)\nนับรวมทุกคนในแผนก เช่น 3 คน x 6 ชม. = 18 ชม./สัปดาห์", 34, "R"),
    ("มูลค่าความเสียหาย/ของเสียต่อปี\n(ของเสีย เคลม งานแก้ ค่าปรับ)", 34, "S"),
    ("ข้อมูลที่มีอยู่ตอนนี้\nData Available\n(มีอะไร เก็บที่ไหน ใครเก็บ)", 38, None),
    ("ตอนนี้ข้อมูลเก็บอยู่ในรูปแบบไหน", 34, "B"),
    ("มีข้อมูลย้อนหลังนานแค่ไหน", 28, "C"),
    ("ปริมาณข้อมูลโดยประมาณ\nนับเป็นรายการ : รูป / ใบ / ไฟล์ / ครั้ง / ล็อต แล้วแต่งาน\n(เอาเฉพาะที่รู้คำตอบแล้ว)", 38, None),
    ("สิทธิ์นำข้อมูลมาใช้\n(เจ้าของข้อมูลยอมให้เอามาพัฒนา AI มั้ย)", 34, "D"),
    ("ต้องเชื่อมต่อกับระบบไอทีเดิมมั้ย\n(ERP / MES / ระบบที่ใช้อยู่)", 40, "E"),
    ("มีคนดูแลระบบต่อหลังส่งมอบมั้ย", 34, "F"),
    ("ถ้า AI ตอบผิด จะเกิดอะไรขึ้น", 36, "G"),
    ("ต้องให้คนตรวจงานที่ AI ทำซ้ำมั้ย", 34, "H"),
    ("ข้อมูลชุดนี้ลับแค่ไหน", 34, "I"),
    ("ความสำคัญที่แผนกอยากได้\n(ความเห็นแผนก ไม่นำไปคิดคะแนน)", 26, "J"),
    ("ต้องการใช้งานภายใน\n(ไตรมาส/ปี)", 16, None),
    ("หมายเหตุ", 26, None),
]
assert sum(g[1] for g in groups1) == len(cols1)
ex1 = [1, "Auto Inspection ตรวจคุณภาพชิ้นงานด้วย Machine Learning",
       "S.K.Polymer", "Production", "คุณสมชาย",
       "พนักงาน QC ตรวจชิ้นงานด้วยตาเปล่า เหนื่อยล้าตอนบ่าย ทำให้ของเสียหลุดไปถึงลูกค้าเดือนละ 2-3 ครั้ง",
       "QC 3 คน ผลัดกันตรวจทุกล็อตด้วยตาเปล่า เทียบกับตัวอย่างมาตรฐาน บันทึกผลลงกระดาษ",
       "ติดกล้องที่สายพาน ให้ AI แยกชิ้นงานดี/เสียอัตโนมัติแบบ Real-time "
       "ชิ้นที่ AI ไม่มั่นใจค่อยส่งให้คนตรวจ และบันทึกผลลงระบบแทนกระดาษ",
       "มากกว่า 40 ชม./สัปดาห์  (มากกว่า 1 คนเต็มเวลา)",
       "5 แสน - 2 ล้านบาท/ปี",
       "มีรูปถ่ายชิ้นงานดี/เสียเก็บไว้ในมือถือหัวหน้ากะ ประมาณ 2 ปี และใบบันทึกผลตรวจเป็นกระดาษ",
       "รูปภาพ / ภาพถ่าย", "1 - 3 ปี - พอใช้ได้", "ประมาณ 4,000 รูป (ดี 3,000 / เสีย 1,000)",
       "ได้ทันที - ข้อมูลอยู่ในมือแผนกเราเอง",
       "ไม่มี - AI ทำงานแยกเดี่ยวได้เลย ไม่ต้องยุ่งกับระบบไหน",
       "มี - มีคนในแผนกรับดูแลต่อได้ (ระบุชื่อในหมายเหตุ)",
       "สูง - ของเสียหลุดถึงลูกค้า / อันตราย / ผิดกฎหมาย",
       "สุ่มตรวจบางส่วน - ตรวจเฉพาะที่ AI ไม่มั่นใจ",
       "ข้อมูลภายใน - ห้ามออกนอกบริษัท", "สูง - ต้องการด่วน", "Q2 2570",
       "คุณสมชายเคยอบรม Python มาแล้ว"]
build_sheet(
    ws1,
    "ชั้นที่ 1 : เก็บ Use Case ทั้งหมด   (ทุก Use Case ต้องกรอก — ระบบใช้ข้อมูลนี้คำนวณคะแนนใน sheet '3-คะแนนและลำดับ')",
    "วิธีกรอก : กรอกเฉพาะช่องสีเหลืองตั้งแต่แถวที่ 7 ลงไป  |  แถวสีเทา (แถวที่ 6) คือตัวอย่าง อย่าลบ  |  "
    "ช่องที่มีลูกศรต้องเลือกจากรายการเท่านั้น เพราะระบบอ่านค่าไปแปลงเป็นคะแนน",
    groups1, cols1, ex1)

# กล่องคำอธิบายของช่องพิมพ์อิสระ (คลิกช่องแล้วเด้งขึ้น)
tip(ws1, "F", "ปัญหาที่ต้องการแก้",
    "ตอนนี้เจอปัญหาอะไร ทำไมถึงอยากได้ AI มาช่วย\nเช่น ตรวจด้วยตาเปล่าทำให้ของเสียหลุดถึงลูกค้าเดือนละ 2-3 ครั้ง")
tip(ws1, "G", "วิธีการทำงานปัจจุบัน (Existing Work)",
    "ตอนนี้ทำงานนี้ยังไง ใครเป็นคนทำ ใช้เครื่องมืออะไร บันทึกผลไว้ที่ไหน\nเขียนแบบที่คนนอกแผนกอ่านแล้วเห็นภาพ")
tip(ws1, "H", "แนวทางหลังนำ AI มาใช้ (Description for Redesign)",
    "อยากให้ AI ทำอะไรให้ และผลลัพธ์ที่อยากได้หน้าตาเป็นยังไง\nเช่น ติดกล้องที่สายพาน ให้ AI แยกดี/เสียอัตโนมัติ\nชิ้นที่ AI ไม่มั่นใจค่อยส่งให้คนตรวจ")
tip(ws1, "K", "ข้อมูลที่มีอยู่ตอนนี้ (Data Available)",
    "มีข้อมูลอะไรบ้าง เก็บไว้ที่ไหน ใครเป็นคนเก็บ\nถ้ามีสัดส่วนที่น่าสนใจให้เขียนด้วย เช่น รูปดี 3,000 / รูปเสีย 1,000")
tip(ws1, "N", "ปริมาณข้อมูลโดยประมาณ",
    "นับเป็น 'รายการ' โดยหน่วยขึ้นกับประเภทงาน :\n   งานตรวจภาพ         =  กี่รูป\n   หาสาเหตุปัญหา/NC   =  กี่เคส กี่ใบ\n   ค้นหา-สรุปเอกสาร   =  กี่ไฟล์ กี่ฉบับ\n   ทำนายเครื่องเสีย   =  เคยเสียมากี่ครั้ง\n   วิเคราะห์การผลิต   =  กี่ล็อต กี่แถวข้อมูล\n   จัดคน / แนะนำ      =  เคยจัดจริงมากี่ครั้ง\n\nนับเฉพาะข้อมูลที่รู้คำตอบแล้ว เช่น รูปที่มาร์คแล้วว่าดี/เสีย\nข้อมูลดิบที่ยังไม่ได้แยกใช้เทรนไม่ได้")

# ======================== SHEET 2 : รายละเอียดเทคนิค (ทุก use case, ลิงก์ชื่อ)
ws2 = wb.create_sheet("2-รายละเอียดเทคนิค")
groups2 = [("ลิงก์อัตโนมัติจากชั้นที่ 1", 3, FILL_BASIC), ("ปริมาณงาน (Workload)", 6, FILL_IMPACT),
           ("ความต้องการเชิงเทคนิค", 6, FILL_TECH), ("ข้อจำกัด / งบประมาณ", 4, FILL_PLAN)]
cols2 = [
    ("No.", 6, None),
    ("ชื่อ Use Case", 32, None),
    ("ฝ่าย / แผนก", 16, None),
    ("ประเภทข้อมูล Input", 22, "K"),
    ("จำนวน Users/วัน", 20, "T"),
    ("Peak concurrent\n(คนใช้พร้อมกันสูงสุด)", 20, "U"),
    ("Requests/วัน\n(ประมาณ)", 22, "V"),
    ("ขนาดข้อมูล Input/วัน", 20, "W"),
    ("ขนาดคลังเอกสารทั้งหมด\n(สำหรับงานค้นหา/RAG)", 30, "X"),
    ("จำนวนรูป/วัน\n(งานภาพเท่านั้น)", 26, "Y"),
    ("ความละเอียดรูป", 18, "O"),
    ("ภาษาที่ต้องการ", 18, "L"),
    ("Latency ที่รับได้", 24, "M"),
    ("Availability ที่ต้องการ", 24, "N"),
    ("ต้องติดตั้งที่ไหน", 28, "P"),
    ("ความต้องการโมเดล", 26, "Q"),
    ("งบประมาณที่ตั้งไว้ (บาท)\n- ถ้ามี", 18, None),
    ("ข้อจำกัดอื่นๆ / หมายเหตุ", 30, None),
    ("ผู้ให้ข้อมูลเทคนิค", 16, None),
]
assert sum(g[1] for g in groups2) == len(cols2)
ex2 = [1, "Auto Inspection ตรวจคุณภาพชิ้นงานด้วย Machine Learning", "Production",
       "รูปภาพ", "1 - 10 คน", "1 - 2 คน", "100 - 1,000 ครั้ง", "1 - 10 GB",
       "ไม่เกี่ยวข้อง - ไม่ใช่งานค้นหาเอกสาร", "100 - 1,000 รูป",
       "1 - 5 MP", "ไม่เกี่ยวกับภาษา",
       "Real-time (0 - 2 วินาที)", "Business Hour (8:00 - 18:00)",
       "On-premise เท่านั้น (ห้ามออกนอก)", "ต้อง Fine-tune ด้วยข้อมูลบริษัท",
       500000, "ต้องติดกล้องที่สายพานเดิม มีพื้นที่จำกัด", "คุณสมชาย"]
build_sheet(
    ws2,
    "ชั้นที่ 2 : รายละเอียดเทคนิค   (กรอกทุก Use Case — ใช้กำหนด Spec Hardware และประเมินราคา Server)",
    "ช่องสีเขียวอ่อน (No. / ชื่อ / แผนก) ระบบดึงมาจากชั้นที่ 1 ให้อัตโนมัติ ห้ามพิมพ์ทับ — "
    "ถ้าอยากแก้ชื่อให้ไปแก้ที่ sheet '1-UseCase ทั้งหมด'  |  ถ้าช่องไหนไม่เกี่ยวข้องให้ใส่ 0",
    groups2, cols2, ex2, auto_cols=(1, 2, 3))

# ลิงก์ No. / ชื่อ / แผนก จากชั้นที่ 1  (รวมแถวตัวอย่าง แถว 6 ด้วย เพื่อให้เห็นว่าลิงก์ทำงานยังไง)
for r in [6] + list(range(FIRST, LAST + 1)):
    ws2.cell(row=r, column=1).value = f'=IF({S1}!B{r}="","",{S1}!A{r})'
    ws2.cell(row=r, column=2).value = f'=IF({S1}!B{r}="","",{S1}!B{r})'
    ws2.cell(row=r, column=3).value = f'=IF({S1}!B{r}="","",{S1}!D{r})'
    if r == 6:   # แถวตัวอย่าง: ให้ 3 ช่องแรกเป็นสีเขียว (ลิงก์) เหมือนแถวจริง
        for col in (1, 2, 3):
            c = ws2.cell(row=6, column=col)
            c.fill = FILL_AUTO
            c.font = Font(name=FONT, size=9, italic=True, color="008000")

# ================================== SHEET 3 : คะแนนและลำดับ (คำนวณอัตโนมัติ)
ws3 = wb.create_sheet("3-คะแนนและลำดับ")
groups3 = [("ข้อมูล Use Case", 3, FILL_BASIC), ("Impact 35%", 3, FILL_IMPACT),
           ("Data 25%", 4, FILL_DATA), ("Feasibility 25%", 3, FILL_FEAS),
           ("Risk 15% (กลับด้าน)", 4, FILL_RISK), ("ผลลัพธ์", 4, FILL_TOTAL)]
cols3 = [
    ("No.", 6, None), ("ชื่อ Use Case", 34, None), ("ฝ่าย / แผนก", 16, None),
    ("ชม./สัปดาห์\n-> คะแนน", 11, None), ("มูลค่าเสียหาย\n-> คะแนน", 11, None),
    ("Impact\n(1-5)", 10, None),
    ("รูปแบบ\nข้อมูล", 10, None), ("ย้อนหลัง", 10, None), ("สิทธิ์ใช้\nข้อมูล", 10, None),
    ("Data\n(1-5)", 10, None),
    ("ระบบเดิม", 10, None), ("คนดูแล\nต่อ", 10, None), ("Feasibility\n(1-5)", 11, None),
    ("AI ผิด", 9, None), ("ตรวจซ้ำ", 9, None), ("ความลับ", 9, None), ("Risk\n(1-5)", 10, None),
    ("คะแนนรวม\n(20-100)", 13, None), ("ระดับ", 16, None), ("อันดับ", 9, None),
    ("ช่องบังคับที่ยัง\nไม่ได้กรอก\n(0 = ครบ)", 14, None),
]
assert sum(g[1] for g in groups3) == len(cols3)
build_sheet(
    ws3,
    "ชั้นที่ 3 : คะแนนและลำดับความสำคัญ   (ระบบคำนวณให้อัตโนมัติ — ห้ามพิมพ์ทับ)",
    "ทุกช่องในหน้านี้เป็นสูตร ดึงข้อมูลจาก sheet '1-UseCase ทั้งหมด' อัตโนมัติ  |  "
    "แถวสีเทา (แถวที่ 6) คือคะแนนของ Use Case ตัวอย่าง ไม่นับเข้าการจัดอันดับ  |  "
    "ถ้าอยากปรับน้ำหนักคะแนนหรือเกณฑ์แบ่งระดับ ให้ไปแก้ที่ sheet 'DATA' คอลัมน์ S-T (ตัวเลขสีน้ำเงิน)",
    groups3, cols3, example=None, auto_cols=tuple(range(1, len(cols3) + 1)))

def lk(col, r):
    """INDEX/MATCH แปลงตัวเลือก dropdown -> คะแนน"""
    return (f'IFERROR(INDEX({score_ref(col)},'
            f'MATCH({S1}!{{c}}{r},{opt_ref(col)},0)),"")')

for r in [6] + list(range(FIRST, LAST + 1)):
    blank = f'{S1}!B{r}=""'
    def F(x):   # ครอบด้วย IF แถวว่าง
        return f'=IF({blank},"",{x})'
    ws3.cell(row=r, column=1).value  = F(f'{S1}!A{r}')
    ws3.cell(row=r, column=2).value  = F(f'{S1}!B{r}')
    ws3.cell(row=r, column=3).value  = F(f'{S1}!D{r}')
    # Impact
    ws3.cell(row=r, column=4).value  = F(lk("R", r).format(c="I"))
    ws3.cell(row=r, column=5).value  = F(lk("S", r).format(c="J"))
    # "-" = เลือก "ประเมินเป็นเงินไม่ได้" -> ใช้เฉพาะชั่วโมง ไม่ถ่วงคะแนนลง
    okE = f'AND(E{r}<>"",E{r}<>"-")'
    ws3.cell(row=r, column=6).value  = F(
        f'IF(D{r}="",IF({okE},E{r},""),IF({okE},ROUND((D{r}+E{r})/2,0),D{r}))')
    # Data
    ws3.cell(row=r, column=7).value  = F(lk("B", r).format(c="L"))
    ws3.cell(row=r, column=8).value  = F(lk("C", r).format(c="M"))
    ws3.cell(row=r, column=9).value  = F(lk("D", r).format(c="O"))
    ws3.cell(row=r, column=10).value = F(f'IFERROR(ROUND(AVERAGE(G{r}:I{r}),0),"")')
    # Feasibility
    ws3.cell(row=r, column=11).value = F(lk("E", r).format(c="P"))
    ws3.cell(row=r, column=12).value = F(lk("F", r).format(c="Q"))
    ws3.cell(row=r, column=13).value = F(f'IFERROR(ROUND(AVERAGE(K{r}:L{r}),0),"")')
    # Risk
    ws3.cell(row=r, column=14).value = F(lk("G", r).format(c="R"))
    ws3.cell(row=r, column=15).value = F(lk("H", r).format(c="S"))
    ws3.cell(row=r, column=16).value = F(lk("I", r).format(c="T"))
    ws3.cell(row=r, column=17).value = F(f'IFERROR(ROUND(AVERAGE(N{r}:P{r}),0),"")')
    # นับช่องบังคับที่ยังไม่ได้กรอกใน sheet 1 (9 ช่องที่ใช้คำนวณคะแนน; มูลค่าความเสียหายไม่บังคับ)
    miss = "+".join(f'COUNTBLANK({S1}!{c}{r})' for c in "IJLMOPQRST")
    ws3.cell(row=r, column=21).value = F(miss)
    # รวม — ถ้ากรอกไม่ครบ ไม่แสดงคะแนน
    ws3.cell(row=r, column=18).value = F(
        f'IF(U{r}>0,"",IFERROR(ROUND((F{r}*{W_IMPACT}+M{r}*{W_FEAS}+J{r}*{W_DATA}'
        f'+(6-Q{r})*{W_RISK})*20,1),""))')
    ws3.cell(row=r, column=19).value = F(
        f'IF(U{r}>0,"ข้อมูลไม่ครบ",'
        f'IF(R{r}="","",IF(R{r}>={TH_HIGH},"สูง",IF(R{r}>={TH_MID},"กลาง","ต่ำ"))))')
    if r == 6:
        # แถวตัวอย่างไม่เข้าการจัดอันดับ (ไม่งั้นจะไปแย่งอันดับกับ Use Case จริง)
        ws3.cell(row=r, column=20).value = "ตัวอย่าง"
    else:
        ws3.cell(row=r, column=20).value = F(f'IF(R{r}="","",RANK(R{r},$R${FIRST}:$R${LAST},0))')

    for col in range(1, 22):
        c = ws3.cell(row=r, column=col)
        if col >= 4:
            c.alignment = Alignment(horizontal="center", vertical="center")
        if r == 6:
            c.fill = FILL_EX
            c.border = BORDER
            c.font = Font(name=FONT, size=9, italic=True, color="808080",
                          bold=col in (6, 10, 13, 17, 18, 19))
        elif col in (6, 10, 13, 17, 18, 19, 20):
            c.font = F_CALC
    ws3.cell(row=r, column=18).number_format = "0.0"
    ws3.row_dimensions[r].height = 22

rng = f"R6:R{LAST}"   # รวมแถวตัวอย่าง เพื่อให้เห็นตัวอย่างการไล่สีด้วย
ws3.conditional_formatting.add(rng, CellIsRule(
    operator="greaterThanOrEqual", formula=["75"],
    fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, size=9, bold=True, color="006100")))
ws3.conditional_formatting.add(rng, CellIsRule(
    operator="between", formula=["50", "74.99"],
    fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, size=9, bold=True, color="9C5700")))
ws3.conditional_formatting.add(rng, CellIsRule(
    operator="lessThan", formula=["50"],
    fill=PatternFill("solid", fgColor="F2F2F2"), font=Font(name=FONT, size=9, color="808080")))

# เตือนแถวที่กรอกข้อมูลไม่ครบ
warn_fill = PatternFill("solid", fgColor="FFC7CE")
warn_font = Font(name=FONT, size=9, bold=True, color="9C0006")
ws3.conditional_formatting.add(f"S6:S{LAST}", CellIsRule(
    operator="equal", formula=['"ข้อมูลไม่ครบ"'], fill=warn_fill, font=warn_font))
ws3.conditional_formatting.add(f"U6:U{LAST}", CellIsRule(
    operator="greaterThan", formula=["0"], fill=warn_fill, font=warn_font))

# สรุปเกณฑ์ใต้ตาราง
r0 = LAST + 2
ws3.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(cols3))
c = ws3.cell(row=r0, column=1, value="เกณฑ์การแบ่งระดับ และสูตรที่ใช้")
c.font = F_H2; c.fill = PatternFill("solid", fgColor="DDEBF7")
for i, txt in enumerate([
    "สูตร :  คะแนนรวม = ( Impact x 35%  +  Feasibility x 25%  +  Data x 25%  +  (6 - Risk) x 15% )  x 20",
    "Risk ถูกกลับด้านด้วย (6 - Risk) เพราะความเสี่ยงสูง = คะแนนควรต่ำ",
    "ระดับ สูง = 75 คะแนนขึ้นไป  |  กลาง = 50 - 74  |  ต่ำ = ต่ำกว่า 50  (คะแนนต่ำสุดที่เป็นไปได้คือ 20 ไม่ใช่ 0)",
    "ถ้าช่องบังคับในชั้นที่ 1 ยังกรอกไม่ครบ ระบบจะไม่แสดงคะแนนรวม แต่ขึ้นว่า 'ข้อมูลไม่ครบ' "
    "และบอกจำนวนช่องที่ยังขาดในคอลัมน์สุดท้าย — กลับไปกรอกให้ครบแล้วคะแนนจะขึ้นเอง",
    "ช่องบังคับ 10 ช่อง = เวลาที่ใช้ / มูลค่าความเสียหาย / รูปแบบข้อมูล / ข้อมูลย้อนหลัง / สิทธิ์ใช้ข้อมูล / "
    "ระบบเดิม / คนดูแลต่อ / ผลกระทบถ้า AI ผิด / การตรวจซ้ำ / ระดับความลับ",
    "ถ้าตีมูลค่าความเสียหายเป็นเงินไม่ได้ ให้เลือก 'ประเมินเป็นเงินไม่ได้' — ระบบจะใช้เฉพาะชั่วโมงคิด Impact "
    "ไม่ได้หักคะแนน (ในตารางจะขึ้นเป็น '-')",
    "คะแนนนี้เป็นตัวช่วยกรองเบื้องต้น ไม่ใช่คำตอบสุดท้าย — ผู้บริหารควรพิจารณา Use Case อันดับต้นๆ ด้วยตนเองอีกครั้ง",
    "ค่าเกณฑ์ชั่วโมง/มูลค่า และน้ำหนักทุกตัว แก้ได้ที่ sheet DATA คอลัมน์ S-T (ตัวเลขสีน้ำเงิน)",
], start=1):
    ws3.merge_cells(start_row=r0 + i, start_column=1, end_row=r0 + i, end_column=len(cols3))
    cc = ws3.cell(row=r0 + i, column=1, value="•  " + txt)
    cc.font = F_BODY; cc.alignment = Alignment(vertical="center")

# ============================================================ SHEET 0 : คู่มือ
ws0 = wb.create_sheet("คู่มือการใช้งาน")
ws0.sheet_view.showGridLines = False
ws0.column_dimensions["A"].width = 4
ws0.column_dimensions["B"].width = 34
ws0.column_dimensions["C"].width = 86
ws0.merge_cells("B2:C2")
c = ws0["B2"]; c.value = "คู่มือการกรอก AI Use Case Template (v2)"
c.font = F_TITLE; c.fill = FILL_TITLE; c.alignment = WRAP_C
ws0.row_dimensions[2].height = 30

rows = [
    ("S", "โครงสร้างไฟล์", ""),
    ("", "1-UseCase ทั้งหมด", "แผนกกรอกเอง — ข้อมูลพื้นฐานและข้อมูลที่ใช้คำนวณคะแนน (กรอกทุก Use Case)"),
    ("", "2-รายละเอียดเทคนิค", "แผนกกรอกเอง — ข้อมูลเชิงเทคนิคสำหรับกำหนด Spec Hardware (กรอกทุก Use Case) "
                               "ช่องชื่อ/แผนกระบบดึงมาจากชั้นที่ 1 ให้อัตโนมัติ"),
    ("", "3-คะแนนและลำดับ", "ระบบคำนวณให้อัตโนมัติ — ไม่ต้องกรอกอะไรเลย ดูอย่างเดียว "
                            "บอกว่าแต่ละ Use Case ได้กี่คะแนน อยู่ระดับไหน อันดับที่เท่าไหร่"),
    ("", "DATA", "รายการตัวเลือก Dropdown + ตารางแปลงตัวเลือกเป็นคะแนน + น้ำหนักและเกณฑ์ (ปรับได้)"),
    ("S", "ขั้นตอนการใช้งาน", ""),
    ("", "ขั้นที่ 1", "แต่ละแผนกกรอก sheet '1-UseCase ทั้งหมด' ให้ครบทุก Use Case ที่อยากได้"),
    ("", "ขั้นที่ 2", "กรอก sheet '2-รายละเอียดเทคนิค' ต่อ โดยดูชื่อ Use Case ที่ระบบดึงมาให้แล้ว"),
    ("", "ขั้นที่ 3", "เปิด sheet '3-คะแนนและลำดับ' ดูผล — ระบบจัดอันดับให้ทันทีโดยไม่ต้องกด"),
    ("", "ขั้นที่ 4", "นำไฟล์เข้าระบบ AI Use Case Tool เพื่อให้ AI ทบทวนคะแนนอีกชั้น "
                      "และวิเคราะห์ Spec Hardware / ประมาณราคา Server ของ Use Case อันดับต้นๆ"),
    ("S", "กติกาการกรอก", ""),
    ("", "ช่องสีเหลือง", "ช่องที่ต้องกรอกเอง"),
    ("", "ช่องสีเขียวอ่อน", "ระบบเติมให้อัตโนมัติ ห้ามพิมพ์ทับ (ถ้าพิมพ์ทับ สูตรจะหายและข้อมูลจะไม่ตรงกัน)"),
    ("", "แถวสีเทา (แถวที่ 6)", "แถวตัวอย่าง อย่าลบทิ้ง ใช้ดูว่าควรกรอกละเอียดแค่ไหน (เป็นตัวเลขสมมติ)"),
    ("", "ช่องที่มีลูกศร", "ต้องเลือกจากรายการเท่านั้น ห้ามพิมพ์เอง เพราะระบบอ่านค่าไปแปลงเป็นคะแนน"),
    ("", "ไม่รู้คำตอบทำยังไง", "เลือก 'ยังไม่แน่ใจ' ดีกว่าเดาตัวเลขมั่ว เพราะตัวเลขที่เดาจะทำให้คะแนนผิดทั้งระบบ"),
    ("", "กรอกได้กี่ Use Case", "แถวที่ 7 ถึง 36 (30 Use Case) ถ้าต้องการมากกว่านี้ให้แจ้งส่วนกลาง"),
    ("S", "คอลัมน์ที่ตรงกับ Excel Template เดิม (ยกมาครบ ไม่ได้ตัดทิ้ง)", ""),
    ("", "Existing Work", "-> คอลัมน์ G 'วิธีการทำงานปัจจุบัน' ในชั้นที่ 1"),
    ("", "Description for Redesign", "-> คอลัมน์ H 'แนวทางหลังนำ AI มาใช้' ในชั้นที่ 1"),
    ("", "Data Available from & by …", "-> คอลัมน์ K 'ข้อมูลที่มีอยู่ตอนนี้' ในชั้นที่ 1"),
    ("", "ปัญหาที่ต้องการแก้ / Output ที่คาดหวัง (แบบ GSK)",
     "-> คอลัมน์ F และ H ในชั้นที่ 1"),
    ("", "ข้อมูลเชิงปริมาณ + ความต้องการเชิงเทคนิค (แบบ GSK)",
     "-> ย้ายไปอยู่ชั้นที่ 2 ทั้งหมด พร้อมเพิ่มที่ขาด เช่น Peak concurrent, ประเภทข้อมูล Input, "
     "ขนาดคลังเอกสาร, ความต้องการ Fine-tune"),
    ("S", "ข้อมูลแต่ละกลุ่มถูกนำไปใช้ทำอะไร", ""),
    ("", "ผลกระทบ (สีส้ม)", "ชั่วโมง/สัปดาห์ + มูลค่าความเสียหาย -> คะแนน Impact (น้ำหนัก 35%)"),
    ("", "ความพร้อมข้อมูล (สีเขียว)", "รูปแบบ + ย้อนหลัง + สิทธิ์ใช้ข้อมูล -> คะแนน Data (น้ำหนัก 25%)"),
    ("", "ความเป็นไปได้ (สีม่วง)", "ระบบเดิมที่ต้องเชื่อม + คนดูแลต่อ -> คะแนน Feasibility (น้ำหนัก 25%)"),
    ("", "ความเสี่ยง (สีแดง)", "ผลถ้า AI ผิด + การตรวจซ้ำ + ความลับข้อมูล -> คะแนน Risk (น้ำหนัก 15% กลับด้าน)"),
    ("", "ชั้นที่ 2 ทั้งหมด", "ใช้กำหนด Spec Hardware : Latency + ประเภทข้อมูล -> ตัดสินว่าต้องใช้ GPU หรือ CPU  |  "
                              "Peak concurrent + Requests -> ตัดสินขนาดเครื่อง  |  "
                              "ระดับความลับ + ที่ตั้งระบบ -> ตัดสินว่า On-premise หรือ Cloud"),
]
r = 4
for kind, label, detail in rows:
    if kind == "S":
        ws0.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws0.cell(row=r, column=2, value=label)
        c.font = F_H2; c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(vertical="center")
        ws0.row_dimensions[r].height = 24
    else:
        a = ws0.cell(row=r, column=2, value=label)
        a.font = Font(name=FONT, size=9, bold=True); a.alignment = WRAP_T
        b = ws0.cell(row=r, column=3, value=detail)
        b.font = F_BODY; b.alignment = WRAP_T
        ws0.row_dimensions[r].height = 32 if len(detail) < 110 else 48
    r += 1
r += 1
ws0.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
c = ws0.cell(row=r, column=2,
             value="หมายเหตุ : ข้อมูลในแถวตัวอย่าง (แถวสีเทา) เป็นตัวเลขสมมติเพื่อแสดงรูปแบบการกรอกเท่านั้น "
                   "ไม่ใช่ข้อมูลจริงของบริษัท")
c.font = F_NOTE; c.alignment = WRAP_T

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb._sheets = [ws0, ws1, ws2, ws3, ws_d]
wb.save(OUT)
print("saved:", OUT)
