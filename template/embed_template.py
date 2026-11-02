"""ฝัง Excel template ว่าง (base64) ลงใน artifact.html และ AI_Use_Case_Program12.html
ให้ปุ่ม "ดาวน์โหลด Template ว่าง" ในเว็บใช้งานได้ (Artifact แนบไฟล์ .xlsx แยกไม่ได้)

รันทุกครั้งหลังแก้ build_template.py:  python template/embed_template.py
สคริปต์สร้าง template แบบว่าง (ไม่ใส่ SEED_JSON) ในโฟลเดอร์ชั่วคราว จึงไม่แตะไฟล์ .xlsx ใน template/
"""
import base64
import os
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(os.path.dirname(HERE), "app")
START, END = "/*XLSX_TEMPLATE_B64*/'", "'/*END_XLSX_TEMPLATE_B64*/"


def build_blank(out_path):
    src = open(os.path.join(HERE, "build_template.py"), "rb").read().decode("utf-8")
    a = src.index(chr(10) + 'OUT = ') + 1
    b = src.index("\n", a)
    src = src[:a] + "OUT = " + repr(out_path) + src[b:]
    os.environ.pop("SEED_JSON", None)
    exec(compile(src, "build_template", "exec"), {"__name__": "__main__"})


def main():
    out = os.path.join(tempfile.mkdtemp(), "AI_UseCase_Template_v2.xlsx")
    build_blank(out)
    b64 = base64.b64encode(open(out, "rb").read()).decode("ascii")
    for name in ("artifact.html", "AI_Use_Case_Program12.html"):
        p = os.path.join(APP, name)
        s = open(p, "rb").read().decode("utf-8")
        i = s.index(START) + len(START)
        j = s.index(END, i)
        s = s[:i] + b64 + s[j:]
        open(p, "wb").write(s.encode("utf-8"))
        print("embedded", len(b64), "chars ->", name)


if __name__ == "__main__":
    main()
