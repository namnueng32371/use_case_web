"""
Test suite for AI Use Case webapp
ทดสอบ: GET /, GET /api/usecases, PUT /api/usecases (add/update/restore)
"""
import json, urllib.request, urllib.error, time

BASE = "http://localhost:8765"
PASS = []
FAIL = []

def check(name, condition, detail=""):
    if condition:
        PASS.append(name)
        print(f"  PASS  {name}")
    else:
        FAIL.append(name)
        print(f"  FAIL  {name}" + (f" — {detail}" if detail else ""))

def get(path):
    req = urllib.request.Request(f"{BASE}{path}")
    with urllib.request.urlopen(req, timeout=5) as r:
        return r.status, r.read().decode("utf-8"), dict(r.headers)

def put(path, data):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(f"{BASE}{path}", data=body, method="PUT",
                                  headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=5) as r:
        return r.status, json.loads(r.read().decode("utf-8"))

# ── 1. หน้าเว็บหลัก ──────────────────────────────────────────────────────────
print("\n[1] ทดสอบหน้าเว็บหลัก (GET /)")
status, body, headers = get("/")
check("HTTP 200 OK", status == 200, f"got {status}")
check("Content-Type เป็น text/html", "text/html" in headers.get("Content-Type",""), headers.get("Content-Type"))
check("มี <title> ใน HTML", "<title>" in body)
check("โหลด HTML ครบ (> 100KB)", len(body) > 100_000, f"{len(body):,} bytes")

# ── 2. GET /api/usecases ──────────────────────────────────────────────────────
print("\n[2] ทดสอบ GET /api/usecases")
status, body, _ = get("/api/usecases")
data = json.loads(body)
check("HTTP 200 OK", status == 200, f"got {status}")
check("Response เป็น JSON Array", isinstance(data, list), type(data).__name__)
check("มีข้อมูลอยู่ใน DB (> 0 รายการ)", len(data) > 0, f"พบ {len(data)} รายการ")
print(f"         พบ use cases ทั้งหมด: {len(data)} รายการ")

# ตรวจ field ครบ
if data:
    sample = data[0]
    for field in ["id", "name", "dept", "problem", "impact", "feasibility", "data", "risk"]:
        check(f"  field '{field}' มีอยู่", field in sample, f"keys: {list(sample.keys())[:5]}")

# ── 3. PUT /api/usecases (เพิ่ม test record) ─────────────────────────────────
print("\n[3] ทดสอบ PUT /api/usecases — เพิ่ม test record")
original = data.copy()
test_record = {
    "id": "__test_record__",
    "name": "TEST — ลบได้เลย",
    "dept": "Test",
    "problem": "test",
    "impact": 1, "feasibility": 1, "data": 1, "risk": 1
}
new_list = original + [test_record]
status, resp = put("/api/usecases", new_list)
check("HTTP 200 OK", status == 200, f"got {status}")
check("ok: true", resp.get("ok") == True, str(resp))
check(f"count = {len(new_list)}", resp.get("count") == len(new_list), str(resp))

# ── 4. ยืนยันว่า GET ได้ข้อมูลใหม่ ───────────────────────────────────────────
print("\n[4] ยืนยัน test record อยู่ใน DB")
_, body2, _ = get("/api/usecases")
data2 = json.loads(body2)
ids = [u["id"] for u in data2]
check("test record อยู่ใน DB", "__test_record__" in ids)
check(f"จำนวนรายการถูกต้อง ({len(new_list)})", len(data2) == len(new_list), f"got {len(data2)}")

# ── 5. Restore ข้อมูลเดิม ────────────────────────────────────────────────────
print("\n[5] Restore ข้อมูลเดิม (ลบ test record)")
status, resp = put("/api/usecases", original)
check("HTTP 200 OK", status == 200)
check(f"count กลับเป็น {len(original)}", resp.get("count") == len(original), str(resp))

_, body3, _ = get("/api/usecases")
data3 = json.loads(body3)
ids3 = [u["id"] for u in data3]
check("test record ถูกลบแล้ว", "__test_record__" not in ids3)
check("ข้อมูลเดิมครบถ้วน", len(data3) == len(original), f"got {len(data3)}, expected {len(original)}")

# ── 6. Error handling ─────────────────────────────────────────────────────────
print("\n[6] ทดสอบ Error Handling")
try:
    body_bad = b"not a json array"
    req = urllib.request.Request(f"{BASE}/api/usecases", data=body_bad, method="PUT",
                                  headers={"Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=5)
    check("PUT ข้อมูลผิด format คืน error", False, "ควรได้ HTTP 4xx แต่ได้ 2xx")
except urllib.error.HTTPError as e:
    check("PUT ข้อมูลผิด format คืน HTTP 4xx", e.code >= 400, f"got {e.code}")

# ── สรุปผล ───────────────────────────────────────────────────────────────────
total = len(PASS) + len(FAIL)
print(f"\n{'='*50}")
print(f"ผลการทดสอบ: {len(PASS)}/{total} PASS")
if FAIL:
    print(f"FAIL ({len(FAIL)} รายการ):")
    for f in FAIL:
        print(f"  - {f}")
else:
    print("ทุก test ผ่าน! ระบบพร้อมใช้งาน")
print('='*50)
