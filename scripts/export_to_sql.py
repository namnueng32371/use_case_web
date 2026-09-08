"""Export use cases from local API to SQL insert statements."""
import json, urllib.request

API_URL = "http://localhost:8765/api/usecases"
OUT_FILE = "docs/seed_data.sql"

resp = urllib.request.urlopen(urllib.request.Request(API_URL))
cases = json.loads(resp.read().decode())

lines = []
lines.append(f"-- Exported {len(cases)} use cases")
lines.append(f"-- Run this on the production MySQL after running schema.sql\n")
lines.append("USE ai_usecase_db;\n")

for uc in cases:
    uid = uc.get("id", "").replace("'", "\\'")
    name = uc.get("name", "").replace("'", "\\'")
    dept = uc.get("dept", "").replace("'", "\\'")
    payload = json.dumps(uc, ensure_ascii=False).replace("'", "\\'")
    lines.append(f"INSERT IGNORE INTO use_cases (id, name, dept, payload) VALUES ('{uid}', '{name}', '{dept}', '{payload}');")

lines.append(f"\n-- Done: {len(cases)} rows")

with open(OUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Export สำเร็จ: {len(cases)} รายการ -> {OUT_FILE}")
