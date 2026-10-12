# ARCHITECTURE.md — สถาปัตยกรรมของ AI Use Case Prioritization Tool

อธิบายจากโครงสร้างและโค้ดจริงที่ตรวจพบ ข้อมูลที่ยืนยันไม่ได้ระบุว่า **ต้องยืนยัน** / **ยังไม่ทราบ**

> ระบบเดิม Flask + MySQL (`server.py`, `schema.sql`, `start_all.bat`) **ถูกลบไปแล้ว** ถ้าเจอเอกสารที่อ้างถึงให้ถือว่าล้าสมัย — `README.md` และ `PRODUCT.md` ปรับให้ตรงสถาปัตยกรรมนี้แล้ว (2026-10-11)

---

## 1. ภาพรวมระบบ

- แอปหน้าเว็บ**ไฟล์เดียว** (HTML + CSS + JavaScript ล้วน) ทำงานเป็น **Claude Artifact** ที่ publish ไว้ ไม่มี backend ไม่มี server ที่ดูแลเอง
- ข้อมูล use case เก็บใน **Artifact database** (`window.claude.use('db')`) ใช้ร่วมกันทุกคนที่เปิด Artifact เดียวกัน
- ฟีเจอร์ AI ใช้ `window.claude.use('sample')` เรียก Claude ผ่านสิทธิ์ของผู้ใช้ที่เปิด Artifact
- มี 2 ไฟล์ที่เนื้อหาเป็นแอปเดียวกัน:

| ไฟล์ | บทบาท | ข้อแตกต่าง |
|---|---|---|
| `artifact.html` | **ตัวหลัก** ที่ publish ขึ้น Artifact | เป็น fragment ไม่มี `<!DOCTYPE>`/`<html>`/`<head>` (แพลตฟอร์มห่อให้) เปิดตรงจากเครื่องไม่ได้ |
| `AI_Use_Case_Program12.html` | ไฟล์สำรองในเครื่อง | เอกสาร HTML สมบูรณ์ (`lang="th"`) ไม่ได้ sync อัตโนมัติ ต้องแก้คู่กันด้วยมือ |

- Artifact ปัจจุบัน: `https://claude.ai/artifact/U5akAKHk25wHyje3oQQU2e` (private เจ้าของเดียว) เวอร์ชันล่าสุดที่ตรวจพบคือ 17

---

## 2. เทคโนโลยีที่ใช้

| ส่วน | ที่ใช้ | หมายเหตุ |
|---|---|---|
| UI | HTML + CSS + JavaScript (สไตล์ `var`/`function` เป็นหลัก มี `async/await`) | ไม่มี framework ไม่มี Tailwind ไม่มี build/bundler |
| กราฟ | Chart.js 4.4.1 (CDN cdnjs) | แท่งแนวนอนเปรียบเทียบคะแนน |
| อ่านไฟล์นำเข้า | JSZip 3.10.1, pdf.js 2.16.105 (CDN cdnjs) | แกะ `.docx` `.pptx` และ PDF |
| ฟอนต์ | IBM Plex Sans Thai (Google Fonts) | ดู `DESIGN_SYSTEM.md` |
| ฐานข้อมูล | Artifact database (`claudeDb`) | ผ่าน `window.claude.use('db')` |
| AI | Claude ผ่าน `window.claude.use('sample')` | `claudeSample.json(prompt, opts)` |
| สร้าง Excel template | Python 3 + `openpyxl` | `docs/build_template.py` |
| เครื่องมือออกแบบ | skill `impeccable` | `.claude/skills/impeccable/` (ไม่เกี่ยวกับ runtime) |

ไม่มี `package.json` / `requirements.txt` / ชุดทดสอบอัตโนมัติ (ตรวจแล้วไม่พบไฟล์ test ใน git)

---

## 3. โครงสร้างโปรเจกต์

```
use_case_web/
├── artifact.html                 # ตัวหลัก (publish ขึ้น Artifact)
├── AI_Use_Case_Program12.html    # ไฟล์สำรองในเครื่อง (ต้องตรงกับตัวหลัก)
├── AGENTS.md · ARCHITECTURE.md · DESIGN_SYSTEM.md · PRD.md   # เอกสารสำหรับ Agent/ทีม
├── CONTEXT.md                    # ข้อตกลงโครงการ + หลัก sizing Spec Hardware
├── IMPLEMENTATION_PLAN.md        # แผนงาน 4 เฟส (task checklist)
├── README.md                     # คู่มือผู้ใช้ (ยุค Artifact)
├── PRODUCT.md                    # บริบทผลิตภัณฑ์สำหรับ skill impeccable
├── data/use_cases_export.json    # สำรองข้อมูลก่อนย้ายมา Artifact
├── backup_ai_usecase_db_*.sql    # สำรองจากยุค MySQL
├── docs/
│   ├── AI_UseCase_Template_v2.xlsx   # Template แจกแผนกกรอก (สร้างจากสคริปต์)
│   ├── build_template.py             # สคริปต์สร้าง template (ห้ามแก้ .xlsx มือ)
│   ├── SCORING.md · TEMPLATE_GUIDE.md · dept_codes.json
│   └── use_cases_source/ · use_cases_source2/   # ไฟล์ต้นฉบับจากแผนก
└── .claude/                      # skill impeccable + agents + launch.json
```

โครงสร้างภายใน `artifact.html` (ไฟล์เดียวแบ่งเป็นส่วน): `<style>` → ส่วน HTML `<header class="app-header">` + `<div id="view-…">` 5 มุมมอง (`usecases`, `org`, `flow`, `spec`, `plan`) + modal 4 ตัว (`doc-modal`, `plan-modal`, `flow-modal`, `org-role-modal`) + footer → `<script>` ตัวเดียวเรียงตาม: ค่าคงที่/ICONS → ตัวช่วย AI → `loadData`/`saveData` → `DATA_LISTS` และสูตรคะแนน → render → สร้างเอกสาร/แผน/Flow → ฟอร์ม → bind event

---

## 4. โครงสร้างข้อมูล

ไม่มี ORM/Model/ตารางแล้ว ข้อมูลเป็น JSON ใน Artifact database

**Collection/Document ที่ใช้ (ตรวจจาก `claudeDb.doc(...)`)**

| Path | เนื้อหา | อ่าน/เขียนที่ |
|---|---|---|
| `usecases/all` | `{ items: [ useCase, … ] }` — array เดียวเก็บทุก use case | `loadData()` / `saveData()` |
| `specAnalysis/{high\|med\|all}` | ผลวิเคราะห์ spec ของแต่ละกลุ่ม (`SPEC_GROUPS`) — **จะลบตอนลงมือเฟส 3** | `loadSpecAnalysisFromDb()` / `runSpecAnalysis()` |

**Document ที่วางแผนไว้ในเฟส 3 (ยังไม่มีในโค้ด):** `config/priceCatalog` (รุ่น ราคา ค่าเงิน วันที่ ที่มา ใช้ร่วมทั้งหน้า Spec และแผนโครงการ) · `config/assumptions` (ค่าสมมติการใช้งานตามประเภทงาน, safety factor, กำลังทีมต่อปี, เดือนเริ่มปีงบ) · `scenarios/{id}` (ชุด case 1–3 ปีที่เตรียมเสนอผู้บริหาร) — รายละเอียดใน `IMPLEMENTATION_PLAN.md` หัวข้อ 3.1

**Use case object** (ไม่มี schema บังคับ สร้างใน `handleSave()`):

- ระบุตัวตน: `id` (`uc_<timestamp>_<random>`), `name`
- ข้อความ 16 ช่อง: `dept` `site` `owner` `qwp` `kqi` `refStep` `problem` `solution` `dataAvailable` `redesignDesc` `dataVolume` `targetTimeline` `remark` `budget` `constraints` `providedBy`
- Dropdown 29 ช่อง (เก็บเป็น**ข้อความตัวเลือกเต็ม** ไม่ใช่รหัส): `company` `timeSpent` `lossValue` `dataFormat` `dataHistory` `dataRights` `needConnect` `adminOwner` `aiErrorImpact` `reviewNeed` `confidentiality` `deptPriority` `triggerType` `inputType` `usersPerDay` `peakConcurrent` `itemsPerDay` `pointsCount` `language` `responseTime` `serviceHours` `slowdownImpact` `retentionPeriod` `dailyInputSize` `corpusSize` `imageResolution` `installLocation` `modelAdaptation` `retrainFrequency`
- ไฟล์แนบ: `attachment` (`{name,type,size,data}` base64 ≤ 10MB) และ `attachRef`
- ค่าที่ระบบคำนวณ/เติม: `aiType` (derive จาก `inputType` + `modelAdaptation` ผ่าน `deriveLegacyAiType()`), `aiPlan` และ `aiFlow` (ผลวิเคราะห์ AI ที่เก็บไว้กับ use case)
- **คะแนนไม่ถูกเก็บ** คำนวณสดทุกครั้งจาก 10 ช่องบังคับ: `computeScoreDetail()` = `(Impact×0.35 + Feasibility×0.25 + Data×0.25 + (6−Risk)×0.15) × 20` (ช่วง 20–100; สูง ≥75, กลาง 50–74, ต่ำ <50) ถ้ากรอกไม่ครบจะแสดง "ข้อมูลไม่ครบ"
- `DATA_LISTS` (29 รายการ) คือแหล่งตัวเลือกและคะแนนของ dropdown ต้องตรงกับ `LISTS` ใน `docs/build_template.py` (ตรวจแล้ว 2026-10-12 ว่าตรงกันทั้ง 3 ที่ ทั้งข้อความและคะแนน; รายการ `P` "ที่ตั้งระบบ" ที่ไม่มีใครใช้ ลบออกจากสคริปต์แล้ว)

**ความสัมพันธ์:** ไม่มี (use case เป็นรายการแบน ไม่มีตารางเชื่อม; `site`/`company` เป็นข้อความอิสระ)

**ข้อมูลในไฟล์อื่น:** โครงสร้างทีม (`ORG_DATA`) ฝังเป็น JSON ในโค้ด; ตารางราคาอ้างอิงฮาร์ดแวร์ฝังเป็นข้อความใน prompt ของ `buildSpecPrompt()` (ไม่ใช่ตารางแยกที่ทีมแก้ได้ ซึ่งขัดกับข้อกำหนดใน `CONTEXT.md`)

---

## 5. การไหลของข้อมูล

```
เปิด Artifact ─► loadData(): claudeDb.doc('usecases/all').get() ─► useCases[] ในหน่วยความจำ
                                                    │
ผู้ใช้เพิ่ม/แก้/ลบ ─► handleSave() / ลบ ─► saveData(): doc('usecases/all').set({items}) ─► renderList()
                                                    │
                           renderList() ─► คำนวณคะแนน ─► การ์ด/สรุป/กราฟ ─► renderSpecTab() · renderPlanTab()
```

- **Fallback ตอนอ่าน/เขียน:** Artifact db → `fetch('/api/usecases')` (ซากของยุค Flask ใช้ไม่ได้แล้ว) → `window.storage` → `localStorage` (คีย์ `ai_usecase_program_data`) ไฟล์ local ที่เปิดนอก Claude จะตกมาที่ `localStorage` ของเบราว์เซอร์นั้นเท่านั้น
- **การเขียนทับทั้งก้อน:** `saveData()` เขียน array ทั้งหมดทุกครั้ง ไม่มีการเช็คเวอร์ชัน ถ้า 2 คนแก้พร้อมกัน คนที่บันทึกทีหลังทับของอีกคน (ความเสี่ยงที่ตรวจพบจากโค้ด)
- **เส้นทาง AI** (ทั้งหมดเรียกผ่าน `sampleJsonWithTimeout()` มี timeout และแปลงข้อความ error เป็นภาษาไทย):

| ฟีเจอร์ | ฟังก์ชัน | เก็บผลที่ไหน |
|---|---|---|
| วิเคราะห์ Spec/ราคา Server | `runSpecAnalysis(key)` (ผ่าน `buildSpecPrompt`) | `specAnalysis/{key}` |
| ประเมินแผนงานต่อ use case | `analyzePlanAI()` | `uc.aiPlan` |
| ประเมิน Flow/n8n | `analyzeFlowAI()` | `uc.aiFlow` |
| นำเข้า use case จากเอกสาร | `runImportAnalysis()` | ผู้ใช้เลือกเพิ่มทีละรายการเข้า `useCases` |

- **เอกสารที่ออกมา:** หน้าจอ/พิมพ์ (`buildDocBody`) และ n8n JSON (`exportN8nJson`, ปุ่มยังอยู่) — ปุ่มส่งออก CSV และดาวน์โหลด Word (เอกสาร/แผนงาน) **ถูกเอาออกจากหน้าเว็บแล้ว (2026-10-12)** เพราะ Artifact ไม่ให้สิทธิ์ดาวน์โหลดแก่ผู้ชม ปุ่มจึงไม่ทำงาน — ฟังก์ชัน `exportCsv` / `exportDocWord` / `exportPlanWord` ยังค้างอยู่ในโค้ดแต่ไม่มีปุ่มเรียก (ถ้าจะนำกลับต้องประกาศ capability `downloads` แล้วบันทึกผ่าน `window.claude.downloads.save`)

**ข้อสังเกตที่ตรวจพบ (ยังไม่แก้ ดูเฟส 3–4 ใน `IMPLEMENTATION_PLAN.md`):** `ucSpecSummaryLine`/`ucSpecSummaryLineShort` ยังอ้างฟิลด์เก่า (`users`, `language`, `dataSensitivity`, `latency`, `availability`) และ `parseTotalWeeks/TeamSize/Budget` อ่านจาก `uc.resources` ซึ่งไม่มีช่องกรอกแล้ว ทำให้ข้อมูล Zone A/B และ `budget` ไม่ถูกส่งเข้า prompt สเปก/แผนงาน

---

## 6. การยืนยันตัวตนและสิทธิ์

- **แอปไม่มีระบบ login/role/middleware/route** ไม่มี API ของตัวเอง
- การเข้าถึงขึ้นกับการแชร์ของ Artifact: ปัจจุบันเป็น private เฉพาะเจ้าของ (จากผล publish) ผู้ใช้อื่นต้องได้รับสิทธิ์จากเมนู Share ของ Artifact ก่อน
- ฟีเจอร์ AI ต้องได้รับอนุญาตใช้ Claude จากผู้เปิดหน้า (โค้ด error `not_granted`, `rate_limited`)
- ผู้ที่เปิด Artifact ได้เห็นข้อมูลชุดเดียวกันและแก้ไขได้ ไม่มีการแยกสิทธิ์ราย use case ปุ่ม "ล้างทั้งหมด" ถูกปิดโดยตั้งใจ (`disabled`) เพราะลบข้อมูลทุกคน
- **ข้อควรระวังด้านความปลอดภัย:** ข้อความที่ผู้ใช้/AI ป้อนต้องผ่าน `escapeHtml()` ก่อนใส่ `innerHTML` (ใช้ทั่วโค้ด) ไม่มี Secrets/API key ในโค้ด (Claude เรียกผ่านสิทธิ์ของผู้ใช้) ไฟล์แนบเก็บเป็น base64 ใน database

---

## 7. การเชื่อมต่อภายนอก

| ปลายทาง | ใช้ทำอะไร | หมายเหตุ |
|---|---|---|
| Claude (`window.claude` db / sample) | เก็บข้อมูล, วิเคราะห์ AI | ใช้ได้เมื่อเปิดผ่าน Claude Artifact เท่านั้น |
| `fonts.googleapis.com` / `fonts.gstatic.com` | ฟอนต์ IBM Plex Sans Thai | |
| `cdnjs.cloudflare.com` | Chart.js, JSZip, pdf.js | เปิดแบบออฟไลน์แล้วกราฟ/นำเข้าไฟล์ใช้ไม่ได้ |
| n8n | รับไฟล์ JSON ที่ export ไปนำเข้าเอง | ไม่มีการเรียก n8n โดยตรง |
| Excel template | แผนกกรอกแบบออฟไลน์ แล้วให้นำข้อมูลเข้าโปรแกรม | ไม่มีฟีเจอร์นำเข้า `.xlsx` ในเว็บ (ที่มีคือนำเข้า PDF/Word/PPT ผ่าน AI) — ตัดสินใจ 2026-10-12 ใช้วิธีให้ AI Agent อ่านไฟล์แล้วเขียน `usecases/all` ผ่านเครื่องมือ ArtifactData โดยเพิ่มต่อท้ายและล็อกเวอร์ชัน |

---

## 8. การขยายในอนาคต

อ้างอิง `IMPLEMENTATION_PLAN.md` (ยังไม่ได้ลงมือ):

1. **เฟส 3 — หน้าประเมิน Spec + หน้าแผนโครงการแบบคำนวณตายตัว** (ออกแบบเสร็จ 2026-10-11)
   - ฟังก์ชันล้วนที่ทดสอบแยกได้: `specEligibility` → `classifyWorkload` → `workloadDemand` → `aggregateSite` → `sizeSite` → `costPlan` / `marginalCost` (ชื่อตามแผน ยังไม่มีในโค้ด)
   - แยกตามสถานที่ตั้ง (`site`) ทุกที่ตั้งมี Server ของตัวเอง แทน `computeSpecGroups()` ที่แบ่งตามระดับคะแนนซ้อนทับกัน
   - แคตตาล็อกราคาเดียวแทนตารางราคาใน prompt; AI ใช้แค่จัดประเภท/อธิบาย ห้ามเปลี่ยนตัวเลข
   - scenario 1–3 ปี เทียบสูงสุด 3 ชุด แทนการ์ด Tier ในหน้าแผนโครงการ
   - แก้ `ucSpecSummaryLine` และ `parseTotalWeeks/TeamSize/Budget` ให้ใช้ฟิลด์ Zone A/B และ `budget` จริง
   - `site` เปลี่ยนเป็น dropdown ตามบริษัทเมื่อได้รายชื่อ
2. **เฟส 4** — รายงานสรุป AI 3 ฉบับ (แผนงาน · Flow · Spec Hardware)
3. **เฟส 2** — Redesign ส่วนอื่นต่อจาก header แล้วสร้าง `DESIGN.md`
4. **ข้อเสนอด้านโครงสร้าง (ต้องยืนยันความต้องการ):** เก็บ use case แยกเอกสารต่อรายการเพื่อกันเขียนทับกัน และเพิ่มการเช็คเวอร์ชันตอนบันทึก (ความเสี่ยงต่ำลงเพราะผู้ใช้มีเฉพาะทีม AI)
