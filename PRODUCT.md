# Product

<!-- impeccable:product-schema 1 -->

## Platform

web (single-file Claude Artifact)

## Users

Primary: the **AI team** (about 2 engineers) who collect, screen, score, and prioritize AI use case submissions across the group companies, and who prepare spec/budget scenarios for leadership. Today they are the only people who use the tool.

Secondary (audience, not operators): **executives** who look at the prioritized list and the spec/plan scenarios and decide which use cases get built each year. Department champions submit ideas upstream via the Excel template; the AI team enters them into the tool.

## Product Purpose

Replace fragmented per-department/per-company Excel sheets with one shared tool. The AI team scores every idea against the same weighted rubric (business impact, technical feasibility, data readiness, risk), ranks it, and produces ready-to-share documentation. Next step (planned, not built yet): turn a chosen set of use cases into a defensible, deterministic on-premise server sizing and a 1–3 year budget plan that executives can compare as scenarios without second-guessing the numbers.

## Positioning

Unlike ad-hoc spreadsheets that each department fills out with no shared scoring standard, this tool applies one formula-based priority score org-wide and turns an idea directly into shareable artifacts (one-page Word proposal, Gantt plan, n8n workflow JSON). The planned spec engine computes hardware needs from the use case inputs with fixed formulas and an editable price catalog, using AI only to classify and explain, so the numbers stay reproducible.

## Operating Context

- Runs as a Claude Artifact; data lives in the Artifact's own shared database (`window.claude.use('db')`, document `usecases/all`). No backend, no MySQL, no LAN server anymore.
- The local file `AI_Use_Case_Program12.html` is a backup copy of the code; opened directly it falls back to the browser's `localStorage`.
- Source material arrives as Excel workbooks (the department template `docs/AI_UseCase_Template_v2.xlsx`, older source sheets in `docs/use_cases_source*/`) or PDF/Word/PowerPoint files that Claude reads through the import feature.
- All AI workloads must run **on-premise**, with **one server per site** (each factory/branch hosts its own). Hardware is datacenter/workstation class only.
- Also hosts a team-structure tab ("โครงสร้างทีม AI Star") and a workflow tab ("Flow การทำงาน AI Start") that exports n8n-compatible JSON.

## Capabilities and Constraints

- Add/edit/delete use cases through a structured form that mirrors the Excel template: company, department, site (สถานที่ตั้ง), owner, QWP/kQI/SQP references, problem, current method, AI solution, data readiness, workload and sizing inputs, and more.
- Company is a dropdown of five group companies (บริษัท แกรนด์เอสเค คอร์ปอเรชั่น, เอส.เค.โพลีเมอร์, ไทยรับเทค, โพลีเมท, จีทีเค สมาร์ท โซลูชั่น) plus "อื่นๆ". Site is currently free text and will become a per-company dropdown once the site list is supplied.
- Automatic priority score: `(impact×0.35 + feasibility×0.25 + data×0.25 + (6−risk)×0.15) × 20`, bucketed สูง ≥75 / กลาง 50–74 / ต่ำ <50. The score is a screening aid; executives make the final pick.
- Exports: one-page Word document per use case, CSV of the whole list, Gantt plan, n8n workflow JSON. Import: PDF/.docx/.pptx parsed by Claude, which proposes only use cases that don't already exist.
- Current spec page asks AI to estimate hardware and prices, so results vary between runs. A replacement is planned (IMPLEMENTATION_PLAN.md phase 3): a deterministic sizing engine (Little's Law concurrency, VRAM budget, utilization headroom, editable safety factor), a shared price catalog (hardware only, excluding VAT, ±20–30% estimates), per-site aggregation, a chassis sized for year 3 with GPUs added yearly, and up to 3 scenarios compared on screen.
- No login or access control inside the app; access is controlled by Artifact sharing.
- The "ล้างทั้งหมด" (clear-all) action is intentionally disabled, because it would wipe the shared data.
- Thai is the primary UI language; entries mix Thai and English.

## Brand Commitments

Header shows the "Polymate" logo (wordmark plus a diamond/cube mark in orange/tan/brown, on a transparent background). UI font is IBM Plex Sans Thai. Design tokens and known deviations are in DESIGN_SYSTEM.md.

## Evidence on Hand

- The Artifact database currently holds 2 demo use cases. Real submissions from the MySQL era are preserved in `backup_ai_usecase_db_*.sql` and `data/use_cases_export.json`.
- Source workbooks live under `docs/use_cases_source/` and `docs/use_cases_source2/`.
- No customer-facing marketing copy, testimonials, or pricing exists. This is an internal operations tool, not a product with external buyers.

## Product Principles

1. One shared source of truth beats per-department spreadsheets: consistency and speed over flexibility.
2. Every use case is scored the same way, so priority calls are comparable across companies. Executives still make the final call.
3. Numbers shown to executives must be reproducible: formulas in code and prices from one catalog, with AI limited to classifying and explaining.
4. Getting from "idea" to "shareable documentation" should take one click, not a manual rebuild in Word/Excel.
5. Preserve existing team/company terminology (QWP, kQI, SQP, บริษัทในเครือ) rather than imposing generic SaaS vocabulary.
