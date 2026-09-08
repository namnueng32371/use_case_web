# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Primary: the **AI Unit / PM team** who screen, score, and prioritize AI use case submissions across the organization's group companies. They open the tool regularly to review incoming ideas, adjust priority scores, and produce documentation for leadership decisions.

Secondary: department champions/managers across the group companies who originate AI use case ideas (currently submitted upstream via Excel workshops, then entered into this tool by the AI Unit team).

## Product Purpose

Replace fragmented, per-department/per-company Excel spreadsheets used to collect and prioritize AI use case ideas with one shared tool. It lets the AI Unit team score each idea consistently against the same weighted rubric (business impact, technical feasibility, data readiness, risk), rank it, and produce ready-to-share documentation without manual reformatting in Word/Excel.

## Positioning

Unlike ad-hoc spreadsheets that each department/company fills out independently with no shared scoring standard, this tool applies one consistent formula-based priority score to every submission org-wide, and turns an approved idea directly into shareable planning artifacts (one-page proposal doc, Gantt implementation plan, n8n workflow export) — something a spreadsheet cannot do on its own.

## Operating Context

- Internal LAN tool for the AI Unit — not internet-facing, no login/auth by design.
- Data lives in a shared MySQL database (`ai_usecase_db`) so everyone on the LAN sees the same list at once.
- Source material for new use cases mostly arrives as Excel workbooks from company workshops/consultations (e.g. departmental idea-collection sheets, engineering ROI case studies) and gets manually entered into the tool.
- Used alongside the org's broader AI Unit structure (a "โครงสร้างทีม AI Star" team-structure tab) and workflow-automation planning (a "Flow การทำงาน AI Start" tab that exports n8n-compatible JSON).

## Capabilities and Constraints

- Add/edit/delete AI use case entries via a structured form: problem, existing method, proposed AI solution, hardware, ROI, risks, QWP/quality-procedure references, and more.
- Automatic priority scoring: `(impact×0.35 + feasibility×0.25 + data readiness×0.25 + (6−risk)×0.15) × 20`, bucketed into สูง/กลาง/ต่ำ (High/Med/Low).
- Generates a one-page Word/PDF summary document and a Gantt-style implementation plan per use case, plus an n8n workflow JSON export.
- Entries span multiple companies in the group (e.g. S.K.Polymer, Thairubbtech, Polymate, GTK) — the form has no dedicated "company" field yet; company is currently folded into department/name text.
- No authentication or access control by design — trusted-LAN only; must never be exposed to the public internet.
- The "ล้างทั้งหมด" (clear-all) action is intentionally disabled in the UI, since it would wipe shared data for every concurrent user.
- Thai is the primary UI language; individual entries mix Thai and English freely.

## Brand Commitments

Branded "Polymate" — wordmark plus a diamond/cube mark in orange/tan/brown, rendered with a transparent background in the app header.

## Evidence on Hand

- Real use case submissions are already live in production (grown from an initial 6-item seed set to 20+ entries from actual department/company input).
- Source workbooks live under `use_caseทั้งหมด/`: departmental AI-idea submission sheets, and detailed engineering ROI case studies (e.g. a CAE mold-design optimization case).
- No customer-facing marketing copy, testimonials, or pricing exists — this is an internal operations tool only, not a product with external buyers.

## Product Principles

1. One shared source of truth beats per-department spreadsheets — consistency and speed over flexibility.
2. Every use case is scored the same way, so priority calls are defensible and comparable across companies.
3. Getting from "idea" to "shareable documentation" should take one click, not a manual rebuild in Word/Excel.
4. Trusted-LAN simplicity over enterprise security — no login friction, but never expose beyond the LAN.
5. Preserve existing team/company terminology (QWP, kQI, บริษัทในเครือ) rather than imposing generic SaaS vocabulary.
