# TDD: Travel Expense Copilot — Copilot Studio (Topics + Actions)

**Owner:** Scot Je  
**Date:** 2026-01-13  
**Status:** Draft

---

## 1. Background and problem statement

We currently have a travel-expense experience implemented as a long, single workflow (YAML) with an external runtime dependency (Foundry agent invocation). Publishing and/or channel execution in Microsoft Teams has produced repeated **HTTP 400** failures during agent calls, making the end-to-end experience unreliable.

This design moves orchestration into **Microsoft Copilot Studio** using **Topics** (YAML-editable via the topic code editor) and calls backend services via **Actions** (custom connector / OpenAPI-backed Azure Functions). This reduces the runtime surface area to a single channel (Copilot Studio connector) and eliminates the Foundry-agent hop.

> Copilot Studio supports editing topics in a YAML **code editor**. (TODO: add Microsoft docs link)
> Copilot Studio connector calls can fail with HTTP 400 when a connector returns too much data; responses are limited to **500 KB**. (TODO: add Microsoft docs link)

---

## 2. Goals

### Functional goals
- Support **Receipt**, **Per Diem**, **Mileage**, **Review draft**, **Submit** flows.
- Ensure **department + activity + account** codes are valid and selected from the authoritative lookup service.
- Submit a report by generating an ImportFormat CSV and sending it to finance with optional receipt bundle.
- Add knowledge sources (Azure AI Search, GSA connector, internal articles) for policy / “how-to” questions without contaminating deterministic coding logic.

### Platform goals
- Implement orchestration as **Copilot Studio topics** (modular, testable).
- Minimize manual build effort by leveraging **topic YAML editing** and node copy/paste. (TODO: add Microsoft docs link)
- Support deployment via **Solutions** export/import across environments. (TODO: add Microsoft docs link)

---

## 3. Non-goals

- Replacing the finance system or changing accounting rules beyond what already exists.
- Building a full receipt OCR + merchant/category extraction pipeline (can be iterative enhancement later).
- Implementing multi-step approvals / manager workflows (out of scope for first migration).

---

## 4. Current assets and constraints

### Existing backend API (Azure Functions via OpenAPI)
From the existing OpenAPI file, the backend provides:
- `GET /api/expense-codes` (`get_expense_codes`)  
  - Required query params: `departmentCode`, `activityCode`
- `GET /api/orgchart-lookup` (`orgchart_lookup`)
- `POST /api/import-csv` (`import_csv`)
- `POST /api/submit-report` (`submit_report`)
- `GET /api/health` (`health_check`)

### Copilot Studio connector compatibility
Power Platform custom connectors commonly expect **OpenAPI 2.0 (Swagger)** for import. If the current OpenAPI is 3.x, conversion is required (or use a supported import path). (TODO: add Microsoft docs link)

### Current implementation note (this repo)
The current topics in `copilot_topics/` call the Azure Functions using **custom connector actions/tools** (`BeginDialog` nodes), with the Functions key stored in the connector connection (not in Topics).

### 400 errors risk
If any action response can exceed 500 KB, Copilot Studio may surface connector request failures (often as HTTP 400). (TODO: add Microsoft docs link)
Design must keep action responses small and filter payloads.

---

## 5. Target architecture

### High level
**Copilot Studio Agent**
- Orchestrates via Topics
- Uses Actions (custom connector) to call Azure Functions
- Uses Knowledge sources (Azure AI Search, GSA, knowledge articles) for non-coding Q&A

**Azure**
- Azure Functions (existing endpoints above)
- Azure AI Search index for OrgChart and knowledge (or use GSA connector where applicable)

### Runtime request chain (target)
Teams / Web -> Copilot Studio -> Action (custom connector) -> Azure Functions -> Response -> Copilot Studio

**Removed:** Teams/Studio -> Foundry agent -> Foundry workflow -> Functions

---

## 6. Topic design (modular)

### Topic 0: Router / Entry
**Triggers:** “travel expense”, “receipt”, “per diem”, “mileage”, “submit”, “expense report”, etc.  
**Responsibilities:**
- Detect intent quickly and redirect to the right topic
- Provide lightweight menu only when needed

### Topic 1: Change Dept / Identity
Collect:
- CORE email (preferred) OR 3-digit dept code  
Resolve:
- In Teams: prefill email from signed-in user when available (e.g., `System.User.Email`).
- If email present: OrgChart lookup (Azure AI Search) with strict exact match, then map OrgChart `department` to a 3-digit department code using the expense-code table (values like `620 - INFORMATION TECHNOLOGY`).
Store:
- `Global.RequesterEmail`
- `Global.DepartmentCode`

### Topic 2: Add Receipt Item
Collect:
- Short category (“hotel”, “parking”, “airfare”, “meal”)
- Amount (required if no parsing)
- Date (optional)
- Attachments (optional; allow “attach now” path)
Compute:
- Activity code from simple keyword mapping
- Account code via `get_expense_codes(departmentCode, activityCode)`
Prompt:
- If multiple matches: present A/B/C options; store selection
Output:
- Append a standardized `DraftItem` object to `Global.DraftItems`

### Topic 3: Add Per Diem Item
Collect:
- Start date
- Number of days
- 5-digit ZIP
- Optional travelers count (default 1)
Compute:
- Call per diem lookup action (see Section 8) and compute total
- Activity + account lookup
Output:
- Append `DraftItem`

### Topic 4: Add Mileage Item
Collect:
- Date
- Miles
- Description (trip purpose)
Compute:
- Amount = miles * configured mileage rate
- Activity + account lookup
Output:
- Append `DraftItem`

### Topic 5: Review Draft
Show:
- Count of draft items
- Totals by category (Receipt/PerDiem/Mileage)
Actions:
- Clear draft
- Submit draft
- Back to add more

### Topic 6: Submit Draft
Preconditions:
- Draft not empty
Collect:
- Confirm requester email (if missing)
- Confirm receipts attachments if required (or warn if none)
Call:
- `submit_report` with:
  - `toEmail = itrenewal@core.coop`
  - `ccRequester = true`
  - `requesterEmail = Global.RequesterEmail`
  - `items = Global.DraftItems` (preferred) OR `draftItemsJson`
  - Attachments (base64) if available
Handle:
- Only claim success if `submit_report` returns ok/sent
Post:
- Clear draft on success

### Topic 7: Help / Policy Q&A
Triggers:
- “what’s reimbursable”, “meal limits”, “receipt required”, “per diem rules”, etc.
Node:
- Generative answers over configured knowledge sources
Redirect:
- Back to Router when the user says “add receipt / submit / etc.”

---

## 7. Data model

### Global variables (conversation scope)
- `Global.RequesterEmail` (string)
- `Global.DepartmentCode` (string)
- `Global.DepartmentName` (string, optional)
- `Global.WorkOrder` (string, optional)
- `Global.IsCapital` (boolean, default false)
- `Global.DraftItems` (array of DraftItem objects)
- `Global.AttachmentCache` (array of {name, contentType, contentBytes} or platform attachment references)

### DraftItem (canonical object)
Minimum fields:
- `mode` ("Receipt" | "PerDiem" | "Mileage")
- `departmentCode`
- `activityCode`
- `accountCode`
- `glAccountOverride` (optional, e.g., "107200")
- `travelDate`
- `reference` (short string)
- `amountTotal` (number)
- `currency` ("USD")
- `lines` (array; at least one)

### Account selection object
- `matches[]` from `get_expense_codes`
- `SelectedAccountCode`

---

## 8. Actions / connectors

### Auth model (Copilot Studio -> Azure Functions)
Two viable patterns:
- **Function key (API key) auth**: stable secret stored in the connector; avoids user sign-in flows in Studio.
  - Requires the target Azure Function routes to use `auth_level=FUNCTION` (route-level `ANONYMOUS` overrides bypass function-key auth).
  - Connector should send `x-functions-key` (recommended) or `code` query param.
- **Entra ID / EasyAuth (Managed Identity) auth**: no shared secret in Studio, but requires platform auth setup and typically a user sign-in / delegated flow depending on the connector.

### A. Expense code lookup
Action: `get_expense_codes`
Inputs:
- `departmentCode` (string)
- `activityCode` (string)

Outputs used in topics:
- `matches` list (filter response to only required fields to reduce payload)

### B. Submit report
Action: `submit_report`
Inputs:
- `toEmail`, `requesterEmail`, `ccRequester`
- `subject`, `bodyText` (optional)
- `items` (preferred) or `draftItemsJson`
- `attachments[]` (optional)
- `zipReceipts` (default true)

Outputs used in topics:
- `ok`/`sent` boolean + a small summary (lineCount, total)

### C. Per diem lookup (gap / addition)
If not already available as an action, implement one of:
- Add a new Azure Function endpoint (e.g., `/api/per-diem?zip=...&date=...`) and include in connector
- Or use an existing service connector

Outputs should be minimal:
- `dailyRate`
- `currency`
- `source`/`asOf` (optional)

### D. OrgChart lookup
Preferred: Azure AI Search index
- Query by email
- Return strict match fields only: email, deptCode, deptName

---

## 9. Business logic rules

### Department resolution
- If dept code not known: ask for CORE email or dept code.
- If email provided: strict exact match from OrgChart; otherwise ask again.

### Activity code mapping (simple)
- Meals/events -> 710
- Training/education -> 770
- Wellness program -> 790
- Default travel -> 700

### GL override
If `Global.IsCapital == true` OR `Global.WorkOrder` present:
- Set `glAccountOverride = "107200"`
- Still pick `accountCode` via lookup service

### Mileage rate
- Store as `Global.MileageRate` (default 0.70 to match existing behavior)
- Make it easy to change annually without topic rewrites (environment variable or single config topic)

### Connector payload size control
- Ensure each action returns only the fields needed by the topic.
- If an API returns large lists, add filtering query params or server-side “projection” to avoid >500 KB responses. (TODO: add Microsoft docs link)

---

## 10. Error handling and observability

### User-facing patterns
- **Validation errors** (missing ZIP, invalid dept code): ask one corrective question.
- **Action failures**: provide one retry prompt and a fallback (e.g., manual entry of dept/account if appropriate).
- **Oversized responses**: adjust API outputs or action schema to reduce size; do not dump huge JSON into chat. (TODO: add Microsoft docs link)

### Telemetry
- Include correlation id in API responses (if possible)
- Add a “Debug” topic that prints:
  - ConversationId (from Studio context)
  - Last action name + correlation id
  - Attachment summary count (not bytes)

---

## 11. Security and compliance

- Use Entra ID auth for connectors; avoid embedding secrets in topics.
- If using Function-key auth instead of Entra: store the key only in the connector configuration (not in Topics), rotate periodically, and ensure all non-public routes use `auth_level=FUNCTION`.
- Don't log raw receipt images/bytes; log only metadata.
- Keep PII exposure minimal in chat transcripts; prompt users not to paste sensitive data beyond what's required.

---

## 12. Build plan checklist

### Phase 0 - Foundations
- [ ] Confirm Studio environment and Teams channel target
- [ ] Enable Generative orchestration if using advanced action filling / MCP
- [ ] Decide connector approach:
  - [ ] Convert OpenAPI 3.x -> OpenAPI 2.0 for custom connector import (TODO: add Microsoft docs link)
  - [ ] Or publish via APIM / supported pipeline
- [ ] Decide authentication approach:
  - [ ] Function key (API key in connector) OR Entra ID / EasyAuth
  - [ ] If Function key: update Azure Function routes to `auth_level=FUNCTION` and configure connector to send `x-functions-key`
- [ ] Create connector connection references (test in Power Apps/Automate if needed)
- [ ] Maintain the connector connection key separately (do not embed secrets in Topics or swagger files)
- [ ] Add knowledge sources:
  - [ ] Azure AI Search index connection
  - [ ] GSA connector
  - [ ] Internal knowledge articles

### Phase 1 - Actions
- [x] Add action: `get_expense_codes`
- [x] Add action: `submit_report`
- [x] Add action: `import_csv` (optional preview)
- [x] Add action: `health_check` (smoke)
- [x] Add/confirm action: `per_diem_lookup` (new if missing)
- [x] Add action: `orgchart_lookup` (Azure AI Search query)

### Phase 2 - Topics (MVP)
- [x] Topic: Router / Entry
- [x] Topic: Change Dept / Identity
- [x] Topic: Add Receipt
- [x] Topic: Add Per Diem
- [x] Topic: Add Mileage
- [x] Topic: Review Draft
- [x] Topic: Submit Draft
- [x] Topic: Help / Policy Q&A

### Phase 3 - Draft model + UX
- [x] Implement DraftItems array and append logic
- [ ] Implement "Approve or Change?" step after item add
- [x] Implement Clear Draft
- [ ] Implement attachment caching strategy (store metadata; bytes only when required)

### Phase 4 — Hardening
- [ ] Add guardrails against oversized connector responses (>500 KB) (TODO: add Microsoft docs link)
- [ ] Add “Debug” topic for diagnostics
- [ ] Add rate/config variables (mileage rate, default activity)
- [ ] Add retry patterns (one retry max per action)

### Phase 5 — Deployment
- [ ] Package agent in a Solution for export/import (TODO: add Microsoft docs link)
- [ ] Validate connection references in target environment
- [ ] Publish to Teams
- [ ] Run acceptance test suite in Teams

---

## 13. Test plan checklist (acceptance criteria)

### Identity / Dept resolution
- [ ] Email resolves to dept code on exact match
- [ ] Non-matching email prompts for correction
- [ ] Manual dept code accepted and reused for subsequent items

### Account lookup
- [ ] Single match auto-selected
- [ ] Multiple matches prompts A/B selection and stores choice
- [ ] No matches triggers helpful fallback prompt

### Receipt item
- [ ] With amount provided -> adds to draft with correct coding
- [ ] Missing amount -> asks for amount
- [ ] Attachment present -> included in cache and submitted (if sendEmail true)

### Per diem
- [ ] Missing ZIP -> asks for ZIP
- [ ] Rate returned -> computes correct total with days * travelers
- [ ] Handles “travel date” formats robustly

### Mileage
- [ ] Computes miles * rate (default 0.70) correctly
- [ ] Adds to draft with correct coding

### Review + Submit
- [ ] Review shows correct count + totals
- [ ] Submit fails when draft empty
- [ ] Submit succeeds and clears draft
- [ ] Submit returns confirmation to user and CCs requester when configured

### Error handling
- [ ] Action returns >500 KB -> reproduced and mitigated by filtering outputs (TODO: add Microsoft docs link)
- [ ] Action timeout -> user sees retry prompt and graceful exit

---

## 14. Migration strategy

- Start with MVP topics that match current behavior (Receipt/PerDiem/Mileage/Submit).
- Run internal pilot in Teams.
- Keep Foundry version as fallback during pilot.
- Cut over once Teams publishing + runtime are stable for 1–2 weeks.

---

## 15. Open questions

- What is the authoritative per diem source and do we already have a callable endpoint?
- Do we want receipt OCR now or phase 2?
- Should WorkOrder/IsCapital be asked globally (once per trip) or per item?
- How should mileage rate be updated each year (env var vs config table)?


