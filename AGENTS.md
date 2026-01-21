# Travel Expense (Copilot Studio) — Agent Handoff

This repo implements a **Teams Copilot Studio** travel-expense "cart" experience: users add **Receipt / Per Diem / Mileage** items, the bot performs deterministic **GL coding** (dept/activity/account + optional override), then submits by email via Azure Functions.

## Topics - v1 vs v2

This repo contains two versions of the Copilot Studio topics:

| Folder | Version | Files | Description |
|--------|---------|-------|-------------|
| `copilot_topics/` | v1 (original) | 10+ files | Complex with many interdependencies |
| `copilot_topics_v2/` | v2 (refactored) | 6 files | **Recommended** - Simplified, focused on core features |

### v2 Topics (Recommended)

The refactored v2 topics provide the core features:
1. **Collect email** - Ask user for CORE email or use signed-in email
2. **Get department code** - Auto-lookup via OrgChart API or manual entry
3. **Build cart** - Add per diem, mileage, or receipt items
4. **Submit** - Submit expense report via API

**v2 Topic Files:**
- `TE_Main.yaml` - Entry point with email/dept collection and routing
- `TE_AddReceipt.yaml` - Add receipt expense
- `TE_AddPerDiem.yaml` - Add per diem (with GSA rate lookup)
- `TE_AddMileage.yaml` - Add mileage expense
- `TE_ReviewCart.yaml` - Review and clear cart
- `TE_Submit.yaml` - Submit expense report

See `copilot_topics_v2/README.md` for setup instructions.

## Source of truth (important)

There are two `gl-lookup-func` folders on disk:
- Repo root: `C:\Users\sjelinski\OneDrive - CORE\AI\Travel Expense\gl-lookup-func` (**DEPLOYED SOURCE OF TRUTH**; publish from here)
- Studio workspace: `C:\Users\sjelinski\OneDrive - CORE\AI\Travel Expense\Studio\` (topics + Swagger docs)

If something works locally but not in Azure, verify you updated/published the **root** `gl-lookup-func`.

## Core components

**Copilot Studio Topics - v2 (Recommended)**
- `copilot_topics_v2/TE_Main.yaml`: entry + email/dept collection + routing
- `copilot_topics_v2/TE_AddReceipt.yaml`: receipt intake + account code lookup
- `copilot_topics_v2/TE_AddPerDiem.yaml`: per diem intake + GSA rate lookup
- `copilot_topics_v2/TE_AddMileage.yaml`: mileage intake + account code lookup
- `copilot_topics_v2/TE_ReviewCart.yaml`: cart review + clear
- `copilot_topics_v2/TE_Submit.yaml`: submit via `POST /api/submit-report`

**Copilot Studio Topics - v1 (Legacy)**
- `copilot_topics/TE_0_Router.yaml`: entry + initializes globals (includes `Global.ApiBaseUrl`)
- `copilot_topics/TE_1_SetIdentity.yaml`: "change department" flow; uses Teams email when available and supports override
- `copilot_topics/TE_2_AddReceipt.yaml`: receipt intake + `GET /api/expense-codes`
- `copilot_topics/TE_3_AddPerDiem.yaml`: per diem intake + `GET /api/per-diem-lookup` (fallback to manual daily rate)
- `copilot_topics/TE_4_AddMileage.yaml`: mileage intake + `GET /api/expense-codes`
- `copilot_topics/TE_5_ReviewDraft.yaml`: draft summary + clear draft
- `copilot_topics/TE_6_SubmitDraft.yaml`: submits draft via `POST /api/submit-report`
- `copilot_topics/TE_8_Debug.yaml`: safe debug (no system identifiers that break in some tenants)

**Azure Functions (DepartmentCodes)**
- Code: repo root `gl-lookup-func/function_app.py`
- Endpoints:
  - `GET /api/health`
  - `GET /api/expense-codes`
  - `GET /api/orgchart-lookup` (Azure AI Search OrgChart lookup; parses JSON stored in Search doc `chunk`)
  - `GET /api/per-diem-lookup` (GSA per diem M&IE daily rate by ZIP/date)
  - `POST /api/import-csv`
  - `POST /api/submit-report`
  - Receipt upload endpoints already present (SharePoint upload flow)

**API specs**
- Swagger (used in Studio Swagger editor): `Azure-DepartmentCodes-Function.swagger.yaml` (in this Studio folder)
- OpenAPI 3 (function repo): `gl-lookup-func/travel-expense-tools.openapi.yaml` (in function repo)

## Runtime model (key design decisions)

- Topics call Azure Functions via **custom connector tools** (Function-key auth via `x-functions-key` stored in the connector connection).
- Teams identity:
  - Topics attempt to prefill `Global.RequesterEmail` from `System.User.Email` when available.
  - Topics call `/api/orgchart-lookup?email=<email>` to map OrgChart department to a 3-digit `Global.DepartmentCode`.
- OrgChart Search index `sharepoint-oc-index` stores user JSON in a `chunk` string; the function parses that JSON and matches on `UPN` / `Mail` for strict email matching.
- GL coding:
  - Activity code is inferred from keywords (700/710/770/790) unless the user overrides.
  - Account code is **never guessed**; always uses `GET /api/expense-codes?departmentCode=...&activityCode=...`.
  - If `IsCapital=true` or a work order is present, `glAccountOverride="107200"`.
- Per diem:
  - `GET /api/per-diem-lookup?zipCode=#####&travelDate=YYYY-MM-DD` returns a daily M&IE rate (`mieRate`).
  - Topic falls back to asking for daily rate if lookup fails.
  - GSA official v2 ZIP endpoint (used by the function implementation):
    - `GET https://api.gsa.gov/travel/perdiem/v2/rates/zip/{ZIP}/year/{FISCAL_YEAR}`
    - `{year}` is **federal fiscal year** (Oct 1–Sep 30): FY = calendar year + 1 for Oct–Dec, else calendar year.

## Azure configuration (must-have app settings)

Function app: `DepartmentCodes` (Resource Group: `AzureAI`)

Auth:
- App Service Authentication/EasyAuth must be set to **Allow unauthenticated requests**, otherwise all calls return `WWW-Authenticate: Bearer` and function keys won’t work.

OrgChart (Azure AI Search):
- `ORGCHART_SEARCH_ENDPOINT=https://foundrysrch.search.windows.net`
- `ORGCHART_SEARCH_INDEX=sharepoint-oc-index`
- `ORGCHART_SEARCH_API_KEY=<secret>`
- `ORGCHART_SEARCH_API_VERSION=2023-11-01`
- Index schema note: for `sharepoint-oc-index`, user JSON is stored in a `chunk` string like `{"DisplayName":...,"UPN":...,"Mail":...,"Department":...}`.
- The function parses `chunk` JSON and matches strict email using `UPN`/`Mail` (top-level `email` is not present).

Per diem (GSA):
- `GSA_API_KEY=<secret>`
- Optional overrides:
  - `GSA_PER_DIEM_BASE_URL` (default `https://api.gsa.gov/travel/perdiem/v2`)
  - `GSA_PER_DIEM_URL_TEMPLATE` (only needed if the official endpoint shape changes)

Copilot Studio:
- Set `Global.ApiBaseUrl` (topics default it).

## Deploy / test

Deploy functions (run from repo root function folder):
- `cd "C:\Users\sjelinski\OneDrive - CORE\AI\Travel Expense\gl-lookup-func"`
- `func azure functionapp publish DepartmentCodes`

Smoke tests:
- `curl "https://departmentcodes-gnfncpgxdueyfdh9.eastus2-01.azurewebsites.net/api/health" -H "x-functions-key: <KEY>"`
- `curl "https://departmentcodes-gnfncpgxdueyfdh9.eastus2-01.azurewebsites.net/api/orgchart-lookup?email=sjelinski@core.coop" -H "x-functions-key: <KEY>"`
- `curl "https://departmentcodes-gnfncpgxdueyfdh9.eastus2-01.azurewebsites.net/api/per-diem-lookup?zipCode=80128&travelDate=2025-12-12" -H "x-functions-key: <KEY>"`

## Tools / Connector (Power Platform)

- Custom connector should use **API key auth** (header `x-functions-key`), not OAuth.
- When adding tools/actions in Studio, set actions to use **connection credentials**, not end-user credentials.
- Swagger source of truth in this Studio workspace: `Azure-DepartmentCodes-Function.swagger.yaml` (use this for Swagger Editor copy/paste + connector re-import).

## Known pitfalls

- Secrets: don’t paste keys into chat; rotate if leaked.
- If `/api/orgchart-lookup` returns 400/502, it’s usually index schema mismatch; for `sharepoint-oc-index`, email is inside the `chunk` JSON.
- If Studio shows “Connection manager”, it’s using a store connector/tool instead of the YAML HTTP request nodes—remove/disable the store tool.
- If HTTP calls return `WWW-Authenticate: Bearer`, App Service Authentication/EasyAuth is still enforcing Entra auth; set unauthenticated action to allow requests for function-key flows.
