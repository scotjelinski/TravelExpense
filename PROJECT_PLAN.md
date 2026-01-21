## Travel Expense Copilot (Studio) - Project Plan

### Target UX (what “good” looks like)
- User starts in Teams chat, adds **Receipt / Per Diem / Mileage** line items to a single draft “cart”.
- Bot does **all GL coding** (dept/activity/account + optional 107200 override) without requiring GL knowledge.
- Bot asks only when required, and always shows the proposed coding with a clear way to correct it.
- Submission emails finance and (optionally) bundles receipts.

### Current implementation (in this repo)
- Topics: `copilot_topics/TE_0_Router.yaml` … `copilot_topics/TE_8_Debug.yaml`
- Functions: `gl-lookup-func/function_app.py` (Function-key auth)
  - `GET /api/expense-codes`
  - `POST /api/import-csv`
  - `POST /api/submit-report`
  - `GET /api/health`
  - `GET /api/orgchart-lookup` (OrgChart mapped `departmentCode`)

### Maker/Admin setup checklist (once per environment)
- Configure the Azure Function key in the **custom connector connection** (so topics/actions never prompt end users for a key).
- Configure OrgChart search in Azure Functions (env vars in `gl-lookup-func/local.settings.json`):
  - `ORGCHART_SEARCH_ENDPOINT`
  - `ORGCHART_SEARCH_INDEX`
  - `ORGCHART_SEARCH_API_KEY`
  - `ORGCHART_SEARCH_EMAIL_FIELD` (default `email`)
- Verify Teams identity token availability in your environment (topics assume `System.User.Email`; insert via Studio variable picker if needed).
- Configure GSA per-diem lookup in Azure Functions (app settings):
  - `GSA_API_KEY`
  - (Optional) `GSA_PER_DIEM_BASE_URL` / `GSA_PER_DIEM_URL_TEMPLATE` if your endpoint shape differs.

### Next improvements (priority order)
1) **“Approve or Change”** UX after adding an item (change dept/activity/account without restarting).
2) **Per-diem lookup** (GSA) is implemented; validate in Teams and refine mapping if needed.
3) **Receipt handling**: simplify the receipt upload/link flow and reduce user friction.
4) **Hardening**: correlation ids, better error messaging, and guardrails for oversized responses.
