# Copilot Studio Topic YAML (Travel Expense)

These files are **Copilot Studio Topic** YAML definitions (`kind: AdaptiveDialog`) that you can paste into the **topic code editor**.

## Configuration

The topics call your Azure Functions via a **custom connector tool** (actions are invoked as `BeginDialog` nodes) so the Functions key lives in the **connector connection**, not in topics.

- Base URL: stored in `Global.ApiBaseUrl` for display/debug (defaults to `https://departmentcodes-gnfncpgxdueyfdh9.eastus2-01.azurewebsites.net`)
- Requester identity: captured in `TE_9_InitContext.yaml` and stored as a plain string in `Global.RequesterEmail` (user can type `me` to use `System.User.Email`).

## Identity auto-fill (Teams)

When running in Teams, the topics:

1) Capture the requester's CORE email/UPN once per conversation (`Global.RequesterEmail`).
2) If `Global.DepartmentCode` is blank, call the connector action for `/api/orgchart-lookup-upn` to map OrgChart department -> 3-digit department code (same mapping table as GL lookup).

If OrgChart lookup fails, the init topic falls back to asking for the 3-digit department code.

Users can override by saying "change department" or "change user" (routes to the identity topic).

## Connection Manager (Tools)

If Teams "Connection manager" shows only a subset of tools (example: only `HealthCheck` + `SubmitReport`), the agent will not be able to call the other actions and you'll see fallback prompts (like asking for a manual per-diem rate).

Fix:
- In Copilot Studio, add the required actions under **Tools** for the `Travel Expense` custom connector (at minimum: `ExpenseCodesLookup`, `PerDiemLookup`, `SubmitReport`, `HealthCheck`).
- For each tool/action, set **Credentials to use** to the connection you created (API key), not "end user credentials".
- After updating the custom connector Swagger/OpenAPI, re-check topic action cards: Copilot Studio can clear previously-bound inputs, which causes connector-native prompts like “Please provide the zip code you would like to look up.”

## Per diem date handling

- `TE_3_AddPerDiem.yaml` accepts a free-text travel date and passes it to the per-diem lookup tool.
- If the per-diem lookup fails or rates look wrong, prompt the user for a standard format (`MM/DD/YYYY` or `YYYY-MM-DD`).
- If the user includes ZIP/days/date in the same message (example: `add per diem 80128 last tuesday 2 days`), the topic attempts to prefill fields before prompting.
- The topic should always prompt for missing ZIP/days/date before calling the per-diem lookup tool (to avoid tool calls with blank `zipCode`).

## Known issues

See `copilot_topics/KNOWN_ISSUES.md` for common YAML + Power Fx pitfalls and error codes we've hit when pasting topics into Copilot Studio.

## Topic files

- `copilot_topics/TE_0_Router.yaml`
- `copilot_topics/TE_1_SetIdentity.yaml` (identity override; router/other topics call `new_agent.topic.TravelExpenseSetIdentity`)
- `copilot_topics/TE_2_AddReceipt.yaml` (called by router)
- `copilot_topics/TE_3_AddPerDiem.yaml` (called by router)
- `copilot_topics/TE_4_AddMileage.yaml` (called by router)
- `copilot_topics/TE_5_ReviewDraft.yaml` (called by router)
- `copilot_topics/TE_6_SubmitDraft.yaml` (called by router)
- `copilot_topics/TE_7_HelpPolicy.yaml` (called by router)
- `copilot_topics/TE_8_Debug.yaml` (called by router)
- `copilot_topics/TE_9_InitContext.yaml` (called by other topics)

## Recommended Studio setup (avoid double-prompts)

To avoid issues where Copilot asks for a value and then a Topic asks again:

- Use `TE_0_Router` as the only topic with trigger phrases.
- Keep the other `TE_*` topics **without trigger phrases** (they are invoked via `BeginDialog` from the router).

## Resetting context (for fast iteration)

If you can't reliably start a new chat thread (Teams 1:1 bot chat) or you're iterating in the Test pane, use:
- `reset` / `start over` / `restart`

This clears draft + identity globals so the next request re-runs init context and uses the newest topic logic.
