# Travel Expense Bot - Agent Instructions (v2)

## Goal

Help users build a travel expense draft (Receipt / Per Diem / Mileage) and submit it via the Azure function.

## Core Flow

1. **Collect Email** - Ask user for their CORE email (or "me" for signed-in email)
2. **Lookup Department** - Call OrgChart API to get department code; fallback to manual entry
3. **Build Cart** - Add receipt, per diem, or mileage items
4. **Submit** - Send the expense report via the submit API

## Behavior Rules

- Ask ONE question at a time
- Don't re-ask for information already provided
- After adding an item, show a summary and confirm before adding to cart
- Always validate inputs (email format, 3-digit dept, 5-digit ZIP, positive numbers)

## Topic Structure (v2)

| Topic | Purpose |
|-------|---------|
| `TE_Main.yaml` | Entry point: collect email, lookup dept, route to item topics |
| `TE_AddReceipt.yaml` | Add receipt expense to cart |
| `TE_AddPerDiem.yaml` | Add per diem expense to cart (with GSA rate lookup) |
| `TE_AddMileage.yaml` | Add mileage expense to cart |
| `TE_ReviewCart.yaml` | Review cart, clear or proceed to submit |
| `TE_Submit.yaml` | Submit expense report via API |

## Global Variables

| Variable | Description |
|----------|-------------|
| `Global.Version` | Topic version (2.0.0) |
| `Global.ApiBaseUrl` | Azure function base URL |
| `Global.MileageRate` | Mileage reimbursement rate (default 0.70) |
| `Global.RequesterEmail` | User's email for the expense report |
| `Global.DepartmentCode` | 3-digit department code |
| `Global.DepartmentName` | Department name (from OrgChart lookup) |
| `Global.CartJson` | JSON array of cart items |
| `Global.CartCount` | Number of items in cart |

## Custom Connector Actions (Tools)

These must be added as Tools in Copilot Studio with API key authentication:

1. **TravelExpense-OrgChartLookupByUpn**
   - Input: `upn` (email)
   - Output: `ok`, `found`, `departmentCode`, `departmentName`, `error`

2. **TravelExpense-ExpenseCodesLookup**
   - Input: `departmentCode`, `activityCode`
   - Output: `matches[]` (array of `{ accountCode, description }`)

3. **TravelExpense-PerDiemLookup**
   - Input: `zipCode`, `travelDate`
   - Output: `ok`, `mieRate`, `error`

4. **TravelExpense-SubmitReport**
   - Input: `requesterEmail`, `toEmail`, `ccRequester`, `draftItemsJson`, `receiptBundleFormat`
   - Output: `ok`, `sent`, `toEmail`, `lineCount`, `amountTotal`, `csvFilename`, `emailError`

## Copilot Studio Setup

### 1. Import Topics

Copy each YAML file into Copilot Studio as a new topic:
- Set `TE_Main` as the entry point with trigger phrases
- Other topics should have NO trigger phrases (called via BeginDialog)

### 2. Configure Custom Connector

1. Create a Custom Connector from the swagger file
2. Set authentication to API Key (header: `x-functions-key`)
3. Test the connection

### 3. Add Tools

For each connector action, add it as a Tool:
1. Go to Tools > Add Tool
2. Select the custom connector action
3. Set it to use connector credentials (NOT end-user credentials)

### 4. Topic Dialog References

Update the `dialog:` references in YAML to match your agent's naming:
```yaml
# Example - update these to match your agent name
dialog: new_agent.topic.TE_AddReceipt
dialog: new_agent.action.TravelExpense-OrgChartLookupByUpn
```

### 5. Orchestration Mode

Set orchestration to **Classic** or disable "AI can use tools/actions" to prevent the AI from calling tools directly (topics should control tool invocation).

## Activity Codes

| Code | Category |
|------|----------|
| 700 | General business travel (default) |
| 710 | Meals and events |
| 770 | Training and education |
| 790 | Wellness and employee engagement |

## User Commands

| Command | Action |
|---------|--------|
| `add receipt` | Add a receipt expense |
| `add per diem` | Add per diem for travel |
| `add mileage` | Add mileage reimbursement |
| `review` | Review cart contents |
| `submit` | Submit expense report |
| `reset` | Clear email, department, and cart |
| `debug` | Show current state (email, dept, cart count) |

## Troubleshooting

### "Connector prompts for input directly"

The topic is not binding inputs correctly. Ensure all tool inputs use `=` prefix:
```yaml
# Correct
input:
  binding:
    upn: =Global.RequesterEmail

# Wrong (sends literal text)
input:
  binding:
    upn: Global.RequesterEmail
```

### "No account codes found"

The expense_codes.csv on the Azure function may not have entries for that department/activity combination. Check the function logs or try a different department.

### "OrgChart lookup failed"

The Azure AI Search index may not have the user's email. Fall back to manual department code entry.

### "Submit failed"

Check that:
1. The Azure function has email sending configured (Mail.Send permission)
2. The function key is correct
3. The toEmail address is valid
