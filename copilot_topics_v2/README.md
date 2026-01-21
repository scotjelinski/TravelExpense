# Travel Expense Bot - Refactored Topics (v2)

This folder contains the refactored Copilot Studio topics for the Travel Expense bot. The v2 topics are simplified and focused on the core features.

## What Changed (v1 -> v2)

### Simplified Architecture

**Before (v1):** 10+ topic files with complex interdependencies
- `TE_0_Router.yaml` - Entry point
- `TE_1_SetIdentity.yaml` - Identity management
- `TE_2_AddReceipt.yaml` - Receipt flow
- `TE_3_AddPerDiem.yaml` - Per diem flow
- `TE_4_AddMileage.yaml` - Mileage flow
- `TE_5_ReviewDraft.yaml` - Review
- `TE_6_SubmitDraft.yaml` - Submit
- `TE_7_HelpPolicy.yaml` - Help
- `TE_8_Debug.yaml` - Debug
- `TE_9_InitContext.yaml` - Context initialization

**After (v2):** 6 focused topic files
- `TE_Main.yaml` - Entry point with email/dept collection and routing
- `TE_AddReceipt.yaml` - Add receipt
- `TE_AddPerDiem.yaml` - Add per diem
- `TE_AddMileage.yaml` - Add mileage
- `TE_ReviewCart.yaml` - Review and clear cart
- `TE_Submit.yaml` - Submit expense report

### Key Improvements

1. **Single entry point** - `TE_Main` handles email collection, department lookup, AND routing (no separate InitContext topic)

2. **Cleaner variable names** - Using `Global.CartJson` and `Global.CartCount` instead of `Global.DraftItemsJson`, `Global.DraftItemCount`, `Global.TotalReceiptAmount`, etc.

3. **Less duplication** - Department code validation is only in Main topic, not repeated in every item topic

4. **Simpler dialog references** - Consistent naming pattern: `TE_Main`, `TE_AddReceipt`, etc.

5. **Focused flows** - Each topic does ONE thing well

## Core Features

The refactored topics support these core features:

1. **Ask for email** - Collects user email or uses signed-in email ("me")
2. **Get department code** - Calls OrgChart API or asks manually
3. **Build cart** - Add per diem, mileage, or receipt items
4. **Review & submit** - Review cart and submit expense report

## Files

| File | Lines | Purpose |
|------|-------|---------|
| `TE_Main.yaml` | ~250 | Entry point, email/dept collection, routing |
| `TE_AddReceipt.yaml` | ~200 | Add receipt expense |
| `TE_AddPerDiem.yaml` | ~280 | Add per diem with GSA rate lookup |
| `TE_AddMileage.yaml` | ~220 | Add mileage expense |
| `TE_ReviewCart.yaml` | ~100 | Review and clear cart |
| `TE_Submit.yaml` | ~150 | Submit expense report |
| `Agent_Instructions.md` | - | Agent behavior instructions |

## Setup in Copilot Studio

1. **Create topics** - Copy each YAML file as a new topic in Copilot Studio
2. **Set triggers** - Only `TE_Main` should have trigger phrases
3. **Update dialog references** - Change `new_agent.topic.X` to match your agent name
4. **Add tools** - Ensure all custom connector actions are added as Tools
5. **Test** - Say "travel expense" or "add receipt" to test

## Global Variables Used

```
Global.Version          - "2.0.0"
Global.ApiBaseUrl       - Azure function URL
Global.MileageRate      - 0.70 (default)
Global.RequesterEmail   - User's email
Global.DepartmentCode   - 3-digit code
Global.DepartmentName   - Department name
Global.CartJson         - JSON array of items
Global.CartCount        - Number of items
```

## Custom Connector Actions Required

1. `TravelExpense-OrgChartLookupByUpn` - Department lookup by email
2. `TravelExpense-ExpenseCodesLookup` - Account code lookup
3. `TravelExpense-PerDiemLookup` - GSA per diem rate lookup
4. `TravelExpense-SubmitReport` - Submit expense report
