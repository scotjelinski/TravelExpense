# Travel Expense GL Coding Logic (Simple Rules)

This document summarizes the **minimal logic** to pick a valid GL coding combo for travel expenses using the `Expense codes.xlsx` mapping table.

## Inputs the workflow/agent should collect

- **Department code** (3 digits, e.g., `220`) — preferably auto-filled from the signed-in user / OrgChart when running in Teams  
- **Expense category** (Receipt / Per Diem / Mileage + what it was for)  
- **Optional:** Work order / capital project indicator  
- **Optional:** User-provided GL/account override (if they already know it)

## Step 0 (Teams) - Auto-fill department from OrgChart (preferred)

When the experience runs in Teams, use the signed-in user's email (e.g., `System.User.Email`) to look up their OrgChart record and map the OrgChart `department` name to a 3-digit department code by matching against the `department` column in the expense-code table (values like `620 - INFORMATION TECHNOLOGY`).

## Step 1 - Map the expense to an Activity Code

Use keywords (and the user’s selection) to map to one of these activity codes:

- **700 – Business Travel**: mileage, per diem, airfare, lodging/hotel, taxi/uber/lyft, rental car, parking, tolls

- **710 – Employee Meals and Events**: meals, team meal, catering, event food

- **770 – Training/Education**: conference, training class, seminar, certification; **include travel for that training**

- **790 – Wellness Program**: wellness reimbursement

- (Rare) **730 – Employment Costs** and **740 – Expense Account** exist in the table; only use if the requester explicitly indicates those categories.


If uncertain, ask one clarifying question: *“Is this Business Travel (700), Meals/Events (710), or Training/Education (770)?”*

## Step 2 — If Capital / Work Order, route to clearing account

- If **IsCapital = yes** OR a **WorkOrder** is provided, the expense should be flagged as capital-related and routed to the **107200 clearing account** (finance reclassifies later).  
- Still pick the most appropriate **dept + activity + account** from the table for consistency, but use `107200` as the GL account override.

## Step 3 — Choose the Account Code from the mapping table

Use `Expense codes.xlsx` as source of truth. Each row is a valid combination:

`(Department, Activity) -> Account`


### Default rule (works for almost everything)

1. Filter rows where **department_code = user’s department** AND **activity_code = selected activity**.

2. If there is **exactly one account**, select it.

3. If there are **multiple accounts**, ask a disambiguation question (see outliers below).


### Known outliers (only 3 cases)

Only the following department+activity combos have more than one possible account in the current table:

- Dept **110 – EXECUTIVE ADMINISTRATION**, Activity **770 – TRAINING/EDUCATION** → choose one of: 921 – OFFICE SUPPLIES; 930.206 – MISC GEN EXP- DIRECTOR EXPENSE
- Dept **610 – SAFETY**, Activity **770 – TRAINING/EDUCATION** → choose one of: 921 – OFFICE SUPPLIES; 925 – INJURIES & DAMAGES
- Dept **620 – INFORMATION TECHNOLOGY**, Activity **770 – TRAINING/EDUCATION** → choose one of: 561 – LOAD DISPATCHING; 921 – OFFICE SUPPLIES

Recommended disambiguation question format:

- “I found two valid accounts for your department and Training/Education. Which one matches this expense?”

## Step 4 — Build a GL string (presentation)

Your org discussed segments roughly as:

- **GL account** (often treated as 3-digit/6-digit), then **Department (3-digit)**, then **Account (3-digit)**, then **Activity**.


In practice, for the agent/workflow you can present a human-readable string like:

- `DEPT-ACCOUNT-ACTIVITY` (e.g., `220-561-700`) or

- `GLACCOUNT-DEPT-ACCOUNT-ACTIVITY` (e.g., `107200-220-561-700` for capital/work order)


## Step 5 — Allow user override & finance fallback

- Always show the default coding and ask: **“Use this coding?”** (Yes / Choose another / I’ll enter it / Send to finance)

- If the user says they don’t know and there’s ambiguity, mark as **needs finance review**.


## Appendix — Quick examples

- **Control Center (220) mileage** → Activity 700, Account 561 (Load Dispatching)

- **Any department training conference travel** → Activity 770; account comes from the table (usually one option)
