# Monthly Operational UAT Checklist Template

Use this checklist for each campaign run, post-schema update, and major engine change.

## Run metadata

- Campaign / Release:
- Environment (Workbook URL/ID):
- Operator:
- Reviewer:
- Run date:
- Engine version (`VDM_CONFIG.VERSION`):

## 10-case verification matrix

| # | Scenario | Input / Setup Summary | Expected Result | Actual Result | Pass/Fail | Evidence (link/screenshot/log) | Owner | Notes |
|---:|---|---|---|---|---|---|---|---|
| 1 | Happy Path |  |  |  |  |  |  |  |
| 2 | Missing Master Tab |  |  |  |  |  |  |  |
| 3 | Missing Catalog Header |  |  |  |  |  |  |  |
| 4 | EEI Row Count < 6 |  |  |  |  |  |  |  |
| 5 | EEI Row 5 Header Mismatch |  |  |  |  |  |  |  |
| 6 | Sales Schema Mismatch |  |  |  |  |  |  |  |
| 7 | Zero-Cost Non-GWP |  |  |  |  |  |  |  |
| 8 | Zero-Cost GWP Exception |  |  |  |  |  |  |  |
| 9 | Queue Precedence & Non-Duplication |  |  |  |  |  |  |  |
| 10 | Midpoint Tie Percentile Math |  |  |  |  |  |  |  |

## Sign-off

- [ ] All 10 scenarios executed
- [ ] All failing scenarios have remediation tickets
- [ ] Queue precedence behavior validated
- [ ] Guardrail and hold behavior validated
- [ ] Pre-flight check passed

Approver name:
Approval date:
