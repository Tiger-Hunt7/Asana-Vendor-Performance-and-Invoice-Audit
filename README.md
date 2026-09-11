# Asana Vendor Performance and Invoice Audit

Management requested this project in response to a vendor payment dispute. The vendor asserted that more than **$40,000 remained unpaid**. The evidence review found **$0 payable for completed work remained unpaid**.

The analysis identified **$23,685 in denied amounts** associated with work that was not performed or did not meet documented requirements. Those amounts were not unpaid obligations. The project provided management with a detailed, traceable basis for the organization's position.

## What the project did

- Structured an exported Asana project history into an auditable dataset.
- Reviewed 147 tasks, including completion, open work, dates, sections, and inspection results.
- Created a 91-row itemized financial ledger.
- Separated reported paid activity from denied work.
- Compared itemized evidence with earlier published summaries.
- Identified data-quality issues and records requiring financial-system reconciliation.

## Key findings

| Measure | Result |
|---|---:|
| Vendor unpaid claim | More than $40,000 |
| Payable unpaid amount found | **$0** |
| Itemized reported paid | $139,948 |
| Denied — work not performed / requirements unmet | $23,685 |
| Itemized total reviewed | $163,633 |
| Asana tasks reviewed | 147 |
| Failed inspections | 45 |
| Tasks without a reporting date | 21 |

The review also found that older summaries exceeded itemized totals by $12,208 for reported paid activity and $1,847 for denied activity. These differences required reconciliation rather than assumption.

## Repository contents

- `Asana_Vendor_Performance_and_Invoice_Audit_Fictionalized.xlsx` — dashboard, performance summaries, itemized ledger, data-quality findings, and sanitized task detail.
- `Asana_Vendor_Performance_and_Invoice_Audit_Project_Report.docx` — management-ready explanation of the request, method, findings, and conclusion.
- `PROJECT_OVERVIEW.md` — browser-readable project narrative.
- `build_workbook.mjs` and `create_report.py` — reproducible artifact-building scripts.

## Important control distinction

Asana was useful evidence for operational dates, task completion, project sections, and inspection history. It was **not** treated as the authoritative payment ledger. Final payment certification still requires reconciliation to the organization's financial system and supporting invoice records.

## Data privacy

This is a fictionalized public portfolio version. Vendor, organization, employee, location, task, invoice, link, and identifying narrative details were changed. The numerical findings and analytical structure were retained to demonstrate the analysis accurately without identifying the original parties.
