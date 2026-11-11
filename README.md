# Leave Checker Web

Apps Script lookup and MONE deduction entry for the VER2 Leave Balance snapshot. The app currently opens without Google sign-in and lets visitors browse names and search balances. It displays VL, SL, FL, WL, SPL and the snapshot's date; empty values are unavailable, not zero.

**Status: deployed test reported a missing Sheets permission; the current deployment and access must be verified. Desktop data is not automatically connected. Test with sample data before publishing personnel records.**

## Private data

This is a public source-code repository. Never commit names, balances, a spreadsheet ID, credentials, allowed staff emails, or a `.clasp.json` associated with a private project. The script owner keeps the balance Sheet private; do not share it with web-app viewers. Anyone who can reach the deployed web app can search names and look up balances. The employee number is a lookup input, not authentication. Duplicate employee IDs are rejected. There is no endpoint returning the full sheet.

## Setup overview

1. Create a private Google Sheet with a tab named `LeaveBalances`. Its first row must contain these headers (column order can vary):

   `employee_id | name | as_of | vl | sl | fl | wl | spl | updated_at`

   Store employee IDs as plain text and keep each ID unique. Leave unavailable balances empty. VL and SL columns must be adjacent for MONE entry.
2. Set the Apps Script project property `SPREADSHEET_ID` to the Sheet ID. The deploying account must be able to edit the Sheet.
3. Deploy as a web app executing as the deploying user. Anyone who can access the deployment can browse the published names and balances while anonymous mode is enabled. Redeploy a new version after code changes.
4. Test with sample data, including duplicate IDs and unavailable balances.

To restore Google sign-in, set `anonymous: false` in `config_()`, configure `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, and `ALLOWED_EMAILS`, configure the OAuth redirect URL, and redeploy. The `ALLOW_UNAUTHENTICATED` Script Property is ignored.

## Recording MONE

Select an employee, then enter the MONE start and end dates and the VL and SL deductions. The dates default to November 11–20, 2026 and can be changed. The app subtracts the specified amounts from the current web sheet balances and records before/after values in a `MoneLog` tab. It does not calculate monthly accruals or update the desktop database. The start month cannot be later than the snapshot month; the snapshot date is not changed. The full deduction is applied immediately even when the end date is later than the snapshot date. Duplicate submissions with the same request ID are not applied twice. A pending log entry requires manual review before another entry for that employee.

A later VER2 import that replaces the balance rows may overwrite these web-only deductions. The log is preserved only if the import does not replace or delete its tab. Reconcile the records before publishing another snapshot.

## Data publishing

This repo does **not** calculate leave credits or automatically read the Windows desktop database. The desktop app must export or publish the selected month's computed VER2 balances into the Sheet. Until that publisher is implemented, populate it manually with reviewed data. A web lookup shows only the most recently published snapshot, not live calculations. Review identity matches and unavailable data before publishing.

## Access

Google sign-in and the MONE PIN are disabled. Anyone who can reach the app can browse names and balances and submit MONE deductions. Missing or duplicate IDs cannot be selected for lookup.
