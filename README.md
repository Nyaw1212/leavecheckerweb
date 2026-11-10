# Leave Checker Web

Read-only Apps Script lookup for the VER2 Leave Balance snapshot. An approved staff member signs in with Google, then enters an employee number. It displays VL, SL, FL, WL, SPL and the snapshot's date; empty values are unavailable, not zero.

**Status: deployed test reported a missing Sheets permission; the current deployment and access must be verified. Desktop data is not automatically connected. Test with sample data before publishing personnel records.**

## Private data

This is a public source-code repository. Never commit names, balances, a spreadsheet ID, credentials, allowed staff emails, or a `.clasp.json` associated with a private project. The script owner keeps the balance Sheet private; do not share it with web-app viewers. Only staff listed in server-side Script Properties may query records. The employee number is a lookup input, not authentication; any approved account can look up any employee number. Duplicate employee IDs are rejected. There is no endpoint returning the full sheet.

## Setup overview

1. Create a private Google Sheet with a tab named `LeaveBalances`. Its first row must contain these headers (column order can vary):

   `employee_id | name | as_of | vl | sl | fl | wl | spl | updated_at`

   Store employee IDs as **plain text** to preserve leading zeros. Use actual calculated month-end VER2 values. Leave missing balances empty; use zero only for an actual zero. Use ISO dates such as `2026-11-30` and record when the snapshot was published. Each ID must occur once. 
2. Create a standalone Apps Script project and add `Code.js`, `Index.html`, and the manifest from this repository. Configure the project to use a Google Cloud project. Create a **Web application** OAuth client for Google sign-in. Configure the consent screen and approved test users if the app is in testing mode. You will need its client ID and secret.
3. In Apps Script project settings, add Script Properties: `SPREADSHEET_ID`, `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, and `ALLOWED_EMAILS` (a comma-separated list of explicitly approved staff Gmail addresses). Never place these values in the repository. The deploying account must have read access to the private Sheet.
4. Deploy as a web app, **execute as the deploying user**. The deployment must be reachable by the intended personal Google accounts; the script itself enforces the email allowlist. Add the exact deployed `/exec` URL as an authorized redirect URI in the Google OAuth client's Cloud Console settings. Redeploy when code changes. An OAuth configuration can require Google verification or test-user configuration depending on its publishing status. Don't publish real data until the full flow and sharing are verified.
5. Test with dummy balances first: an approved and an unapproved Google account, incorrect IDs, duplicate IDs, unavailable balances, expired sessions, and an ordinary viewer lacking access to the underlying Sheet. Verify that no private Sheet or project is accessible to the viewer.

To update who can search, edit `ALLOWED_EMAILS`; an existing session is checked against the current allowlist on every lookup. Sessions expire after 30 minutes. Changing the Sheet or its contents does not require a web-app deployment.

## Data publishing

This repo does **not** calculate leave credits or automatically read the Windows desktop database. The desktop app must export or publish the selected month's computed VER2 balances into the Sheet. Until that publisher is implemented, populate it manually with reviewed data. A web lookup shows only the most recently published snapshot, not live calculations. Review identity matches and unavailable data before publishing.

## Temporary unauthenticated access

Set Script Property `ALLOW_UNAUTHENTICATED` to the exact value `true` to disable Google sign-in. Anyone who can reach the deployment can look up balances by employee number; employee IDs are not secrets. Use dummy data only. Set the property to `false` or delete it to restore sign-in. Do not remove existing OAuth configuration. The server checks this property for each lookup.

Personnel names are masked in lookup responses. Masking is not authentication and does not prevent someone who knows an employee number from accessing that employee's balances.
