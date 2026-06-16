This repo contains files for automatically tracking portco financials inside Google Drive. Companies can email documents to a dedicated CFC reporting inbox, and Apps Script saves attachments into the correct Drive folder while updating the tracker.

1. Make a new Google Sheet
2. Go to Extensions -> Apps Script
3. Paste all .gs files in
4. Go to Settings -> Show "appscript.json" ...
5. Paste in the appscript.json file
6. Set `INTAKE_EMAIL` in `Code.gs` to the dedicated inbox, for example `reporting@cofounderscapital.com`
7. Run setupTracker() in [Setup.gs](http://Setup.gs) and auth
8. Run `installTriggers()` in [Code.gs](http://Code.gs)
9. On refresh, go to CFC Reporting -> Onboard company from prompts

Email intake:

- The Apps Script must run as the Google account that owns or receives mail for the dedicated inbox.
- CEOs should send attachments to the dedicated inbox with the company and reporting month in the subject, for example `Acme Corp May 2026 financials`.
- `gmailInboxWatcher()` searches for unprocessed attachment emails, matches the sender/subject to a tracker row, saves recognized attachments into `Portfolio Reporting/<Company>/<YYYY-MM>/`, and marks matching required docs as `received`.
- A single attachment can satisfy multiple required docs when the filename clearly signals that, for example `Acme model forecast.xlsx` can mark both `model` and `forecast` received.
- Emails that cannot be matched or classified are labeled `CFC_Reporting_Needs_Review`, and attachments are saved in `_email_needs_review`.
- Successfully processed threads are labeled `CFC_Reporting_Processed`.

Google Forms are now optional. Google Forms can contain file upload questions, but Google's APIs do not currently support creating those questions programmatically. The generated form therefore asks CEOs to upload files into their shared Drive folder and paste Drive file links into the form. The `driveFolderWatcher()` trigger also marks files received when they are dropped directly into the company/month folder.

Still need to test - waiting for my CFC account 2FA to be fixed. 
