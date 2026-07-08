This repo contains files for automatically tracking portco financials inside Google Drive. Companies can email documents to a dedicated CFC reporting inbox, and Apps Script saves attachments into the correct Drive folder while updating the tracker.

1. Make a new Google Sheet
2. Go to Extensions -> Apps Script
3. Paste all .gs files in
4. Go to Settings -> Show "appscript.json" ...
5. Paste in the appscript.json file
6. Set `INTAKE_EMAIL` in `Code.gs` to the dedicated inbox, for example `reporting@cofounderscapital.com`
7. Add an Apps Script property named `OPENAI_API_KEY` with your OpenAI API key
8. Run setupTracker() in [Setup.gs](http://Setup.gs) and auth
9. Run `installTriggers()` in [Code.gs](http://Code.gs)
10. On refresh, go to CFC Reporting -> Onboard company from prompts

Email intake:

- The Apps Script must run as the Google account that owns or receives mail for the dedicated inbox.
- CEOs should send attachments to the dedicated inbox with the company and reporting month in the subject, for example `Acme Corp May 2026 financials`.
- `gmailInboxWatcher()` searches for unprocessed attachment emails, matches the sender/subject to a tracker row, uses the OpenAI Responses API to classify each attachment, saves recognized attachments into `Portfolio Reporting/<Company>/<YYYY-MM>/`, and marks matching required docs as `received`.
- A single file can satisfy multiple required docs when the LLM finds multiple materials in it, for example a workbook with financials, model, and forecast tabs can mark all three docs received. Multi-type files are stored once with a composite name such as `financials_forecast_model.xlsx`.
- Emails that cannot be matched or classified are labeled `CFC_Reporting_Needs_Review`, and attachments are saved in `_email_needs_review`.
- Successfully processed threads are labeled `CFC_Reporting_Processed`.
- Each LLM classification writes per-document-type audit rows to the `Document Classification Log` sheet, including confidence, reasoning, whether the LLM matched the type, and whether the pipeline accepted that type for the tracker.

Google Forms are now optional. Google Forms can contain file upload questions, but Google's APIs do not currently support creating those questions programmatically. The generated form therefore asks CEOs to upload files into their shared Drive folder and paste Drive file links into the form. The `driveFolderWatcher()` trigger also marks files received when they are dropped directly into the company/month folder.

Document classification and flash report generation use the OpenAI Responses API. The default model is configured in `Code.gs` as `OPENAI_MODEL`.

Still need to test - waiting for my CFC account 2FA to be fixed. 
