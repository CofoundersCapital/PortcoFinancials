This repo contains files for automatically tracking portco financials inside Google Drive. Companies can email documents to a dedicated CFC reporting inbox, and Apps Script saves attachments into the correct Drive folder while updating the tracker.

1. Make a new Google Sheet
2. Go to Extensions -> Apps Script
3. Paste all .gs files in
4. Go to Settings -> Show "appscript.json" ...
5. Paste in the appscript.json file
6. Run setupTracker() in [Setup.gs](http://Setup.gs) and auth
7. Add an Apps Script property named `OPENAI_API_KEY` with your OpenAI API key
8. In the `Master Config` sheet, confirm or edit `INTAKE_EMAIL`, `CFC_TEAM_EMAIL`, reminder schedule, notification toggles, and folder sharing settings
9. Run `installTriggers()` in [Code.gs](http://Code.gs)
10. On refresh, go to CFC Reporting -> Onboard company from prompts

Master Config:

- `setupTracker()` creates a `Master Config` tab with editable `value` cells and visible `default_value` cells. Defaults are seeded from the current Apps Script code, including `INTAKE_EMAIL=cofoundersreporting@gmail.com`, `OPENAI_MODEL=gpt-5.4-nano`, deadline/reminder timing, Gmail labels, and extraction limits.
- Use `CFC Reporting -> Refresh master config` to add newly introduced settings without overwriting existing value cells.
- Blank or invalid setting values fall back to the visible default, then the code default, and are logged in `Logs`.
- Feature toggles can disable email intake, Drive watcher processing, form intake, overdue reminders, escalations, flash report generation, and automatic flash report generation on completion.
- `GMAIL_WATCHER_INTERVAL_MINUTES` and `DRIVE_WATCHER_INTERVAL_MINUTES` control how often automatic watcher triggers run. After changing either value, run `CFC Reporting -> Install triggers` so Apps Script recreates those triggers at the new frequency.
- Notification toggles control founder and assigned board member document/completion emails, team warning emails, classification review alerts, and flash-report-ready alerts.
- `CEO_FOLDER_ACCESS` and `BOARD_MEMBER_FOLDER_ACCESS` control whether those users are added as `none`, `viewer`, or `editor` on company folders. Use `CFC Reporting -> Apply folder sharing settings` to sync current settings to existing company folders.

Email intake:

- The Apps Script must run as the Google account that owns or receives mail for the dedicated inbox.
- CEOs should send attachments to the dedicated inbox with the company and reporting month in the subject, for example `Acme Corp May 2026 financials`.
- `gmailInboxWatcher()` searches for unprocessed attachment emails, matches the sender/subject to a tracker row, uses the OpenAI Responses API to classify each attachment, saves recognized attachments into `Portfolio Reporting/<Company>/<YYYY-MM>/`, and marks matching required docs as `received`.
- A single file can satisfy multiple required docs when the LLM finds multiple materials in it, for example a workbook with financials, model, and forecast tabs can mark all three docs received. Multi-type files are stored once with a composite name such as `financials_forecast_model.xlsx`.
- Emails that cannot be matched or classified are labeled `CFC_Reporting_Needs_Review`, and attachments are saved in `_email_needs_review`.
- Successfully processed threads are labeled `CFC_Reporting_Processed`.
- Each LLM classification writes per-document-type audit rows to the `Document Classification Log` sheet, including confidence, reasoning, whether the LLM matched the type, and whether the pipeline accepted that type for the tracker.
- Spreadsheet classification previews divide the classification character budget evenly across tabs, so later workbook tabs are still represented in the LLM prompt.

Google Forms are now optional. Google Forms can contain file upload questions, but Google's APIs do not currently support creating those questions programmatically. The generated form therefore asks CEOs to upload files into their shared Drive folder and paste Drive file links into the form. The `driveFolderWatcher()` trigger also marks files received when they are dropped directly into the company/month folder.

The `deadline` column in the Submissions sheet is editable. New rows default it from the reporting month, and changing it recalculates `days_overdue` while resetting reminder/escalation timing for incomplete rows. Use `CFC Reporting -> Refresh tracker columns/formulas` to apply deadline validation and backfill missing deadlines on an existing tracker.

Document classification and flash report generation use the OpenAI Responses API. The default model is configured in `Code.gs` as `OPENAI_MODEL`.
