This repo contains files for automatically tracking portco financials inside Google Drive. To test:

1. Make a new Google Sheet
2. Go to Extensions -> Apps Script
3. Paste all .gs files in
4. Go to Settings -> Show "appscript.json" ...
5. Paste in the appscript.json file
6. Run setupTracker() in [Setup.gs](http://Setup.gs) and auth
7. Run installTriggers() in [Code.gs](http://Code.gs)
8. On refresh, go to CFC Reporting -> Onboard company from prompts

Still need to test - waiting for my CFC account 2FA to be fixed. 

