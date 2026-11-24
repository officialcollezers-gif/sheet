# Center Admissions

Enter a center code to see the January and July 2026 SVGU, VGU, and total admissions from Google Sheets. January 2027 displays no data.

## Deploy

1. Create a standalone Google Apps Script project using an account with access to the source sheet.
2. Paste the contents of Code.gs into the script file. Add an HTML file named Index and paste in Index.html.
3. In Project Settings → Script properties, add ADMISSIONS_SHEET_ID using the source spreadsheet ID.
4. Deploy as a web app, execute as yourself, and start with Only myself for verification. Authorize Google Sheets access and open the deployment URL.

A center code selects a record. Anyone granted access to the web app can look up any code they know. Updates to the source sheet appear on the next lookup.
