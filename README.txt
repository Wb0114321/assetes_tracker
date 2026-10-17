ASSET MANAGEMENT SYSTEM — GITHUB PAGES + GOOGLE SHEETS

FILES
- Code.gs: Apps Script backend
- index.html: GitHub Pages frontend

A) APPS SCRIPT
1. Open the existing Google Spreadsheet containing Assets, Departments and Users.
2. Extensions > Apps Script. Replace Code.gs with the supplied file and Save.
3. Run setupAssetApp once and authorize it. Existing rows are preserved; missing tabs are created; headers are added only if row 1 is blank. If Users has no data rows, demo login is created.
4. Deploy > New deployment > Web app.
5. Execute as: Me.
6. Who has access: Anyone.
7. Deploy and copy the /exec URL.

B) CONFIGURE FRONTEND
Open index.html and replace:
const API_URL="PASTE_YOUR_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE";
with your deployed /exec URL, keeping the quotation marks.

C) GITHUB PAGES
1. Create a repository, e.g. asset-management.
2. Upload index.html to repository root.
3. Repository Settings > Pages > Deploy from branch > main > /(root).
4. Open the published HTTPS URL.

FIRST LOGIN
If Users was empty: admin / DemoAdmin@123.
If Users already has users, use an existing account. Change demo password before real use.

CAMERA
Use the GitHub Pages HTTPS URL, not the Apps Script /exec page. Click Scan and allow camera permission. If previously blocked, Chrome site settings must be changed to Allow. JavaScript cannot forcibly grant browser permission. html5-qrcode loads from a CDN, so internet is needed.

DATA LAYOUT
Assets: A S.no, B Item Name, C Serial Number, D Specification, E Qty, F Department, G Location, H Other Remarks.
Departments: A Dept_ID, B DeptName.
Users: A User_id, B User_name, C Password.

SECURITY
This starter preserves the existing plain-text Users schema. Use a strong unique password and restrict spreadsheet access. The web app must be publicly reachable for GitHub Pages to call it; actions are gated by login tokens, but this starter is not a hardened enterprise authentication system. Do not commit passwords or private data to GitHub.
