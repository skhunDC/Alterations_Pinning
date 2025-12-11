# Alterations Pinning Certification Web App

Google Apps Script + HTMLService single-page experience that delivers Dublin Cleaners' five-module Alterations Pinning Certification and logs quiz outcomes to Google Sheets.

## Files
- `Code.gs` — Apps Script backend (HTMLService entrypoint, Sheet helpers, quiz persistence, certification status API).
- `appsscript.json` — Manifest configured for web app deployment.
- `index.html` — Main UI shell with navigation, status, and modules.
- `styles.html` — Inline CSS for the modern, responsive layout.
- `scripts.html` — Client-side JS for module rendering, quizzes, Sheet saves, and status updates.
- `print.html` — Print-friendly certificate view that reflects module completion.
- `docs/` — Trainer guidance and project overview docs.
- `package.json` + `tests/` — Minimal local test harness placeholder.

## Deploying as a Web App
1. Create a new Apps Script project and add all files above (File → New → HTML file for each partial; `Code.gs` as a script file; replace manifest with `appsscript.json`).
2. Publish → Deploy as web app → Execute as **User accessing the web app** → Access: **Anyone** (adjust later for SSO).
3. On first save, the script creates a spreadsheet named **Alterations Pinning Certification** with sheet **ModuleResults**.

## How Data Is Stored
- Each quiz submission calls `saveModuleResult(moduleId, employeeName, employeeLocationOrId, score, passed)` in `Code.gs`.
- Rows include: timestamp, employee name, location/ID, module ID (M1–M5), score, passed (TRUE/FALSE).
- `getEmployeeCertificationStatus(employeeName)` aggregates passing attempts to report completed and missing modules plus overall certification.

## Using the UI
1. Enter employee name (required) and location/ID (optional) at the top.
2. Navigate modules via sidebar; read objectives, content, visuals, and take the quiz (80% pass threshold).
3. Submit a quiz to save to Sheets and refresh status. The banner shows remaining modules and progress bar.
4. Use **Print Certificate** to open `print.html` with the employee name and completion table for HR files and supervisor sign-off.

## Recertification + Sign-off Notes
- Program suggests recertification every 18–24 months and immediate refreshers after documentation issues.
- Supervisor sign-off for live pinning is captured outside the app but noted on the printable certificate.

## Testing
- Run `npm test` to execute the placeholder sanity check in `tests/app.test.js` (no dependencies required).
