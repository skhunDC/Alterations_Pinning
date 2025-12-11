# Alterations Pinning Certification Web App (Apps Script)

Google Apps Script + HTMLService training hub for Dublin Cleaners CSRs. The app delivers the five-module Alterations Pinning Certification, tracks quiz completions in Google Sheets, and provides a print-friendly view for HR records.

## What’s Included
- **index.html**: Main UI with module cards, employee info fields, and progress grid.
- **scripts.html**: Module content, quizzes, and client-side logic for saving and rendering progress.
- **styles.html**: Clean, responsive styling for desktop and mobile.
- **Code.gs**: Apps Script backend (sheet bootstrap, web entry, save/load progress helpers).
- **print.html**: Printable completion summary (connect to `getLatestRecordsForPrint`).
- **appsscript.json**: Manifest with Sheets/Drive scopes and public web app settings.
- **docs/**: Trainer notes and agent guidance.
- **package.json** + **tests**: Lightweight local test harness placeholder.

## Running & Deploying
1. Create an Apps Script project and add the files above.
2. Deploy as a web app: Execute as **User accessing the web app**, access **Anyone** (adjust if SSO later).
3. On first load the app creates a spreadsheet named **Dublin Cleaners Training** with tab **TrainingProgress** and headers: Timestamp, EmployeeName, LocationOrID, ModuleID, ModuleTitle, QuizScore, PassFail, Notes.
4. Employees enter their name/location, complete module quizzes, and submissions append rows to the sheet. The progress grid shows the latest attempt per module.

## Certification Flow
- Modules reflect the Alterations Pinning curriculum: objective measurements, safety pins & horizontal orientation, garment-specific rules, SPOT POS clarity, and escalation criteria.
- Pass threshold: 80% per module micro-quiz. Quick note saves log practice/review without a score.
- Recertification suggested every 18–24 months; supervisors can export the print view for HR files.

## Testing
- Minimal placeholder test script included: `npm test` (prints guidance). Expand with linting or content checks as needed.
