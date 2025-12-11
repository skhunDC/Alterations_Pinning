# Dublin Cleaners Alterations Pinning Certification - Trainer Notes

## Purpose
- Certify CSRs to pin safely, record final measurements, and write clear SPOT POS notes.
- Store module completions in Google Sheets for HR visibility and recertification tracking.

## App Structure
- **Code.gs**: Web entry (`doGet`), sheet bootstrap, and endpoints (`saveModuleResult`, `getModuleStatus`, `getModulesCatalog`, `getLatestRecordsForPrint`).
- **index.html**: Main UI with module cards, employee inputs, and progress grid.
- **scripts.html**: Frontend logic for rendering modules, quizzes, saving results, and syncing progress.
- **styles.html**: Responsive styling tuned for mobile and print-friendly layouts.
- **print.html**: Optional print/PDF view; wire to `getLatestRecordsForPrint` for HR exports.
- **appsscript.json**: Manifest with web app settings and scopes.

## Sheet Design
- Spreadsheet name: **Dublin Cleaners Training**; tab: **TrainingProgress**.
- Header: Timestamp, EmployeeName, LocationOrID, ModuleID, ModuleTitle, QuizScore, PassFail, Notes.
- The app auto-creates the file/tab and rewrites the header if missing.

## Module Logging Rules
- Quiz submissions default to PASS at ≥80% and store score + note.
- "Save Note Only" writes a NOTE status without a score (useful for demos or supervisor sign-off steps).
- Progress grid shows the latest attempt per module per employee/location.

## Deployment Tips
- Publish as web app: **Execute as User accessing the web app**, **Anyone** access (adjust if SSO is added later).
- To preview locally, use clasp + `clasp push`/`clasp open`, or run the HTML in Apps Script IDE preview.
- Add HR/QA reviewers to the spreadsheet for auditing and recertification checks every 18–24 months.

## Extending Modules
- To add modules, update the `modules` array in `scripts.html` and mirror IDs/titles in `getModulesCatalog()` if needed.
- Keep quizzes 2–4 questions with explicit references to safety pins, objective measurements, and escalation rules.
- Avoid +/- shorthand in any new content; always store final measurements.
