# Alterations Pinning Certification Web App

Google Apps Script + HTMLService single-page experience that delivers Dublin Cleaners' five-module Alterations Pinning Certification and logs quiz outcomes to Google Sheets.

## Files
- `Code.gs` — Apps Script backend (HTMLService entrypoint, Sheet helpers, quiz persistence, certification status API).
- `appsscript.json` — Manifest configured for web app deployment.
- `index.html` — Main UI shell with navigation, status, and modules.
- `styles.html` — Inline CSS for the modern, responsive layout (sticky header, active navigation state, progress styles).
- `scripts.html` — Client-side JS for module rendering, quizzes, Sheet saves, status updates, navigation focus, localStorage helpers, and toasts.
- `print.html` — Print-friendly certificate view that reflects module completion and highlights the employee name.
- `docs/` — Trainer guidance and project overview docs.
- `package.json` + `tests/` — Minimal local test harness placeholder.

## Deploying as a Web App
1. Create a new Apps Script project and add all files above (File → New → HTML file for each partial; `Code.gs` as a script file; replace manifest with `appsscript.json`).
2. Publish → Deploy as web app → Execute as **User accessing the web app** → Access: **Anyone** (adjust later for SSO).
3. On first save, the script creates a spreadsheet named **Alterations Pinning Certification** with sheets **ModuleResults** (scores) and **QuizAttempts** (question-level detail).

## How Data Is Stored
- Each quiz submission calls `saveModuleAttempt(moduleId, employeeName, employeeLocationOrId, score, passed, correctCount, totalQuestions, answers)` in `Code.gs`, which logs both summary scores and detailed answers while keeping `saveModuleResult` intact for compatibility.
- **ModuleResults** rows include: timestamp, employee name, location/ID, module ID (M1–M5), score, passed (TRUE/FALSE).
- **QuizAttempts** rows include: timestamp, attemptId, employee identity, module ID, question number/id/text, selected answer (id/label), correct answer (id/label), per-question correctness, overall correct count, total questions, score %, and pass/fail.
- `getEmployeeCertificationStatus(employeeName)` aggregates passing attempts to report completed and missing modules plus overall certification.

## Using the UI
1. Enter employee name (required) and location/ID (optional) at the top. The app remembers these via localStorage when available.
2. Navigate modules via sticky top navigation or sidebar; read objectives, content, visuals, and take the quiz (83% pass threshold — 5 of 6 correct). Active navigation is highlighted and focusable.
3. Submit a quiz to save to Sheets and refresh status. The dashboard shows modules remaining, a progress bar with estimated minutes, and a checklist.
4. Use **Print Certificate** to open `print.html` with the employee name and completion table for HR files and supervisor sign-off (only when certified).
5. Admins can reference the underlying ModuleResults sheet directly in Google Sheets for summary reporting.

## Recertification + Sign-off Notes
- Program suggests recertification every 18–24 months and immediate refreshers after documentation issues.
- Supervisor sign-off for live pinning is captured outside the app but noted on the printable certificate.

## Testing
- Run `npm test` to execute lightweight helper tests in `tests/app.test.js` (no dependencies required).

## Extending Reporting
- `Code.gs#getAllModuleResults` returns each ModuleResults row.
- `Code.gs#getSummaryByEmployee` returns aggregated completion by employee name.
