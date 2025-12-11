# Dublin Cleaners Alterations Pinning Training

Use this employee-facing web app to deliver the five-module Alterations Pinning Certification and track quiz results in Google Sheets.

## Trainer Quick Start
- Open the deployed web app, enter the employee’s name (and location/ID if needed), and click **Save/Refresh**.
- Coach learners through each module card (objectives → content → visuals → quiz). Passing score is 80%.
- After each quiz, results save to the **Alterations Pinning Certification** spreadsheet, sheet **ModuleResults**.
- Check the status banner: remaining modules listed until all five pass. Then print the certificate for HR files.

## How Progress Is Calculated
- Backend function: `getEmployeeCertificationStatus(employeeName)` returns passed modules, missing modules, and `isCertified` flag.
- The progress bar fills based on M1–M5 pass results (latest passing attempt counts).

## Tips for Running Sessions
- Reinforce objective measurements: “shorten inseam to 30 inches,” never “hem 2 inches.”
- Require safety pins and horizontal pinning for any garment traveling to the plant.
- Verify customer-pinned garments with a tape measure and record final measurements.
- Use SPOT examples in Module 4 to model clear notes.
- Reference escalation scripts in Module 5 for delicate or high-risk garments.

## Updating Content Later
- Module copy and quizzes live in `scripts.html` within the `modules` array. Update text there for new policies or visuals.
- Styling updates can be made in `styles.html`; keep the logo and brand colors intact.
- Sheet headers are set in `Code.gs` (`MODULE_HEADERS`). Avoid altering order unless you also update downstream reporting.

## Data & Records
- Spreadsheet: **Alterations Pinning Certification** → **ModuleResults**. Columns: Timestamp, EmployeeName, LocationOrID, ModuleID, Score, Passed.
- Certificate print view (`print.html`) mirrors module completion; supervisor sign-off remains an external step.
