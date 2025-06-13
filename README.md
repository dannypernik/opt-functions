## Summary of `findNewScoreReports()`

The `findNewScoreReports()` function identifies recently updated SAT admin spreadsheets within a specified time limit (default: 5 days). It processes these files to check for new completed tests and generates score reports for them. If no `students` parameter is provided, it retrieves student data from the "Clients" sheet in the `clientDataSs` spreadsheet.

---

## Instructions for Using `findNewScoreReports()` with a Daily Time-Based Trigger

### 1. Set Up the Script Properties

Ensure the `clientDataSsId` property is set in the script properties. This should point to the ID of the spreadsheet containing the "Clients" sheet.

- Go to **Project Settings** (the gear icon in the left sidebar).
- Scroll to the bottom and click **Add script property**.
- Name the property `clientDataSsId` and add the ID of your client data spreadsheet (the part of the Sheets URL between `/d/` and the trailing slash).
- You can make a copy of this spreadsheet:
  [Client Data Spreadsheet Template](https://docs.google.com/spreadsheets/d/1U4SzaEwcFEMoJEqb0U3G08e92qFuAlwQBNudEc06esU/edit?usp=sharing)

---

### 2. Remove or update code as needed

- `getOPTPermissionsList()` is designed to send a group email to any @openpathtutoring.com email addresses that have access to a student's admin spreadsheet. This is useful if more than one person in the organization should receive the score report. The tutor or admin can then forward the email to a parent and/or student. You can either switch the domain name within getOPTPermissionsList to your own domain name or replace and set `const email = 'youremail@yourdomain.com`.
- Remove to `const studentData = getStudentsSpreadsheetData()` in `sendPdfScoreReport()` and simplify the `const message` to exclude `studentData` references.

---

### 2. Create Daily Triggers for `updateOPTStudentFolderData` and `findNewScoreReports`

- In the left sidebar of the Apps Script editor, click **Triggers** (the stopwatch icon).
- Click **Add trigger** in the bottom right corner.
- Select `updateOPTStudentFolderData` from the function list.
- Select **Time-driven** from event source.
- Leave the deployment set as **Head**.
- Choose **Daily timer** and the hour during which you’d like the function to run.
- Repeat the process for `findNewScoreReports`.

**Recommendation:**
Run `updateOPTStudentFolderData` before `findNewScoreReports` each day, since the latter searches for new tests from the folder data that is saved to the `clientDataSs` by the update function.

---

### 3. Verify Dependencies

Ensure all required functions are present and working correctly. These include:

- `findNewCompletedTests()`
- `getTestCodes()`
- `createSatScoreReports()`
- `sendPdfScoreReport()`
- `savePdfSheet()`
- `mergePDFs()`
- `getLastFilledRow()`
- `getOPTPermissionsList()`

---

### 4. Test the Function

Run `updateOPTStudentFolderData()` and `findNewScoreReports()` manually to ensure they work as expected before relying on the triggers. When `updateOPTStudentFolderData()` completes, you should see values entered into the "Clients" sheet. `findNewScoreReports()` should run from the Apps Script editor without errors.

---

## List of Functions Needed for `findNewScoreReports()`

### `scoreReports.gs`

- **`findNewScoreReports()`**: Main function to identify updated SAT admin spreadsheets.
- **`findNewCompletedTests(fileList)`**: Processes the list of updated files to check for new completed tests.
- **`createSatScoreReports(spreadsheetId, scores)`**: Generates SAT score reports for completed tests.
- **`sendPdfScoreReport(spreadsheetId, studentName, scoresUpToCurrent)`**: Sends the PDF score report via email.
- **`savePdfSheet(spreadsheetId, sheetId, studentName)`**: Exports a specific sheet as a PDF.
- **`mergePDFs(fileIds, destinationFolderId, name)`**: Merges multiple PDF files into one.
- **`getStudentsSpreadsheetData(studentName)`**: Retrieves student data from the "Summary" sheet.

### `helpers.gs`

- **`getLastFilledRow(sheet, col)`**: Finds the last filled row in a column.
- **`getTestCodes()`**: Retrieves unique test codes from the "Practice test data" sheet.
- **`getOPTPermissionsList(id)`**: Retrieves a list of editors for a file.

## File Structure and Function Reference

- `findNewScoreReports(students, folderName)`
- `updateOPTStudentFolderData()`
- `findNewCompletedTests(fileList)`
- `createSatScoreReports(spreadsheetId, scores)`
- `sendPdfScoreReport(spreadsheetId, studentName, scoresUpToCurrent = [])`
- `mergePDFs(fileIds, destinationFolderId, name="merged.pdf")`
- `savePdfSheet(spreadsheetId, sheetId, studentName)`

---

### `helpers.gs`

- `getLastFilledRow(sheet, col)`
- `getOPTPermissionsList(id)`
- `getTestCodes()`

---

## Example Workflow

1. **Daily Trigger Execution**: The trigger runs `updateOPTStudentFolderData()` and `findNewScoreReports()` every day at the specified time.
2. **File Identification**: `findNewScoreReports()` identifies updated SAT admin spreadsheets.
3. **Test Processing**: `findNewCompletedTests()` checks for new completed tests in the identified files.
4. **Score Report Generation**: `createSatScoreReports()` generates score reports for completed tests.
5. **PDF Export and Email**: `sendPdfScoreReport()` exports the score report as a PDF and emails it to the relevant recipients.

---

## Notes

- Ensure all required permissions (Drive, Sheets, Mail) are enabled in the script manifest ([appsscript.json](https://github.com/dannypernik/opt-functions/blob/main/appsscript.json)).
- Adjust the time limit in `findNewScoreReports()` if needed:
  ```javascript
  const msInTimeLimit = 5 * 24 * 60 * 60 * 1000; // 5 days
  ```
