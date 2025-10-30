# Attendance Tracker Library – Google Forms to Master Spreadsheet

## Library Links
- [Primary Library](https://script.google.com/macros/library/d/1DyLX-aHpvoPibzjSfcEmZhKRpfo6hpl4EEWZrcQYf0046Gz1COTmhbtK/10)
- [Alternative/Secondary Library](https://script.google.com/macros/library/d/AKfycbzsE8mGxN4txYdxXrHrw-O3KSHJzyMo1fRLYLsBKGIoadwNhDRRjqzXCFB-AWqn9y1I)

This library aggregates attendance data from multiple Google Forms into a single master Google Spreadsheet. It tracks participation counts, prevents duplicate processing, and supports configurable folders.

---

## Features
- Aggregate multiple Google Forms responses into a single sheet.
- Automatically tracks attendance counts per email.
- Prevents double-counting using a log sheet.
- Supports optional folder/subfolder processing.
- Skips RSVP forms automatically (optional).
- Detailed logging for monitoring processing progress.

---

## Installation
1.  Open [Google Apps Script](https://script.google.com/) and create a new project.
2.  Go to **Libraries** → **Add a Library**.
3.  Paste the library ID:
    `AKfycbzsE8mGxN4txYdxXrHrw-O3KSHJzyMo1fRLYLsBKGIoadwNhDRRjqzXCFB-AWqn9y1I`
4.  Select the latest version and click **Add**.
5.  Save your project.

---

## Usage
```javascript
function runAttendanceImport() {
  const summary = importFormsToSpreadsheet({
    parentFolderName: "Class Attendance Forms",
    outputSpreadsheetName: "Master Attendance",
    outputSheetName: "Attendance",
    subfolderName: "Week 1", // optional
    logSheetName: "ProcessedForms", // optional
    skipRSVP: true, // optional
    verbose: true // optional
  });

  console.log("Attendance import complete:", summary);
}
