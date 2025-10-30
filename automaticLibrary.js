/**
 * Attendance Tracker Library – Google Forms to Master Spreadsheet
 * GitHub: https://github.com/aamoghS/attendanceAutomator
 *
 * Aggregates attendance data from multiple Google Forms into a single master spreadsheet.
 * Tracks participation counts, prevents duplicate processing, and supports configurable folders.
 *
 * @version 12
 * @author Aamogh Sawant
 */

/**
 * Main function to aggregate attendance data from multiple Google Forms into a single master spreadsheet.
 * Tracks participation counts, prevents duplicate processing, and supports configurable folders.
 *
 * @param {Object} config Configuration object containing settings for the import process
 * @param {string} config.parentFolderName Name of the parent folder containing forms (required)
 * @param {string} config.outputSpreadsheetName Name of the master output spreadsheet (required)
 * @param {string} config.outputSheetName Name of the sheet within the spreadsheet (required)
 * @param {string} [config.subfolderName] Optional subfolder name to process
 * @param {string} [config.logSheetName="Processed"] Name of log sheet
 * @param {boolean} [config.skipRSVP=true] Whether to skip forms with "RSVP" in name
 * @param {boolean} [config.verbose=true] Enable detailed logging
 * @return {Object} Summary object containing processing statistics and spreadsheet URL
 */
function importFormsToSpreadsheet(config) {
  if (!config || !config.parentFolderName || !config.outputSpreadsheetName || !config.outputSheetName) {
    throw new Error("Missing required config: parentFolderName, outputSpreadsheetName, and outputSheetName are required");
  }
  
  const CONFIG = {
    parentFolderName: config.parentFolderName,
    subfolderName: config.subfolderName || null,
    outputSpreadsheetName: config.outputSpreadsheetName,
    outputSheetName: config.outputSheetName,
    logSheetName: config.logSheetName || "Processed",
    skipRSVP: config.skipRSVP !== false,
    verbose: config.verbose !== false,
    masterHeaders: ["Name", "Email", "Count"]
  };
  
  log("Starting attendance import process", CONFIG.verbose);
  
  // Get or create output spreadsheet
  let outputSS;
  const existingFiles = DriveApp.getFilesByName(CONFIG.outputSpreadsheetName);
  if (existingFiles.hasNext()) {
    outputSS = SpreadsheetApp.open(existingFiles.next());
    log("Found existing spreadsheet: " + CONFIG.outputSpreadsheetName, CONFIG.verbose);
  } else {
    outputSS = SpreadsheetApp.create(CONFIG.outputSpreadsheetName);
    log("Created new spreadsheet: " + CONFIG.outputSpreadsheetName, CONFIG.verbose);
  }

  // Get or create output sheet
  let outputSheet = outputSS.getSheetByName(CONFIG.outputSheetName);
  if (!outputSheet) {
    outputSheet = outputSS.insertSheet(CONFIG.outputSheetName);
    outputSheet.appendRow(CONFIG.masterHeaders);
    log("Created new sheet: " + CONFIG.outputSheetName, CONFIG.verbose);
  } else {
    log("Using existing sheet: " + CONFIG.outputSheetName, CONFIG.verbose);
  }
  
  // Get or create log sheet
  let logSheet = outputSS.getSheetByName(CONFIG.logSheetName);
  if (!logSheet) {
    logSheet = outputSS.insertSheet(CONFIG.logSheetName);
    logSheet.appendRow(["Processed Form IDs"]);
    logSheet.hideSheet();
    log("Created log sheet: " + CONFIG.logSheetName, CONFIG.verbose);
  }

  // Load existing data into memory
  const emailMap = new Map();
  const lastRow = outputSheet.getLastRow();
  if (lastRow > 1) {
    const existingData = outputSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    existingData.forEach(row => {
      const name = row[0];
      const email = row[1].toString().toLowerCase().trim();
      const count = row[2];
      if (email) {
        emailMap.set(email, { name: name, count: count });
      }
    });
    log("Loaded " + emailMap.size + " existing entries", CONFIG.verbose);
  }

  // Load processed forms log
  const processedForms = new Set();
  const lastLogId = logSheet.getLastRow();
  if (lastLogId > 1) {
    const existingIds = logSheet.getRange(2, 1, lastLogId - 1, 1).getValues();
    existingIds.forEach(row => {
      processedForms.add(row[0].toString());
    });
    log("Found " + processedForms.size + " previously processed forms", CONFIG.verbose);
  }

  const newlyProcessedFormIds = new Set();

  // Find parent folder
  log("Searching for parent folder: " + CONFIG.parentFolderName, CONFIG.verbose);
  const parentFolders = DriveApp.getFoldersByName(CONFIG.parentFolderName);
  if (!parentFolders.hasNext()) {
    throw new Error("Parent folder not found: " + CONFIG.parentFolderName);
  }
  const parentFolder = parentFolders.next();
  log("Found parent folder", CONFIG.verbose);
  
  // Find subfolder if specified
  let targetFolder = parentFolder;
  if (CONFIG.subfolderName) {
    log("Searching for subfolder: " + CONFIG.subfolderName, CONFIG.verbose);
    const subfolders = parentFolder.getFoldersByName(CONFIG.subfolderName);
    if (!subfolders.hasNext()) {
      throw new Error("Subfolder not found: " + CONFIG.subfolderName);
    }
    targetFolder = subfolders.next();
    log("Found subfolder", CONFIG.verbose);
  }

  // Process all forms in folder tree
  const stats = {
    formsProcessed: 0,
    newForms: 0,
    skippedForms: 0,
    responsesProcessed: 0,
    newEmails: 0
  };
  
  processFolder(targetFolder, emailMap, processedForms, newlyProcessedFormIds, CONFIG, stats, "");
  
  // Write all data back to sheet
  const allRows = [];
  emailMap.forEach((data, email) => {
    allRows.push([data.name, email, data.count]);
  });
  
  const oldLastRow = outputSheet.getLastRow();
  
  if (allRows.length > 0) {
    outputSheet.getRange(2, 1, allRows.length, 3).setValues(allRows);
    const newLastRow = allRows.length + 1;
    if (oldLastRow > newLastRow) {
      outputSheet.getRange(newLastRow + 1, 1, oldLastRow - newLastRow, outputSheet.getLastColumn()).clearContent();
    }
    log("Wrote " + allRows.length + " records to output sheet", CONFIG.verbose);
  } else {
    if (oldLastRow > 1) {
      outputSheet.getRange(2, 1, oldLastRow - 1, outputSheet.getLastColumn()).clearContent();
    }
  }
  
  // Log newly processed form IDs
  if (newlyProcessedFormIds.size > 0) {
    const newIdsArray = Array.from(newlyProcessedFormIds).map(id => [id]);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newIdsArray.length, 1).setValues(newIdsArray);
    log("Logged " + newIdsArray.length + " new processed forms", CONFIG.verbose);
  }
  
  // Return summary
  const summary = {
    totalEmails: emailMap.size,
    formsProcessed: stats.formsProcessed,
    newFormsProcessed: stats.newForms,
    skippedForms: stats.skippedForms,
    responsesProcessed: stats.responsesProcessed,
    newEmailsAdded: stats.newEmails,
    spreadsheetUrl: outputSS.getUrl()
  };
  
  log("\n=== PROCESSING COMPLETE ===", CONFIG.verbose);
  log("Total emails tracked: " + summary.totalEmails, CONFIG.verbose);
  log("Forms processed: " + summary.formsProcessed, CONFIG.verbose);
  log("New forms: " + summary.newFormsProcessed, CONFIG.verbose);
  log("Skipped forms: " + summary.skippedForms, CONFIG.verbose);
  log("Responses processed: " + summary.responsesProcessed, CONFIG.verbose);
  log("New emails added: " + summary.newEmailsAdded, CONFIG.verbose);
  log("Spreadsheet: " + summary.spreadsheetUrl, CONFIG.verbose);
  
  return summary;
}

/**
 * Recursively processes a folder and all its subfolders to find and process Google Forms.
 *
 * @private
 * @param {GoogleAppsScript.Drive.Folder} folder The folder to process
 * @param {Map} emailMap Map of email addresses to attendance data objects
 * @param {Set} processedForms Set of form IDs already processed in previous runs
 * @param {Set} newlyProcessedFormIds Set to collect form IDs processed in current run
 * @param {Object} config Configuration object with settings
 * @param {Object} stats Statistics object tracking processing metrics
 * @param {string} indent Indentation string for log output hierarchy
 */
/**
 * Recursively processes a folder and all its subfolders to find and process Google Forms.
 *
 * @private
 * @param {GoogleAppsScript.Drive.Folder} folder The folder to process
 * @param {Map} emailMap Map of email addresses to attendance data objects
 * @param {Set} processedForms Set of form IDs already processed in previous runs
 * @param {Set} newlyProcessedFormIds Set to collect form IDs processed in current run
 * @param {Object} config Configuration object with settings
 * @param {Object} stats Statistics object tracking processing metrics
 * @param {string} indent Indentation string for log output hierarchy
 */
function processFolder(folder, emailMap, processedForms, newlyProcessedFormIds, config, stats, indent) {
  const folderName = folder.getName();
  log(indent + "Processing folder: " + folderName, config.verbose); // <-- CORRECTED LINE
  
  const forms = folder.getFilesByType(MimeType.GOOGLE_FORMS);
  let formCount = 0;
  
  while (forms.hasNext()) {
    const formFile = forms.next();
    formCount++;
    
    try {
      const formId = formFile.getId();
      const formName = formFile.getName();
      
      if (processedForms.has(formId)) {
        formCount--;
        continue;
      }
      
      if (config.skipRSVP && formName.toLowerCase().includes("rsvp")) {
        log(indent + "  Skipping RSVP form: \"" + formName + "\"", config.verbose);
        stats.skippedForms++;
        formCount--;
        continue;
      }
      
      const form = FormApp.openById(formId);
      const formResponses = form.getResponses();
      const responseCount = formResponses.length;
      
      log(indent + "  Processing form: \"" + formName + "\" (" + responseCount + " responses)", config.verbose);
      stats.formsProcessed++;
      stats.newForms++;
      
      if (responseCount === 0) {
        log(indent + "    No responses found", config.verbose);
        newlyProcessedFormIds.add(formId);
        continue;
      }

      let newEmailsInForm = 0;
      formResponses.forEach(response => {
        stats.responsesProcessed++;
        const itemResponses = response.getItemResponses();
        let fullName = "", firstName = "", lastName = "";
        let email = response.getRespondentEmail() || "";
        
        itemResponses.forEach(itemResponse => {
          const title = itemResponse.getItem().getTitle().toLowerCase();
          const answer = itemResponse.getResponse();
          if (!answer) return;
          const answerString = answer.toString().trim();
          
          if (title.includes("email")) email = answerString || email;
          else if (title.includes("first") && title.includes("name")) firstName = answerString;
          else if (title.includes("last") && title.includes("name")) lastName = answerString;
          else if (title.includes("name") && !title.includes("email")) fullName = answerString;
        });
        
        let finalName = firstName && lastName ? firstName + " " + lastName : fullName || firstName || lastName;
        if (!email) return;
        email = email.toLowerCase().trim();
        
        if (emailMap.has(email)) {
          const existing = emailMap.get(email);
          existing.count++;
          if (finalName.length > existing.name.length) existing.name = finalName;
          emailMap.set(email, existing);
        } else {
          emailMap.set(email, { name: finalName, count: 1 });
          newEmailsInForm++;
        }
      });
      
      newlyProcessedFormIds.add(formId);
      stats.newEmailsAdded += newEmailsInForm;
      
      log(indent + (newEmailsInForm > 0 
          ? "    Added " + newEmailsInForm + " new emails (Total: " + emailMap.size + ")"
          : "    Updated existing entries"), config.verbose);
      
    } catch (e) {
      log(indent + "  Error processing \"" + formFile.getName() + "\": " + e.message, config.verbose);
    }
  }
  
  if (formCount === 0) log(indent + "  No new forms in this folder", config.verbose);
  
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    processFolder(subfolders.next(), emailMap, processedForms, newlyProcessedFormIds, config, stats, indent + "  ");
  }
}
