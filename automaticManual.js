//run inside google script apps as a .gs file
function importFormsToSpreadsheet() {
  const PARENT_FOLDER_NAME = "";
  const SUBFOLDER_NAME = "";
  const MASTER_HEADERS = ["Name", "Email", "Count"]; //can be anything 
  const LOG_SHEET_NAME = "Processed"; //hidden feature to not repeat 
  
  console.log("its gerbing time");
  
  let outputSS;
  const existingFiles = DriveApp.getFilesByName("Master Attendance");
  if (existingFiles.hasNext()) {
    outputSS = SpreadsheetApp.open(existingFiles.next());
    console.log("found old");
  }

  let outputSheet = outputSS.getSheetByName("TestFinal"); //can rename to your own 
  if (!outputSheet) {
    outputSheet = outputSS.insertSheet("TestFinal");
    outputSheet.appendRow(MASTER_HEADERS);
  } else {
    console.log("exist");
  }
  
  let logSheet = outputSS.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = outputSS.insertSheet(LOG_SHEET_NAME);
    console.log(`made sumn '${LOG_SHEET_NAME}' plans`);
    logSheet.appendRow(["Processed Form IDs"]);
    logSheet.hideSheet(); //hide
  }

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
  }

  const processedForms = new Set();
  const lastLogId = logSheet.getLastRow();
  if (lastLogId > 1) {
    const existingIds = logSheet.getRange(2, 1, lastLogId - 1, 1).getValues();
    existingIds.forEach(row => {
      processedForms.add(row[0].toString());
    });
  }

  const newlyProcessedFormIds = new Set();

  console.log(`trying to run up'${PARENT_FOLDER_NAME}'`);
  const parentFolders = DriveApp.getFoldersByName(PARENT_FOLDER_NAME);
  if (!parentFolders.hasNext()) {
    console.log(`parent not found`);
    return;
  }
  const parentFolder = parentFolders.next();
  
  const subfolders = parentFolder.getFoldersByName(SUBFOLDER_NAME);
  if (!subfolders.hasNext()) {
    console.log(`subfolder not found`);
    return;
  }
  const eventsFolder = subfolders.next();
  console.log("checking events");

  processFolder(eventsFolder, emailMap, processedForms, newlyProcessedFormIds, "");
  const allRows = [];
  emailMap.forEach((data, email) => {
    allRows.push([data.name, email, data.count]);
  });
  
  const oldLastRow = outputSheet.getLastRow();
  
  if (allRows.length > 0) {
    outputSheet.getRange(2, 1, allRows.length, 3)
               .setValues(allRows);
    
    const newLastRow = allRows.length + 1;
    if (oldLastRow > newLastRow) {
      console.log('cleaning upp sumn nows');
      outputSheet.getRange(newLastRow + 1, 1, oldLastRow - newLastRow, outputSheet.getLastColumn()).clearContent();
    }
  } else {
    console.log("nun new");
     if (oldLastRow > 1) {
      console.log(`clear test ${oldLastRow}`);
      outputSheet.getRange(2, 1, oldLastRow - 1, outputSheet.getLastColumn()).clearContent();
    }
  }
  
  if (newlyProcessedFormIds.size > 0) {
    const newIdsArray = Array.from(newlyProcessedFormIds).map(id => [id]);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newIdsArray.length, 1)
           .setValues(newIdsArray);
    console.log(`done with sumn new `);
  } else {
    console.log("there was nothing");
  }
}

function processFolder(folder, emailMap, processedForms, newlyProcessedFormIds, indent) {
  const folderName = folder.getName();
  console.log(`im running up ${folderName}`);
  
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
      
      if (formName.toLowerCase().includes("rsvp")) {
        console.log(`fade rsvp"${formName}"`);
        formCount--;
        continue;
      }
      
      const form = FormApp.openById(formId);
      const formResponses = form.getResponses();
      const responseCount = formResponses.length;
      
      console.log(`new form "${formName}"`);
      
      if (responseCount === 0) {
        console.log(`nah theres nothing foo`);
        newlyProcessedFormIds.add(formId);
        continue;
      }

      let newEmails = 0;
      formResponses.forEach(response => {
        const itemResponses = response.getItemResponses();
        let fullName = "", firstName = "", lastName = "";
        let email = response.getRespondentEmail() || "";
        
        itemResponses.forEach(itemResponse => {
          const title = itemResponse.getItem().getTitle().toLowerCase();
          const answer = itemResponse.getResponse();
          if (!answer) return;
          const answerString = answer.toString().trim();
          
          if (title.includes("email")) {
            email = answerString || email;
          } else if (title.includes("first") && title.includes("name")) {
            firstName = answerString;
          } else if (title.includes("last") && title.includes("name")) {
            lastName = answerString;
          } else if (title.includes("name") && !title.includes("email")) {
            fullName = answerString;
          }
        });
        
        let finalName = "";
        if (firstName && lastName) finalName = `${firstName} ${lastName}`;
        else if (fullName) finalName = fullName;
        else finalName = firstName || lastName;
        
        if (!email) return;
        email = email.toLowerCase().trim();
        
        if (emailMap.has(email)) {
          const existing = emailMap.get(email);
          existing.count++;
          if (finalName.length > existing.name.length) {
            existing.name = finalName;
          }
          emailMap.set(email, existing);
        } else {
          emailMap.set(email, { name: finalName, count: 1 });
          newEmails++;
        }
      });
      
      newlyProcessedFormIds.add(formId);
      
      if(newEmails > 0) {
        console.log(`Added new: ${emailMap.size}`);
      } else {
        console.log(`already there`);
      }
      
    } catch (e) {
      console.log(`u dumb as HELL boy"${formFile.getName()}": ${e.message}`);
    }
  }
  
  if (formCount === 0) {
    console.log(`no new forms found in this folder`);
  }
  
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    processFolder(subfolders.next(), emailMap, processedForms, newlyProcessedFormIds, indent + "  ");
  }

  
}
