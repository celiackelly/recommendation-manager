/**
 * @OnlyCurrentDoc
 */

//Run the function createTriggers() in the InstallableTriggers.gs file ONCE when setting up the project, to set up the 'on edit' and 'on form submit' triggers for the functions below

function sortSheetsAlphabetically () {
  // Sort sheets in alpha order; keep 'Form Responses' first and 'Template' last

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //get array of all sheets (tabs) in the spreadsheet
  const sheetsArray = ss.getSheets()

  //map sheetsArray onto the name of each sheet (tab)
  const sheetNames = sheetsArray.map(sheet => sheet.getSheetName())

  sheetNames.sort((a, b) => {
    //Keep 'Form Responses' first and 'Template' last
    if (a === 'Form Responses' || b === 'Form Responses') {
      return a === 'Form Responses' ? -1 : 1
    } else if (a === 'Template' || b === 'Template') {
      return a === 'Template' ? 1 : -1
    }
    //Sort the rest alphabetically
    return a.localeCompare(b)
  });
    
  for (let i = 0; i < sheetsArray.length; i++) {
    ss.setActiveSheet(ss.getSheetByName(sheetNames[i]));
    ss.moveActiveSheet(i + 1);
  }

  //Return to the 'Form Responses' sheet
  ss.setActiveSheet(ss.getSheetByName('Form Responses'))
}


function createNewSheetOnSubmit(e) {
  // Create new teacher tab from 'Template' on form submit, if no tab exists for that teacher

  const ss = e.source;  //A Spreadsheet object, representing the Google Sheets file to which the script is bound.
  const formResponsesSheet = ss.getSheetByName('Form Responses');

  //get array of all sheets (tabs) in the spreadsheet
  const sheetsArray = ss.getSheets()

  //map sheetsArray onto the name of each sheet (tab)
  const sheetNames = sheetsArray.map(sheet => sheet.getSheetName())

  //get the new row created by form submission 
  const newRow = e.range.getRow();

  //get the cell ranges for each teacher submission 
  const mathTeacherCell = formResponsesSheet.getRange(e.range.getRow(), 4)
  const laTeacherCell = formResponsesSheet.getRange(e.range.getRow(), 5)
  const otherTeacherCell = formResponsesSheet.getRange(e.range.getRow(), 6)
  
  const regex = /[A-Za-z.]*(?=@)/   //matches all alphanumeric characters that precede '@'
  //function to get teacher initial + last name from email in cell value; example: 'JoMarie Broccoli (jbroccoli@nysmith.com)' => 'jbroccoli'
  const getTeacherName = teacherCell => teacherCell.getValue().match(regex) ? teacherCell.getValue().match(regex)[0] : null     //String.match(regex) returns array of matches or null if no matches

  //array of submitted teacher cells (ranges, not values)
  const teacherCells = [mathTeacherCell, laTeacherCell, otherTeacherCell]

  //creates a new sheet (tab) from the template for any teacher in the form submission (teacherCells) that does not already have a tab
  teacherCells.forEach(cell => {
    
    //if there's no tab for that teacher and a teacher's name was submitted (instead of e.g. 'No Math Teacher Recommendation Required', in which case getTeacherName(cell) === null)
    if (!sheetNames.includes(getTeacherName(cell)) && getTeacherName(cell)) {
      const templateSheet = ss.getSheetByName('Template')

      //insert new sheet named with teacher's name, at end of sheet, based on 'Template' sheet. 
      ss.insertSheet(getTeacherName(cell), sheetsArray.length, {template: templateSheet})

      const sheet = ss.getActiveSheet()

      //add the query formula to cell A2, with parameters set to that teacher
      sheet.getRange('A2').setFormula(`=query('Form Responses'!A2:K, "select A, B, C, G, H, J where D='${cell.getValue()}' or E='${cell.getValue()}' or F='${cell.getValue()}'")`)
    
      // Protect the active sheet except for columns G and H.
      const protection = sheet.protect().setDescription('Teachers: Edit only G:H');
      const unprotected = sheet.getRange('G:H');
      protection.setUnprotectedRanges([unprotected]);

      // Ensure Celia and Brian are editors before removing others in the Nysmith domain from editing the protected ranges. 
      // Otherwise, if the user's edit permission comes from a group, the script throws an exception upon removing the group.
      // For this to work, make sure that permissions in the spreadsheet GUI are set to 'Anyone at Nysmith can edit'; 
      //NO ONE besides Celia and Brian should be given explicit editing rights to the spreadsheet as a whole
      protection.addEditors(['ckelly@nysmithschool.com', 'bschrembs@nysmithschool.com']);
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }
  })

  // sort the sheets in alpha order, except for 'Form Responses' and 'Template'
  sortSheetsAlphabetically()
}


function addUuidAndCheckbox(e) {
  //For each form submission, add universal unique id to column J (10) of 'Form Responses' sheet, and add checkbox to column K (11)

  const row = e.range.getRow();

  //Generate a universal unique id (Uuid) for the response in column J (10) of 'Form Responses'
  const responseSheet = e.range.getSheet();
  responseSheet.getRange(row, 10).setValue(Utilities.getUuid())

  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).setHelpText('Please click the cell to check or uncheck the box.').build();
  responseSheet.getRange(row, 11).setDataValidation(rule);
}


function markCompletion(e) {
  // When a recommendation is checked off as complete on any teacher's sheet, find the corresponding entry in 'Form Responses' and make that cell green. 
  // Reset the cell background if a recommendation is unchecked
  // This function runs EVERY TIME the spreadsheet is edited by a user, but checks where the edit was made before doing anything 

  const ss = e.source;  //A Spreadsheet object, representing the Google Sheets file to which the script is bound.
  const sheet = ss.getActiveSheet();  //the sheet currently being edited
  const range = e.range   //the cell that was edited (in this case, the date completed cell)

  //if the edit is not in column G or the edit is in the 'Form Responses' sheet, end the function
  if (range.getColumn() !== 7 || sheet.getName() === 'Form Responses') { 
    Logger.log('outside range')
    return 
  }

  //set background of this recommendation line on the teacher's sheet to green
  const recommendationEntry = sheet.getRange(range.getRow(), 1, 1, 8)
  recommendationEntry.setBackground('#d9ead3') 

  //get the universal unique id (Uuid) of the response (from column F, next to the date completed)
  const checkedUuid = sheet.getRange(e.range.getRow() ,6).getValue();

  //get name of sheet (tab) being edited - which teacher? 
  const sheetName = sheet.getName()

  const formResponsesSheet = ss.getSheetByName('Form Responses')

  //get array of response Uuids in 'Form Responses' sheet
  const lastRow = formResponsesSheet.getLastRow();  //get the number of the last row with content
  const responseUuidArray = formResponsesSheet.getRange(1, 10, lastRow).getValues()

  //find the index of the response Uuid that matches the Uuid of the response checked off on the edited tab; add 1 to get the row number of that response (arrays are zero-indexed; ranges are not)
  const responseRow = responseUuidArray.findIndex(id => id[0] === checkedUuid) + 1

  //get columns D:F of the response row
  const teacherNameCells = formResponsesSheet.getRange(responseRow, 4, 1, 3)  
  
  //get an array of the values (teacher names) from columns D:F of the response row
  const teacherNameCellValues = teacherNameCells.getValues()[0] //ex: [JoMarie Broccoli (jbroccoli@nysmith.com), Emily Stephens (estephens@nysmith.com), No Supplemental Recommendation Required]

  //in teacherNameCellValues array, find the index of the teacher name that matches the edited tab; add 4 to get the correct column (D, E, or F) in 'Form Responses'
    //get this to throw an error and alert spreadsheet admins if not found? e.g. in case someone accidentally edited the tab names, which would break this function?
  const responseColumn = teacherNameCellValues.findIndex(entry => entry.includes(sheetName)) + 4

  //get the cell in 'Form Responses' that corresponds to the completed recommendation 
  const cellToFormat = formResponsesSheet.getRange(responseRow, responseColumn)  

  //if the date completed cell on the teacher sheet is filled out, change the background of the corresponding cell in 'Form Responses' to light green 
  if (range.getValue()) {
    cellToFormat.setBackground('#d9ead3')

  } else {  //if date completed is deleted, reset background on 'Form Responses' page and teacher sheet
    cellToFormat.setBackground(null)
    recommendationEntry.setBackground(null) 
  }
}

function queueCompletionEmail(e) {
  // When Brian checks a box in the 'Ready to send email'? column (K) in 'Form Responses,' save information about that recommendation entry, to queue for the sendQueuedCompletionEmails() function
  // This function runs EVERY TIME the spreadsheet is edited by a user, but checks where the edit was made before doing anything 
  // What if you check a bunch of boxes very quickly, before the previous function call has finished? Need to test and account for this

  const ss = e.source;  //A Spreadsheet object, representing the Google Sheets file to which the script is bound.
  const sheet = ss.getActiveSheet();  //the sheet currently being edited
  const range = e.range   //the cell that was edited  

  //if the edit is not in the 'Form Responses' sheet and not in column L, end the function
  if (sheet.getName() !== 'Form Responses' && range.getColumn() !== 11) { 
    Logger.log('outside range')
    return 
  } 

  // DocumentProperties in the PropertiesService allows you to save string data related to the spreadsheet. We can then access this saved data when running sendQueuedCompletionEmails() 
  // We need this workaround because Google Apps Script does not support global variables.
  // https://developers.google.com/apps-script/guides/properties 

  //Check PropertiesService for the 'queuedEmailInfo' property. Set queuedEmailInfo equal to this property if it exists, or an empty array otherwise
  let queuedEmailInfo = JSON.parse(PropertiesService.getDocumentProperties().getProperty('queuedEmailInfo')) || []

  //the row number of the checkbox
  const responseRow = e.range.getRow()

  //if the box was UNCHECKED, delete that email info from the 'queuedEmailInfo' property (i.e. remove that email from the queue)
  if (range.getValue() === false) { 
    Logger.log('"Ready to send" box unchecked')
    let index = queuedEmailInfo.findIndex(emailInfo => emailInfo.responseRow === responseRow) 
    Logger.log(index)
    queuedEmailInfo.splice(index, 1)

    let queuedEmailInfoJSON = JSON.stringify(queuedEmailInfo)
    PropertiesService.getDocumentProperties().setProperty('queuedEmailInfo', queuedEmailInfoJSON) 
    return 

  //else, the box was CHECKED. Capture the info for that email and reassign the 'queuedEmailInfo' property to include that email (i.e. queue up the email)
  } else {
    Logger.log('"Ready to send" box checked')
  
        const addressBook = ss.getSheetByName('Address Book')
    const studentName = sheet.getRange(responseRow, 2).getValue()

    const addressBookNames = addressBook.getRange(1, 1, 1, addressBook.getLastRow()).getValues()
    const studentRowInAddressBook = addressBookNames.findIndex(name => name === studentName) + 3
  
    const emailInfo = {
      responseRow: responseRow,
      parentEmails: `${addressBook.getRange(studentRowInAddressBook, 2).getValue()}, ${addressBook.getRange(studentRowInAddressBook, 3).getValue()}`,
      studentName: studentName,
      school: sheet.getRange(responseRow, 3).getValue()
    }
    queuedEmailInfo.push(emailInfo)
  }

  let queuedEmailInfoJSON = JSON.stringify(queuedEmailInfo)

  PropertiesService.getDocumentProperties().setProperty('queuedEmailInfo', queuedEmailInfoJSON) 
  Logger.log(PropertiesService.getDocumentProperties().getProperty('queuedEmailInfo'))
}


function sendQueuedCompletionEmails() {
//This might fail if you click the 'Send emails' button too quickly after checking the boxes; if queueCompletionEmail() hasn't finished, the queue will still be empty. Any way to fix this? 

  // get the queued email info from the PropertiesService; parse the JSON data back into an array of objects
  let queuedEmailInfo = JSON.parse(PropertiesService.getDocumentProperties().getProperty('queuedEmailInfo'))

  Logger.log(queuedEmailInfo)

  //if there are no queued emails, end the function 
  if (!queuedEmailInfo) { return }

  //for each emailInfo object in the array/queue, send an email 
  queuedEmailInfo.forEach(emailInfo => {
      MailApp.sendEmail({
        to: emailInfo.parentEmails,
        replyTo: 'bschrembs@nysmith.com',   //Emails will be sent from Celia's account, but if parents reply, the replies will default to Brian
        subject: `Completed Recommendation for ${emailInfo.studentName}`,
        body: `All the recommendations for ${emailInfo.studentName} for ${emailInfo.school} have been completed. Please contact Brian Schrembs with any questions.`
    })

  })

  resetQueuedEmailInfo()
}

function resetQueuedEmailInfo() {

  //delete any queued email info
  PropertiesService.getDocumentProperties().deleteProperty('queuedEmailInfo') 
  Logger.log('Queued emails reset; list is empty')
}


//PROBLEM- I can't restrict access to this sidebar, can I? So anyone could send emails from here if they have domain edit access on the spreadsheet as a whole?
//Maybe just display the queued emails in the sidebar, but still trigger them from the button on the speadsheet (so only Brian and I can trigger)
// function onOpen() {
//   SpreadsheetApp
//     .getUi()
//     .createMenu('Admin Controls')
//     .addItem('Admin Controls', 'showAdminSidebar')
//     .addToUi();
// }

// function showAdminSidebar() {
//   const sidebar = HtmlService.createHtmlOutputFromFile('sidebar.html');
//   sidebar.setTitle('Admin Controls')
//   SpreadsheetApp.getUi().showSidebar(sidebar);
// }



//Test of serving templated HTML
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Admin Controls')
    .addItem('Admin Controls', 'showAdminSidebar')
    .addToUi();
}

function doGet() {
  const sidebar = HtmlService.createTemplateFromFile('sidebar.html')

  return sidebar.evaluate();
}

function showAdminSidebar() {
  const sidebar = doGet();
  sidebar.setTitle('Admin Controls')
  SpreadsheetApp.getUi().showSidebar(sidebar);
}

