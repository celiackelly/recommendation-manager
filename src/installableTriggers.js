/**
 * @OnlyCurrentDoc
 */
//Run the function createTriggers() ONCE when setting up the project, to set up the 'on edit' and 'on form submit' triggers for the functions below

function createTriggers() {
  ScriptApp.newTrigger("createNewSheetsOnSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("addUuidAndEmailCheckbox")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("markCompletion").forSpreadsheet(ss).onEdit().create();
}

function sortSheetsAlphabetically() {
  const sheetsArray = ss.getSheets(); //get array of all sheets (tabs) in the spreadsheet
  const sheetNames = sheetsArray.map((sheet) => sheet.getSheetName()); //map sheetsArray onto the name of each sheet (tab)

  // Sort sheets in alpha order; keep 'Form Responses 1, 'Due Dates', 'Address Book', and 'Template' first
  sheetNames.sort((a, b) => {
    const sortOrder = [
      "Form Responses 1",
      "Due Dates",
      "Address Book",
      "Template",
    ];
    const index1 = sortOrder.indexOf(a);
    const index2 = sortOrder.indexOf(b);

    // If tab name is in the sortOrder array, sort based on index in sortOrder
    if (index1 > -1 || index2 > -1) {
      return (
        (index1 > -1 ? index1 : Infinity) - (index2 > -1 ? index2 : Infinity)
      );
    }
    // Sort the rest of the tabs alphabetically
    return a.localeCompare(b);
  });

  for (let i = 0; i < sheetsArray.length; i++) {
    ss.setActiveSheet(ss.getSheetByName(sheetNames[i]));
    ss.moveActiveSheet(i + 1);
  }

  //Return to the 'Form Responses 1' sheet
  ss.setActiveSheet(formResponsesSheet);
}

function createNewSheetsOnSubmit(e) {
  // Create new teacher tab from 'Template' on form submit, if no tab exists for that teacher
  // This version is refactored to use the Sheets API and batch the updates to the spreadsheet, to minimize function calls to Google services and speed up the run time

  const sheetsArray = ss.getSheets(); //get array of all sheets (tabs) in the spreadsheet
  const sheetNames = sheetsArray.map((sheet) => sheet.getSheetName()); //map sheetsArray onto the name of each sheet (tab)

  //get the new row created by form submission
  const newRow = e.range.getRow();

  // get the cell values for the teacher submissions from 'Form Responses 1' sheet
  const teacherCellValues = formResponsesSheet
    .getRange(
      `${formResponses.columnLetters.mathTeacher}${newRow}:${formResponses.columnLetters.principalRec}${newRow}`
    )
    .getValues()[0]; //[mathTeacherCell, laTeacherCell, principalRecCell]

  //map teacherCellValues onto teacher names => 'jbroccoli'
  const teacherNames = teacherCellValues.map((value, i) => {
    const regex = /[A-Za-z.]*(?=@)/; //matches all alphanumeric characters that precede '@'
    // get teacher initial + last name from email in cell value; example: 'JoMarie Broccoli (jbroccoli@nysmith.com)' => 'jbroccoli'
    return value.match(regex) ? value.match(regex)[0] : null;
  });

  let requests = [];

  //create a new sheet (tab) from the template for any teacher in the form submission (teacherCells) that does not already have a tab
  //queue requests to update sheet protections and add query formulas to each sheet
  teacherNames.forEach((name, i) => {
    //if there's no tab for that teacher and a teacher's name was submitted (instead of e.g. 'No Math Teacher Recommendation Required', in which case teacher name value === null)
    if (name !== null && !sheetNames.includes(name)) {
      //insert new sheet named with teacher's name, as the last sheet in the doc, based on 'Template' sheet.
      const sheetId = ss
        .insertSheet(name, sheetsArray.length, { template: templateSheet })
        .getSheetId();

      let sheetProtection = {
        range: {
          sheetId: sheetId,
        },
        description: "Except for F:G, only Celia and Brian can edit the sheet",
        warningOnly: false,
        unprotectedRanges: [
          {
            sheetId: sheetId,
            startColumnIndex: 5, //column F
            endColumnIndex: 7, //column G, exclusive
          },
        ],
        editors: {
          users: ["ckelly@nysmithschool.com", "bschrembs@nysmithschool.com"],
          domainUsersCanEdit: false,
        },
      };

      let columnFGProtection = {
        range: {
          sheetId: sheetId,
          startColumnIndex: 5, //column F
          endColumnIndex: 7, //column G, exclusive
        },
        description: `Only ${name}, Celia, and Brian can edit F:G`,
        warningOnly: false,
        editors: {
          users: [
            "ckelly@nysmithschool.com",
            "bschrembs@nysmithschool.com",
            `${name}@nysmithschool.com`,
          ],
          domainUsersCanEdit: false,
        },
      };

      let queryFormula = {
        rows: [
          {
            values: [
              {
                userEnteredValue: {
                  formulaValue: `=iferror(query('Form Responses 1'!A2:I, "select A, B, C, G, I where D='${teacherCellValues[i]}' or E='${teacherCellValues[i]}' or F='${teacherCellValues[i]}'"), "")`,
                },
              },
            ],
          },
        ],
        fields: "userEnteredValue",
        range: {
          sheetId: sheetId,
          startRowIndex: 1,
          endRowIndex: 2,
          startColumnIndex: 0, //range A2
          endColumnIndex: 1,
        },
      };

      requests.push(
        { addProtectedRange: { protectedRange: sheetProtection } },
        { addProtectedRange: { protectedRange: columnFGProtection } },
        { updateCells: queryFormula }
      );
    }
  });

  //send all updates to Sheets API
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);

  // sort the teacher sheets in alpha order, keeping the set-up sheets like 'Form Responses 1' at the front
  sortSheetsAlphabetically();
}

//this could be optimized with Sheets API
function addUuidAndEmailCheckbox(e) {
  //For each form submission, add universal unique id, a checkbox for queuing emails, and query formula to the response row in 'Form Responses 1' sheet
  const newRow = e.range.getRow();

  //Generate a universal unique id (Uuid) for the response and add to 'Form Responses 1'
  formResponsesSheet
    .getRange(newRow, formResponses.columnNumbers.uuId)
    .setValue(Utilities.getUuid());

  //add checkbox for queuing emails
  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .setHelpText("Please click the cell to check or uncheck the box.")
    .build();
  formResponsesSheet
    .getRange(newRow, formResponses.columnNumbers.queueEmails)
    .setDataValidation(rule);

  //add query formula for parent contact emails to 'Form Responses 1'
  const studentName = formResponsesSheet
    .getRange(newRow, formResponses.columnNumbers.studentName)
    .getValue();
  formResponsesSheet
    .getRange(newRow, formResponses.columnNumbers.primaryContactEmail)
    .setFormula(
      `=iferror(query('Address Book'!A2:E, "SELECT C, E WHERE A='${studentName}' AND A IS NOT NULL", 0), "")`
    );

  //get the cell ranges for each teacher submission
  const mathTeacherCell = formResponsesSheet.getRange(
    newRow,
    formResponses.columnNumbers.mathTeacher
  );
  const laTeacherCell = formResponsesSheet.getRange(
    newRow,
    formResponses.columnNumbers.laTeacher
  );
  const principalRecCell = formResponsesSheet.getRange(
    newRow,
    formResponses.columnNumbers.principalRec
  );
  //array of submitted teacher cells (ranges, not values)
  const teacherCells = [
    mathTeacherCell,
    laTeacherCell, 
    principalRecCell
  ];

  //if teacher cell value is empty (meaning it's for a public school and the form bypassed the Choose Recommenders section), replace with "No Recommendation Required- Public School"
  teacherCells.forEach((cell) => {
    if (cell.getValue() === "") {
      cell.setValue("No Recommendation Required- Public School");
    }
  });
}

function markCompletion(e) {
  // When a recommendation is checked off as complete on any teacher's sheet, find the corresponding entry in 'Form Responses 1' and make that cell green.
  // Reset the cell background if a recommendation is unchecked
  // This function runs EVERY TIME the spreadsheet is edited by a user, but checks where the edit was made before doing anything

  const currentSheet = ss.getActiveSheet(); //the sheet currently being edited
  const range = e.range; //the cell that was edited (in this case, the date completed cell)

  //if the edit is not in column F or the edit is in the 'Form Responses 1' sheet, end the function
  //what about the other set-up tabs? FIX- this runs when I edit "Due Dates," for example
  if (
    range.getColumn() !== 6 ||
    currentSheet.getName() === "Form Responses 1"
  ) {
    Logger.log("outside range");
    return;
  }

  //set background of this recommendation line on the teacher's sheet to green
  const recommendationEntry = currentSheet.getRange(range.getRow(), 1, 1, 7);
  recommendationEntry.setBackground("#d9ead3");

  //get the universal unique id (Uuid) of the response (from column E, next to the date completed)
  const checkedUuid = currentSheet.getRange(e.range.getRow(), 5).getValue();

  //get name of sheet (tab) being edited - which teacher?
  const sheetName = currentSheet.getName();

  //get array of response Uuids in 'Form Responses 1' sheet
  const lastRow = formResponsesSheet.getLastRow(); //get the number of the last row with content
  const responseUuidArray = formResponsesSheet
    .getRange(1, formResponses.columnNumbers.uuId, lastRow)
    .getValues();

  //find the index of the response Uuid that matches the Uuid of the response checked off on the edited tab; add 1 to get the row number of that response (arrays are zero-indexed; ranges are not)
  const responseRow =
    responseUuidArray.findIndex((id) => id[0] === checkedUuid) + 1;

  //get math teacher, la teacher, and supplemental teacher cells from response row
  const teacherNameCells = formResponsesSheet.getRange(
    responseRow,
    formResponses.columnNumbers.mathTeacher,
    1,
    3
  );

  //get an array of the values (teacher names) from columns D:F of the response row
  const teacherNameCellValues = teacherNameCells.getValues()[0]; //ex: [JoMarie Broccoli (jbroccoli@nysmith.com), Emily Stephens (estephens@nysmith.com), No Supplemental Recommendation Required]

  //in teacherNameCellValues array, find the index of the teacher name that matches the edited tab; add 4 to get the correct column in 'Form Responses 1'
  //get this to throw an error and alert spreadsheet admins if not found? e.g. in case someone accidentally edited the tab names, which would break this function?
  const responseColumn =
    teacherNameCellValues.findIndex((entry) => entry.includes(sheetName)) + 4;

  //get the cell in 'Form Responses 1' that corresponds to the completed recommendation
  const cellToFormat = formResponsesSheet.getRange(responseRow, responseColumn);

  //if the date completed cell on the teacher sheet is filled out, change the background of the corresponding cell in 'Form Responses 1' to light green
  if (range.getValue()) {
    cellToFormat.setBackground("#d9ead3");
  } else {
    //if date completed is deleted, reset background on 'Form Responses 1' page and teacher sheet
    cellToFormat.setBackground(null);
    recommendationEntry.setBackground(null);
  }
}
