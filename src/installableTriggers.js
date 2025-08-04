/**
 * @OnlyCurrentDoc
 */
//Run the function createTriggers() ONCE when setting up the project, to set up the 'on edit' and 'on form submit' triggers for the functions below

function createTriggers() {
  ScriptApp.newTrigger('createNewSheetsOnSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create()

  ScriptApp.newTrigger('formatResponseRow')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create()

  ScriptApp.newTrigger('markCompletion').forSpreadsheet(ss).onEdit().create()
}

function sortSheetsAlphabetically() {
  const sheetsArray = ss.getSheets() //get array of all sheets (tabs) in the spreadsheet
  const sheetNames = sheetsArray.map(sheet => sheet.getSheetName()) //map sheetsArray onto the name of each sheet (tab)

  // Sort sheets in alpha order; keep 'Form Responses 1, 'Due Dates', 'Address Book', and 'Template' first
  sheetNames.sort((a, b) => {
    const sortOrder = [
      'Form Responses 1',
      'Due Dates',
      'Summary',
      'By School',
      'Address Book',
      'Deleted Requests',
      'Template',
    ]
    const index1 = sortOrder.indexOf(a)
    const index2 = sortOrder.indexOf(b)

    // If tab name is in the sortOrder array, sort based on index in sortOrder
    if (index1 > -1 || index2 > -1) {
      return (
        (index1 > -1 ? index1 : Infinity) - (index2 > -1 ? index2 : Infinity)
      )
    }
    // Sort the rest of the tabs alphabetically
    return a.localeCompare(b)
  })

  for (let i = 0; i < sheetsArray.length; i++) {
    ss.setActiveSheet(ss.getSheetByName(sheetNames[i]))
    ss.moveActiveSheet(i + 1)
  }

  //Return to the 'Form Responses 1' sheet
  ss.setActiveSheet(formResponsesSheet)
}

function createNewSheetsOnSubmit(e) {
  // Create new teacher tab from 'Template' on form submit, if no tab exists for that teacher
  // This version is refactored to use the Sheets API and batch the updates to the spreadsheet, to minimize function calls to Google services and speed up the run time
  // Updated in Aug 2025 to include supplemental teacher rec

  const sheetsArray = ss.getSheets() //get array of all sheets (tabs) in the spreadsheet
  const sheetNames = sheetsArray.map(sheet => sheet.getSheetName()) //map sheetsArray onto the name of each sheet (tab)

  //get the new row created by form submission
  const newRow = e.range.getRow()

  // get the cell values for the teacher submissions from 'Form Responses 1' sheet
  const teacherCellValues = formResponsesSheet
    .getRange(
      `${formResponses.columnLetters.mathTeacher}${newRow}:${formResponses.columnLetters.principalRec}${newRow}`,
    )
    .getValues()[0]
    .filter(el => el) //[mathTeacherCell, laTeacherCell, principalRecCell, supplementalTeacherCell] - filter out any empty cells

  //map teacherCellValues onto teacher names => 'jbroccoli'
  const teacherNames = teacherCellValues.map((value, i) => {
    const regex = /[A-Za-z.]*(?=@)/ //matches all alphanumeric characters that precede '@'
    // get teacher initial + last name from email in cell value; example: 'JoMarie Broccoli (jbroccoli@nysmith.com)' => 'jbroccoli'
    return value.match(regex) ? value.match(regex)[0] : null      // if no regex match, return null
  })

  let requests = []

  //create a new sheet (tab) from the template for any teacher in the form submission (teacherCells) that does not already have a tab
  //queue requests to update sheet protections and add query formulas to each sheet
  teacherNames.forEach((name, i) => {
    //if there's no tab for that teacher and a teacher's name was submitted (instead of e.g. 'No Math Teacher Recommendation Required', in which case teacher name value === null)
    if (name !== null && !sheetNames.includes(name)) {
      //insert new sheet named with teacher's name, as the last sheet in the doc, based on 'Template' sheet.
      const sheetId = ss
        .insertSheet(name, sheetsArray.length, { template: templateSheet })
        .getSheetId()

      let sheetProtection = {
        range: {
          sheetId: sheetId,
        },
        description: 'Except for F:G, only Celia and Brian can edit the sheet',
        warningOnly: false,
        unprotectedRanges: [
          {
            sheetId: sheetId,
            startColumnIndex: 5, //column F
            endColumnIndex: 7, //column G, exclusive
          },
        ],
        editors: {
          users: ['ckelly@nysmithschool.com', 'bschrembs@nysmithschool.com'],
          domainUsersCanEdit: false,
        },
      }

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
            'ckelly@nysmithschool.com',
            'bschrembs@nysmithschool.com',
            `${name}@nysmithschool.com`,
          ],
          domainUsersCanEdit: false,
        },
      }

      const selectStatement = `${formResponses.columnLetters.timeStamp}, ${formResponses.columnLetters.studentName}, ${formResponses.columnLetters.school}, ${formResponses.columnLetters.source}, ${formResponses.columnLetters.uuId}`
      const whereStatement = `${formResponses.columnLetters.mathTeacher}='${teacherCellValues[i]}' or ${formResponses.columnLetters.laTeacher}='${teacherCellValues[i]}' or ${formResponses.columnLetters.principalRec}='${teacherCellValues[i]}' or ${formResponses.columnLetters.supplementalTeacher}='${teacherCellValues[i]}'`

      let queryFormulaRequest = {
        rows: [
          {
            values: [
              {
                userEnteredValue: {
                  formulaValue: `=iferror(query('Form Responses 1'!${formResponses.columnLetters.timeStamp}2:${formResponses.columnLetters.uuId}, "select ${selectStatement} where ${whereStatement}", 0), "")`,
                },
              },
            ],
          },
        ],
        fields: 'userEnteredValue',
        range: {
          sheetId: sheetId,
          startRowIndex: 1,
          endRowIndex: 2,
          startColumnIndex: 0, //range A2
          endColumnIndex: 1,
        },
      }

      requests.push(
        { addProtectedRange: { protectedRange: sheetProtection } },
        { addProtectedRange: { protectedRange: columnFGProtection } },
        { updateCells: queryFormulaRequest },
      )
    }
  })

  //if there are requests (meaning new teacher sheets have been created), send all update requests to Sheets API
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId)

    // sort the teacher sheets in alpha order, keeping the set-up sheets like 'Form Responses 1' at the front
    sortSheetsAlphabetically()
  }
}

function createRemoveDataValidationRequest(sheetId, row) {
  //Create request to remove any existing data validation in the new row. This is necessary because the sheet automatically copies down the checkboxes from the previous row into the new response row.
  let removeDataValidationRequest = {
    range: {
      sheetId: sheetId,
      startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
      endRowIndex: row,
    },
  }
  return { setDataValidation: removeDataValidationRequest }
}

function createQueueDeletionCheckboxRequest(sheetId, row) {
  const checkboxRule = {
    condition: {
      type: 'BOOLEAN',
      values: [],
    },
    strict: true,
    showCustomUi: false,
  }

  let queueDeletionCheckboxRequest = {
    range: {
      sheetId: sheetId,
      startRowIndex: row - 1, //subtract one because this is an index, not a row/col in a range
      endRowIndex: row,
      startColumnIndex: formResponses.columnIndex.deleteRecord, 
      endColumnIndex: formResponses.columnIndex.deleteRecord + 1,
    },
    rule: checkboxRule,
  }

  return { setDataValidation: queueDeletionCheckboxRequest }
}

function createAddUuidRequest(sheetId, row) {
  const uuId = Utilities.getUuid()
  let addUuIdRequest = {
    rows: [
      {
        values: [
          {
            userEnteredValue: {
              stringValue: uuId,
            },
          },
        ],
      },
    ],
    fields: 'userEnteredValue',
    range: {
      sheetId: sheetId,
      startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
      endRowIndex: row,
      startColumnIndex: formResponses.columnIndex.uuId,
      endColumnIndex: formResponses.columnIndex.uuId + 1,
    },
  }
  return { updateCells: addUuIdRequest }
}

function createQueueEmailsCheckboxRequest(sheetId, row) {
  const checkboxRule = {
    condition: {
      type: 'BOOLEAN',
      values: [],
    },
    strict: true,
    showCustomUi: false,
  }

  let queueEmailsCheckboxRequest = {
    range: {
      sheetId: sheetId,
      startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
      endRowIndex: row,
      startColumnIndex: formResponses.columnIndex.queueEmails, 
      endColumnIndex: formResponses.columnIndex.queueEmails + 1,
    },
    rule: checkboxRule,
  }

  return { setDataValidation: queueEmailsCheckboxRequest }
}

function createAddParentEmailsQueryRequest(sheetId, row) {
  const studentName = formResponsesSheet
    .getRange(row, formResponses.columnNumbers.studentName)
    .getValue()

  let queryFormulaRequest = {
    rows: [
      {
        values: [
          {
            userEnteredValue: {
              formulaValue: `=iferror(query('Address Book'!A2:E, "SELECT C, E WHERE A='${studentName}' AND A IS NOT NULL", 0), "")`,
            },
          },
        ],
      },
    ],
    fields: 'userEnteredValue',
    range: {
      sheetId: sheetId,
      startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
      endRowIndex: row,
      startColumnIndex: formResponses.columnIndex.primaryContactEmail,
      endColumnIndex: formResponses.columnIndex.primaryContactEmail + 1,
    },
  }

  return { updateCells: queryFormulaRequest }
}

function createAddRecommendationCheckboxesRequests(
  sheetId,
  row,
  recommendationCellValues,
) {
  let addRecommendationCheckboxesRequests = []

  const isNotRequired = recommendationValue => {
    if (
      !recommendationValue ||
      recommendationValue === 'No' ||
      recommendationValue === 'No Math Teacher Recommendation Required' ||
      recommendationValue === 'No Language Arts Teacher Recommendation Required' ||
      recommendationValue === 'No Supplemental Recommendation Required'
    ) {
      return true
    } else {
      return false
    }
  }

  // for math teacher, la teacher, supplemental teacher, and principal rec completion columns, if rec is not required, set completion cell value to 'n/a
  for (let i = 0; i < recommendationCellValues.length; i += 2) {
    //iterate through teacher rec cells by skipping odd indices (the completion cells)
    const teacherRec = recommendationCellValues[i]

    if (isNotRequired(teacherRec)) {
      let request = {
        rows: [
          {
            values: [
              {
                userEnteredValue: {
                  stringValue: 'n/a',
                },
              },
            ],
          },
        ],
        fields: 'userEnteredValue',
        range: {
          sheetId: sheetId,
          startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
          endRowIndex: row,
          startColumnIndex: formResponses.columnIndex.mathTeacherCompletion + i,
          endColumnIndex: formResponses.columnIndex.mathTeacherCompletion + 1 + i,
        },
      }

      addRecommendationCheckboxesRequests.push({ updateCells: request })

      //otherwise add checkbox;
    } else {
      const checkboxRule = {
        condition: {
          type: 'BOOLEAN',
          values: [],
        },
        strict: true,
        showCustomUi: false,
      }

      let checkboxRequest = {
        range: {
          sheetId: sheetId,
          startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
          endRowIndex: row,
          startColumnIndex: formResponses.columnIndex.mathTeacherCompletion + i,
          endColumnIndex: formResponses.columnIndex.mathTeacherCompletion + 1 + i,
        },
        rule: checkboxRule,
      }

      addRecommendationCheckboxesRequests.push({
        setDataValidation: checkboxRequest,
      })
    }
  }

  return addRecommendationCheckboxesRequests
}

function createNoRecRequiredRequests(sheetId, row, recommendationCellValues) {
  //if teacher cell value is empty (meaning it's for a public school and the form bypassed the Choose Recommenders section), replace with "No Recommendation Required- Public School"

  let requests = []

  for (let i = 0; i <= recommendationCellValues.length; i += 2) {
    //iterate through teacher rec cells by skipping odd indices (the completion cells)
    const teacherRec = recommendationCellValues[i]
    if (teacherRec === '') {
      let request = {
        rows: [
          {
            values: [
              {
                userEnteredValue: {
                  stringValue: 'No Recommendation Required- Public School',
                },
              },
            ],
          },
        ],
        fields: 'userEnteredValue',
        range: {
          sheetId: sheetId,
          startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
          endRowIndex: row,
          startColumnIndex: formResponses.columnIndex.mathTeacher + i,
          endColumnIndex: formResponses.columnIndex.mathTeacher + 1 + i,
        },
      }

      requests.push({ updateCells: request })
    }
  }

  return requests
}

function createAddDuplicatesQueryRequest(sheetId, row) {
  //add helper formula to column P, so that conditional formatting can highlight duplicates in red

  let queryFormulaRequest = {
    rows: [
      {
        values: [
          {
            userEnteredValue: {
              formulaValue: `=${formResponses.columnLetters.studentName}${row}&"- " &${formResponses.columnLetters.school}${row}`,
            },
          },
        ],
      },
    ],
    fields: 'userEnteredValue',
    range: {
      sheetId: sheetId,
      startRowIndex: row - 1, //subtract one from all values, because this is an index, not a row/col in a range
      endRowIndex: row,
      startColumnIndex:
        formResponses.columnIndex.findDuplicatesHelperQuery,
      endColumnIndex: formResponses.columnIndex.findDuplicatesHelperQuery + 1,
    },
  }

  return { updateCells: queryFormulaRequest }
}

function formatResponseRow(e) {
  //For each form submission row, add universal unique id, a checkbox for queuing emails, query formula to the response row in 'Form Responses 1' sheet; format recommendation cells and add completion checkboxes

  const newRow = e.range.getRow()
  const formResponsesSheetId = formResponsesSheet.getSheetId()

  let requests = []

  //Remove any existing data validation in the new row. This is necessary because the sheet automatically copies down the checkboxes from the previous row into the new response row.
  const removeDataValidationRequest = createRemoveDataValidationRequest(
    formResponsesSheetId,
    newRow,
  )

    //create request to add checkbox for queuing emails
    const queueDeletionCheckboxRequest = createQueueDeletionCheckboxRequest(
      formResponsesSheetId,
      newRow,
    )

  //Generate a universal unique id (Uuid) for the response; create request to add to Form Responses 1 sheet
  const addUuidRequest = createAddUuidRequest(formResponsesSheetId, newRow)

  //create request to add checkbox for queuing emails
  const queueEmailsCheckboxRequest = createQueueEmailsCheckboxRequest(
    formResponsesSheetId,
    newRow,
  )

  //create request to add query formula for parent emails from 'Address Book' tab
  const addParentEmailsQueryRequest = createAddParentEmailsQueryRequest(
    formResponsesSheetId,
    newRow,
  )

  //get the range for each teacher submission and recommendation completion checkoff column
  const recommendationCellValues = formResponsesSheet
    .getRange(newRow, formResponses.columnNumbers.mathTeacher, 1, 8)
    .getValues()[0]

  //create requests: for math teacher, la teacher, supplemental teacher, and principal rec completion columns, if rec is not required, set completion cell value to 'n/a; otherwise, add a checkbox for completion
  const addRecommendationCheckboxesRequests =
    createAddRecommendationCheckboxesRequests(
      formResponsesSheetId,
      newRow,
      recommendationCellValues,
    )

  //create requests: if teacher cell value is empty (meaning it's for a public school and the form bypassed the Choose Recommenders section), replace with "No Recommendation Required- Public School"
  const noRecRequiredRequests = createNoRecRequiredRequests(
    formResponsesSheetId,
    newRow,
    recommendationCellValues,
  )

  // create request to add helper formula to column P, so that conditional formatting can highlight duplicates in red
  const addDuplicatesQueryRequest = createAddDuplicatesQueryRequest(
    formResponsesSheetId,
    newRow,
  )

  requests.push(
    removeDataValidationRequest,
    addUuidRequest,
    queueDeletionCheckboxRequest,
    queueEmailsCheckboxRequest,
    addParentEmailsQueryRequest,
    ...addRecommendationCheckboxesRequests,
    ...noRecRequiredRequests,
    addDuplicatesQueryRequest,
  )

  Logger.log(requests)

  //send all updates to Sheets API
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId)
}

function markCompletion(e) {
  // When a recommendation is checked off as complete on any teacher's sheet, find the corresponding entry in 'Form Responses 1' and check the checkbox.
  // Reset the cell background if a recommendation is unchecked
  // This function runs EVERY TIME the spreadsheet is edited by a user, but checks where the edit was made before doing anything

  const currentSheet = ss.getActiveSheet() //the sheet currently being edited
  const range = e.range //the cell that was edited (in this case, the date completed cell)

  //if the edit is not in column F or the edit is in one of the admin sheets (not a teacher tab), end the function
  const setupSheetsNames = [
    'Form Responses 1',
    'Due Dates',
    'Summary', 
    'By School',
    'Deleted Requests',
    'Address Book',
    'Template',
  ]
  if (
    range.getColumn() !== 6 ||
    setupSheetsNames.includes(currentSheet.getName())
  ) {
    Logger.log('outside range')
    return
  }

  //get the universal unique id (Uuid) of the response (from column E of the teacher tab, next to the date completed)
  const checkedUuid = currentSheet.getRange(e.range.getRow(), teacherTabs.columnNumbers.uuId).getValue()

  //get name of sheet (tab) being edited - which teacher completed the rec?
  const sheetName = currentSheet.getName()

  //get array of response Uuids in 'Form Responses 1' sheet
  const lastRow = formResponsesSheet.getLastRow() //get the number of the last row with content
  const responseUuidArray = formResponsesSheet
    .getRange(1, formResponses.columnNumbers.uuId, lastRow)
    .getValues()

  //find the index of the response Uuid that matches the Uuid of the response checked off on the edited tab; add 1 to get the row number of that response (arrays are zero-indexed; ranges are not)
  const responseRow =
    responseUuidArray.findIndex(id => id[0] === checkedUuid) + 1

  const recommendationCells = formResponsesSheet.getRange(
    responseRow,
    formResponses.columnNumbers.mathTeacher,
    1,
    8,
  ) //range of the 4 teacher cells and their completion checkoff columns (8 columns total)

  //get an array of the values (teacher names) from columns D:F of the response row
  const recommendationCellsValues = recommendationCells.getValues()[0] //ex: [JoMarie Broccoli (jbroccoli@nysmith.com), Emily Stephens (estephens@nysmith.com), No Supplemental Recommendation Required]

  //in recommendationCellsValues array, find the index of the teacher name that matches the edited tab; add formResponses.columnNumbers.mathTeacherCompletion to get the correct column in 'Form Responses 1'
  //get this to throw an error and alert spreadsheet admins if not found? e.g. in case someone accidentally edited the tab names, which would break this function?
  const responseColumn =
    recommendationCellsValues.findIndex(
      entry => typeof entry === 'string' && entry.includes(sheetName),
    ) + formResponses.columnNumbers.mathTeacherCompletion

  //get the cell in 'Form Responses 1' that corresponds to the completed recommendation
  const cellToFormat = formResponsesSheet.getRange(responseRow, responseColumn)

  //if the date completed cell on the teacher sheet is filled out, change the background of the corresponding cell in 'Form Responses 1' to light green
  if (range.getValue()) {
    cellToFormat.setValue(true)
  } else {
    //if date completed is deleted, reset background on 'Form Responses 1' page and teacher sheet
    cellToFormat.setValue(false)
  }
}
