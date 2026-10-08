//Functions for managing the "Add Teacher Tab" modal, to manually add a teacher tab for a supplemental recommender to the spreadsheet

function createTeacherTabFromEmail(email) {
  email = String(email || '').trim().toLowerCase()

  // Restrict the input to the school's domain and supported username characters.
email = String(email || '').trim().toLowerCase()

const match = email.match(
  /^([a-z0-9]+(?:[._+-][a-z0-9]+)*)@nysmith\.com$/
)

if (!match) {
  throw new Error('Enter a valid @nysmith.com email address.')
}

const name = match[1]

const protectionEmail = `${name}@nysmithschool.com`

  if (name.length > 100) {
    throw new Error('The email username is too long for a tab name.')
  }

  // Prevent simultaneous modal submissions from creating the same tab.
  const lock = LockService.getDocumentLock()
  lock.waitLock(30000)

  try {
    const alreadyExists = ss.getSheets().some(
      sheet => sheet.getName().toLowerCase() === name
    )

    if (alreadyExists) {
      throw new Error(`A tab named "${name}" already exists.`)
    }

    if (!templateSheet) {
      throw new Error('The Template tab could not be found.')
    }

    const admins = [
      'ckelly@nysmithschool.com',
      'bschrembs@nysmithschool.com',
    ]

    const columns = formResponses.columnLetters

    const selectStatement = [
      columns.timeStamp,
      columns.studentName,
      columns.school,
      columns.source,
      columns.earlyDeadline,
      columns.dueDate,
      columns.uuId,
    ].join(', ')

const teacherEmails = [email]

    const recommenderColumns = [
      columns.mathTeacher,
      columns.laTeacher,
      columns.principalRec,
      columns.supplementalTeacher,
    ]

    // Match either "Teacher Name (email)" or a cell containing only the email.
    // Parentheses prevent matching another username that ends with this one.
    const whereStatement = recommenderColumns
      .flatMap(column =>
        teacherEmails.map(address =>
          `(lower(${column}) contains '(${address})' ` +
          `or lower(${column}) = '${address}')`
        )
      )
      .join(' or ')

    const formula =
      `=IFERROR(QUERY('Form Responses 1'!` +
      `${columns.timeStamp}2:${columns.uuId}, ` +
      `"select ${selectStatement} where ${whereStatement}", 0), "")`

    const newSheet = ss.insertSheet(
      name,
      ss.getSheets().length,
      { template: templateSheet }
    )

    const sheetId = newSheet.getSheetId()

    const editableRange = {
      sheetId: sheetId,
      startColumnIndex: teacherTabs.columnNumbers.dateCompleted - 1,
      endColumnIndex: teacherTabs.columnNumbers.notes,
    }

    const requests = [
      // Protect the sheet, except Date Completed and Notes.
      {
        addProtectedRange: {
          protectedRange: {
            range: { sheetId: sheetId },
            description:
              'Except for Date Completed and Notes, only Celia and Brian can edit the sheet',
            warningOnly: false,
            unprotectedRanges: [editableRange],
            editors: {
              users: admins,
              domainUsersCanEdit: false,
            },
          },
        },
      },

      // Restrict Date Completed and Notes to this teacher and the admins.
      {
        addProtectedRange: {
          protectedRange: {
            range: editableRange,
            description:
              `Only ${name}, Celia, and Brian can edit Date Completed and Notes`,
            warningOnly: false,
            editors: {
              users: [...admins, protectionEmail],
              domainUsersCanEdit: false,
            },
          },
        },
      },

      // Add the teacher's QUERY formula to A2.
      {
        updateCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: 2,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          rows: [
            {
              values: [
                {
                  userEnteredValue: {
                    formulaValue: formula,
                  },
                },
              ],
            },
          ],
          fields: 'userEnteredValue',
        },
      },
    ]

    try {
      Sheets.Spreadsheets.batchUpdate(
        { requests: requests },
        spreadsheetId
      )
    } catch (error) {
      // Remove the newly created tab if its setup fails.
      try {
        ss.deleteSheet(newSheet)
      } catch (cleanupError) {
        throw new Error(
          `Setup failed: ${error.message}. ` +
          `The tab "${name}" could not be removed; check its protections before using it.`
        )
      }

      throw new Error(`Could not configure the teacher tab: ${error.message}`)
    }

    // Use the same alphabetical sorting as the form-submit function.
    try {
      sortSheetsAlphabetically()
    } catch (error) {
      return `Tab "${name}" was created, but sorting failed: ${error.message}`
    }

    return `Teacher tab "${name}" created successfully.`
  } finally {
    lock.releaseLock()
  }
}