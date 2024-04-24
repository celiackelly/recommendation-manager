/**
 * @OnlyCurrentDoc
 */

const ss = SpreadsheetApp.getActive()
const spreadsheetId = ss.getId()

const formResponsesSheet = ss.getSheetByName('Form Responses 1')
const templateSheet = ss.getSheetByName('Template')

const formResponses = (() => {
  const columnNumbers = {
    timeStamp: 1, //col A
    deleteRecord: 2,
    studentName: 3,
    school: 4,
    mathTeacher: 5,
    mathTeacherCompletion: 6,
    laTeacher: 7,
    principalRec: 9,
    source: 11,
    uuId: 13,
    queueEmails: 14,
    emailsSent: 15,
    primaryContactEmail: 16,
    secondaryContactEmail: 17,
    findDuplicatesHelperQuery: 20,
  }

  const convertColNumstoLetters = () => {
    const columnLetters = {}

    for (let column in columnNumbers) {
      columnLetters[column] = String.fromCharCode(columnNumbers[column] + 64)
    }
    return columnLetters
  }

  const columnLetters = convertColNumstoLetters()

  const converColNumsToIndex = () => {
    const columnIndex = {}

    for (let column in columnNumbers) {
      columnIndex[column] = columnNumbers[column] - 1
    }
    return columnIndex
  }

  const columnIndex = converColNumsToIndex()

  return {
    columnNumbers,
    columnLetters,
    columnIndex,
  }
})()

const teacherTabs = (() => {
  const columnNumbers = {
    timeStamp: 1, //col A
    studentName: 2,
    school: 3,
    source: 4,
    uuId: 5,
    dateCompleted: 6,
    notes: 7,
  }

  const convertColNumstoLetters = () => {
    const columnLetters = {}

    for (let column in columnNumbers) {
      columnLetters[column] = String.fromCharCode(columnNumbers[column] + 64)
    }
    return columnLetters
  }

  const columnLetters = convertColNumstoLetters()

  const converColNumsToIndex = () => {
    const columnIndex = {}

    for (let column in columnNumbers) {
      columnIndex[column] = columnNumbers[column] - 1
    }
    return columnIndex
  }

  const columnIndex = converColNumsToIndex()

  return {
    columnNumbers,
    columnLetters,
    columnIndex,
  }
})()

