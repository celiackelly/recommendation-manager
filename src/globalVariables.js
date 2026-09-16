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
    deleteRecord: 2, //col B
    studentName: 3, //col C
    school: 4, //col D
    mathTeacher: 5, //col E
    mathTeacherCompletion: 6, //col F
    laTeacher: 7, //col G
    supplementalTeacher: 9, //col I
    principalRec: 11, //col K
    source: 14, //col N
    uuId: 17, //col Q
    queueEmails: 18, //col R
    emailsSent: 19, //col S
    primaryContactEmail: 20, //col T
    secondaryContactEmail: 21, //col U
    findDuplicatesHelperQuery: 24, //col X
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

