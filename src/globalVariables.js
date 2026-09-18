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
    privateOrPublic: 4, //col D
    school: 5, //col E
    mathTeacher: 6, //col F
    mathTeacherCompletion: 7, //col G
    laTeacher: 8, //col H
    supplementalTeacher: 10, //col J
    principalRec: 12, //col L
    source: 15, //col O
    publicSchoolName: 18, //col R
    uuId: 19, //col S
    queueEmails: 20, //col T
    emailsSent: 21, //col U
    primaryContactEmail: 22, //col V
    secondaryContactEmail: 23, //col W
    findDuplicatesHelperQuery: 26, //col Z
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

