/**
 * @OnlyCurrentDoc
 */

const ss = SpreadsheetApp.getActive();
const spreadsheetId = ss.getId();

const formResponsesSheet = ss.getSheetByName("Form Responses 1");
const templateSheet = ss.getSheetByName("Template");

const formResponses = (() => {
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

