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
    mathTeacher: 4, 
    laTeacher: 6, 
    principalRec: 8, 
    source: 10, 
    uuId: 12, 
    queueEmails: 13, 
    emailsSent: 14, 
    primaryContactEmail: 15, 
    secondaryContactEmail: 18, 
    findDuplicatesHelperQuery: 19,
  };

  const convertColNumstoLetters = () => {
    const columnLetters = {};

    for (let column in columnNumbers) {
      columnLetters[column] = String.fromCharCode(columnNumbers[column] + 64);
    }
    return columnLetters;
  };

  const columnLetters = convertColNumstoLetters();

  return {
    columnNumbers,
    columnLetters,
  };
})();
