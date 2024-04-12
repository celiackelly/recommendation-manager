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
    studentName: 2, //col B
    school: 3, //col C
    mathTeacher: 4, //col D
    laTeacher: 6, //col F
    principalRec: 8, //col H
    source: 10, //col J
    uuId: 12, //col L
    queueEmails: 13, //col M
    emailsSent: 14, //col N
    primaryContactEmail: 15, //col O
    secondaryContactEmail: 16, //col P
    findDuplicatesHelperQuery: 19, //col S
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
