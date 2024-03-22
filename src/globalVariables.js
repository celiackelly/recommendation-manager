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
    laTeacher: 5, //col E
    supplementalTeacher: 6, //col F; this will change to principalRec
    uuId: 9, //col I
    queueEmails: 10, //col J
    emailsSent: 11, //col K
    primaryContactEmail: 12, //col L
    secondaryContactEmail: 13, //col M
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
