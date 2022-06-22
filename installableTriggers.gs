function createTriggers() {

  const ss = SpreadsheetApp.getActive()

  ScriptApp.newTrigger('createNewSheetOnSubmit')
  .forSpreadsheet(ss)
  .onFormSubmit()
  .create();

  ScriptApp
  .newTrigger('addUuidAndCheckbox')
  .forSpreadsheet(ss)
  .onFormSubmit()
  .create();

  ScriptApp.newTrigger('markCompletion')
  .forSpreadsheet(ss)
  .onEdit()
  .create();

  ScriptApp.newTrigger('queueCompletionEmail')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
}
