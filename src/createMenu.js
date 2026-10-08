//When the spreadsheet is opened, create a menu " Admin Controls" with a button to open the "Email Sidebar"
//Add an event listener, so that when the 'Email Sidebar' button in the menu is clicked, the sidebar is shown
//Also add a button to open the "Add Teacher Tab" modal
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Admin Controls')
    .addItem('Open Email Sidebar', 'showEmailSidebar')
    .addItem('Add Teacher Tab', 'showAddTeacherTabModal')
    .addToUi()
}

function showAddTeacherTabModal() {
  const html = HtmlService
    .createTemplateFromFile('addTeacherTabModal')
    .evaluate()
    .setWidth(500)
    .setHeight(350)

  SpreadsheetApp.getUi().showModalDialog(html, 'Add Teacher Tab')
}