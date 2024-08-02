//When the spreadsheet is opened, create a menu " Admin Controls" with a button to open the "Email Sidebar"
//Add an event listener, so that when the 'Email Sidebar' button in the menu is clicked, the sidebar is shown
//PROBLEM- I can't restrict access to this sidebar, can I? So anyone could send emails from here if they have domain edit access on the spreadsheet as a whole?
//Maybe just display the queued emails in the sidebar, but still trigger them from the button on the speadsheet (so only Brian and I can trigger)
//Not a problem- looks like the permissions on the sheet take care of this.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Admin Controls')
    .addItem('Open Email Sidebar', 'showEmailSidebar')
    .addToUi()
}

//Render the HTML for the sidebar from the 'sidebar.html' template
function doGet() {
  const sidebar = HtmlService.createTemplateFromFile('sidebar.html')
  return sidebar.evaluate()
}

//Define an include() function, which lets you include another file in the HTML template
//This function is called in the <head> of sidebar.html, to include the 'style.html' file when the HTML is rendered
//Separation of concerns in Google Apps Script: https://developers.google.com/apps-script/guides/html/best-practices#code.gs
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

//When the 'Email Controls' button in the menu is clicked, show the sidebar
function showEmailSidebar() {
  const sidebar = doGet()
  sidebar.setTitle('Email Sidebar')
  SpreadsheetApp.getUi().showSidebar(sidebar)
}

function queueEmails() {
  // When Brian clicks the Update Email Queue button in column J, save information about the checked recommendation entries, to queue for the sendEmails() function

  //get all data
  try {
    const data = formResponsesSheet.getDataRange().getValues()

    //filter just the rows with "Queue Emails" column checked 
    const checkedRowsData = data.filter(
      row => row[formResponses.columnIndex.queueEmails] === true,
    )

    const emailsInfo = checkedRowsData.map(row => {
      const info = {
        studentName: row[formResponses.columnNumbers.studentName - 1],
        parentEmails: [
          row[formResponses.columnNumbers.primaryContactEmail - 1],
          row[formResponses.columnNumbers.secondaryContactEmail - 1],
        ].filter(email => email !== ''),
        school: row[formResponses.columnNumbers.school - 1],
        uuId: row[formResponses.columnNumbers.uuId - 1],
      }
      return info
    })

    //save queued emails info in the Properties Service
    const emailsInfoJSON = JSON.stringify(emailsInfo)
    PropertiesService.getDocumentProperties().setProperty(
      'queuedEmailInfo',
      emailsInfoJSON,
    )
    SpreadsheetApp.getUi().alert(`Email queue successfully updated.`)
  } catch (err) {
    SpreadsheetApp.getUi().alert(`Error: ${err}`)
  }
  showEmailSidebar()
}

function clearCheckboxes() {
  //get checkboxes range
  try {
    const lastRow = formResponsesSheet.getLastRow()
    const checkboxes = formResponsesSheet.getRange(
      `${formResponses.columnLetters.queueEmails}2:${formResponses.columnLetters.queueEmails}${lastRow}`,
    )
    checkboxes.clearContent()
    // SpreadsheetApp.getUi().alert(`Checkboxes cleared.`)
  } catch (err) {
    SpreadsheetApp.getUi().alert(`Error: ${err}`)
  }
}

function getEmailQueue() {
  const queuedEmailInfo =
    PropertiesService.getDocumentProperties().getProperty('queuedEmailInfo')
  return JSON.parse(queuedEmailInfo) || []
}

function clearEmailQueue() {
  //delete any queued email info
  PropertiesService.getDocumentProperties().deleteProperty('queuedEmailInfo')
  return []
}

function markEmailsSent(queuedEmailInfo) {
  //get array of response Uuids in 'Form Responses 1' sheet
  const lastRow = formResponsesSheet.getLastRow() //get the number of the last row with content
  const responseUuidArray = formResponsesSheet
    .getRange(1, formResponses.columnNumbers.uuId, lastRow)
    .getValues()

  queuedEmailInfo.forEach(emailInfo => {
    const uuId = emailInfo.uuId

    //find the index of the response Uuid that matches the Uuid of the response checked off on the edited tab; add 1 to get the row number of that response (arrays are zero-indexed; ranges are not)
    const responseRow = responseUuidArray.findIndex(id => id[0] === uuId) + 1

    //In "Email Sent? column," mark this row as sent; colors of columns M and N will change based on conditional formatting set in the sheet
    formResponsesSheet
      .getRange(responseRow, formResponses.columnNumbers.emailsSent)
      .setValue('SENT')
  })
}

function sendEmails(queuedEmailInfo) {
  try {
    //if there are no queued emails, end the function
    if (!queuedEmailInfo) {
      throw new Error('There are no emails in the queue.')
    }

    const template = HtmlService.createTemplateFromFile('emailTemplate')
    const emailLogo = DriveApp.getFileById(
      '1aY0TjgEmqU8DtvYffn2aps5JOty6R_eq',
    ).getAs('image/jpeg')
    const emailImages = { 'nysmith-email-logo': emailLogo }

    //for each emailInfo object in the array/queue, send an email
    queuedEmailInfo.forEach(emailInfo => {
      ;[template.studentLastName, template.studentFirstName] =
        emailInfo.studentName.split(', ')
      template.school = emailInfo.school
      const message = template.evaluate().getContent()

      MailApp.sendEmail({
        to: emailInfo.parentEmails.join(','),
        subject: `Completed Recommendation for ${emailInfo.studentName}`,
        body: `Dear Parents,\n\nAll the high school recommendations for ${template.studentFirstName} ${template.studentLastName} for ${template.school} have been completed. Please contact Celia Kelly or Brian Schrembs with any questions. \n\nSincerely, \nCelia Kelly \nAdministrative Assistant \nRegistrar | Student Outplacement Coordinator \n13625 EDS Drive | Herndon, VA 20171 \noffice 703-713-3332 ext.1151 | fax 703-713-3336`,
        htmlBody: message,
        inlineImages: emailImages,
      })
    })
    //return null because there were no errors
    return null
  } catch (err) {
    return err
  }
}

async function handleSendEmails() {
  const queuedEmailInfo = await getEmailQueue()

  const emailError = await sendEmails(queuedEmailInfo) //do I need this to be await? or not b/c it's not a promise?

  if (emailError) {
    SpreadsheetApp.getUi().alert(`Error sending emails: ${emailError}`)
    return
  } else {
    SpreadsheetApp.getUi().alert('Emails sent successfully.')
    markEmailsSent(queuedEmailInfo)
    clearEmailQueue()
    clearCheckboxes()

    showEmailSidebar()
  }
}