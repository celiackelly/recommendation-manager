//When the spreadsheet is opened, create a menu " Admin Controls" with a button to open the "Email Sidebar" 
//Add an event listener, so that when the 'Email Sidebar' button in the menu is clicked, the sidebar is shown
      //PROBLEM- I can't restrict access to this sidebar, can I? So anyone could send emails from here if they have domain edit access on the spreadsheet as a whole?
      //Maybe just display the queued emails in the sidebar, but still trigger them from the button on the speadsheet (so only Brian and I can trigger)
      //Not a problem- looks like the permissions on the sheet take care of this. 
      function onOpen() {
        SpreadsheetApp
          .getUi()
          .createMenu('Admin Controls')
          .addItem('Open Email Sidebar', 'showEmailSidebar')
          .addToUi();
      }
      
      //Render the HTML for the sidebar from the 'sidebar.html' template
      function doGet() {
        const sidebar = HtmlService.createTemplateFromFile('sidebar.html')
        return sidebar.evaluate();
      }
      
      //Define an include() function, which lets you include another file in the HTML template
      //This function is called in the <head> of sidebar.html, to include the 'style.html' file when the HTML is rendered
      //Separation of concerns in Google Apps Script: https://developers.google.com/apps-script/guides/html/best-practices#code.gs
      function include(filename) {
        return HtmlService.createHtmlOutputFromFile(filename)
            .getContent();
      }
      
      //When the 'Email Controls' button in the menu is clicked, show the sidebar
      function showEmailSidebar() {
        const sidebar = doGet();
        sidebar.setTitle('Email Sidebar')
        SpreadsheetApp.getUi().showSidebar(sidebar);
      }
      
      function queueEmails() {
        // When Brian clicks the Update Email Queue button in column J, save information about the checked recommendation entries, to queue for the sendEmails() function
      
        //get all data 
        try {
          const data = formResponsesSheet.getDataRange().getValues()
      
        //filter just the rows with "Queue Emails" column checked (subtract 1 because arrays are zero-indexed)
        const checkedRowsData = data.filter(row => row[formResponses.columnNumbers.queueEmails - 1] === true)
      
        const emailsInfo = checkedRowsData.map(row => {
          const info = {
            studentName: row[formResponses.columnNumbers.studentName - 1], 
            parentEmails: [row[formResponses.columnNumbers.primaryContactEmail - 1], row[formResponses.columnNumbers.secondaryContactEmail - 1]].filter(email => email !== ''), 
            school: row[formResponses.columnNumbers.school - 1],
            uuId: row[formResponses.columnNumbers.uuId - 1]
          }
          return info
        })
      
        //save queued emails info in the Properties Service
        const emailsInfoJSON = JSON.stringify(emailsInfo)
        PropertiesService.getDocumentProperties().setProperty('queuedEmailInfo', emailsInfoJSON) 
        SpreadsheetApp.getUi().alert(`Email queue successfully updated.`) 
      
        } catch(err) {
            SpreadsheetApp.getUi().alert(`Error: ${err}`) 
        }
        showEmailSidebar() 
      }
      
      function clearCheckboxes() {
        //get checkboxes range 
        try {
          const lastRow = formResponsesSheet.getLastRow()
          const checkboxes = formResponsesSheet.getRange(`${formResponses.columnLetters.queueEmails}2:${formResponses.columnLetters.queueEmails}${lastRow}`)
          checkboxes.clearContent()
          // SpreadsheetApp.getUi().alert(`Checkboxes cleared.`) 
        } catch(err) {
            SpreadsheetApp.getUi().alert(`Error: ${err}`) 
        }
      }
      
      function getEmailQueue() {
        const queuedEmailInfo = PropertiesService.getDocumentProperties().getProperty('queuedEmailInfo')
        return JSON.parse(queuedEmailInfo) || []
      }
      
      function clearEmailQueue() {
        //delete any queued email info
        PropertiesService.getDocumentProperties().deleteProperty('queuedEmailInfo') 
        return []
      }
      
      //if you update column structure, you must update this function. refactoring didn't work
      function markEmailsSent(queuedEmailInfo) {        //can you optimize this? 
      
        Logger.log(formResponses.columnNumbers.uuId)
        //get array of response Uuids in 'Form Responses 1' sheet
        const lastRow = formResponsesSheet.getLastRow();  //get the number of the last row with content
        const responseUuidArray = formResponsesSheet.getRange(1, formResponses.columnNumbers.uuId, lastRow).getValues()   
      
        queuedEmailInfo.forEach(emailInfo => {
          const uuId = emailInfo.uuId
          
          //find the index of the response Uuid that matches the Uuid of the response checked off on the edited tab; add 1 to get the row number of that response (arrays are zero-indexed; ranges are not)
          const responseRow = responseUuidArray.findIndex(id => id[0] === uuId) + 1
      
          //In "Email Sent? column," mark this row as sent and change background color to green
          formResponsesSheet.getRange(responseRow, formResponses.columnNumbers.emailsSent).setValue("SENT").setBackground('#d9ead3')
      
          //change color of "Queue Email" checkbox column in this row to grey 
          formResponsesSheet.getRange(responseRow, formResponses.columnNumbers.queueEmails).setBackground('#d9d9d9')
        }) 
      
      }
      
      function sendEmails(queuedEmailInfo) {
      
          try {
            //if there are no queued emails, end the function 
            if (!queuedEmailInfo) { 
            throw new Error("There are no emails in the queue.")
          }
      
            const template = HtmlService.createTemplateFromFile('emailTemplate'); 
            const emailLogo = DriveApp.getFileById("1aY0TjgEmqU8DtvYffn2aps5JOty6R_eq").getAs("image/jpeg");
            const emailImages = {"nysmith-email-logo": emailLogo}; 
      
            //for each emailInfo object in the array/queue, send an email 
            queuedEmailInfo.forEach(emailInfo => {
      
                [template.studentLastName, template.studentFirstName] = emailInfo.studentName.split(', ')
                template.school = emailInfo.school;
                const message = template.evaluate().getContent();
      
                MailApp.sendEmail({
                  to: emailInfo.parentEmails.join(","),
                  replyTo: 'bschrembs@nysmith.com',   //Emails will be sent from Celia's account, but if parents reply, the replies will default to Brian
                  subject: `Completed Recommendation for ${emailInfo.studentName}`,
                  body: `Dear Parents,\n\nAll the high school recommendations for ${template.studentFirstName} ${template.studentLastName} for ${template.school} have been completed. Please contact Brian Schrembs with any questions. \n\nSincerely, \nBrian Schrembs \nFaculty Coordinator for Student Outplacement \n13625 EDS Drive | Herndon, VA 20171 \noffice 703-713-3332 ext.1064 | fax 703-713-3336`, 
                  htmlBody: message, 
                  inlineImages: emailImages
              })
          })
          //return null because there were no errors
          return null
      
          } catch(err) {
              return err
          }
      }
      
      async function handleSendEmails() {
      
        const queuedEmailInfo = await getEmailQueue()
        
        const emailError = await sendEmails(queuedEmailInfo)    //do I need this to be await? or not b/c it's not a promise?
      
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
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      // function markEmailSent(emailInfo) {
      //     const uuId = emailInfo.uuId
      //     Logger.log(uuId)
          
      //     //get array of response Uuids in 'Form Responses 1' sheet
      //     const lastRow = formResponsesSheet.getLastRow();  //get the number of the last row with content
      //     const responseUuidArray = formResponsesSheet.getRange(1, 9, lastRow).getValues()    // 9 = column I (uuid column) 
      
      //     Logger.log(responseUuidArray)
      //     //find the index of the response Uuid that matches the Uuid of the response checked off on the edited tab; add 1 to get the row number of that response (arrays are zero-indexed; ranges are not)
      //     const responseRow = responseUuidArray.findIndex(id => id[0] === uuId) + 1
      //     Logger.log(responseRow)
      //     //In column K "Email Sent?," mark this row as sent and change background color to green
      //     formResponsesSheet.getRange(responseRow, 11).setValue("SENT").setBackground('#d9ead3')
      
      //     //change color of Column J in this row to grey 
      //     formResponsesSheet.getRange(responseRow, 10).setBackground('#d9d9d9')
      // }
      
      // function sendEmails() {
      
      //   try {
      //     const queuedEmailInfo = getEmailQueue()
      //     Logger.log(queuedEmailInfo)
      
      //     //if there are no queued emails, end the function 
      //     if (!queuedEmailInfo) { 
      //       throw new Error("There are no emails in the queue.")
      //     }
      
      //     //for each emailInfo object in the array/queue, send an email 
      //     queuedEmailInfo.forEach(emailInfo => {
      //         MailApp.sendEmail({
      //           to: emailInfo.parentEmails.join(","),
      //           replyTo: 'bschrembs@nysmith.com',   //Emails will be sent from Celia's account, but if parents reply, the replies will default to Brian
      //           subject: `Completed Recommendation for ${emailInfo.studentName}`,
      //           body: `All the recommendations for ${emailInfo.studentName} for ${emailInfo.school} have been completed. Please contact Brian Schrembs with any questions.`
      //       })
      
      //       markEmailSent(emailInfo)
      //     })
      
      //     SpreadsheetApp.getUi().alert('Emails sent successfully.')
      //     clearEmailQueue()
      //     clearCheckboxes()
          
      //   } catch(err) {
      //       SpreadsheetApp.getUi().alert(`Error: ${err}`) 
      //   }
      
      //   showEmailSidebar()
      
      // }
      
      
      
      
      
      
      
      
      
      