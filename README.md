# Recommendation Manager
Recommendation Manager is an add-on built in Google Apps Script to manage student recommendation requests and send automated completion emails.

The add-on was built for the school where I work, which had been manually managing all of these tasks. It significantly reduces the time and labor needed to track student recommendations, as well as the potential for human error.

![Screenshot of Recommendation Manager](https://github.com/celiackelly/recommendation-manager/blob/1958cc6ed8a11d076026f383f6b71235c849b1ab/recommendation-manager-cover.png)

## How It Works:

- Parents submit a Google Form, listing the school their child is applying to and 2-3 teachers they are requesting recommendations from. 
- Each time the form is submitted, the associated Google Sheet checks the teachers' names against the existing tabs in the spreadsheet. If there is not already a tab for that teacher, one is created from a template sheet. 
- Each teacher's tab includes a query function to automatically populate the recommendation requests for that teacher, along with a "Date Completed" column. 
- When a teacher marks a recommendation as complete, the corresponding cell in the main 'Form Reponses' tab turns green. 
- When all three recommendations are complete, the administrator checks a box next to the entry on the 'Form Responses' tab to add the information to the email queue. 
- The administrator can view the email queue by opening the custom "Admin Controls" sidebar from the menu. 
- Clicking the "send emails" button sends a "recommendations completed" notification email to the parents, for each completed entry in the email queue.  

## How It's Made:

**Tech used:** Google Apps Script, HTML, CSS

The core of the add-on is written in a container-bound Google Apps Script file, `code.gs`. The program calls several Google Apps Script services, including the Spreadsheet, Forms, Mail, Properties, and Utilities services.
- The Utilities service is called to create universal unique identifiers (uuIDs), which associate form submissions on the 'Form Responses' tab with their corresponding entries on the individual teachers' tabs. 
- The Properties service is used to store the list of queued emails as it is being built up, so that the sendQueuedEmails() function can access this data when it is time to send. 

The `installableTriggers.gs` file programmatically creates triggers to execute certain functions from code.gs when the form is submitted or the spreadsheet is updated. This file needs to be run once when setting up the project.  

The custom sidebar is coded using templated HTML and CSS. When the sidebar is opened, it gets the queuedEmailInfo property from the Properties Service, parses it, and dynamically creates a table to display the information for each email. In this way, the user can verify that the information is correct before clicking the "send emails" button. 

## Lessons Learned:

I spent hours reading the Google Apps Script documentation while developing this add-on. I learned how to do a number of things programmatically that I had only ever done through the UI, including how to generate a new sheet from a template, set data validation rules, and build installable triggers. A big win was when I realized that the new tabs my program was generating did not have the same protections as the template, and I had to figure out how to set protected ranges and user permissions programmatically. 

The biggest challenge I encountered was figuring out how to send queued emails from the spreadsheet UI. My original plan was to trigger
Properties Service and the weirdness of global variables in Apps Script 

This was also my first time serving templated HTML within Google Sheets to create a custom sidebar, which opens up an exciting world of possibilities for building other custom add-ons for Google products. 

## Optimizations: 

Since there is no way to add an event listener to a specific range in Goole Apps Script, the markCompletion() and queueEmailCompletion() functions have to run each time the spreadsheet is edited, which is not ideal for performance. To optimize the speed of these functions, a conditional first checks whether the edit was in the the "Date Completed" or "Ready to Send?" columns; if not in that range, the function immediately returns. 

## Next Steps:

