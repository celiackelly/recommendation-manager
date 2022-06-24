# Recommendation Manager
Recommendation Manager is an add-on built in Google Apps Script to manage student recommendation requests and send automated completion emails.

The add-on was built for the school where I work, which had been manually managing all of these tasks. It significantly reduces the time and labor needed to track student recommendations, as well as the potential for human error.

![Screenshot of Recommendation Manager](https://github.com/celiackelly/recommendation-manager/blob/1958cc6ed8a11d076026f383f6b71235c849b1ab/recommendation-manager-cover.png)

## How It Works:

- Parents submit a Google Form, listing the school their child is applying to and 2-3 teachers they are requesting recommendations from. 
- Each time the form is submitted, the associated Google Sheet checks the teachers' names against the existing tabs in the sheet. If there is not already  a tab for that teacher, one is created from a template sheet. 
- Each teacher's tab automatically populates the recommendation requests for that teacher, along with a "Date Completed" column. 
- When a teacher marks a recommendation as complete, the corresponding cell in the main 'Form Reponses' tab turns green. 
- When all three recommendations are complete, the administrator checks a box next to the entry on the 'Form Responses' tab to add the information to the email queue. 
- The administrator can view the email queue by opening the custom "Admin Controls" sidebar from the menu. 
- Clicking the "send emails" button sends a "recommendations completed" notification email to the parents, for each completed entry in the email queue.  

## How It's Made:

**Tech used:** Google Apps Script, HTML, CSS



Here's where you can go to town on how you actually built this thing. Write as much as you can here, it's totally fine if it's not too much just make sure you write *something*. If you don't have too much experience on your resume working on the front end that's totally fine. This is where you can really show off your passion and make up for that ten fold.

## Optimizations
*(optional)*

You don't have to include this section but interviewers *love* that you can not only deliver a final product that looks great but also functions efficiently. Did you write something then refactor it later and the result was 5x faster than the original implementation? Did you cache your assets? Things that you write in this section are **GREAT** to bring up in interviews and you can use this section as reference when studying for technical interviews!

## Lessons Learned:

No matter what your experience level, being an engineer means continuously learning. Every time you build something you always have those *whoa this is awesome* or *fuck yeah I did it!* moments. This is where you should share those moments! Recruiters and interviewers love to see that you're self-aware and passionate about growing.

