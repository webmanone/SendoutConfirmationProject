# Sendout Confirmation Creator

## Why?

### Why C#?
This is my first C# project. I'm very interested in C# and .Net since it seems there is a great need for people with this experience in the industry at them oment, and from what I've seen online it has a very promising future due to the investment from Microsoft. 

### Why this project in particular?
I wanted to create something simple but something that would be able to help me in the real world. <br>

In my current job in recruitment, each time I arrange an interview I need to create a calendar invite with teams and send it to the client and candidate. <br>

Currently, I copy an old calendar invite and paste it and then have to delete and replace a lot of the information, which can be unnecessarily time consuming. <br>

The idea came to me one night after I glanced at my phone at 11pm and saw that a client had asked me to schedule an in-person interview as a first stage interview for the next day. (I work in a specialist construction industry in North America that can move very fast at times due to project requirements). <br>

I was there working at 11pm wishing I was in bed, wishing there was a quicker way to arrange the meeting that would allow me to go back to sleep faster, so I thought of this project idea. <br>

C# is perfect for this because of its integration with Microsoft products and should be easy to create an application that's compatible with Teams.

## Implementation

### What does it need to have?

The project needs to first check whether it's a first stage phone call, teams interview, an in-person interview or a second/further stage interview. <br>

It then needs to take in all the necessary user data: Candidate name, client name, company name, candidate phone number, candidate email, client email, day and time of the interview. It should also contain an optional box to add additional information if needed.<br>

If it's a teams invite, it needs to automatically change the settings so that both the client and candidate can access the interview without me being present.<br>

Because I'm scared of things going wrong and sending the wrong email to the wrong person, it needs to show the calendar invite and allow me to manually review it before I send it at first. <br>

The user needs to be able to customise the message template, include thier email signature etc. <br>

Would be good if the user has the option to save client information and candidate information so they can select them quickly later.<br>

GUI checklist:
- Template (first call, teams, etc.)
- Candidate name, number, email, time zone.
- Client name, email, company, time zone.
- Date & time of the interview.
- Optional box for additional information.
- Make sure user signature is added (should automatically be added on the invite).
- Preview & ability to make chages before sending.
- Ability to save client or candidate information and then use it in the future.
- Add ability to edit template.
- Add option to add extra clients and copy people in on the invite.

Code/functionality checklist:
- Takes the user input and puts it in a calendar invite.
- Stores hidden templates for each different template so that the strings/information is slotted into the empty spaces when previewed.
- Validates the data.
- Saves client/candidate when data is clean.
- Loads candidate into the boxes for additional editing.
- Pressing preview will bring up the teams invite with everything ready to send.
- Lets the user edit a template.
- Ensure that if it's a teams invite, the settings are changed so that the client and candidate can enter the call without you being present.

## Current status - 21/05/2023

### What the program does

Overall, the program allows me to quickly generate an interview confirmation in outlook.
Breakdown:
- Allows the user to select a template depending on which interview they need to arrange. (First stage phone call, teams, in-person or other).
- Allows the user to select the date/time of the interview in their local time. Depending on the time zones of the client and caniddates, the template will be autimatically updated to account for different time zones and daylight saving time.
- Allows the user to enter all the client and candidate information necessary to arrange the interview.
- Allows the user to include additional information/accommodations if needed.
- Allows me to save and load client information locally
- Pressing the preview button takes all the data and puts it in an outlook meeting request.

### Future improvements
- Need to allow user to update the file path that they save candidates to, or save candidates to a database stored online.
- Need to include the edit template functionality. Currently the template is very specilised towards my style of recruitment.
- Since all my clients are in the 4 time zones specified, I don't need any more, but it would be helpful to include the option of selecting any time zone if other recruiters were to use the software.
- Adding additional clients doesn't work currently. In the rare event that I do need to include more than 1 client, I would manually do it currently in the email. However, I could restructure the program so that the clients are stored in a list, and then all clients are automatically included in the invite.
- Visual - currently quite bland and generic, could do with a design improvement.
- Limited to Outlook - currently, this only works with the Outlook desktop app. This is fine for personal use, but in another version I should make use of the Microsoft Graph API so that the program works for whatever platform the user is using, whether it's outlook in the web browser, google calendar, etc.
- Currently the Teams Interview template doesn't automatically create a Teams link and set it so that the users can automatically join. An oversight I had was that I assumed the outlook on my personal laptop would be the same as on my work laptop, but you have to pay to get the business version of Outlook which is different. As of my personal Outlook, the default is set to Skype. It may work on my work on my work laptop, but it should be easier to make sure it sets up the online meeting regardlesss of if you're using Teams, Skype, Zoom, etc. using the Graph API.
- The current entry fields may be limited for other recruiters, for example some may need to include the client's number, candidate's company, etc.


In its current state, the code is very useful to me in my personal job, but there are a number of improvements that need to be made before it could be shippable commercially and used by other recruiters.
