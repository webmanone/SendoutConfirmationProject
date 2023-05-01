# Sendout Confirmation Creator

## Why?

### Why C#?
This is my first C# project. I'm very interested in C# and .Net since it seems there are an abundance of jobs for C# online and from what I've seen online it has a very promising future due to the investment from Microsoft. 

### Why this project in particular?
I wanted to create something simple but something that would be able to help me in the real world. <br>

In my current job in recruitment, each time I arrange an interview I need to create a calendar invite with teams and send it to the client and candidate. <br>

Currently, I copy an old calendar invite and paste it and then have to delete and replace a lot of the information, which can be unnecessarily time consuming. <br>

The idea came to me since last night I was copying and pasting after I glanced at my phone and saw that a client had asked me to schedule an in-person interview as a first stage interview for tomorrow. (I work in a specialist construction industry that can move very fast at times due to project requirements). <br>

I was there at 11pm wishing I was in bed and wishing there was a quicker way to arrange the meeting that would allow me to go back to sleep faster, so I thought of this project idea. <br>

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
