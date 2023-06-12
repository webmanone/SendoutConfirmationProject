# Sendout Confirmation Creator

## Why?

### Why C#?
This is my first C# project. I'm very interested in C# and .Net since it seems there is a great need for people with this experience in the industry at the  oment, and from what I've seen online it has a very promising future due to the investment from Microsoft. 

### Why this project in particular?
I wanted to create something simple but something that would be able to help me in the real world. <br>

In my current job in recruitment, each time I arrange an interview I need to create a calendar invite with teams and send it to the client and candidate. <br>

Currently, I copy an old calendar invite and paste it and then have to delete and replace a lot of the information, which can be unnecessarily time consuming. <br>

The idea came to me one night after I glanced at my phone at 11pm and saw that a client had asked me to schedule an in-person interview as a first stage interview for the next day. (I work in a specialist construction industry in North America that can move very fast at times due to project requirements). <br>

I was there working at 11pm wishing I was in bed, wishing there was a quicker way to arrange the meeting that would allow me to go back to sleep faster, so I thought of this project idea. <br>

C# is perfect for this because of its integration with Microsoft products and should be easy to create an application that's compatible with outlook/Teams.

## Implementation

### What the program does

Initial starting screen:

![Initial Screen](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/b2debf70-205d-42e9-ab23-14d788747f29)


Overall, the program allows me to quickly generate an interview confirmation in the outlook web app.
Breakdown:
- Allows the user to select a template depending on which interview they need to arrange. (First stage phone call, teams, in-person or other).
- Allows the user to select the date/time of the interview in their local time. Depending on the time zones of the client and caniddates, the template will be autimatically updated to account for different time zones and daylight saving time.
- Allows the user to enter all the client and candidate information necessary to arrange the interview.
- Allows the user to include additional information/accommodations if needed.
- Allows the user to save and load client and candidate information locally.
- Pressing the preview button takes all the data and puts it in an outlook calendar invite using the Microsoft Graph API.

#### Examples screenshots:

Loading a client:

![Load client button click](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/5565af67-e176-424a-8cab-10b5a4372a8c)

![After load client selection](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/844db3bf-2b3f-4c58-bb85-09b8e60d5c98)

Loading a candidate:

![Load candidate button](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/44478841-d52f-4602-a9f6-7b711ce037f0)

![After load candidate button click](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/821f44e3-c557-4c46-acde-8d53558487df)

Error handling when fields are empty:

![Error handling when all fields aren't filled in](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/f2892a74-7860-426f-b490-fbc93a09e0b7)

Date/Time selector:

![DateTime selector](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/1d7e7b55-1475-4894-8753-91d7b9b01fa8)

Template selector:

![Template selector](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/71316c7e-1cad-4b73-b659-b27eed9d776e)

TimeZone selector:

![Time Zone Change](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/36695948-8b32-42a9-bfbf-bf3e3d465a72)

Preview button pressed, leading to the calendar invite being created:

![Calendar Invite Created Example Elon and Joe](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/4d521fa9-9777-4238-ada2-c057df259b87)

Saving example:

![Christopher nolan saveclient](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/d75e0c3f-ea3d-49a9-9c97-567ec9b97c66)

Additional info example:

![Additional info example](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/2ad5de25-0e19-44ce-ae47-4d03bcfa0576)

Calendar invite example:

![Leo and Christopher calendar invite](https://github.com/webmanone/SendoutConfirmationProject/assets/66835665/5076a5d2-15c3-4933-a3e0-1494b0f16220)

### Examples of OOP

#### Abstraction
I have the classes 'Person', 'Client', and 'Candidate in my code:

```
 public class Person
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string TimeZone { get; set; }

        public void SavePerson(string filePath)
        {
            // Serializes the person to JSON
            string json = JsonConvert.SerializeObject(this);

            try
            {
                // Appends the JSON string to the existing file
                File.AppendAllText(filePath, json + Environment.NewLine);
                System.Windows.MessageBox.Show("Information saved successfully.");
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"An error occurred while saving the person information: {ex.Message}");
            }
        }
    }
 ```
 ```
 public class Client : Person
    {
        public string Company { get; set; }
    }
  ```
  
  ```
  public class Candidate: Person
    {
        public string Phone { get; set; }
    }
  ```
  
  This showcases abstraction because I define common attributes in the person class, but have the company and phone seperate for the client and candidate. Although it may be important information in other contexts to know the Candidate's company and the Client's phone number, it's not necessary for my program to work. 

#### Inheritance
The code above also shows inheritance, since Client and Candidate inherit from the Person class.

Then, later in my code I can create instance of client and candidate whilst inheriting properties from the person class:
```
            Client client = new Client
            {
                Name = clientName,
                Email = clientEmail,
                Company = clientCompany
            };

            // Create candidate object
            Candidate candidate = new Candidate
            {
                Name = candidateName,
                Email = candidateEmail,
                Phone = candidatePhone
            };
```
#### Polymorphism
By creating classes in this way, I have been able to create the following functions:
```
 private void SavePerson_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            string filePath = null;

            //Creates a person object and then creates either a client or candidate depending on the button pressed.
            Person person = null;
            if (button.Name == SaveClientButton.Name)
            {
                filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\clients.json";
                person = new Client
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = ClientNameTextBox.Text,
                    Email = ClientEmailTextBox.Text,
                    Company = ClientCompanyTextBox.Text,
                    TimeZone = ClientComboBox.Text
                };
            }
            else if (button.Name == SaveCandidateButton.Name)
            {
                filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\candidates.json";
                person = new Candidate
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = CandidateNameTextBox.Text,
                    Email = CandidateEmailTextBox.Text,
                    Phone = CandidatePhoneTextBox.Text,
                    TimeZone = CandidateComboBox.Text
                };
            }
            //calls SavePerson method
            person.SavePerson(filePath);
        }
```
This allows me to have 1 function to save a person, instead of creating seperate functions for saving a candidate and saving a client, showing the advantage of having classes which can have many forms.

#### Encapsulation
Encapsulation is especially helpful when working on larger projects where multiple people are working on the same codebase and might not need to know the internal details of a module or class.

Since I've been working on this project on my own, I haven't necessarily needed to do any information hiding, but I have tried to make my code as modular as possible to increase the maintainability of my code.

For example, the following code gets the correct time zone and time for the client and candidate based on the user's local time, accounting for daylight saving time. I have seperated this into 3 functions to seperate the concerns and so it's clear what each section does:

```
private void UpdatePersonTimeZone(Person person)
        {
            string timeZoneString = (person is Client) ? ClientComboBox.Text : CandidateComboBox.Text;
            TimeZoneInfo personTimeZone = null;

            if (timeZoneString == "Eastern")
            {
                personTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            }
            else if (timeZoneString == "Central")
            {
                personTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
            }
            else if (timeZoneString == "Mountain")
            {
                personTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time");
            }
            else if (timeZoneString == "Pacific")
            {
                personTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
            }

            DateTimeOffset selectedDateTimeOffset = DaylightConvert(selectedDateTime, personTimeZone);
            personTimeZone = TimeZoneInfo.CreateCustomTimeZone(personTimeZone.Id, selectedDateTimeOffset.Offset, personTimeZone.DisplayName, personTimeZone.StandardName);
            string personTime = ConvertTimeZone(selectedDateTime, personTimeZone);

            if (person is Client)
            {
                clientTimeZoneString = timeZoneString;
                clientTimeZone = personTimeZone;
                clientTime = personTime;
            } else if (person is Candidate)
            {
                candidateTimeZoneString = timeZoneString;
                candidateTimeZone = personTimeZone;
                candidateTime = personTime;
            }
        }

        //Function that calculates the offset for daylight saving time
        public static DateTimeOffset DaylightConvert(DateTime selectedDateTime, TimeZoneInfo timeZone)
        {
            DateTimeOffset selectedDateTimeOffset = new DateTimeOffset(selectedDateTime, TimeSpan.Zero);
            TimeSpan offset = timeZone.GetUtcOffset(selectedDateTimeOffset);

            if (timeZone.IsDaylightSavingTime(selectedDateTime))
            {
                TimeZoneInfo.AdjustmentRule[] rules = timeZone.GetAdjustmentRules();
                foreach (TimeZoneInfo.AdjustmentRule rule in rules)
                {
                    if (rule.DateStart <= selectedDateTime && selectedDateTime < rule.DateEnd)
                    {
                        offset += rule.DaylightDelta;
                        break;
                    }
                }
            }

            return new DateTimeOffset(selectedDateTime, offset);
        }
        
        //Converts the date/time from local time to the target time zone
        private string ConvertTimeZone(DateTime dateTime, TimeZoneInfo targetZone)
            {
            TimeZoneInfo localZone = TimeZoneInfo.Local;
            
            DateTimeOffset selectedDateTimeOffset = DaylightConvert(dateTime, localZone);
            localZone = TimeZoneInfo.CreateCustomTimeZone(localZone.Id, selectedDateTimeOffset.Offset, localZone.DisplayName, localZone.StandardName);
            DateTime convertedTime = TimeZoneInfo.ConvertTime(dateTime, localZone, targetZone);
            
            string formattedTime = convertedTime.ToString("h:mm tt", new CultureInfo("en-US"));
            return formattedTime;
            }
```
### Future improvements
- Need to allow user to update the file path that they save candidates to, or save candidates to a database stored online.
- Need to include the edit template functionality. Currently the template is very specilised towards my style of recruitment.
- Since all my clients are in the 4 time zones specified, I don't need any more, but it would be helpful to include the option of selecting any time zone if other recruiters were to use the software.
- Adding the option to add multiple clients may be helpful in the rare event that I need to include more than 1 client for an interview, I would manually do it currently in the email. However, I could restructure the program so that the clients are stored in a list, and then all clients are automatically included in the invite.
- Visual - currently quite bland and generic, could do with a design improvement.
- Limited to outlook web app. Using the Microsoft API, my code currently creates a calendar inite in the outlook web app, but some people may want to use gmail or another software.
- Currently the Teams Interview template doesn't automatically create a Teams link and set it so that the users can automatically join. An oversight I had was that I assumed the outlook on my personal laptop would be the same as on my work laptop, but you have to pay to get the business version of Outlook which is different. As of my personal Outlook, the default is set to Skype. It may work on my work on my work laptop, but it should be easier to make sure it sets up the online meeting regardlesss of if you're using Teams, Skype, Zoom, etc.
- The current entry fields may be limited for other recruiters, for example some may need to include other information to arrange an interview.
- Enable searching/filtering/ordering list of saved clients/candidates as the list may get quite big.
- Better handle saving duplicate candidates and clients.

In its current state, the code is very useful to me in my personal job, but there are a number of improvements that need to be made before it could be shippable commercially and used by other recruiters.

