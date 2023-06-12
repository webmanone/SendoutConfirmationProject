using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Xceed.Wpf.Toolkit;
using System.Globalization;
using Newtonsoft.Json;
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Diagnostics;
using System.Web;
using System.Windows.Forms;
using System.Security.AccessControl;

namespace Sendout_Calendar_Invite_Project
{
    public partial class MainWindow : Window
    {
        private string selectedTemplate = "First stage phone call";
        private string selectedDate; //Stores the selected date to use in the template
        private DateTime selectedDateTime; //Stores the selected date and time for the purpose of using in the calendar invite
        private TimeZoneInfo clientTimeZone; //Time zone to be used in time zone conversion calculations.
        private TimeZoneInfo candidateTimeZone; //""
        private string clientTime; //Used in the templates
        private string candidateTime; //""
        private string clientTimeZoneString = "Eastern"; //""
        private string candidateTimeZoneString = "Eastern"; //""
        private string location = ""; //Stores location of the event, which will change between Microsoft Teams or the candidate's phone number depending on what's selected.
        private static readonly HttpClient graphClient = new HttpClient();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Preview_Click(object sender, RoutedEventArgs e)
        {


            //Initialise variables to store user input
            string eventTitle = null;
            string clientName = ClientNameTextBox.Text;
            string clientEmail = ClientEmailTextBox.Text;
            string clientCompany = ClientCompanyTextBox.Text;
            string candidateName = CandidateNameTextBox.Text;
            string candidateEmail = CandidateEmailTextBox.Text;
            string candidatePhone = CandidatePhoneTextBox.Text;
            string additionalInfo = AdditionalInfoTextBox.Text;
            string clientFirstName = clientName.Split(' ')[0];
            string candidateFirstName = candidateName.Split(' ')[0];
            string emailTemplate = "";
            string differentTimeZone = "";

            //Validation to ensure fields aren't empty
            if (string.IsNullOrWhiteSpace(ClientNameTextBox.Text) ||
                string.IsNullOrWhiteSpace(ClientEmailTextBox.Text) ||
                string.IsNullOrWhiteSpace(ClientCompanyTextBox.Text) ||
                string.IsNullOrWhiteSpace(CandidateNameTextBox.Text) ||
                string.IsNullOrWhiteSpace(CandidateEmailTextBox.Text) ||
                string.IsNullOrWhiteSpace(CandidatePhoneTextBox.Text) ||
                ClientComboBox.SelectedItem == null ||
                CandidateComboBox.SelectedItem == null ||
                selectedDateTime == DateTime.MinValue)
            {
                System.Windows.Forms.MessageBox.Show("Please fill in all required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //Create client object
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

            UpdatePersonTimeZone(client);
            UpdatePersonTimeZone(candidate);

            // Create calendar invite object using client and candidate objects
            CalendarInvite invite = new CalendarInvite
            {
                EventTitle = eventTitle,
                Location = location,
                EventType = selectedTemplate,
                Date = selectedDate,
                StartTime = selectedDateTime,
                EndTime = selectedDateTime.AddMinutes(30),
                Client = client,
                Candidate = candidate,
                AdditionalInfo = additionalInfo
            };

            //Update client and candidate time zones
            UpdatePersonTimeZone(client);
            UpdatePersonTimeZone(candidate);

            //Construct event title
            eventTitle = $"{candidate.Name}/{client.Company} - {invite.EventType}";

            //Determine if candidate and client time zones are different, and edit template information accordingly
            if (clientTimeZoneString != candidateTimeZoneString)
            {
                differentTimeZone += $"{clientTime} {clientTimeZoneString}/{candidateTime} {candidateTimeZoneString}";
            } else
            {
                differentTimeZone += $"{clientTime} {clientTimeZoneString}";
            }

            //Append additional information if provided
            if (invite.AdditionalInfo != "") {
                invite.AdditionalInfo = $"<p>{invite.AdditionalInfo}</p> <br>";
            }

            //Generate text in the body of the calendar invite based on selected template
            if (selectedTemplate == "First stage phone call")
            {
                emailTemplate += $"<html><body>" +
                    $"<p>{clientFirstName}/{candidateFirstName},</p> <br>" +
                    $"<p>I'm pleased to confirm the following call at {differentTimeZone} on {selectedDate}.</p> <br>" +
                    $"<p>Client: {client.Name} - {client.Company}</p>" +
                    $"<p>Candidate: {candidate.Name}</p>" +
                    $"<p>Date: {invite.Date}</p>" +
                    $"<p>Time: {differentTimeZone}</p> <br>" +
                    $"<p>{clientFirstName} - Please call {candidateFirstName} on {candidate.Phone} at the arranged time.</p> <br>" +
                    $"<p>I'm looking forward to discussing feedback following the call.</p> <br>" +
                    $"<p>If anything comes up and we need to re-arrange the call, please let me know.</p> <br>" +
                    $"{invite.AdditionalInfo}" +
                    $"<p>Best regards,</p>" +
                    $"</body></html>";

                invite.Location = candidatePhone;

            } else if (selectedTemplate == "Teams interview")
            {
                emailTemplate += $"<html><body>" +
                    $"<p>{clientFirstName}/{candidateFirstName},</p> <br>" +
                    $"<p>I'm pleased to confirm the following {invite.EventType} at {differentTimeZone} on {selectedDate}.</p> <br>" +
                    $"<p>Client: {client.Name} - {client.Company}</p>" +
                    $"<p>Candidate: {candidate.Name}</p>" +
                    $"<p>Date: {invite.Date}</p>" +
                    $"<p>Time: {differentTimeZone}</p> <br>" +
                    $"<p>Please join the Teams meeting at the arranged time.</p> <br>" +
                    $"<p>I'm looking forward to discussing feedback following the call.</p> <br>" +
                    $"<p>If anything comes up and we need to re-arrange the call, please let me know.</p> <br>" +
                    $"{invite.AdditionalInfo}" +
                    $"<p>Best regards,</p>" +
                    $"</body></html>";

            } else if (selectedTemplate == "In-person interview")
            {
                emailTemplate += $"<html><body>" +
                    $"<p>{clientFirstName}/{candidateFirstName},</p> <br>" +
                    $"<p>I'm pleased to confirm the following in-person interview at {differentTimeZone} on {selectedDate}.</p> <br>" +
                    $"<p>Client: {client.Name} - {client.Company}</p>" +
                    $"<p>Candidate: {candidate.Name}</p>" +
                    $"<p>Date: {invite.Date}</p>" +
                    $"<p>Time: {differentTimeZone}</p> <br>" +
                    $"<p>{clientFirstName} - Please reach out to {candidateFirstName} to arrange the meeting location and details. They can be reached on {candidate.Phone} or at {candidate.Email}.</p> <br>" +
                    $"<p>I'm looking forward to discussing feedback following the call.</p> <br>" +
                    $"<p>If anything comes up and we need to re-arrange the call, please let me know.</p> <br>" +
                    $"{invite.AdditionalInfo}" +
                    $"<p>Best regards,</p>" +
                    $"</body></html>";

            } else if (selectedTemplate == "Other")
            {
                emailTemplate += $"<html><body>" +
                    $"<p>{clientFirstName}/{candidateFirstName},</p> <br>" +
                    $"<p>I'm pleased to confirm the following interview at {differentTimeZone} on {selectedDate}.</p> <br>" +
                    $"<p>Client: {client.Name} - {client.Company}</p>" +
                    $"<p>Candidate: {candidate.Name}</p>" +
                    $"<p>Date: {invite.Date}</p>" +
                    $"<p>Time: {differentTimeZone}</p> <br>" +
                    $"<p>{clientFirstName} - Please reach out to {candidateFirstName} to arrange the meeting location and details. They can be reached on {candidate.Phone} or at {candidate.Email}.</p> <br>" +
                    $"<p>I'm looking forward to discussing feedback following the call.</p> <br>" +
                    $"<p>If anything comes up and we need to re-arrange the call, please let me know.</p> <br>" +
                    $"{invite.AdditionalInfo}" +
                    $"<p>Best regards,</p>" +
                    $"</body></html>";
            }

            //Function that opens a calendar invite in the user's browser based on the information provided and to authenticate the user.
            PerformAuthentication(eventTitle, invite.Location, emailTemplate, client.Email, candidate.Email, selectedDateTime);
        }

        private async Task PerformAuthentication(string eventTitle, string location, string emailTemplate, string clientEmail, string candidateEmail, DateTime selectedDateTime)
        {
            try
            {
                //Sets up the necessary variables for authentication
                string clientId = "bfb8ba7b-9b57-4315-b865-764a4980d9d4";
                string clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET");
                string authority = "https://login.microsoftonline.com/common";
                string[] scopes = { "https://graph.microsoft.com/.default" };

                //Builds the confidential client application with the provided credentials
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                    .Create(clientId)
                    .WithClientSecret(clientSecret)
                    .WithAuthority(authority)
                    .Build();

                //Acquires an access token for the client application
                AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                //Sets the authorization header for the graph client using the access token
                graphClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);

                //Creates the URL to open the calendar invite creation page in Outlook
                string baseUrl = "https://outlook.office.com/calendar/action/compose";
                string subject = HttpUtility.UrlEncode(eventTitle);
                string body = HttpUtility.UrlEncode(emailTemplate);
                string startDateTime = HttpUtility.UrlEncode(selectedDateTime.ToString("o"));
                string endDateTime = HttpUtility.UrlEncode(selectedDateTime.AddMinutes(30).ToString("o"));
                string encodedLocation = HttpUtility.UrlEncode(location);
                string clientEmailAddress = HttpUtility.UrlEncode(clientEmail);
                string candidateEmailAddress = HttpUtility.UrlEncode(candidateEmail);

                //Combines the URL parameters to form the complete URL
                string url = $"{baseUrl}?subject={subject}&startdt={startDateTime}&enddt={endDateTime}&body={body}&location={encodedLocation}&to={clientEmailAddress};{candidateEmailAddress}";

                //Opens the URL in the default browser to allow the user to create the calendar invite
                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });

            }
            catch (System.Exception ex)
            {
             System.Windows.MessageBox.Show($"An error occurred while creating the calendar invite: {ex.Message}");
            }

            }
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

        private void LoadPerson_Click(object sender, RoutedEventArgs e)
        {
            string filePath = null;
            List<Client> clientsList = new List<Client>();
            List<Candidate> candidatesList = new List<Candidate>();

            if (sender == LoadClientButton)
            {
                filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\clients.json";
            }
            else if (sender == LoadCandidateButton)
            {
                filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\candidates.json";
            }

            try
            {
                // Reads the JSON data from the file
                string jsonData = File.ReadAllText(filePath);

                //Splits JSON data into individual lines so they can be processed individually
                string[] jsonLines = jsonData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                // Deserialize JSON data into list of clients or candidates
                if (sender == LoadClientButton)
                {
                    foreach (string json in jsonLines)
                    {
                        Client client = JsonConvert.DeserializeObject<Client>(json);
                        clientsList.Add(client);
                    }
                }
                else if (sender == LoadCandidateButton)
                {
                    foreach (string json in jsonLines)
                    {
                        Candidate candidate = JsonConvert.DeserializeObject<Candidate>(json);
                        candidatesList.Add(candidate);
                    }
                }

                // Creates an instance of the DataViewer window and displays the list of either candidates or clients
                DataViewer dataViewer = null;
                if (sender == LoadClientButton)
                {
                    dataViewer = new DataViewer(clientsList); // Show clients
                }
                else if (sender == LoadCandidateButton)
                {
                    dataViewer = new DataViewer(candidatesList); // Show candidates
                }
                dataViewer.ShowDialog();

                //Populates the textboxes/combobox with the selected person's data
                if (sender == LoadClientButton)
                {
                    ClientNameTextBox.Text = dataViewer.SelectedClientName;
                    ClientEmailTextBox.Text = dataViewer.SelectedClientEmail;
                    ClientCompanyTextBox.Text = dataViewer.SelectedClientCompany;
                    ClientComboBox.Text = dataViewer.SelectedClientTimeZone;
                }
                else if (sender == LoadCandidateButton)
                {
                    CandidateNameTextBox.Text = dataViewer.SelectedCandidateName;
                    CandidateEmailTextBox.Text = dataViewer.SelectedCandidateEmail;
                    CandidatePhoneTextBox.Text = dataViewer.SelectedCandidatePhone;
                    CandidateComboBox.Text = dataViewer.SelectedCandidateTimeZone;
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"An error occurred while loading the information: {ex.Message}");
            }
        }
        
        //Function that updates the selectedTemplate based on which template is selected
        private void TemplateComboBox_DropDownClosed(object sender, EventArgs e)
            {
        
                selectedTemplate = TemplateComboBox.Text;    

                if (selectedTemplate == "First stage phone call" || selectedTemplate == "Other"){
                location = CandidatePhoneTextBox.Text;

                } else if (selectedTemplate == "Teams interview"){
                location = "Microsoft Teams";
                }
                else if (selectedTemplate == "In-person interview"){
                location = selectedTemplate;
                }
            }

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
        //Updates the date/time of the appointment based on the user entry
        private void DateTimePicker_ValueChanged(object sender, EventArgs e)
            {
            Xceed.Wpf.Toolkit.DateTimePicker dateTimePicker = (Xceed.Wpf.Toolkit.DateTimePicker)sender;
            selectedDateTime = dateTimePicker.Value ?? DateTime.Now;
            selectedDate = selectedDateTime.ToString("dddd, MMMM dd", new CultureInfo("en-US"));
        }

    }
}
