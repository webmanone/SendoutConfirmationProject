using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Wpf.Toolkit;
using Xceed.Wpf.Toolkit.Primitives;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Globalization;
using Newtonsoft.Json;
using System.Data;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Identity.Client;

namespace Sendout_Calendar_Invite_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string selectedTemplate = "First stage phone call";
        private string selectedDate;
        private DateTime selectedDateTime;
        private TimeZoneInfo clientTimeZone;
        private TimeZoneInfo candidateTimeZone;
        private string clientTime;
        private string candidateTime;
       // private string dateString;
        private string startTimeString;
        private string clientTimeZoneString = "Eastern";
        private string candidateTimeZoneString = "Eastern";
        private string location = "";

        public MainWindow()
        {
            InitializeComponent();

            string clientId = "bfb8ba7b-9b57-4315-b865-764a4980d9d4";
            string authority = "https://login.microsoftonline.com/common";
            string redirectUri = "https://localhost/8080";
            IPublicClientApplication app = PublicClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(authority)
                .WithRedirectUri(redirectUri)
                .Build();

            DateTimePicker dateTimePicker = new DateTimePicker();
            dateTimePicker.Value = DateTime.Now;
        }
        private void Preview_Click(object sender, RoutedEventArgs e)
        {
            // Handle preview button click

            string eventTitle = null;
            //string startTime = selectedDateTime.Value.ToString("h:mm tt");
            //DateTime endTime = startTime.AddHours(1);
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
            //template, candidate time zone and client time zone should already be stored by the event handlers

            Client client = new Client
            {
                Name = clientName,
                Email = clientEmail,
                Company = clientCompany,
                TimeZone = clientTimeZoneString
            };

            // Create candidate object
            Candidate candidate = new Candidate
            {
                Name = candidateName,
                Email = candidateEmail,
                Phone = candidatePhone,
                TimeZone = candidateTimeZoneString
            };

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

            if (clientTimeZone == null)
            {
                clientTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                DateTimeOffset selectedDateTimeOffset = DaylightConvert(selectedDateTime, clientTimeZone);
                clientTimeZone = TimeZoneInfo.CreateCustomTimeZone(clientTimeZone.Id, selectedDateTimeOffset.Offset, clientTimeZone.DisplayName, clientTimeZone.StandardName);
                clientTime = ConvertTimeZone(selectedDateTime, clientTimeZone);
            }

            if (candidateTimeZone == null)
            {
                candidateTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                DateTimeOffset selectedDateTimeOffset = DaylightConvert(selectedDateTime, candidateTimeZone);
                candidateTimeZone = TimeZoneInfo.CreateCustomTimeZone(candidateTimeZone.Id, selectedDateTimeOffset.Offset, candidateTimeZone.DisplayName, candidateTimeZone.StandardName);
                candidateTime = ConvertTimeZone(selectedDateTime, candidateTimeZone);
            }

            eventTitle = $"{candidate.Name}/{client.Company} - {invite.EventType}";

            if (clientTimeZoneString != candidateTimeZoneString)
            {
                differentTimeZone += $"{clientTime} {clientTimeZoneString}/{candidateTime} {candidateTimeZoneString}";
            } else
            {
                differentTimeZone += $"{clientTime} {clientTimeZoneString}";
            }

            if (invite.AdditionalInfo != "") {
                invite.AdditionalInfo = $"{invite.AdditionalInfo} \n \n";
            }

            if (selectedTemplate == "First stage phone call")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following call at {differentTimeZone} on {selectedDate}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" + //will need to edit this to cater for if there are multiple clients
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"{clientFirstName} - Please call {candidateFirstName} on {candidate.Phone} at the arranged time. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo}" +
                    $"Best regards, \n";

                invite.Location = candidatePhone;

            } else if (selectedTemplate == "Teams interview")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following {invite.EventType} at {differentTimeZone} on {selectedDate}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" +
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"Please join the Teams meeting at the arranged time. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo}" +
                    $"Best regards, \n";
            } else if (selectedTemplate == "In-person interview")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following in-person interview at {differentTimeZone} on {selectedDate}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" +
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"{clientFirstName} - Please reach out to {candidateFirstName} to arrange the meeting location and details. They can be reached on {candidate.Phone} or at {candidate.Email}. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo}" +
                    $"Best regards, \n";
            } else if (selectedTemplate == "Other")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following interview at {differentTimeZone} on {selectedDate}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" + //will need to edit this to cater for if there are multiple clients
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"{clientFirstName} - Please call {candidateFirstName} on {candidate.Phone} at the arranged time. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo}" +
                    $"Best regards, \n";
            }

            // create a new appointment item
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.AppointmentItem appointment = outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);

            // set the properties of the appointment
            appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            Outlook.Recipient clientRecipient = appointment.Recipients.Add(client.Email);
            clientRecipient.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
            Outlook.Recipient candidateRecipient = appointment.Recipients.Add(candidate.Email);
            candidateRecipient.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
            appointment.Subject = eventTitle;
            appointment.Location = invite.Location;
            appointment.Body = emailTemplate;
            DateTime start = new DateTime(invite.StartTime.Year, invite.StartTime.Month, invite.StartTime.Day, invite.StartTime.Hour, invite.StartTime.Minute, 0);
            DateTime end = new DateTime(invite.StartTime.Year, invite.StartTime.Month, invite.StartTime.Day, invite.EndTime.Hour, invite.EndTime.Minute, 0);

            appointment.Start = start;
            appointment.End = end;

            appointment.Display(true); //need to edit as causes an error if calendar invite is already up
        }

            private void SaveClient_Click(object sender, RoutedEventArgs e)
            {
            // Create a new instance of the Client class with the entered information
            Client client = new Client
            {
                Id = Guid.NewGuid().ToString(),  // Generate a unique identifier for the client
                Name = ClientNameTextBox.Text,
                Email = ClientEmailTextBox.Text,
                Company = ClientCompanyTextBox.Text,
                TimeZone = clientTimeZoneString
            };

            // Define the file path to save the client information
            string filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\clients.json";

            try
            {
                // Serialize the new client to JSON
                string json = JsonConvert.SerializeObject(client);

                // Append the JSON string to the existing file
                File.AppendAllText(filePath, json + Environment.NewLine);

                // Show a success message
                System.Windows.MessageBox.Show("Client information saved successfully.");
            }
            catch (Exception ex)
            {
                // Handle any potential exceptions that occurred during the save operation
                System.Windows.MessageBox.Show($"An error occurred while saving the client information: {ex.Message}");
            }
        }

            private void LoadClient_Click(object sender, RoutedEventArgs e)
        {
            string filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\clients.json";

            try
            {
                // Read the JSON data from the file
                string jsonData = File.ReadAllText(filePath);

                string[] jsonLines = jsonData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                // Deserialize each JSON object and add it to the list of clients
                List<Client> clientsList = new List<Client>();
                foreach (string json in jsonLines)
                {
                    Client client = JsonConvert.DeserializeObject<Client>(json);
                    clientsList.Add(client);
                }

                DataViewer dataViewer = new DataViewer(clientsList);  // Create an instance of the DataViewer window
                // Set the ItemsSource of the DataGrid

                dataViewer.ShowDialog();

                ClientNameTextBox.Text = dataViewer.SelectedClientName;
                ClientEmailTextBox.Text = dataViewer.SelectedClientEmail;
                ClientCompanyTextBox.Text = dataViewer.SelectedClientCompany;
                ClientComboBox.Text = dataViewer.SelectedClientTimeZone;
            }
            catch (Exception ex)
            {
                // Handle any potential exceptions
                System.Windows.MessageBox.Show($"An error occurred while loading the clients: {ex.Message}");
            }
        }

            private void SaveCandidate_Click(object sender, RoutedEventArgs e)
            {
            // Create a new instance of the Client class with the entered information
            Candidate candidate = new Candidate
            {
                Id = Guid.NewGuid().ToString(),  // Generate a unique identifier for the client
                Name = CandidateNameTextBox.Text,
                Email = CandidateEmailTextBox.Text,
                Phone = CandidatePhoneTextBox.Text,
                TimeZone = candidateTimeZoneString
            };

            // Define the file path to save the client information
            string filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\candidates.json";

            try
            {
                // Serialize the new client to JSON
                string json = JsonConvert.SerializeObject(candidate);

                // Append the JSON string to the existing file
                File.AppendAllText(filePath, json + Environment.NewLine);

                // Show a success message
                System.Windows.MessageBox.Show("Candidate information saved successfully.");
            }
            catch (Exception ex)
            {
                // Handle any potential exceptions that occurred during the save operation
                System.Windows.MessageBox.Show($"An error occurred while saving the candidate information: {ex.Message}");
            }
        }

            private void LoadCandidate_Click(object sender, RoutedEventArgs e)
            {
            string filePath = @"C:\Users\lukem\source\repos\Sendout Calendar Invite Project\Data\candidates.json";

            try
            {
                // Read the JSON data from the file
                string jsonData = File.ReadAllText(filePath);

                string[] jsonLines = jsonData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                // Deserialize each JSON object and add it to the list of clients
                List<Candidate> candidatesList = new List<Candidate>();
                foreach (string json in jsonLines)
                {
                    Candidate candidate = JsonConvert.DeserializeObject<Candidate>(json);
                    candidatesList.Add(candidate);
                }

                DataViewer dataViewer = new DataViewer(candidatesList);  // Create an instance of the DataViewer window
                // Set the ItemsSource of the DataGrid

                dataViewer.ShowDialog();

                CandidateNameTextBox.Text = dataViewer.SelectedCandidateName;
                CandidateEmailTextBox.Text = dataViewer.SelectedCandidateEmail;
                CandidatePhoneTextBox.Text = dataViewer.SelectedCandidatePhone;
                CandidateComboBox.Text = dataViewer.SelectedCandidateTimeZone;
            }
            catch (Exception ex)
            {
                // Handle any potential exceptions
                System.Windows.MessageBox.Show($"An error occurred while loading the clients: {ex.Message}");
            }
        }

            private void Edit_Click(object sender, RoutedEventArgs e)
            {
                // Handle edit button click
            }

            private void AddClient_Click(object sender, RoutedEventArgs e)
            {
                // Add a new client to the list of clients
            }

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

            private void ClientComboBox_DropDownClosed(object sender, EventArgs e)
            {
                
                clientTimeZoneString = ClientComboBox.Text;
               
                if (clientTimeZoneString == "Eastern"){
                    clientTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                }
                else if (clientTimeZoneString == "Central"){
                    clientTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
                }
                else if (clientTimeZoneString == "Mountain"){
                    clientTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time");
                }
                else if (clientTimeZoneString == "Pacific"){
                    clientTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                }

            DateTimeOffset selectedDateTimeOffset = DaylightConvert(selectedDateTime, clientTimeZone);
            clientTimeZone = TimeZoneInfo.CreateCustomTimeZone(clientTimeZone.Id, selectedDateTimeOffset.Offset, clientTimeZone.DisplayName, clientTimeZone.StandardName);
            clientTime = ConvertTimeZone(selectedDateTime, clientTimeZone);
        }

            private void CandidateComboBox_DropDownClosed(object sender, EventArgs e)
            {
                candidateTimeZoneString = CandidateComboBox.Text;
            
                if (candidateTimeZoneString == "Eastern"){
                    candidateTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                }
                else if (candidateTimeZoneString == "Central"){
                    candidateTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
                }
                else if (candidateTimeZoneString == "Mountain"){
                    candidateTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Mountain Standard Time");
                }
                else if (candidateTimeZoneString == "Pacific"){
                    candidateTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                }

            DateTimeOffset selectedDateTimeOffset = DaylightConvert(selectedDateTime, candidateTimeZone);
            candidateTimeZone = TimeZoneInfo.CreateCustomTimeZone(candidateTimeZone.Id, selectedDateTimeOffset.Offset, candidateTimeZone.DisplayName, candidateTimeZone.StandardName);
            candidateTime = ConvertTimeZone(selectedDateTime, candidateTimeZone);
        }
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
        private string ConvertTimeZone(DateTime dateTime, TimeZoneInfo targetZone)
            {
            TimeZoneInfo localZone = TimeZoneInfo.Local;
            
            DateTimeOffset selectedDateTimeOffset = DaylightConvert(dateTime, localZone);
            localZone = TimeZoneInfo.CreateCustomTimeZone(localZone.Id, selectedDateTimeOffset.Offset, localZone.DisplayName, localZone.StandardName);
            DateTime convertedTime = TimeZoneInfo.ConvertTime(dateTime, localZone, targetZone);
            
            string formattedTime = convertedTime.ToString("h:mm tt", new CultureInfo("en-US"));
            return formattedTime;
            }

        private void DateTimePicker_ValueChanged(object sender, EventArgs e)
            {
            DateTimePicker dateTimePicker = (DateTimePicker)sender;
            selectedDateTime = dateTimePicker.Value ?? DateTime.Now;
            selectedDate = selectedDateTime.ToString("dddd, MMMM dd", new CultureInfo("en-US"));
        }

    }
}
