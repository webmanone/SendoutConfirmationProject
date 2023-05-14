using System;
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
        public MainWindow()
        {
            InitializeComponent();
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
                TimeZone = clientTimeZone
            };

            // Create candidate object
            Candidate candidate = new Candidate
            {
                Name = candidateName,
                Email = candidateEmail,
                Phone = candidatePhone,
                TimeZone = candidateTimeZone
            };

            // Create calendar invite object using client and candidate objects
            CalendarInvite invite = new CalendarInvite
            {
                EventTitle = eventTitle,
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

            if (client.TimeZone != candidate.TimeZone)
            {
                differentTimeZone += $"{clientTime} {clientTimeZoneString}/{candidateTime} {candidateTimeZoneString}";
            } else
            {
                differentTimeZone += $"{clientTime} {clientTimeZoneString}";
            }


            if (selectedTemplate == "First stage phone call")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following {invite.EventType} at {differentTimeZone}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" + //will need to edit this to cater for if there are multiple clients
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"{clientFirstName} - Please call {candidateFirstName} on {candidate.Phone} at the arranged time. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo} \n \n" +
                    $"Best regards, \n";

            } else if (selectedTemplate == "Teams interview")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following {invite.EventType} at {differentTimeZone}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" +
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"Please join the Teams meeting at the arranged time \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo} \n \n" +
                    $"Best regards, \n";
            } else if (selectedTemplate == "In-person interview")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following {invite.EventType} at {differentTimeZone}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" +
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"{clientFirstName} - Please reach out to {candidateFirstName} to arrange the meeting location and details. They can be reached at {candidate.Phone} or {candidate.Email}. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo} \n \n" +
                    $"Best regards, \n";
            } else if (selectedTemplate == "Other")
            {
                emailTemplate += $"{clientFirstName}/{candidateFirstName}, \n \n" +
                    $"I'm pleased to confirm the following {invite.EventType} at {differentTimeZone}. \n \n" +
                    $"Client: {client.Name} - {client.Company} \n" + //will need to edit this to cater for if there are multiple clients
                    $"Candidate: {candidate.Name} \n" +
                    $"Date: {invite.Date} \n" +
                    $"Time: {differentTimeZone} \n \n" +
                    $"{clientFirstName} - Please call {candidateFirstName} on {candidate.Phone} at the arranged time. \n \n" +
                    $"I'm looking forward to discussing feedback following the call. \n \n" +
                    $"If anything comes up and we need to re-arrange the call, please let me know. \n \n" +
                    $"{invite.AdditionalInfo} \n \n" +
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
            appointment.Location = "Microsoft Teams";
            appointment.Body = emailTemplate;
            DateTime start = new DateTime(invite.StartTime.Year, invite.StartTime.Month, invite.StartTime.Day, invite.StartTime.Hour, invite.StartTime.Minute, 0);
            DateTime end = new DateTime(invite.StartTime.Year, invite.StartTime.Month, invite.StartTime.Day, invite.EndTime.Hour, invite.EndTime.Minute, 0);

            appointment.Start = start;
            appointment.End = end;

            appointment.Display(true); //need to edit as causes an error if calendar invite is already up
        }

            private void SaveClient_Click(object sender, RoutedEventArgs e)
            {
                // Handle save client button click
            }

            private void LoadClient_Click(object sender, RoutedEventArgs e)
            {
                // Handle load client button click
            }

            private void SaveCandidate_Click(object sender, RoutedEventArgs e)
            {
                // Handle save candidate button click
            }

            private void LoadCandidate_Click(object sender, RoutedEventArgs e)
            {
                // Handle load candidate button click
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
            
            string formattedTime = convertedTime.ToString("h:mm tt");
            return formattedTime;
            }

        private void DateTimePicker_ValueChanged(object sender, EventArgs e)
            {
           

            DateTimePicker dateTimePicker = (DateTimePicker)sender;

            selectedDateTime = dateTimePicker.Value ?? DateTime.Now;

            string selectedDate = selectedDateTime.ToString("dddd, dd MMMM");

            // DateTimePicker dateTimePicker = new DateTimePicker();
            //DateTime selectedDateTime = dateTimePicker.Value;
            // dateString = selectedDate.ToString("dddd MMMM d");
            //startTimeString = selectedDateTime.Value.ToString("h:mm tt");

        }

    }
}
