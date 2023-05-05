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

namespace Sendout_Calendar_Invite_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string selectedTemplate;
        private DateTime? selectedDateTime;
        private string clientTimeZone;
        private string candidateTimeZone;
        private string dateString;
        private string startTimeString;
        public MainWindow()
        {
            InitializeComponent();

        }
            private void Preview_Click(object sender, RoutedEventArgs e)
            {
            // Handle preview button click

            //string eventTitle = 
            //string startTime = selectedDateTime.Value.ToString("h:mm tt");
            //DateTime endTime = startTime.AddHours(1);
            string clientName = ClientNameTextBox.Text;
            string clientEmail = ClientEmailTextBox.Text;
            string clientCompany = ClientCompanyTextBox.Text;
            string candidateName = CandidateNameTextBox.Text;
            string candidateEmail = CandidateEmailTextBox.Text;
            string candidatePhone = CandidatePhoneTextBox.Text;
            string additionalInfo = AdditionalInfoTextBox.Text;
            string emailTemplate = "";
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
                // EventTitle = "Interview",
                Date = dateString,
                StartTime = selectedDateTime.Value,
                EndTime = selectedDateTime.Value.AddMinutes(30),
                Client = client,
                Candidate = candidate,
                AdditionalInfo = additionalInfo
            };

            if (selectedTemplate == "First stage phone call")
            {
                emailTemplate = $"{client.Name}/{candidate.Name}, I'm pleased to confirm the following initial phone call at {startTimeString}";

            } else if (selectedTemplate == "Teams Interview")
            {

            } else if (selectedTemplate == "In-person interview")
            {

            } else if (selectedTemplate == "Other")
            {

            }
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

            private void TemplateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                ComboBox comboBox = (ComboBox)sender;
                if (comboBox.SelectedItem != null)
                {
                    selectedTemplate = comboBox.SelectedItem.ToString();
                }
                else
                {
                    System.Windows.MessageBox.Show("Please select a template.");
                }
            }

            private void ClientComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                ComboBox comboBox = (ComboBox)sender;
                clientTimeZone = comboBox.SelectedItem.ToString();
            }

            private void CandidateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                ComboBox comboBox = (ComboBox)sender;
                candidateTimeZone = comboBox.SelectedItem.ToString();
        }
        
            private void DateTimePicker_SelectedDateTimeChanged(object sender, RoutedEventArgs e)
            {
                DateTime? selectedDateTime = ((Xceed.Wpf.Toolkit.DateTimePicker)sender).Value;

                DateTime selectedDate = (DateTime)selectedDateTime;
                dateString = selectedDate.ToString("dddd MMMM d");
                startTimeString = selectedDateTime.Value.ToString("h:mm tt");

        }

    }
}
