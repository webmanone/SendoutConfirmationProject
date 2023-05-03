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
        public MainWindow()
        {
            InitializeComponent();

        }
            private void Preview_Click(object sender, RoutedEventArgs e)
            {
            // Handle preview button click

            //string eventTitle = 
           // DateTime startTime = (DateTime)DateTimePicker.Value;
            //DateTime endTime = startTime.AddHours(1);
            string clientName = ClientNameTextBox.Text;
            string clientEmail = ClientEmailTextBox.Text;
            string clientCompany = ClientCompanyTextBox.Text;
            string candidateName = CandidateNameTextBox.Text;
            string candidateEmail = CandidateEmailTextBox.Text;
            string candidatePhone = CandidatePhoneTextBox.Text;
            string additionalInfo = AdditionalInfoTextBox.Text;
            //template, candidate time zone and client time zone should already be stored by the event handlers

            // Create a new calendar invite object with the input values
            CalendarInvite invite = new CalendarInvite
            {
               // EventTitle = eventTitle,
                //StartTime = startTime,
               // EndTime = endTime,
                ClientName = clientName,
                ClientEmail = clientEmail,
                ClientCompany = clientCompany,
                CandidateName = candidateName,
                CandidateEmail = candidateEmail,
                CandidatePhone = candidatePhone,
                AdditionalInfo = additionalInfo
            };
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
                // Handle selection change event
            }

            private void ClientComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                // Handle selection change event
            }

            private void CandidateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                // Handle selection change event
            }
        
        private void DateTimePicker_SelectedDateTimeChanged(object sender, RoutedEventArgs e)
            {
                DateTime? selectedDateTime = ((Xceed.Wpf.Toolkit.DateTimePicker)sender).Value;
                // Do something with the selected date/time
            }

    }
}
