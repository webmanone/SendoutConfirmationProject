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
using System.Windows.Shapes;

namespace Sendout_Calendar_Invite_Project
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class DataViewer : Window
    {
        public List<Client> ClientsList { get; set; }
        public string SelectedClientName { get; set; }
        public string SelectedClientEmail { get; set; }
        public string SelectedClientCompany { get; set; }
        public string SelectedClientTimeZone { get; set; }
        public string SelectedCandidateName { get; set; }
        public string SelectedCandidateEmail { get; set; }
        public string SelectedCandidatePhone { get; set; }
        public string SelectedCandidateTimeZone { get; set; }
        public DataViewer(List<Client> clients)
        {
            InitializeComponent();

            // Set the ItemsSource of the DataGrid to the passed-in clients
            ClientsList = clients;

            DataContext = this;
        }

        public DataViewer(List<Candidate> candidates)
        {
            InitializeComponent();

            // Set the ItemsSource of the DataGrid to the passed-in candidates
            dataViewer.ItemsSource = candidates;
        }
        public DataViewer()
        {
            InitializeComponent();
        }

        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Get the selected item from the DataGrid
            var selectedItem = dataViewer.SelectedItem;

            // Check if an item is selected
            if (selectedItem != null)
            {
                // Retrieve the selected JSON object
                if (selectedItem is Client selectedClient)
                {
                    // Populate the form fields with the selected client's data
                    SelectedClientName = selectedClient.Name;
                    SelectedClientEmail = selectedClient.Email;
                    SelectedClientCompany = selectedClient.Company;
                    SelectedClientTimeZone = selectedClient.TimeZone;
                }
                else if (selectedItem is Candidate selectedCandidate)
                {
                    SelectedCandidateName = selectedCandidate.Name;
                    SelectedCandidateEmail = selectedCandidate.Email;
                    SelectedCandidatePhone = selectedCandidate.Phone;
                    SelectedCandidateTimeZone = selectedCandidate.TimeZone;
                }
            }
            Close();
        }
    }
}
