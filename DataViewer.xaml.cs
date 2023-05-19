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
        public DataViewer(List<Client> clients)
        {
            InitializeComponent();

            // Set the ItemsSource of the DataGrid to the passed-in clients
            ClientsList = clients;
            //dataViewer.ItemsSource = clients;
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
    }
}
