using Newtonsoft.Json;
using System;
using System.IO;

namespace Sendout_Calendar_Invite_Project
{
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


}
