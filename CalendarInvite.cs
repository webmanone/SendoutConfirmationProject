using System;

namespace Sendout_Calendar_Invite_Project
{
    public class CalendarInvite
    {
        public string EventTitle { get; set; }
        public string EventType { get; set; }
        public string Location { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Date { get; set; }
        public Client Client { get; set; }
        public Candidate Candidate { get; set; }
        public string AdditionalInfo { get; set; }
    }
}
