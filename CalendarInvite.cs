using System;

namespace Sendout_Calendar_Invite_Project
{
    public class CalendarInvite
    {
       // public string EventTitle { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public Client Client { get; set; }
        public Candidate Candidate { get; set; }
        public string AdditionalInfo { get; set; }
    }
}
