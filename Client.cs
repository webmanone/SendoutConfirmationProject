﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sendout_Calendar_Invite_Project
{
    public class Client
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Company { get; set; }
        public TimeZoneInfo TimeZone { get; set; }
    }
}