using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MicrosoftOutlookAppointments.Models
{
    public class clsAppointment
    {
        public string Subject { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}