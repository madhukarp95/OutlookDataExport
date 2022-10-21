using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExport.Models
{
    public class CalendarModel : MailModel
    {
        public bool AllDayEvent { get; set; }

        public int Duration { get; set; }

        public bool IsRecurring { get; set; }

        public string Location { get; set; }

        public string Organizer { get; set; }

        public string RequiredAttendees { get; set; }

        public string OptionalAttendees { get; set; }
    }
}
