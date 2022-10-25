using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExport
{
    public class ConfigOptions
    {
        public List<string> InboxColumns { get; set; }

        public List<string> CalendarColumns { get; set; }

        public List<string> SentItemColumns { get; set; }
    }

    public class FolderCount
    {
        public int InboxItems { get; set; }
        public int CalendarItems { get; set; }
        public int SentItems { get; set; }
    }
}
