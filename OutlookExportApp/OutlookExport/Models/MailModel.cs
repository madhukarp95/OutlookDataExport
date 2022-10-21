using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExport.Models
{
    public abstract class MailModel
    {
        private string? _mailBody;
        public string Subject { get; set; }
        public string? MailBody
        {
            get
            {
                return _mailBody;
            }
            set
            {
                _mailBody = value?.Length > 1000 ? value.Substring(0, 1000) : value;
            }
        }

        public MailRecipients Recipients { get; set; }

        public MailAddress SenderMail { get; set; }
        public DateTime CreatedTime { get; set; }
        public string Importance { get; set; }
        public string Category { get; set; }
        public string Sensitivity { get; set; }

    }
}
