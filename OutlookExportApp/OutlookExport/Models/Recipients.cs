using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExport.Models
{
    public class MailRecipients
    {
        public List<MailAddress> ToAddress { get; set; }

        public List<MailAddress> To_CC_Address { get; set; }

        public List<MailAddress> To_BCCAddress { get; set; }

        public MailRecipients()
        {
            ToAddress = new();
            To_BCCAddress = new();
            To_CC_Address = new();
        }
    }

    public class MailAddress
    {
        public string UserEmail { get; set; }

        public string Username { get; set; }

        public string EmailType { get; set; }
    }


}
