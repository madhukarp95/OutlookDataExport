using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using OutlookExport.Models;
using System.Text.RegularExpressions;

namespace OutlookExport.Services
{
    public abstract class BaseService<T>
    {
        private readonly ILogger<BaseService<T>> _logger;
        
        public BaseService(ILogger<BaseService<T>> logger)
        {
            _logger = logger;
        }

        internal const string RECALL = "IPM.Outlook.Recall";
        internal const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public abstract void UpdateHeaders(ref Worksheet worksheet);

        public abstract void UpdateRows(List<T> modelList, ref Worksheet worksheet);

        public abstract void CreateWorkSheet(MAPIFolder myItems, ref Worksheet worksheet);

        internal MailAddress GetSenderMail(MailItem mailItem)
        {
            MailAddress senderMail = new();

            string? mailAddress;
            if (mailItem.SenderEmailType == "SMTP")
            {
                mailAddress = mailItem.SenderEmailAddress;
            }
            else
            {
                mailAddress = mailItem.Sender.GetExchangeUser()?.PrimarySmtpAddress;
            }

            if (!string.IsNullOrEmpty(mailAddress))
            {
                senderMail.UserEmail = MaskEmail(mailAddress);
            }
            senderMail.Username = mailItem.SenderName;
            senderMail.EmailType = mailItem.SenderEmailType;

            return senderMail;
        }

        public string MaskEmail(string emailAddress)
        {
            string pattern = @"(?<=[\w]{1})[\w-\._\+%]*(?=[\w]{1}@)";
            string result = Regex.Replace(emailAddress, pattern, m => new string('*', m.Length));
            return result;
        }

        internal MailRecipients GetRecepients(Recipients recipients)
        {
            MailRecipients recipientsList = new();

            if (recipients != null && recipients.Count > 0)
            {
                foreach (Recipient recipient in recipients)
                {
                    MailAddress mailAddress = new();

                    switch (recipient.Type)
                    {


                        case (int)OlMailRecipientType.olCC:
                            mailAddress.UserEmail = GetEmailAddress(recipient);
                            mailAddress.Username = recipient.Name;
                            recipientsList.To_CC_Address.Add(mailAddress);
                            break;
                        case (int)OlMailRecipientType.olBCC:
                            mailAddress.UserEmail = GetEmailAddress(recipient);
                            mailAddress.Username = recipient.Name;
                            recipientsList.To_BCCAddress.Add(mailAddress);
                            break;
                        default:
                            mailAddress.UserEmail = GetEmailAddress(recipient);
                            mailAddress.Username = recipient.Name;
                            recipientsList.ToAddress.Add(mailAddress);
                            break;
                    }
                }
            }
            return recipientsList;
        }

        private string GetEmailAddress(Recipient recipient)
        {
            PropertyAccessor pa = recipient.PropertyAccessor;
            string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();

            return MaskEmail(smtpAddress);
        }

        internal int GetItemsCount(int UserSelected, int TotalCount)
        {
            return Math.Min(UserSelected, TotalCount);
        }
    }
}
