using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using OutlookExport.Models;
using Serilog.Core;

namespace OutlookExport.Services
{
    public class InboxService : BaseService<SentMailModel>
    {
        private readonly ConfigOptions _options;
        private readonly ILogger<InboxService> _logger;
        public InboxService(IOptions<ConfigOptions> options,
            ILogger<InboxService> logger) : base(logger)
        {
            _options = options.Value;
            _logger = logger;
        }

        public override void CreateWorkSheet(MAPIFolder myItems, ref Worksheet worksheet)
        {
            _logger.LogInformation("Started Processing Inbox Items");
            UpdateHeaders(ref worksheet);

            if (myItems.Items.Count > 0)
            {
                List<SentMailModel> inboxModelList = new();

                try
                {
                    for (int j = 1; j <= myItems.Items.Count; j++)
                    {
                        IncrementLog("Inbox", j, myItems.Items.Count);

                        if (myItems.Items[j] is MailItem)
                        {
                            var outlookXcell = ((MailItem)myItems.Items[j]);

                            if (outlookXcell.MessageClass == RECALL || outlookXcell.Sent == false)
                                continue;

                            SentMailModel inboxModel = new();

                            inboxModel.Subject = outlookXcell.Subject;
                            inboxModel.MailBody = outlookXcell.Body;
                            inboxModel.SenderMail = GetSenderMail(outlookXcell);
                            inboxModel.Recipients = GetRecepients(outlookXcell.Recipients);

                            inboxModel.Sensitivity = outlookXcell.Sensitivity.ToString();
                            inboxModel.Importance = outlookXcell.Importance.ToString();
                            inboxModel.Category = outlookXcell.Categories;

                            inboxModel.CreatedTime = outlookXcell.CreationTime;

                            inboxModelList.Add(inboxModel);
                        }
                    }

                    _logger.LogInformation("Adding Inbox items to worksheet");

                    UpdateRows(inboxModelList, ref worksheet);

                    //worksheet = RemoveUnwantedColumns(worksheet, _options.SentItemColumns);
                }
                catch (System.Exception ex)
                {
                    _logger.LogError(ex, "Sorry some error occurred");
                }
            }
        }

        public override void UpdateHeaders(ref Worksheet worksheet)
        {
            int columnIndex = 1;

            worksheet.Cells[1, columnIndex++] = "Subject";
            worksheet.Cells[1, columnIndex++] = "Body";
            worksheet.Cells[1, columnIndex++] = "From (Name)";
            worksheet.Cells[1, columnIndex++] = "From (Address)";
            worksheet.Cells[1, columnIndex++] = "From (Type)";
            worksheet.Cells[1, columnIndex++] = "To (Name)";
            worksheet.Cells[1, columnIndex++] = "To (Address)";
            worksheet.Cells[1, columnIndex++] = "CC (Name)";
            worksheet.Cells[1, columnIndex++] = "CC (Address)";
            worksheet.Cells[1, columnIndex++] = "BCC (Name)";
            worksheet.Cells[1, columnIndex++] = "BCC (Address)";
            worksheet.Cells[1, columnIndex++] = "Category";
            worksheet.Cells[1, columnIndex++] = "Sensitivity";
            worksheet.Cells[1, columnIndex++] = "Importance";
            worksheet.Cells[1, columnIndex++] = "CreatedTime";
        }

        public override void UpdateRows(List<SentMailModel> inboxModelList, ref Worksheet worksheet)
        {
            int xlrow = 2;

            foreach (SentMailModel inboxModel in inboxModelList)
            {
                int i = 1;
                worksheet.Cells[xlrow, i++] = inboxModel.Subject;
                worksheet.Cells[xlrow, i++] = inboxModel.MailBody;
                worksheet.Cells[xlrow, i++] = inboxModel.SenderMail.Username;
                worksheet.Cells[xlrow, i++] = inboxModel.SenderMail.UserEmail;
                worksheet.Cells[xlrow, i++] = inboxModel.SenderMail.EmailType;
                worksheet.Cells[xlrow, i++] = string.Join(',', inboxModel.Recipients.ToAddress.Select(x => x.Username));
                worksheet.Cells[xlrow, i++] = string.Join(',', inboxModel.Recipients.ToAddress.Select(x => x.UserEmail));
                worksheet.Cells[xlrow, i++] = string.Join(',', inboxModel.Recipients.To_CC_Address.Select(x => x.Username));
                worksheet.Cells[xlrow, i++] = string.Join(',', inboxModel.Recipients.To_CC_Address.Select(x => x.UserEmail));
                worksheet.Cells[xlrow, i++] = string.Join(',', inboxModel.Recipients.To_BCCAddress.Select(x => x.Username));
                worksheet.Cells[xlrow, i++] = string.Join(',', inboxModel.Recipients.To_BCCAddress.Select(x => x.UserEmail));
                worksheet.Cells[xlrow, i++] = inboxModel.Category;
                worksheet.Cells[xlrow, i++] = inboxModel.SenderMail;
                worksheet.Cells[xlrow, i++] = inboxModel.Importance;
                worksheet.Cells[xlrow, i++] = inboxModel.CreatedTime;
                xlrow++;
            }
        }
    }
}
