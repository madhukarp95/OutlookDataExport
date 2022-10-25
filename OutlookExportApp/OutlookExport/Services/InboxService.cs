using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using OutlookExport.Models;
using Serilog.Core;
using System.Runtime.InteropServices;

namespace OutlookExport.Services
{
    public class InboxService : BaseService<InboxModel>
    {
        private readonly ConfigOptions _options;
        private readonly FolderCount _folderCountOptions;
        private readonly ILogger<InboxService> _logger;
        public InboxService(IOptions<ConfigOptions> options,
            IOptions<FolderCount> folderCountOptions,
            ILogger<InboxService> logger) : base(logger)
        {
            _options = options.Value;
            _folderCountOptions = folderCountOptions.Value;
            _logger = logger;
        }

        public override void CreateWorkSheet(MAPIFolder myItems, ref Worksheet worksheet)
        {
            _logger.LogInformation("Started Processing Inbox Items");
            UpdateHeaders(ref worksheet);

            if (myItems.Items.Count > 0)
            {
                int itemCount = GetItemsCount(myItems.Items.Count, _folderCountOptions.InboxItems);

                List<InboxModel> inboxModelList = new();

                try
                {
                    for (int j = 1; j <= itemCount; j++)
                    {
                        _logger.LogInformation(Logger.InformationLog, "Inbox", j, myItems.Items.Count);

                        if (myItems.Items[j] is MailItem outlookXcell)
                        {

                            // TODO: Need to decide to either Allow or Ignore, Currenly being ignored
                            if (outlookXcell.MessageClass == RECALL || outlookXcell.Sent == false)
                            {
                                Marshal.ReleaseComObject(outlookXcell);
                                continue;
                            }
                            
                            InboxModel inboxModel = new();

                            inboxModel.Subject = outlookXcell.Subject;
                            inboxModel.MailBody = outlookXcell.Body;
                            inboxModel.SenderMail = GetSenderMail(outlookXcell);
                            inboxModel.Recipients = GetRecepients(outlookXcell.Recipients);

                            inboxModel.Sensitivity = outlookXcell.Sensitivity.ToString();
                            inboxModel.Importance = outlookXcell.Importance.ToString();
                            inboxModel.Category = outlookXcell.Categories;

                            inboxModel.CreatedTime = outlookXcell.CreationTime;

                            inboxModelList.Add(inboxModel);

                            if (outlookXcell != null)
                            {
                                Marshal.ReleaseComObject(outlookXcell);
                            }
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
                finally
                {
                    if(myItems != null)
                    {
                        Marshal.ReleaseComObject(myItems);
                    }
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

        public override void UpdateRows(List<InboxModel> inboxModelList, ref Worksheet worksheet)
        {
            int xlrow = 2;

            foreach (InboxModel inboxModel in inboxModelList)
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
