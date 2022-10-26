using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using OutlookExport.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExport.Services
{
    public class SentItemService : BaseService<SentMailModel>
    {
        private readonly ILogger<SentItemService> _logger;
        private readonly FolderCount _folderCountOptions;
        public SentItemService(ILogger<SentItemService> logger, IOptions<FolderCount> folderCountOptions) : base(logger)
        {
            _logger = logger;
            _folderCountOptions = folderCountOptions.Value;
        }
        public override void CreateWorkSheet(MAPIFolder myItems, ref Worksheet worksheet)
        {
            _logger.LogInformation("Started Processing SentMail Items");
            UpdateHeaders(ref worksheet);

            if (myItems.Items.Count > 0)
            {
                List<SentMailModel> sentMailModelList = new();
                int itemCount = GetItemsCount(myItems.Items.Count, _folderCountOptions.SentItems);
                try
                {
                    for (int j = 1; j <= itemCount; j++)
                    {
                        _logger.LogInformation(Logger.InformationLog, "SentMail", j, myItems.Items.Count);

                        if (myItems.Items[j] is MailItem outlookXcell)
                        {
                            if (outlookXcell.MessageClass == RECALL || outlookXcell.Sent == false)
                            {
                                Marshal.ReleaseComObject(outlookXcell);
                                continue;
                            }

                            SentMailModel sentMailModel = new();

                            sentMailModel.Subject = outlookXcell.Subject;
                            sentMailModel.SenderMail = GetSenderMail(outlookXcell);
                            sentMailModel.Recipients = GetRecepients(outlookXcell.Recipients);
                            sentMailModel.Sensitivity = outlookXcell.Sensitivity.ToString();
                            sentMailModel.Importance = outlookXcell.Importance.ToString();
                            sentMailModel.Category = outlookXcell.Categories;
                            sentMailModel.CreatedTime = outlookXcell.CreationTime;
                            sentMailModel.MailBody = outlookXcell.Body;

                            sentMailModelList.Add(sentMailModel);

                            if (outlookXcell != null)
                            {
                                Marshal.ReleaseComObject(outlookXcell);
                            }
                        }
                    }

                    _logger.LogInformation("Adding Sent mail items to worksheet");

                    UpdateRows(sentMailModelList, ref worksheet);

                    //worksheet = RemoveUnwantedColumns(worksheet, _options.SentItemColumns);
                }
                catch (System.Exception ex)
                {
                    _logger.LogError(ex, "Sorry some error occurred");
                }
                finally
                {
                    if (myItems != null)
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
            worksheet.Cells[1, columnIndex++] = "Created Date";
            worksheet.Cells[1, columnIndex++] = "Category";
            worksheet.Cells[1, columnIndex++] = "Sensitivity";
            worksheet.Cells[1, columnIndex++] = "Importance";
        }

        public override void UpdateRows(List<SentMailModel> modelList, ref Worksheet worksheet)
        {
            int xlrow = 2;
            foreach (SentMailModel sentMailModel in modelList)
            {
                int i = 1;
                worksheet.Cells[xlrow, i++] = sentMailModel.Subject;
                worksheet.Cells[xlrow, i++] = sentMailModel.MailBody;
                worksheet.Cells[xlrow, i++] = sentMailModel.SenderMail.Username;
                worksheet.Cells[xlrow, i++] = sentMailModel.SenderMail.UserEmail;
                worksheet.Cells[xlrow, i++] = sentMailModel.SenderMail.EmailType;
                worksheet.Cells[xlrow, i++] = string.Join(',', sentMailModel.Recipients.ToAddress.Select(x => x.Username));
                worksheet.Cells[xlrow, i++] = string.Join(',', sentMailModel.Recipients.ToAddress.Select(x => x.UserEmail));
                worksheet.Cells[xlrow, i++] = string.Join(',', sentMailModel.Recipients.To_CC_Address.Select(x => x.Username));
                worksheet.Cells[xlrow, i++] = string.Join(',', sentMailModel.Recipients.To_CC_Address.Select(x => x.UserEmail));
                worksheet.Cells[xlrow, i++] = string.Join(',', sentMailModel.Recipients.To_BCCAddress.Select(x => x.Username));
                worksheet.Cells[xlrow, i++] = string.Join(',', sentMailModel.Recipients.To_BCCAddress.Select(x => x.UserEmail));
                worksheet.Cells[xlrow, i++] = sentMailModel.CreatedTime;
                worksheet.Cells[xlrow, i++] = sentMailModel.Category;
                worksheet.Cells[xlrow, i++] = sentMailModel.SenderMail;
                worksheet.Cells[xlrow, i++] = sentMailModel.Importance;
                xlrow++;
            }
        }
    }
}
