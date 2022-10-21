﻿using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using OutlookExport.Models;
using OutlookExport.Services;
using System.DirectoryServices;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Exception = System.Exception;

namespace OutlookExport
{
    // https://stackoverflow.com/questions/26816144/import-outlook-to-excel-using-c-sharp
    // https://www.codeproject.com/Articles/1060078/Extracting-Email-Addresses-from-Outlook-Mailboxes
    public class ProcessReport
    {
        private readonly ConfigOptions _options;
        private readonly InboxService _inboxService;
        private readonly CalendarService _calendarService;
        private readonly SentItemService _sentItemService;
        private ILogger<ProcessReport> _logger;

        public ProcessReport(IOptions<ConfigOptions> options,
            ILogger<ProcessReport> logger,
            InboxService inboxService,
            CalendarService calendarService,
            SentItemService sentItemService)
        {
            _options = options.Value;
            _inboxService = inboxService;
            _calendarService = calendarService;
            _sentItemService = sentItemService;
            _logger = logger;
        }

        public void Download_data()
        {
            _logger.LogInformation("Export Process Started");

            Microsoft.Office.Interop.Excel.Application app = null;
            Workbook workbook = null;

            try
            {
                Microsoft.Office.Interop.Outlook.Application myApp = new();
                NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
                MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                MAPIFolder myCalendar = mapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

                MAPIFolder mySentItems = mapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);

                app = new();
                workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                workbook.Worksheets.Add();
                workbook.Worksheets.Add();


                Worksheet InboxWorksheet = (Worksheet)workbook.Worksheets[1];
                InboxWorksheet.Name = "Inbox";


                _inboxService.CreateWorkSheet(myInbox, ref InboxWorksheet);

                Worksheet calendarWorksheet = (Worksheet)workbook.Worksheets[2];
                calendarWorksheet.Name = "Calendar";

                _calendarService.CreateWorkSheet(myCalendar, ref calendarWorksheet);

                Worksheet SentItemsWorksheet = (Worksheet)workbook.Worksheets[3];
                SentItemsWorksheet.Name = "Sent Items";

                _sentItemService.CreateWorkSheet(mySentItems, ref SentItemsWorksheet);

                SaveReport(workbook);


                for (int i = 1; i <= workbook.Sheets.Count; i++)
                {

                    Worksheet sheet = workbook.Sheets.Item[i];
                    Console.WriteLine(sheet.Name);
                    if (sheet != null) Marshal.ReleaseComObject(sheet);
                }

                workbook.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Sorry some error occurred");
                throw;
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (app != null) Marshal.ReleaseComObject(app);
            }

            return;
        }

        private void SaveReport(Workbook workbook)
        {
            string sTemplatePath = AppDomain.CurrentDomain.BaseDirectory;
            string reportDateTime = DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss");
            string workbookName = Path.Combine(sTemplatePath, "Outlook_Dataset_" + reportDateTime + ".xlsx");
            workbook.SaveAs(workbookName);
        }


        private Worksheet RemoveUnwantedColumns(Worksheet worksheet, List<string> _options)
        {
            int columnCount = worksheet.UsedRange.Columns.Count;

            for (int i = 1; i < columnCount; i++)
            {
                if (!_options.Contains(worksheet.Cells[1, i].Value2))
                {
                    string columnName = worksheet.Columns[i].Address;

                    Regex reg = new Regex(@"(\$)(\w*):");
                    if (reg.IsMatch(columnName))
                    {
                        Match match = reg.Match(columnName);
                        string test = match.Groups[2].Value;

                        Microsoft.Office.Interop.Excel.Range objRange =
                            (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range($"{test}1", Missing.Value);
                        objRange.EntireColumn.Delete(Missing.Value);
                        i--;
                    }
                }
            }

            return worksheet;
        }
    }
}