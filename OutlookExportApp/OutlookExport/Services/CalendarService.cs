using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using OutlookExport.Models;
using Exception = System.Exception;

namespace OutlookExport.Services
{
    public class CalendarService : BaseService<CalendarModel>
    {
        private readonly ILogger<CalendarService> _logger;
        public CalendarService(ILogger<CalendarService> logger) : base(logger)
        {
            _logger = logger;
        }

        public override void CreateWorkSheet(MAPIFolder myItems, ref Worksheet worksheet)
        {
            _logger.LogInformation("Started Processing Calendar Items");

            UpdateHeaders(ref worksheet);

            if (myItems.Items.Count > 0)
            {
                List<CalendarModel> calendarModelList = new();

                try
                {
                    for (int j = 1; j <= myItems.Items.Count; j++)
                    {
                        IncrementLog("Calendar", j, myItems.Items.Count);

                        if (myItems.Items[j] is AppointmentItem)
                        {
                            var outlookXcell = ((AppointmentItem)myItems.Items[j]);

                            CalendarModel calendarModel = new();
                            calendarModel.Subject = outlookXcell.Subject;
                            calendarModel.MailBody = outlookXcell.Body;
                            calendarModel.SenderMail = GetOrganizerEmail(outlookXcell.GetOrganizer());
                            calendarModel.Recipients = GetRecepients(outlookXcell.Recipients);
                            calendarModel.RequiredAttendees = outlookXcell.RequiredAttendees;
                            calendarModel.OptionalAttendees = outlookXcell.OptionalAttendees;
                            calendarModel.AllDayEvent = outlookXcell.AllDayEvent;
                            calendarModel.Duration = outlookXcell.Duration;
                            calendarModel.IsRecurring = outlookXcell.IsRecurring;
                            calendarModel.Location = outlookXcell.Location;
                            calendarModel.CreatedTime = outlookXcell.CreationTime;
                            calendarModelList.Add(calendarModel);
                        }

                        if (j == 10) break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Sorry some error occurred");
                }

                _logger.LogInformation("Adding Calendar items to worksheet");

                UpdateRows(calendarModelList, ref worksheet);
            }
        }

        public override void UpdateHeaders(ref Worksheet worksheet)
        {
            int columnIndex, rowIndex;
            columnIndex = rowIndex = 1;

            worksheet.Cells[rowIndex, columnIndex++] = "Subject";
            worksheet.Cells[rowIndex, columnIndex++] = "Body";
            worksheet.Cells[rowIndex, columnIndex++] = "Organizer (Name)";
            worksheet.Cells[rowIndex, columnIndex++] = "Organizer (Address)";
            worksheet.Cells[rowIndex, columnIndex++] = "Organizer (Type)";
            worksheet.Cells[rowIndex, columnIndex++] = "To (Name)";
            worksheet.Cells[rowIndex, columnIndex++] = "To (Address)";
            //worksheet.Cells[rowIndex, columnIndex++] = "To (Type)";
            worksheet.Cells[rowIndex, columnIndex++] = "Required Attendees";
            worksheet.Cells[rowIndex, columnIndex++] = "Optional Attendees";
            worksheet.Cells[rowIndex, columnIndex++] = "All day event";
            worksheet.Cells[rowIndex, columnIndex++] = "Duration";
            worksheet.Cells[rowIndex, columnIndex++] = "Is Recurring";
            worksheet.Cells[rowIndex, columnIndex++] = "Location";
            worksheet.Cells[rowIndex, columnIndex++] = "Creation Date";
        }

        public override void UpdateRows(List<CalendarModel> modelList, ref Worksheet worksheet)
        {
            int rowIndex = 2;
            foreach (CalendarModel calendarModel in modelList)
            {
                int columnIndex = 1;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.Subject;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.MailBody;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.SenderMail.Username;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.SenderMail.UserEmail;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.SenderMail.EmailType;

                worksheet.Cells[rowIndex, columnIndex++] = string.Join(',', calendarModel.Recipients.ToAddress.Select(x => x.Username));
                worksheet.Cells[rowIndex, columnIndex++] = string.Join(',', calendarModel.Recipients.ToAddress.Select(x => x.UserEmail));

                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.RequiredAttendees;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.OptionalAttendees;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.AllDayEvent;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.Duration;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.IsRecurring;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.Location;
                worksheet.Cells[rowIndex, columnIndex++] = calendarModel.CreatedTime;
                rowIndex++;
            }
        }

        private MailAddress GetOrganizerEmail(AddressEntry addressEntry)
        {

            string? mailAddress;

            if (addressEntry.Type == "SMTP")
            {
                mailAddress = addressEntry.Address;
            }
            else
            {
                mailAddress = addressEntry.GetExchangeUser()?.PrimarySmtpAddress;
            }

            MailAddress mail = new()
            {
                Username = addressEntry.Name,
                EmailType = addressEntry.Type,
                UserEmail = String.IsNullOrEmpty(mailAddress) ? String.Empty : MaskEmail(mailAddress)
            };

            return mail;
        }
    }
}
