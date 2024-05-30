using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;

namespace ReportSys.Pages.PageAccess0
{
    public class IndexModel : PageModel
    {

        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        public void OnGet()
        {
        }


        public async Task<IActionResult> OnPostAsync()
        {

            var employeeNumber = HttpContext.Session.GetString("EmployeeNumber");

            var employee = await _context.Employees.Include(e => e.WorkSchedule)
                .Include(e => e.Events).ThenInclude(s => s.EventType)
                .Include(e => e.Unavailabilitys).ThenInclude(s => s.UnavailabilityType)
                .FirstOrDefaultAsync(e => e.Id.ToString() == employeeNumber);

            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Employee");

                // ��������� ��������
                worksheet.Cells[1, 1].Value = "����";
                worksheet.Cells[1, 2].Value = "�����";
                worksheet.Cells[1, 3].Value = "�������";
                worksheet.Cells[1, 4].Value = "����������";
                worksheet.Cells[1, 5].Value = "���������� �� ���";
                worksheet.Cells[1, 8].Value = "�� ������ �������� �������";
                worksheet.Cells[1, 9].Value = "������ ������";
                worksheet.Cells[2, 5].Value = "c";
                worksheet.Cells[2, 6].Value = "��";
                worksheet.Cells[2, 7].Value = "���������";

                // ����������� ����� ��� ����������
                worksheet.Cells["A1:A2"].Merge = true; // ����������� �� ��������� ��� "����"
                worksheet.Cells["B1:B2"].Merge = true; // ����������� �� ��������� ��� "�����"
                worksheet.Cells["C1:C2"].Merge = true; // ����������� �� ��������� ��� "�������"
                worksheet.Cells["D1:D2"].Merge = true; // ����������� �� ��������� ��� "����������"
                worksheet.Cells["H1:H2"].Merge = true; // ����������� �� ��������� ��� "�� ������ �������� �������"
                worksheet.Cells["I1:I2"].Merge = true; // ����������� �� ��������� ��� "������ ������"
                worksheet.Cells["E1:G1"].Merge = true;



                // ������
                //for (int i = 0; i < employee.Events.Count; i++)
                //{
                //    worksheet.Cells[i + 3, 1].Value = employee.Events[i].Date;
                //    worksheet.Cells[i + 3, 2].Value = employee.Events[i].Time;
                //    worksheet.Cells[i + 3, 3].Value = employee.Events[i].EventType.Name;
                //    worksheet.Cells[i + 3, 4].Value = employee.Events[i].Territory;



                //}

                int rowIndex = 3; // �������� � ������� ������, ��� ��� ������ ��� ������ �����������

                // ����������� ������� �� ����
                var eventsGroupedByDate = employee.Events.GroupBy(e => e.Date);
                var count = employee.Events.Count;
                foreach (var eventGroup in eventsGroupedByDate)
                {
                    // �������� ������ ������� ��� ������� ����
                    var eventsForDate = eventGroup.ToList();

                    // �������� �� ������� ���������� ��� ������ ����
                    var unavailabilityForDate = employee.Unavailabilitys
                        .FirstOrDefault(u => u.Date == eventGroup.Key);

                    // ������������ ������ ������� ��� ������� ����
                    foreach (var eventItem in eventsForDate)
                    {
                        worksheet.Cells[rowIndex, 1].Value = eventItem.Date;
                        worksheet.Cells[rowIndex, 2].Value = eventItem.Time;
                        worksheet.Cells[rowIndex, 3].Value = eventItem.EventType.Name;
                        worksheet.Cells[rowIndex, 4].Value = eventItem.Territory;

                        if (unavailabilityForDate != null)
                        {
                            if (unavailabilityForDate.UnavailabilityType.Id == 4)
                            {
                                worksheet.Cells[rowIndex, 5].Value = unavailabilityForDate.UnavailabilityFrom.ToShortTimeString();
                                worksheet.Cells[rowIndex, 6].Value = unavailabilityForDate.UnavailabilityBefore.ToShortTimeString();
                                worksheet.Cells[rowIndex, 7].Value = unavailabilityForDate.Reason;
                            }
                            else
                            {
                                worksheet.Cells[rowIndex, 1].Value = eventItem.Date;
                                worksheet.Cells[rowIndex, 8].Value = unavailabilityForDate.UnavailabilityType.Name;
                            }
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, 5].Value = "-";
                            worksheet.Cells[rowIndex, 6].Value = "-";
                            worksheet.Cells[rowIndex, 7].Value = "-";
                        }

                        rowIndex++;
                    }

                }

                var str = employee.WorkSchedule.GetScheduleString();
                worksheet.Cells[3, 9].Value = str;
                worksheet.Cells[$"I3:I{count+2}"].Merge = true;

                package.Save();
            }
            stream.Position = 0;
            var fileName = "Employee.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(stream, contentType, fileName);
        }
    }
}
