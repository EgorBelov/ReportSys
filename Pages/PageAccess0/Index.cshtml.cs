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

                // Заголовки столбцов
                worksheet.Cells[1, 1].Value = "Дата";
                worksheet.Cells[1, 2].Value = "Время";
                worksheet.Cells[1, 3].Value = "Событие";
                worksheet.Cells[1, 4].Value = "Территория";
                worksheet.Cells[1, 5].Value = "Отсутствие по ЖМК";
                worksheet.Cells[1, 8].Value = "По табелю рабочего времени";
                worksheet.Cells[1, 9].Value = "Личный график";
                worksheet.Cells[2, 5].Value = "c";
                worksheet.Cells[2, 6].Value = "по";
                worksheet.Cells[2, 7].Value = "основание";

                // Объединение ячеек для заголовков
                worksheet.Cells["A1:A2"].Merge = true; // Объединение по вертикали для "Дата"
                worksheet.Cells["B1:B2"].Merge = true; // Объединение по вертикали для "Время"
                worksheet.Cells["C1:C2"].Merge = true; // Объединение по вертикали для "Событие"
                worksheet.Cells["D1:D2"].Merge = true; // Объединение по вертикали для "Территория"
                worksheet.Cells["H1:H2"].Merge = true; // Объединение по вертикали для "По табелю рабочего времени"
                worksheet.Cells["I1:I2"].Merge = true; // Объединение по вертикали для "Личный график"
                worksheet.Cells["E1:G1"].Merge = true;



                // Данные
                //for (int i = 0; i < employee.Events.Count; i++)
                //{
                //    worksheet.Cells[i + 3, 1].Value = employee.Events[i].Date;
                //    worksheet.Cells[i + 3, 2].Value = employee.Events[i].Time;
                //    worksheet.Cells[i + 3, 3].Value = employee.Events[i].EventType.Name;
                //    worksheet.Cells[i + 3, 4].Value = employee.Events[i].Territory;



                //}

                int rowIndex = 3; // Начинаем с третьей строки, так как первые две заняты заголовками

                // Группировка событий по дате
                var eventsGroupedByDate = employee.Events.GroupBy(e => e.Date);
                var count = employee.Events.Count;
                foreach (var eventGroup in eventsGroupedByDate)
                {
                    // Получаем список событий для текущей даты
                    var eventsForDate = eventGroup.ToList();

                    // Проверка на наличие отсутствий для данной даты
                    var unavailabilityForDate = employee.Unavailabilitys
                        .FirstOrDefault(u => u.Date == eventGroup.Key);

                    // Обрабатываем каждое событие для текущей даты
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
