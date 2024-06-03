using System;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using ReportSys.DAL;
using ReportSys.DAL.Entities;

namespace ReportSys.Pages.Services
{
    public class ServicesPage : PageModel
    {

        public SelectList EmployeesSL { get; set; }
        public SelectList DepartmentsSL { get; set; }



        public async Task<List<int>> GetSubordinateDepartmentsAsync(ReportSysContext context, int departmentId)
        {
            var result = new List<int> { departmentId };

            var subDepartments = await context.Hierarchies
                .Where(h => h.UpperDepartmentId == departmentId)
                .Select(h => h.LowerDepartmentId)
                .ToListAsync();

            foreach (var subDepartmentId in subDepartments)
            {
                result.AddRange(await GetSubordinateDepartmentsAsync(context, subDepartmentId));
            }

            return result;
        }

        public async Task EmployeesFromDepartAsync(ReportSysContext context, Employee emp, object value = null)
        {
            var departmentIds = await GetSubordinateDepartmentsAsync(context, emp.DepartmentId);

            var query = context.Employees
                .Where(x => departmentIds.Contains(x.DepartmentId))
                .OrderBy(x => x.FirstName);

            var employees = await query.AsNoTracking().ToListAsync();

            // Отладочная информация
            Console.WriteLine($"Found {employees.Count} employees");

            EmployeesSL = new SelectList(employees, "Id", "FirstName", value);
        }

        public async Task DepartmentsFromDepartAsync(ReportSysContext context, int departmentId, object value = null)
        {
            var departmentIds = await GetSubordinateDepartmentsAsync(context, departmentId);

            var query = context.Departments
                .Where(d => departmentIds.Contains(d.Id))
                .OrderBy(d => d.Name);

            var departments = await query.AsNoTracking().ToListAsync();

            // Отладочная информация
            Console.WriteLine($"Found {departments.Count} departments");

            DepartmentsSL = new SelectList(departments, "Id", "Name", value);
        }

        public async Task<IActionResult> CreateXlsxFirst(ReportSysContext _context, List<string> employeeNumbers, DateOnly startDate, DateOnly endDate)
        {
            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                foreach (var employeeNumber in employeeNumbers)
                {
                    var employee = await _context.Employees
                       .Include(e => e.WorkSchedule)
                       .Include(e => e.Events).ThenInclude(s => s.EventType)
                       .Include(e => e.Unavailabilitys).ThenInclude(s => s.UnavailabilityType)
                       .FirstOrDefaultAsync(e => e.Id.ToString() == employeeNumber);

                    if (employee == null)
                    {
                        continue; // Пропускаем, если сотрудник не найден
                    }

                    var star_time = employee.WorkSchedule.Arrival;
                    var end_time = employee.WorkSchedule.Exit;

                    var worksheet = package.Workbook.Worksheets.Add(employeeNumber);

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
                    worksheet.Cells["A1:A2"].Merge = true;
                    worksheet.Cells["B1:B2"].Merge = true;
                    worksheet.Cells["C1:C2"].Merge = true;
                    worksheet.Cells["D1:D2"].Merge = true;
                    worksheet.Cells["H1:H2"].Merge = true;
                    worksheet.Cells["I1:I2"].Merge = true;
                    worksheet.Cells["E1:G1"].Merge = true;

                    int rowIndex = 3; // Начинаем с третьей строки, так как первые две заняты заголовками

                    // Проход по дням в выбранном промежутке, пропуская выходные
                    for (var date = startDate; date <= endDate; date = date.AddDays(1))
                    {
                        if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                        {
                            continue; // Пропускаем субботу и воскресенье
                        }

                        var eventsForDate = employee.Events.Where(e => e.Date == date).ToList();
                        var unavailabilityForDate = employee.Unavailabilitys
                            .FirstOrDefault(u => u.Date == date && u.EmployeeId == employee.Id);

                        if (eventsForDate.Count == 0 && unavailabilityForDate == null)
                        {
                            continue; // Пропускаем дни, которых нет
                        }

                        var startRow = rowIndex;
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
                                worksheet.Cells[rowIndex, 8].Value = unavailabilityForDate.UnavailabilityType.Name;
                            }
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, 5].Value = "-";
                            worksheet.Cells[rowIndex, 6].Value = "-";
                            worksheet.Cells[rowIndex, 7].Value = "-";
                        }
                        var firstEventType0 = eventsForDate.FirstOrDefault(e => e.EventType.Id == 1);
                        var lastEventType1 = eventsForDate.LastOrDefault(e => e.EventType.Id == 2);

                        foreach (var eventItem in eventsForDate)
                        {
                            worksheet.Cells[rowIndex, 1].Value = eventItem.Date.ToShortDateString();
                            worksheet.Cells[rowIndex, 2].Value = eventItem.Time;
                            if (eventItem.Time == firstEventType0.Time)
                            {
                                if ((star_time - eventItem.Time > TimeSpan.FromMinutes(3)) && eventItem.Time < star_time)
                                {
                                    // Устанавливаем цвет фона для ячейки
                                    worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }


                            }
                            if (eventItem.Time == lastEventType1.Time)
                            {
                                if ((eventItem.Time - end_time > TimeSpan.FromMinutes(3)) && eventItem.Time > end_time)
                                {
                                    // Устанавливаем цвет фона для ячейки
                                    worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                                }
                            }

                            //отсутствие
                            // c - worksheet.Cells[rowIndex, 5]
                            // по - worksheet.Cells[rowIndex, 6]

                            if (worksheet.Cells[rowIndex, 5].Value ==  "-" && worksheet.Cells[rowIndex, 5].Value == "-")
                            {
                                if (eventItem.Time == firstEventType0.Time)
                                {
                                    if ((star_time - eventItem.Time > TimeSpan.FromMinutes(3)) && eventItem.Time < star_time)
                                    {
                                        // Устанавливаем цвет фона для ячейки
                                        worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                                    }
                                }
                                if (eventItem.Time == lastEventType1.Time)
                                {
                                    if ((eventItem.Time - end_time > TimeSpan.FromMinutes(3)) && eventItem.Time > end_time)
                                    {
                                        // Устанавливаем цвет фона для ячейки
                                        worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                                    }
                                }
                            }
                            else
                            {

                            }
                                                    

                            //if (eventItem.Time == firstEventType0.Time && )
                            //{
                            //    if ((star_time - eventItem.Time > TimeSpan.FromMinutes(3)) && eventItem.Time < star_time)
                            //    {
                            //        // Устанавливаем цвет фона для ячейки
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                            //    }
                            //}

                            //if (eventItem.Time == lastEventType1.Time &&)
                            //{
                            //    if ((eventItem.Time - end_time > TimeSpan.FromMinutes(3)) && eventItem.Time > end_time)
                            //    {
                            //        // Устанавливаем цвет фона для ячейки
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                            //    }
                            //}

                            worksheet.Cells[rowIndex, 3].Value = eventItem.EventType.Name;
                            worksheet.Cells[rowIndex, 4].Value = eventItem.Territory;

                            rowIndex++;
                        }
                        
                        if (eventsForDate.Count > 0)
                        {
                            worksheet.Cells[$"E{startRow}:E{rowIndex - 1}"].Merge = true;
                            worksheet.Cells[$"F{startRow}:F{rowIndex - 1}"].Merge = true;
                            worksheet.Cells[$"G{startRow}:G{rowIndex - 1}"].Merge = true;
                        }


                        // Если нет событий для даты, все равно добавляем строку
                        if (eventsForDate.Count == 0)
                        {
                            worksheet.Cells[rowIndex, 1].Value = date.ToString("yyyy-MM-dd");
                            worksheet.Cells[rowIndex, 2].Value = "-";
                            worksheet.Cells[rowIndex, 3].Value = "-";
                            worksheet.Cells[rowIndex, 4].Value = "-";

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

                    if (rowIndex != 3)
                    {
                        worksheet.Cells[$"I3:I{rowIndex - 1}"].Merge = true;
                    }
                    
                }

                package.Save();
            }

            stream.Position = 0;
            var fileName = "Employees.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(stream, contentType, fileName);
        }

        public List<Department> FindTopLevelDepartments(List<int> departmentIds, ReportSysContext _context)
        {
            // Получаем все департаменты из списка с их иерархией
            var departments = _context.Departments
                .Include(d => d.UpperDepartments)
                .Include(d => d.LowerDepartments)
                .Where(d => departmentIds.Contains(d.Id))
                .ToList();


            var deps = _context.Hierarchies
                        .Where(e => departmentIds.Contains(e.UpperDepartmentId))
                        .ToList();




            // Создаем хэшсет для хранения всех департаментов, которые являются подчиненными
            var lowerDepartmentIds = new HashSet<int>();

            // Добавляем все подчиненные департаменты в хэшсет
            foreach (var department in departments)
            {
                foreach (var lower in department.LowerDepartments)
                {
                    lowerDepartmentIds.Add(lower.LowerDepartmentId);
                }
            }

            // Ищем все департаменты, которые не являются подчиненными ни одному другому департаменту
            var topLevelDepartments = new List<Department>();
            foreach (var department in departments)
            {
                if (!lowerDepartmentIds.Contains(department.Id))
                {
                    topLevelDepartments.Add(department); // Добавляем департамент высшего уровня
                }
            }

            return topLevelDepartments; // Возвращаем список департаментов высшего уровня
        }

    }

}


