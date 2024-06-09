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
using System.Globalization;
using static OfficeOpenXml.ExcelErrorValue;

namespace ReportSys.Pages.Services
{
    public class ServicesPage : PageModel
    {

        public SelectList EmployeesSL { get; set; }
        public SelectList DepartmentsSL { get; set; }

        public SelectList AllEmployeesSL { get; set; }



        public async Task GetAllEmployeesAsync(ReportSysContext context, object value = null)
        {
            var query = context.Employees
                .OrderBy(x => x.FirstName)
                .Select(x => new
                {
                    x.Id,
                    DisplayText = x.Id + " " + x.FirstName + " " + x.SecondName + " " + x.LastName
                });

            var employees = await query.AsNoTracking().ToListAsync();

            AllEmployeesSL = new SelectList(employees, "Id", "DisplayText", value);
        }


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
                .OrderBy(x => x.FirstName)
                 .Select(x => new
                  {
                      x.Id,
                      DisplayText = x.FirstName + " " + x.SecondName + " " + x.LastName
                  });

            var employees = await query.AsNoTracking().ToListAsync();

            // Отладочная информация
            Console.WriteLine($"Found {employees.Count} employees");

            EmployeesSL = new SelectList(employees, "Id", "DisplayText", value);
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

                    var worksheet = package.Workbook.Worksheets.Add(employee.FirstName);

                    worksheet.Cells[1, 1].Value = $"Сведения по событиям доступа с {startDate.ToString("dd-MM-yyyy")} по {endDate.ToString("dd-MM-yyyy")} по {employee.FirstName + " " + employee.SecondName + " " + employee.LastName}";
                    worksheet.Cells[2, 1].Value = $"Дата составления: {DateOnly.FromDateTime(DateTime.Now).ToString("dd-MM-yyyy")} {TimeOnly.FromDateTime(DateTime.Now).ToString("HH:mm:ss")}";
                    // Заголовки столбцов
                    worksheet.Cells[3, 1].Value = "Дата";
                    worksheet.Cells[3, 2].Value = "Время";
                    worksheet.Cells[3, 3].Value = "Событие";
                    worksheet.Cells[3, 4].Value = "Территория";
                    worksheet.Cells[3, 5].Value = "Отсутствие по ЖМК";
                    worksheet.Cells[3, 8].Value = "По табелю рабочего времени";
                    worksheet.Cells[3, 9].Value = "Личный график";
                    worksheet.Cells[4, 5].Value = "c";
                    worksheet.Cells[4, 6].Value = "по";
                    worksheet.Cells[4, 7].Value = "основание";

                    // Объединение ячеек для заголовков
                    worksheet.Cells["A3:A4"].Merge = true;
                    worksheet.Cells["B3:B4"].Merge = true;
                    worksheet.Cells["C3:C4"].Merge = true;
                    worksheet.Cells["D3:D4"].Merge = true;
                    worksheet.Cells["H3:H4"].Merge = true;
                    worksheet.Cells["I3:I4"].Merge = true;
                    worksheet.Cells["E3:G3"].Merge = true;

                    // Форматирование ячеек заголовков
                    worksheet.Cells["A3:I4"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A3:I4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // Добавление границ к заголовкам
                    worksheet.Cells["A3:I4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A3:I4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A3:I4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A3:I4"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                    worksheet.Column(1).Width = 15;
                    worksheet.Column(2).Width = 15;
                    worksheet.Column(3).Width = 15;
                    worksheet.Column(4).Width = 25;
                    worksheet.Column(5).Width = 15;
                    worksheet.Column(6).Width = 15;
                    worksheet.Column(7).Width = 20;
                    worksheet.Column(8).Width = 20;
                    worksheet.Column(8).Style.WrapText = true;
                    worksheet.Column(9).Width = 25;

                    int rowIndex = 5; // Начинаем с третьей строки, так как первые две заняты заголовками

                    // Проход по дням в выбранном промежутке, пропуская выходные
                    for (var date = startDate; date <= endDate; date = date.AddDays(1))
                    {
                        if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                        {
                            continue; // Пропускаем субботу и воскресенье
                        }

                        var eventsForDate = employee.Events.Where(e => e.Date == date).OrderBy(e => e.Time).ToList();
                        var unavailabilityForDate = employee.Unavailabilitys
                            .FirstOrDefault(u => u.Date == date && u.EmployeeId == employee.Id);

                        if (eventsForDate.Count == 0 && unavailabilityForDate == null)
                        {
                            //continue; // пропускаем дни, которых нет
                            colorCell(worksheet, rowIndex, Color.SandyBrown);
                        }

                        var startRow = rowIndex;
                        if (unavailabilityForDate != null)
                        {
                            if (unavailabilityForDate.UnavailabilityType.Id == 4)
                            {
                                worksheet.Cells[rowIndex, 5].Value = unavailabilityForDate.UnavailabilityFrom.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[rowIndex, 6].Value = unavailabilityForDate.UnavailabilityBefore.ToString(@"hh\:mm\:ss");
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
                            worksheet.Cells[rowIndex, 1].Value = eventItem.Date.ToString("dd-MM-yyyy");
                            worksheet.Cells[rowIndex, 2].Value = eventItem.Time.ToString("HH:mm:ss");

                            if (unavailabilityForDate != null)
                            {
                                if (unavailabilityForDate.UnavailabilityType.Id == 4)
                                {
                                    worksheet.Cells[rowIndex, 5].Value = unavailabilityForDate.UnavailabilityFrom.ToShortTimeString();
                                    worksheet.Cells[rowIndex, 6].Value = unavailabilityForDate.UnavailabilityBefore.ToShortTimeString();
                                    worksheet.Cells[rowIndex, 7].Value = unavailabilityForDate.Reason;
                                    //worksheet.Cells[rowIndex, 8].Value = unavailabilityForDate.UnavailabilityType.Name;
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
                            if (worksheet.Cells[rowIndex, 8].Value != null)
                            {
                                colorCell(worksheet, rowIndex, Color.SkyBlue);
                            }
                            else
                            {
                                if((star_time - eventItem.Time >= TimeSpan.FromMinutes(3) && eventItem.Time < star_time) || (eventItem.Time - end_time >= TimeSpan.FromMinutes(3) && eventItem.Time > end_time))
                                {
                                    colorCell(worksheet, rowIndex, Color.LightGreen);
                                }
                                else
                                {
                                    if (worksheet.Cells[rowIndex, 5].Value != "-" && worksheet.Cells[rowIndex, 6].Value != "-" && worksheet.Cells[rowIndex, 5].Value != null && worksheet.Cells[rowIndex, 6].Value != null)
                                    {
                                        if (toTimeOnly(worksheet.Cells[rowIndex, 5].Value.ToString()) <= eventItem.Time && eventItem.Time <= toTimeOnly(worksheet.Cells[rowIndex, 6].Value.ToString()))
                                        {
                                            colorCell(worksheet, rowIndex, Color.Yellow);
                                        }
                                        else if ((eventItem.Time - star_time >= TimeSpan.FromMinutes(3) && eventItem.Time > star_time) && (end_time - eventItem.Time >= TimeSpan.FromMinutes(3) && eventItem.Time < end_time) && (eventItem.Time <= employee.WorkSchedule.LunchStart || eventItem.Time >= employee.WorkSchedule.LunchEnd))
                                        {
                                            colorCell(worksheet, rowIndex, Color.SandyBrown);
                                        }
                                    }
                                    else if ( (eventItem.Time - star_time >= TimeSpan.FromMinutes(3) && eventItem.Time > star_time) && (end_time - eventItem.Time >= TimeSpan.FromMinutes(3) && eventItem.Time < end_time) && (eventItem.Time <= employee.WorkSchedule.LunchStart || eventItem.Time >= employee.WorkSchedule.LunchEnd) )
                                    {
                                        colorCell(worksheet, rowIndex, Color.SandyBrown);
                                    }
                                }
                            }


                            //else
                            //{
                            //    if (eventItem.Time == firstEventType0.Time)
                            //    {
                            //        if ((star_time - eventItem.Time > TimeSpan.FromMinutes(3)) && eventItem.Time < star_time)
                            //        {
                            //            colorCell(worksheet, rowIndex, Color.Green);
                            //        }
                            //        else
                            //        {
                            //            if (worksheet.Cells[rowIndex, 5].Value != " - " && worksheet.Cells[rowIndex, 6].Value != "-" && worksheet.Cells[rowIndex, 5].Value != null && worksheet.Cells[rowIndex, 6].Value != null)
                            //            {
                            //                if (toTimeOnly(worksheet.Cells[rowIndex, 5].Value.ToString()) <= eventItem.Time && eventItem.Time <= toTimeOnly(worksheet.Cells[rowIndex, 6].Value.ToString()))
                            //                {
                            //                    colorCell(worksheet, rowIndex, Color.Yellow);
                            //                }
                            //            }
                            //            else
                            //            {
                            //                colorCell(worksheet, rowIndex, Color.Orange);
                            //            }
                            //        }
                            //    }
                            //    else
                            //    {
                            //        if (worksheet.Cells[rowIndex, 5].Value != "-" && worksheet.Cells[rowIndex, 6].Value != "-" && worksheet.Cells[rowIndex, 5].Value != null && worksheet.Cells[rowIndex, 6].Value != null)
                            //        {
                                        
                            //            if (toTimeOnly(worksheet.Cells[rowIndex, 5].Value.ToString()) <= eventItem.Time && eventItem.Time <= toTimeOnly(worksheet.Cells[rowIndex, 6].Value.ToString()))
                            //            {
                            //                colorCell(worksheet, rowIndex, Color.Yellow);
                            //            }
                                        
                            //        }
                            //        else
                            //        {
                            //            if (eventItem.Time > star_time && (eventItem.Time <= employee.WorkSchedule.LunchStart || eventItem.Time >= employee.WorkSchedule.LunchEnd) && eventItem.Time <= end_time)
                            //            {
                            //                colorCell(worksheet, rowIndex, Color.Orange);
                            //            }
                            //            if (eventItem.Time < star_time || eventItem.Time > end_time)
                            //            {
                            //                colorCell(worksheet, rowIndex, Color.Green);
                            //            }
                            //        }
                            //    }
                            //}
                           

                            //if (eventItem.Time == firstEventType0.Time)
                            //{
                            //    if ((star_time - eventItem.Time > TimeSpan.FromMinutes(3)) && eventItem.Time < star_time)
                            //    {
                            //        // Устанавливаем цвет фона для ячейки
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                            //    }
                            //}

                            //if (eventItem.Time == lastEventType1.Time)
                            //{
                            //    if ((eventItem.Time - end_time > TimeSpan.FromMinutes(3)) && eventItem.Time > end_time)
                            //    {
                            //        // Устанавливаем цвет фона для ячейки
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //        worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Green);
                            //    }
                            //}

                            //if (worksheet.Cells[rowIndex, 5].Value == "-" && worksheet.Cells[rowIndex, 5].Value == "-")
                            //{
                            //    if (eventItem.Time == firstEventType0.Time)
                            //    {
                            //        if ((star_time - eventItem.Time > TimeSpan.FromMinutes(3)) && eventItem.Time < star_time)
                            //        {
                            //            // Устанавливаем цвет фона для ячейки
                            //            worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //            worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                            //        }
                            //    }

                            //    if (eventItem.Time == lastEventType1.Time)
                            //    {
                            //        if ((eventItem.Time - end_time > TimeSpan.FromMinutes(3)) && eventItem.Time > end_time)
                            //        {
                            //            // Устанавливаем цвет фона для ячейки
                            //            worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            //            worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                            //        }
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

                        
                        
                        if (startRow == rowIndex)
                        {
                            // Форматирование строк данных
                            worksheet.Cells[$"A{startRow}:I{rowIndex }"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[$"A{startRow}:I{rowIndex }"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[$"A{startRow}:I{rowIndex }"].Style.WrapText = true;

                            // Добавление бордера к диапазону строк данных
                            worksheet.Cells[$"A{startRow}:I{rowIndex }"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[$"A{startRow}:I{rowIndex}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[$"A{startRow}:I{rowIndex}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[$"A{startRow}:I{rowIndex}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        }
                        else
                        {
                            // Форматирование строк данных
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.WrapText = true;

                            // Добавление бордера к диапазону строк данных
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[$"A{startRow}:I{rowIndex - 1}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
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
                                    //worksheet.Cells[rowIndex, 8].Value = unavailabilityForDate.UnavailabilityType.Name;
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



                            // Добавление бордера к каждой заполненной ячейке
                            for (int col = 1; col <= 9; col++)
                            {
                                worksheet.Cells[rowIndex, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[rowIndex, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[rowIndex, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[rowIndex, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            }

                            rowIndex++;
                        }

                    }

                    var str = employee.WorkSchedule.GetScheduleString();
                    worksheet.Cells[5, 9].Value = str;

                    if (rowIndex != 5)
                    {
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Merge = true;
                        // Форматирование столбца с личным графиком
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.WrapText = true;
                        // Форматирование столбца с личным графиком
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[$"I5:I{rowIndex - 1}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }

                    // Форматирование столбца с личным графиком
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.WrapText = true;
                    // Форматирование столбца с личным графиком
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[$"I5:I{rowIndex}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
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

        public static TimeOnly toTimeOnly(string row)
        {
            string[] formatsTime = { "H.mm.ss", "h:mm:ss tt", "HH:mm:ss", "h:mm tt", "h:mm", "HH:mm","h:mm tt" };
            if (TimeOnly.TryParseExact(row, formatsTime, CultureInfo.InvariantCulture, DateTimeStyles.None, out TimeOnly result1))
            {
                return result1;
            }
            else
            {
                return new TimeOnly(0);
            }
        }

        public static void colorCell(ExcelWorksheet worksheet, int rowIndex, Color color)
        {
            worksheet.Cells[rowIndex, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex, 2].Style.Fill.BackgroundColor.SetColor(color);
        }
        public static void colorCell(ExcelWorksheet worksheet, int rowIndex, int column, Color color)
        {
            worksheet.Cells[rowIndex, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex, column].Style.Fill.BackgroundColor.SetColor(color);
        }

        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static string FormatTimeSpan(TimeSpan timeSpan)
        {
            int totalHours = (int)timeSpan.TotalHours; // Получаем общее количество часов
            int minutes = timeSpan.Minutes;
            int seconds = timeSpan.Seconds;

            // Форматируем строку
            return $"{totalHours:D2}:{minutes:D2}:{seconds:D2}";
        }


        public static double CalculatePercentageDevs(int numDevsS, int numWorkDays)
        {
            return Math.Round((double)numDevsS / numWorkDays * 100, 2);
        }

        public static double CalculatePercentageTimeDevs(TimeSpan timDevsS, int numWorkDays)
        {
            TimeSpan eightHours = TimeSpan.FromHours(8);

            return Math.Round(timDevsS.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);
        }


        public async Task<ReportData> Smth(
                                                ReportSysContext _context,
                                                List<DateOnly> Dates,
                                                int depId,
                                                ReportData data)
        {
            var datesSet = new HashSet<DateOnly>(Dates); // Преобразуем список дат в HashSet для быстрой проверки

            var dep = await _context.Departments
                .Include(d => d.Employees).ThenInclude(e => e.Unavailabilitys)
                .Include(d => d.Employees).ThenInclude(e => e.Position)
                .Include(d => d.Employees).ThenInclude(e => e.WorkSchedule)
                .Include(d => d.Employees).ThenInclude(e => e.Events.Where(e => datesSet.Contains(e.Date)))
                .FirstOrDefaultAsync(d => d.Id == depId);

            data.Worksheet.Cells[data.Row, 1].Value = dep.Name;
            data.Worksheet.Cells[data.Row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            data.Worksheet.Row(data.Row).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            data.Row++;
            data.Worksheet.Cells[$"A{data.Row - 1}:{GetExcelColumnName(12 + Dates.Count())}{data.Row - 1}"].Merge = true;
            data.Worksheet.Cells[$"A{data.Row - 1}:{GetExcelColumnName(12 + Dates.Count())}{data.Row - 1}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            data.Worksheet.Cells[$"A{data.Row - 1}:{GetExcelColumnName(12 + Dates.Count())}{data.Row - 1}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            data.Worksheet.Cells[$"A{data.Row - 1}:{GetExcelColumnName(12 + Dates.Count())}{data.Row - 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            data.Worksheet.Cells[$"A{data.Row - 1}:{GetExcelColumnName(12 + Dates.Count())}{data.Row - 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            var numPP = 0;
            var numWorkDays = 0;
            var numNegDevsS = 0;
            var numPosDevsS = 0;
            var numNegDevsE = 0;
            var numPosDevsE = 0;
            var timNegDevsS = new TimeSpan();
            var timPosDevsS = new TimeSpan();
            var timNegDevsE = new TimeSpan();
            var timPosDevsE = new TimeSpan();
            // Создание временного интервала в 8 часов
            TimeSpan eightHours = TimeSpan.FromHours(8);

            foreach (var emp in dep.Employees)
            {
                var startTime = emp.WorkSchedule.Arrival;
                var endTime = emp.WorkSchedule.Exit;

                data.Worksheet.Cells[data.Row, 1].Value = numPP;
                data.Row++;
                data.Worksheet.Cells[$"A{data.Row - 1}:A{data.Row}"].Merge = true;

                numPP++;

                data.Worksheet.Cells[data.Row - 1, 2].Value = emp.Position.Name;
                data.Worksheet.Cells[$"B{data.Row - 1}:B{data.Row}"].Merge = true;

                data.Worksheet.Cells[data.Row - 1, 3].Value = emp.FirstName + " " + emp.SecondName + " " + emp.LastName;
                data.Worksheet.Cells[$"C{data.Row - 1}:C{data.Row}"].Merge = true;

                data.Worksheet.Cells[data.Row - 1, 4].Value = "приход";
                data.Worksheet.Cells[data.Row, 4].Value = "уход";
                var i = 5;
                numWorkDays = 0;
                numNegDevsS = 0;
                numPosDevsS = 0;
                numNegDevsE = 0;
                numPosDevsE = 0;
                timNegDevsS = new TimeSpan();
                timPosDevsS = new TimeSpan();
                timNegDevsE = new TimeSpan();
                timPosDevsE = new TimeSpan();
                foreach (var date in Dates)
                {
                    var events = await _context.Events
                                .Include(e => e.EventType)
                                .Where(e => e.EmployeeId == emp.Id && e.Date == date)
                                .ToListAsync();

                    var unf = await _context.Unavailabilitys
                                .Include(e => e.UnavailabilityType)
                                .Include(e => e.Employee)
                                .Where(e => e.Date == date && e.Employee == emp)
                                .FirstOrDefaultAsync();

                    if (events != null && events.Count != 0)
                    {
                        // Найти первый евент с EventTypeId == 0
                        var firstEventType0 = events.FirstOrDefault(e => e.EventType.Id == 1);

                        // Найти последний евент с EventTypeId == 1
                        var lastEventType1 = events.LastOrDefault(e => e.EventType.Id == 2);

                        if (firstEventType0 != null)
                        {
                            data.Worksheet.Cells[data.Row - 1, i].Value = firstEventType0.Time.ToString("HH:mm:ss");

                            if (firstEventType0.Time - startTime > TimeSpan.FromMinutes(3) && firstEventType0.Time > startTime)
                            {
                                if (unf != null)
                                {
                                    if (unf.UnavailabilityTypeId != 4)
                                    {
                                        colorCell(data.Worksheet, data.Row - 1, i, Color.SkyBlue);
                                    }
                                    else if (firstEventType0.Time < unf.UnavailabilityFrom || firstEventType0.Time > unf.UnavailabilityBefore)
                                    {
                                        colorCell(data.Worksheet, data.Row - 1, i, Color.SandyBrown);
                                        numNegDevsS++;
                                        timNegDevsS = timNegDevsS.Add(firstEventType0.Time - startTime);
                                    }
                                    else
                                    {
                                        colorCell(data.Worksheet, data.Row - 1, i, Color.Khaki);
                                    }
                                }
                                else
                                {
                                    colorCell(data.Worksheet, data.Row - 1, i, Color.SandyBrown);
                                    numNegDevsS++;
                                    timNegDevsS = timNegDevsS.Add(firstEventType0.Time - startTime);
                                }
                            }
                            if (startTime - firstEventType0.Time > TimeSpan.FromMinutes(3) && startTime > firstEventType0.Time)
                            {
                                if (unf != null)
                                {
                                    if (unf.UnavailabilityTypeId != 4)
                                    {
                                        colorCell(data.Worksheet, data.Row - 1, i, Color.SkyBlue);
                                    }
                                }
                                else
                                {
                                    colorCell(data.Worksheet, data.Row - 1, i, Color.LightGreen);
                                    numPosDevsS++;
                                    timPosDevsS = timPosDevsS.Add(startTime - firstEventType0.Time);
                                }
                            }
                        }
                        else
                        {
                            colorCell(data.Worksheet, data.Row - 1, i, Color.Pink);
                        }
                        if (lastEventType1 != null)
                        {
                            data.Worksheet.Cells[data.Row, i].Value = lastEventType1.Time.ToString("HH:mm:ss");

                            if (endTime - lastEventType1.Time > TimeSpan.FromMinutes(3) && endTime > lastEventType1.Time)
                            {
                                if (unf != null)
                                {
                                    if (unf.UnavailabilityTypeId != 4)
                                    {
                                        colorCell(data.Worksheet, data.Row, i, Color.SkyBlue);
                                    }
                                    else if (lastEventType1.Time < unf.UnavailabilityFrom || lastEventType1.Time > unf.UnavailabilityBefore)
                                    {
                                        colorCell(data.Worksheet, data.Row, i, Color.SandyBrown);
                                        numNegDevsE++;
                                        timNegDevsE = timNegDevsE.Add(endTime - lastEventType1.Time);
                                    }
                                    else
                                    {
                                        colorCell(data.Worksheet, data.Row, i, Color.Khaki);
                                    }
                                }
                                else
                                {
                                    colorCell(data.Worksheet, data.Row, i, Color.SandyBrown);
                                    numNegDevsE++;
                                    timNegDevsE = timNegDevsE.Add(endTime - lastEventType1.Time);
                                }
                            }
                            if (lastEventType1.Time - endTime > TimeSpan.FromMinutes(3) && lastEventType1.Time > endTime)
                            {
                                if (unf != null)
                                {
                                    if (unf.UnavailabilityTypeId != 4)
                                    {
                                        colorCell(data.Worksheet, data.Row, i, Color.SkyBlue);
                                    }
                                }
                                else
                                {
                                    colorCell(data.Worksheet, data.Row, i, Color.LightGreen);
                                    numPosDevsE++;
                                    timPosDevsE = timPosDevsE.Add(lastEventType1.Time - endTime);
                                }
                            }
                        }
                        else
                        {
                            colorCell(data.Worksheet, data.Row, i, Color.Pink);
                        }
                        if (lastEventType1 != null || firstEventType0 != null)
                        {
                            numWorkDays++;
                        }

                        i++;
                    }
                    else
                    {
                        if (unf != null)
                        {
                            if (unf.UnavailabilityTypeId != 4)
                            {
                                colorCell(data.Worksheet, data.Row - 1, i, Color.Khaki);
                                colorCell(data.Worksheet, data.Row, i, Color.Khaki);
                            }
                        }
                        else
                        {
                            colorCell(data.Worksheet, data.Row - 1, i, Color.SandyBrown);
                            colorCell(data.Worksheet, data.Row, i, Color.SandyBrown);
                        }
                        i++;
                    }
                }

                if (numWorkDays != 0)
                {
                    data.Worksheet.Cells[data.Row - 1, 5 + Dates.Count()].Value = numNegDevsS;
                    data.Worksheet.Cells[data.Row - 1, 6 + Dates.Count()].Value = CalculatePercentageDevs(numNegDevsS, numWorkDays);

                    data.Worksheet.Cells[data.Row - 1, 7 + Dates.Count()].Value = FormatTimeSpan(timNegDevsS);
                    data.Worksheet.Cells[data.Row - 1, 8 + Dates.Count()].Value = CalculatePercentageTimeDevs(timNegDevsS, numWorkDays);

                    data.Worksheet.Cells[data.Row - 1, 9 + Dates.Count()].Value = numPosDevsS;
                    data.Worksheet.Cells[data.Row - 1, 10 + Dates.Count()].Value = CalculatePercentageDevs(numPosDevsS, numWorkDays);

                    data.Worksheet.Cells[data.Row - 1, 11 + Dates.Count()].Value = FormatTimeSpan(timPosDevsS);
                    data.Worksheet.Cells[data.Row - 1, 12 + Dates.Count()].Value = CalculatePercentageTimeDevs(timPosDevsS, numWorkDays);

                    data.Worksheet.Cells[data.Row, 5 + Dates.Count()].Value = numNegDevsE;
                    data.Worksheet.Cells[data.Row, 6 + Dates.Count()].Value = CalculatePercentageDevs(numNegDevsE, numWorkDays);

                    data.Worksheet.Cells[data.Row, 7 + Dates.Count()].Value = FormatTimeSpan(timNegDevsE);
                    data.Worksheet.Cells[data.Row, 8 + Dates.Count()].Value = CalculatePercentageTimeDevs(timNegDevsE, numWorkDays);

                    data.Worksheet.Cells[data.Row, 9 + Dates.Count()].Value = numPosDevsE;
                    data.Worksheet.Cells[data.Row, 10 + Dates.Count()].Value = CalculatePercentageDevs(numPosDevsE, numWorkDays);

                    data.Worksheet.Cells[data.Row, 11 + Dates.Count()].Value = FormatTimeSpan(timPosDevsE);
                    data.Worksheet.Cells[data.Row, 12 + Dates.Count()].Value = CalculatePercentageTimeDevs(timPosDevsE, numWorkDays);

                    data.MinSumS += numNegDevsS;
                    data.MinSumE += numNegDevsE;
                    data.PlusSumS += numPosDevsS;
                    data.PlusSumE += numPosDevsE;
                    data.MinTimeS = data.MinTimeS.Add(timNegDevsS);
                    data.MinTimeE = data.MinTimeE.Add(timNegDevsE);
                    data.PlusTimeS = data.PlusTimeS.Add(timPosDevsS);
                    data.PlusTimeE = data.PlusTimeE.Add(timPosDevsE);
                }
                // Добавление бордера к каждой заполненной ячейке
                for (int col = 1; col <= 12 + Dates.Count(); col++)
                {
                    data.Worksheet.Cells[data.Row - 1, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row - 1, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row - 1, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row - 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    data.Worksheet.Cells[data.Row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                data.Row++;
            }
            return data;
        }

    }

}


