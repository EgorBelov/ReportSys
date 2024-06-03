using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.Pages.Services;

namespace ReportSys.Pages.PageAccess1
{
    public class IndexModel : ServicesPage
    {


        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public List<SelectListItem> EmployeeList { get; set; }


        [BindProperty]
        public List<int> SelectedEmployeeIds { get; set; }



        [BindProperty]
        public List<SelectListItem> DepartList { get; set; }


        [BindProperty]
        public List<int> SelectedDepartIds { get; set; }



        [BindProperty]
        public DateOnly StartDate { get; set; }

        [BindProperty]
        public DateOnly EndDate { get; set; }

        [BindProperty]
        public List<DateOnly> Dates { get; set; }


        [BindProperty]
        public string Action { get; set; }

        public async Task<IActionResult> OnGetAsync()
        {
            var employeeNumber = HttpContext.Session.GetString("EmployeeNumber");

            if (string.IsNullOrEmpty(employeeNumber))
            {
                return RedirectToPage("/Error"); // Перенаправление на страницу ошибки, если нет номера сотрудника
            }

            var employee = await _context.Employees
                .Include(e => e.Department)
                .FirstOrDefaultAsync(e => e.Id.ToString() == employeeNumber);

            if (employee == null)
            {
                return RedirectToPage("/Error"); // Перенаправление на страницу ошибки, если сотрудник не найден
            }

            await EmployeesFromDepartAsync(_context, employee);
            await DepartmentsFromDepartAsync(_context, employee.DepartmentId);

            // Заполнение свойств EmployeeList и DepartList
            EmployeeList = EmployeesSL.ToList();
            DepartList = DepartmentsSL.ToList();

            return Page();
        }




        public async Task<IActionResult> OnPostAsync()
        {
            if (Action == "Action1")
            {
                return await HandleAction1();
            }
            else if (Action == "Action2")
            {
                return await HandleAction2();
            }

            return Page();
        }

        private async Task<IActionResult> HandleAction1()
        {
            // Логика для Action1
            // Например, перенаправление на другую страницу или возврат данных
            var employeeNumbers = new List<string>();

            foreach(var empId in SelectedEmployeeIds)
            {
                employeeNumbers.Add(empId.ToString());
            }
            return await CreateXlsxFirst(_context, employeeNumbers, StartDate, EndDate);
        }

        private async Task<IActionResult> HandleAction2()
        {

           


            if (StartDate > EndDate)
            {
                TempData["Message"] = "Start date cannot be later than end date.";
                return Page();
            }

            Dates = new List<DateOnly>();
            for (var date = StartDate; date <= EndDate; date = date.AddDays(1))
            {
                if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                {
                    continue; // Пропускаем субботу и воскресенье
                }
                Dates.Add(date);
            }

            var employeeNumber = HttpContext.Session.GetString("EmployeeNumber");

            var employee = await _context.Employees.Include(e => e.WorkSchedule)
               .Include(e => e.Events).ThenInclude(s => s.EventType)
               .Include(e => e.Unavailabilitys).ThenInclude(s => s.UnavailabilityType)
               .FirstOrDefaultAsync(e => e.Id.ToString() == employeeNumber);

            var deps = await _context.Employees.Include(e => e.Department)
                                                    .Include(e => e.Position)
                                                    .Include(e => e.Events).ThenInclude(s => s.EventType)
                                                    .Where(e => e.DepartmentId == employee.DepartmentId).ToListAsync();



            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Employee");

                
                worksheet.Cells[1, 1].Value = "ПП";
                worksheet.Cells[1, 2].Value = "Наименование штатной должности";
                worksheet.Cells[1, 3].Value = "ФИО";
                worksheet.Cells[1, 4].Value = "Событие";
                worksheet.Cells[1, 5].Value = "Время прихода ухода";
               

                
                worksheet.Cells["A1:A2"].Merge = true; 
                worksheet.Cells["B1:B2"].Merge = true; 
                worksheet.Cells["C1:C2"].Merge = true; 
                worksheet.Cells["D1:D2"].Merge = true; 
               
                for (int i = 0; i < Dates.Count(); i++)
                {
                    worksheet.Cells[2, i + 5].Value = Dates[i];
                }




                int baseColumnIndex = 5 + Dates.Count();

                worksheet.Cells[1, baseColumnIndex].Value = "Кол-во \"-\" откл.";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex)}1:{GetExcelColumnName(baseColumnIndex + 1)}1"].Merge = true;

                worksheet.Cells[1, baseColumnIndex + 2].Value = "Общее время";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex + 2)}1:{GetExcelColumnName(baseColumnIndex + 3)}1"].Merge = true;

                worksheet.Cells[1, baseColumnIndex + 4].Value = "Кол-во \"+\" откл.";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex + 4)}1:{GetExcelColumnName(baseColumnIndex + 5)}1"].Merge = true;

                worksheet.Cells[1, baseColumnIndex + 6].Value = "Общее время";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex + 6)}1:{GetExcelColumnName(baseColumnIndex + 7)}1"].Merge = true;

                worksheet.Cells[2, 5 + Dates.Count()].Value = "ед";
                worksheet.Cells[2, 6 + Dates.Count()].Value = "%";
                worksheet.Cells[2, 7 + Dates.Count()].Value = "ч";
                worksheet.Cells[2, 8 + Dates.Count()].Value = "%";
                worksheet.Cells[2, 9 + Dates.Count()].Value = "ед";
                worksheet.Cells[2, 10 + Dates.Count()].Value = "%";
                worksheet.Cells[2, 11 + Dates.Count()].Value = "ч";
                worksheet.Cells[2, 12 + Dates.Count()].Value = "%";



                //var UpDepIds = FindTopLevelDepartments(SelectedDepartIds, _context);
                var UpDepIds = await _context.Hierarchies
                        .Where(e => SelectedDepartIds.Contains(e.UpperDepartmentId))
                        .ToListAsync();

                int row = 3;


                if (UpDepIds.Count() == 0)
                {
                    foreach(var depId in SelectedDepartIds)
                    {
                        var datesSet = new HashSet<DateOnly>(Dates); // Преобразуем список дат в HashSet для быстрой проверки

                        var dep = await _context.Departments
                            .Include(d => d.Employees).ThenInclude(e => e.Unavailabilitys)
                            .Include(d => d.Employees).ThenInclude(e => e.Position)
                            .Include(d => d.Employees).ThenInclude(e => e.WorkSchedule)
                            .Include(d => d.Employees).ThenInclude(e => e.Events.Where(e => datesSet.Contains(e.Date)))
                            .FirstOrDefaultAsync(d => d.Id == depId);


                        worksheet.Cells[row, 1].Value = dep.Name;
                        row++;
                        worksheet.Cells[$"A{row-1}:{GetExcelColumnName(12 + Dates.Count())}{row-1}"].Merge = true;
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

                            
                            worksheet.Cells[row, 1].Value = numPP;
                            row++;
                            worksheet.Cells[$"A{row-1}:A{row}"].Merge= true;

                            numPP++;

                            
                            worksheet.Cells[row-1, 2].Value = emp.Position.Name;

                            worksheet.Cells[$"B{row - 1}:B{row}"].Merge = true;

                          
                            worksheet.Cells[row-1, 3].Value = emp.FirstName + " " + emp.SecondName + " " + emp.LastName;

                            worksheet.Cells[$"C{row - 1}:C{row}"].Merge = true;

                            worksheet.Cells[row - 1, 4].Value = "приход";
                            worksheet.Cells[row, 4].Value = "уход";
                            var i = 5;
                            numWorkDays = 0;
                            numWorkDays = 0;
                            numWorkDays = 0;
                            numWorkDays = 0;
                            numPosDevsE = 0;
                            timNegDevsS = new TimeSpan();
                            timPosDevsS = new TimeSpan();
                            timNegDevsS = new TimeSpan();
                            timNegDevsS = new TimeSpan();
                            foreach (var date in Dates)
                            {
                                var events = await _context.Events
                                            .Include(e => e.EventType)
                                            .Where(e => e.EmployeeId == emp.Id && e.Date == date)
                                            .ToListAsync();

                                


                                if (events != null && events.Count != 0)
                                {
                                    // Найти первый евент с EventTypeId == 0
                                    var firstEventType0 = events.FirstOrDefault(e => e.EventType.Id == 1);

                                    // Найти последний евент с EventTypeId == 1
                                    var lastEventType1 = events.LastOrDefault(e => e.EventType.Id == 2);



                                    if (firstEventType0 != null)
                                    {
                                        worksheet.Cells[row - 1, i].Value = firstEventType0.Time.ToString("HH:mm:ss"); ;
                                        if (firstEventType0.Time - startTime > TimeSpan.FromMinutes(3) && firstEventType0.Time > startTime)
                                        {
                                            numNegDevsS++;
                                            timNegDevsS = timNegDevsS.Add(firstEventType0.Time - startTime);
                                        }
                                        if (startTime - firstEventType0.Time > TimeSpan.FromMinutes(3) && startTime > firstEventType0.Time)
                                        {
                                            numPosDevsS++;
                                            timPosDevsS = timPosDevsS.Add(startTime - firstEventType0.Time);
                                        }
                                    }
                                    
                                    if (lastEventType1 != null)
                                    {
                                        worksheet.Cells[row, i].Value = lastEventType1.Time.ToString("HH:mm:ss");
                                        if ( lastEventType1.Time - endTime > TimeSpan.FromMinutes(3) && lastEventType1.Time > endTime)
                                        {
                                            numPosDevsE++;
                                            timPosDevsE = timPosDevsE.Add(lastEventType1.Time - endTime);
                                        }
                                        if (endTime - lastEventType1.Time > TimeSpan.FromMinutes(3) && endTime > lastEventType1.Time)
                                        {
                                            numNegDevsE++;
                                            timNegDevsE = timNegDevsE.Add(endTime - lastEventType1.Time);
                                        }

                                    }
                                    
                                    if (lastEventType1 != null || firstEventType0 != null)
                                    {
                                        numWorkDays++;
                                    }
                                                                       
                                    i++;
                                }
                                
                                
                            }

                            if (numWorkDays != 0)
                            {
                                worksheet.Cells[row - 1, 5 + Dates.Count()].Value = numNegDevsS;
                                worksheet.Cells[row - 1, 6 + Dates.Count()].Value = Math.Round((double)numNegDevsS / numWorkDays * 100, 2);

                                worksheet.Cells[row - 1, 7 + Dates.Count()].Value = timNegDevsS.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row - 1, 8 + Dates.Count()].Value = Math.Round(timNegDevsS.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);

                                worksheet.Cells[row - 1, 9 + Dates.Count()].Value = numPosDevsS;
                                worksheet.Cells[row - 1, 10 + Dates.Count()].Value = Math.Round((double)numPosDevsS / numWorkDays * 100, 2);

                                worksheet.Cells[row - 1, 11 + Dates.Count()].Value = timPosDevsS.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row - 1, 12 + Dates.Count()].Value = Math.Round(timPosDevsS.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);



                                worksheet.Cells[row, 5 + Dates.Count()].Value = numNegDevsE;
                                worksheet.Cells[row, 6 + Dates.Count()].Value = Math.Round((double)numNegDevsE / numWorkDays * 100, 2);

                                worksheet.Cells[row, 7 + Dates.Count()].Value = timNegDevsE.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row, 8 + Dates.Count()].Value = Math.Round(timNegDevsE.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);

                                worksheet.Cells[row, 9 + Dates.Count()].Value = numPosDevsE;
                                worksheet.Cells[row, 10 + Dates.Count()].Value = Math.Round((double)numPosDevsE / numWorkDays * 100, 2);

                                worksheet.Cells[row - 1, 11 + Dates.Count()].Value = timPosDevsE.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row, 12 + Dates.Count()].Value = Math.Round(timPosDevsE.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);
                            }

                            row++;
                        }
                    }







                }
                else
                {
                    foreach (var deP in UpDepIds)
                    {
                        var datesSet = new HashSet<DateOnly>(Dates); // Преобразуем список дат в HashSet для быстрой проверки
                        var depId = deP.Id;
                        var dep = await _context.Departments
                            .Include(d => d.Employees).ThenInclude(e => e.Unavailabilitys)
                            .Include(d => d.Employees).ThenInclude(e => e.Position)
                            .Include(d => d.Employees).ThenInclude(e => e.WorkSchedule)
                            .Include(d => d.Employees).ThenInclude(e => e.Events.Where(e => datesSet.Contains(e.Date)))
                            .FirstOrDefaultAsync(d => d.Id == depId);


                        worksheet.Cells[row, 1].Value = dep.Name;
                        row++;
                        worksheet.Cells[$"A{row - 1}:{GetExcelColumnName(12 + Dates.Count())}{row - 1}"].Merge = true;
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


                            worksheet.Cells[row, 1].Value = numPP;
                            row++;
                            worksheet.Cells[$"A{row - 1}:A{row}"].Merge = true;

                            numPP++;


                            worksheet.Cells[row - 1, 2].Value = emp.Position.Name;

                            worksheet.Cells[$"B{row - 1}:B{row}"].Merge = true;


                            worksheet.Cells[row - 1, 3].Value = emp.FirstName + " " + emp.SecondName + " " + emp.LastName;

                            worksheet.Cells[$"C{row - 1}:C{row}"].Merge = true;

                            worksheet.Cells[row - 1, 4].Value = "приход";
                            worksheet.Cells[row, 4].Value = "уход";
                            var i = 5;
                            numWorkDays = 0;
                            numWorkDays = 0;
                            numWorkDays = 0;
                            numWorkDays = 0;
                            numPosDevsE = 0;
                            timNegDevsS = new TimeSpan();
                            timPosDevsS = new TimeSpan();
                            timNegDevsS = new TimeSpan();
                            timNegDevsS = new TimeSpan();
                            foreach (var date in Dates)
                            {
                                var events = await _context.Events
                                            .Include(e => e.EventType)
                                            .Where(e => e.EmployeeId == emp.Id && e.Date == date)
                                            .ToListAsync();




                                if (events != null && events.Count != 0)
                                {
                                    // Найти первый евент с EventTypeId == 0
                                    var firstEventType0 = events.FirstOrDefault(e => e.EventType.Id == 1);

                                    // Найти последний евент с EventTypeId == 1
                                    var lastEventType1 = events.LastOrDefault(e => e.EventType.Id == 2);



                                    if (firstEventType0 != null)
                                    {
                                        worksheet.Cells[row - 1, i].Value = firstEventType0.Time.ToString("HH:mm:ss"); ;
                                        if (firstEventType0.Time - startTime > TimeSpan.FromMinutes(3) && firstEventType0.Time > startTime)
                                        {
                                            numNegDevsS++;
                                            timNegDevsS = timNegDevsS.Add(firstEventType0.Time - startTime);
                                        }
                                        if (startTime - firstEventType0.Time > TimeSpan.FromMinutes(3) && startTime > firstEventType0.Time)
                                        {
                                            numPosDevsS++;
                                            timPosDevsS = timPosDevsS.Add(startTime - firstEventType0.Time);
                                        }
                                    }

                                    if (lastEventType1 != null)
                                    {
                                        worksheet.Cells[row, i].Value = lastEventType1.Time.ToString("HH:mm:ss");
                                        if (lastEventType1.Time - endTime > TimeSpan.FromMinutes(3) && lastEventType1.Time > endTime)
                                        {
                                            numPosDevsE++;
                                            timPosDevsE = timPosDevsE.Add(lastEventType1.Time - endTime);
                                        }
                                        if (endTime - lastEventType1.Time > TimeSpan.FromMinutes(3) && endTime > lastEventType1.Time)
                                        {
                                            numNegDevsE++;
                                            timNegDevsE = timNegDevsE.Add(endTime - lastEventType1.Time);
                                        }

                                    }

                                    if (lastEventType1 != null || firstEventType0 != null)
                                    {
                                        numWorkDays++;
                                    }

                                    i++;
                                }


                            }

                            if (numWorkDays != 0)
                            {
                                worksheet.Cells[row - 1, 5 + Dates.Count()].Value = numNegDevsS;
                                worksheet.Cells[row - 1, 6 + Dates.Count()].Value = Math.Round((double)numNegDevsS / numWorkDays * 100, 2);

                                worksheet.Cells[row - 1, 7 + Dates.Count()].Value = timNegDevsS.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row - 1, 8 + Dates.Count()].Value = Math.Round(timNegDevsS.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);

                                worksheet.Cells[row - 1, 9 + Dates.Count()].Value = numPosDevsS;
                                worksheet.Cells[row - 1, 10 + Dates.Count()].Value = Math.Round((double)numPosDevsS / numWorkDays * 100, 2);

                                worksheet.Cells[row - 1, 11 + Dates.Count()].Value = timPosDevsS.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row - 1, 12 + Dates.Count()].Value = Math.Round(timPosDevsS.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);



                                worksheet.Cells[row, 5 + Dates.Count()].Value = numNegDevsE;
                                worksheet.Cells[row, 6 + Dates.Count()].Value = Math.Round((double)numNegDevsE / numWorkDays * 100, 2);

                                worksheet.Cells[row, 7 + Dates.Count()].Value = timNegDevsE.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row, 8 + Dates.Count()].Value = Math.Round(timNegDevsE.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);

                                worksheet.Cells[row, 9 + Dates.Count()].Value = numPosDevsE;
                                worksheet.Cells[row, 10 + Dates.Count()].Value = Math.Round((double)numPosDevsE / numWorkDays * 100, 2);

                                worksheet.Cells[row - 1, 11 + Dates.Count()].Value = timPosDevsE.ToString(@"hh\:mm\:ss");
                                worksheet.Cells[row, 12 + Dates.Count()].Value = Math.Round(timPosDevsE.TotalHours / (numWorkDays * eightHours.TotalHours) * 100, 2);
                            }

                            row++;
                        }
                    }
                }
                package.Save();
            }
            stream.Position = 0;
            var fileName = "Employee.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(stream, contentType, fileName);
        
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }



    }

}
