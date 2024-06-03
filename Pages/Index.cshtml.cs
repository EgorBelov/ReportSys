using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.DAL.Entities;
using System.Data;
using System.Globalization;
using System.Net;

namespace ReportSys.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public IFormFile Upload { get; set; }


        public string RemoveExtraSpaces(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            // Разбиваем строку на слова, удаляем пустые строки и объединяем обратно с одним пробелом
            var words = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", words);
        }

        //public async Task LoadExcelFile()
        //{
        //    DataTable dataTable = new DataTable();

        //    // Копируем загруженный файл в поток
        //    using (var stream = new MemoryStream())
        //    {
        //        await Upload.CopyToAsync(stream);
        //        using (ExcelPackage package = new ExcelPackage(stream))
        //        {
        //            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Используем первый лист

        //            // Добавляем колонки
        //            foreach (var firstRowCell in worksheet.Cells[4, 1, 4, worksheet.Dimension.End.Column])
        //            {
        //                dataTable.Columns.Add(firstRowCell.Text);
        //            }

        //            // Добавляем строки
        //            for (int rowNum = 5; rowNum <= worksheet.Dimension.End.Row; rowNum++)
        //            {
        //                var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
        //                DataRow row = dataTable.NewRow();
        //                foreach (var cell in wsRow)
        //                {
        //                    row[cell.Start.Column - 1] = cell.Text;
        //                }
        //                dataTable.Rows.Add(row);
        //            }
        //        }
        //    }

        //    var uniqueEmployeeNames = GetUniqueColumnValues(dataTable, "Сотрудник (Посетитель)");

        //    foreach (var employeeName in uniqueEmployeeNames)
        //    {
        //        string[] words = employeeName.Split(' ');

        //        if (words.Length < 3)
        //        {
        //            // Обработка случая, когда ФИО сотрудника имеет менее 3 частей
        //            continue;
        //        }


        //        int id = int.Parse(RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Карта №")).Trim());
        //        string positionName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Должность")).Trim();
        //        string divisionOrDepartmentName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Подразделение")).Trim();

        //        try
        //        {
        //            var position = await _context.Positions
        //                                         .FirstOrDefaultAsync(x => x.Name == positionName);

        //            if (position == null)
        //            {
        //                // Обработка случая, когда позиция не найдена
        //                Console.WriteLine($"Position not found: {positionName}");
        //                continue;
        //            }

        //            var employee = new Employee
        //            {
        //                Id = id,
        //                FirstName = words[0],
        //                SecondName = words[1],
        //                LastName = words[2]
        //            };

        //            var department = await _context.Departments
        //                                            .FirstOrDefaultAsync(x => x.Name == divisionOrDepartmentName);

        //            if (department != null)
        //            {
        //                department.Employees.Add(employee);
        //            }
        //            else
        //            {
        //                Console.WriteLine($"Department not found: {divisionOrDepartmentName}");
        //            }



        //            position.Employees.Add(employee);

        //            var workSchedules = await _context.WorkSchedules.ToListAsync();
        //            if (workSchedules.Any())
        //            {
        //                workSchedules[0].Employees.Add(employee);
        //            }

        //            var needrows = GetRowsByColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName);

        //            var events = new List<Event>();
        //            foreach (var row in needrows) 
        //            {
        //                var eventtype = await _context.EventTypes
        //                                             .FirstOrDefaultAsync(x => x.Name == row[10].ToString());
        //                // Формат даты
        //                string format = "d.M.yyyy";

        //                // Попытка парсинга строки в объект DateOnly
        //                if (DateOnly.TryParseExact(row[3].ToString(), format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly result))
        //                {
        //                    Console.WriteLine($"Дата: {result}");
        //                }
        //                else
        //                {
        //                    Console.WriteLine("Невозможно преобразовать строку в дату.");
        //                }


        //                // Формат времени
        //                string format1 = "H:mm:ss";

        //                // Попытка парсинга строки в объект TimeOnly
        //                if (TimeOnly.TryParseExact(row[4].ToString(), format1, out TimeOnly result1))
        //                {
        //                    Console.WriteLine($"Время: {result1}");
        //                }
        //                else
        //                {
        //                    Console.WriteLine("Невозможно преобразовать строку в время.");
        //                }


        //                events.Add(

        //                    new Event
        //                    {
        //                        Date = result,
        //                        Time = result1,
        //                        Territory = row[8].ToString(),
        //                        EventType = eventtype,
        //                        Employee = employee
        //                    }

        //                    );

        //            }


        //            await _context.Events.AddRangeAsync(events);
        //            await _context.Employees.AddAsync(employee);
        //            await _context.SaveChangesAsync();
        //        }
        //        catch (Exception ex)
        //        {
        //            // Логирование исключения
        //            Console.WriteLine($"An error occurred: {ex.Message}");
        //        }
        //    }
        //}

        // Метод для получения уникальных значений столбца 

        public async Task LoadExcelFile()
        {
            DataTable dataTable = new DataTable();

            // Copy the uploaded file to a stream
            using (var stream = new MemoryStream())
            {
                await Upload.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Use the first sheet

                    // Add columns
                    foreach (var firstRowCell in worksheet.Cells[4, 1, 4, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

                    // Add rows
                    for (int rowNum = 5; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                        DataRow row = dataTable.NewRow();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                        dataTable.Rows.Add(row);
                    }
                }
            }

            var uniqueEmployeeNames = GetUniqueColumnValues(dataTable, "Сотрудник (Посетитель)");

            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                var employeesToAdd = new List<Employee>();
                var eventsToAdd = new List<Event>();

                foreach (var employeeName in uniqueEmployeeNames)
                {
                    string[] words = employeeName.Split(' ');

                    if (words.Length < 3)
                    {
                        // Handle case where employee name has less than 3 parts
                        continue;
                    }

                    int id = int.Parse(RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Карта №")).Trim());
                    string positionName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Должность")).Trim();
                    string divisionOrDepartmentName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Подразделение")).Trim();

                    try
                    {
                        var position = await _context.Positions
                                                     .FirstOrDefaultAsync(x => x.Name == positionName);

                        if (position == null)
                        {
                            // Handle case where position is not found
                            Console.WriteLine($"Position not found: {positionName}");
                            continue;
                        }

                        var employee = new Employee
                        {
                            Id = id,
                            FirstName = words[0],
                            SecondName = words[1],
                            LastName = words[2]
                        };

                        var department = await _context.Departments
                                                        .FirstOrDefaultAsync(x => x.Name == divisionOrDepartmentName);

                        if (department != null)
                        {
                            department.Employees.Add(employee);
                        }
                        else
                        {
                            Console.WriteLine($"Department not found: {divisionOrDepartmentName}");
                        }

                        position.Employees.Add(employee);

                        var workSchedules = await _context.WorkSchedules.ToListAsync();
                        if (workSchedules.Any())
                        {
                            workSchedules[0].Employees.Add(employee);
                        }

                        var needrows = GetRowsByColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName);

                        foreach (var row in needrows)
                        {
                            var eventtype = await _context.EventTypes
                                                         .FirstOrDefaultAsync(x => x.Name == row[10].ToString());

                            if (DateOnly.TryParseExact(row[3].ToString(), "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly dateResult) &&
                                TimeOnly.TryParseExact(row[4].ToString(), "H:mm:ss", out TimeOnly timeResult))
                            {
                                eventsToAdd.Add(new Event
                                {
                                    Date = dateResult,
                                    Time = timeResult,
                                    Territory = row[8].ToString(),
                                    EventType = eventtype,
                                    Employee = employee
                                });
                            }
                            else
                            {
                                Console.WriteLine("Failed to parse date or time.");
                            }
                        }

                        employeesToAdd.Add(employee);
                    }
                    catch (Exception ex)
                    {
                        // Log exception
                        Console.WriteLine($"An error occurred: {ex.Message}");
                    }
                }

                // Add employees and events in batches
                const int batchSize = 100;
                for (int i = 0; i < employeesToAdd.Count; i += batchSize)
                {
                    var employeeBatch = employeesToAdd.Skip(i).Take(batchSize);
                    await _context.Employees.AddRangeAsync(employeeBatch);
                }

                for (int i = 0; i < eventsToAdd.Count; i += batchSize)
                {
                    var eventBatch = eventsToAdd.Skip(i).Take(batchSize);
                    await _context.Events.AddRangeAsync(eventBatch);
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();
            }
        }


        private IEnumerable<string> GetUniqueColumnValues(DataTable dataTable, string columnName)
        {
            return dataTable.AsEnumerable()
                            .Select(row => row.Field<string>(columnName))
                            .Where(value => !string.IsNullOrEmpty(value)) // Отфильтровываем пустые значения
                            .Distinct();
        }



        // Метод для получения значения другого столбца в первой строке, где найдено указанное значение столбца
        public string GetOtherColumnValue(DataTable dataTable, string searchColumn, string searchValue, string resultColumn)
        {
            return dataTable.AsEnumerable()
                            .Where(row => row.Field<string>(searchColumn) == searchValue)
                            .Select(row => row.Field<string>(resultColumn))
                            .FirstOrDefault();
        }
        public IEnumerable<DataRow> GetRowsByColumnValue(DataTable dataTable, string searchColumn, string searchValue)
        {
            return dataTable.AsEnumerable()
                            .Where(row => row.Field<string>(searchColumn) == searchValue);
        }

        public async Task<IActionResult> OnPostAsync()
        {

            await LoadExcelFile();

            //foreach (DataRow row in data)
            //{
            //    var employee = new Employee
            //    {
            //        // Заполняем свойства модели данными из строки
            //        Name = row["NameColumnName"].ToString(), // Замените на реальное имя колонки
            //        Position = row["PositionColumnName"].ToString() // Замените на реальное имя колонки
            //                                                        // Добавьте другие свойства по необходимости
            //    };

            //    _context.Employees.Add(employee);
            //}

            await _context.SaveChangesAsync();

            // Setting success message
            TempData["SuccessMessage"] = "File uploaded successfully.";
            return RedirectToPage("/PageUnavailability/Index");
        }
    }


}