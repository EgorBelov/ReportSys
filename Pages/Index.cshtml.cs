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

        public async Task LoadExcelFile()
        {
            DataTable dataTable = new DataTable();

            // Копируем загруженный файл в поток
            using (var stream = new MemoryStream())
            {
                await Upload.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Используем первый лист

                    // Добавляем колонки
                    foreach (var firstRowCell in worksheet.Cells[4, 1, 4, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

                    // Добавляем строки
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

            foreach (var employeeName in uniqueEmployeeNames)
            {
                string[] words = employeeName.Split(' ');

                if (words.Length < 3)
                {
                    // Обработка случая, когда ФИО сотрудника имеет менее 3 частей
                    continue;
                }

                string positionName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Должность")).Trim();
                string divisionOrDepartmentName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Подразделение")).Trim();

                try
                {
                    var position = await _context.Positions
                                                 .FirstOrDefaultAsync(x => x.Name == positionName);

                    if (position == null)
                    {
                        // Обработка случая, когда позиция не найдена
                        Console.WriteLine($"Position not found: {positionName}");
                        continue;
                    }

                    var employee = new Employee
                    {
                        FirstName = words[0],
                        SecondName = words[1],
                        PatronymicName = words[2]
                    };

                    if (divisionOrDepartmentName.Contains("Отдел"))
                    {
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
                    }
                    else
                    {
                        var division = await _context.Divisions
                                                     .FirstOrDefaultAsync(x => x.Name == divisionOrDepartmentName);

                        if (division != null)
                        {
                            division.Employees.Add(employee);
                        }
                        else
                        {
                            Console.WriteLine($"Division not found: {divisionOrDepartmentName}");
                        }
                    }

                    position.Employees.Add(employee);

                    var workSchedules = await _context.WorkSchedules.ToListAsync();
                    if (workSchedules.Any())
                    {
                        workSchedules[0].Employees.Add(employee);
                    }

                    var needrows = GetRowsByColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName);

                    var events = new List<Event>();
                    foreach (var row in needrows) 
                    {
                        var eventtype = await _context.EventTypes
                                                     .FirstOrDefaultAsync(x => x.Name == row[10].ToString());
                        // Формат даты
                        string format = "dd.MM.yyyy";

                        // Попытка парсинга строки в объект DateOnly
                        if (DateOnly.TryParseExact(row[3].ToString(), format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly result))
                        {
                            Console.WriteLine($"Дата: {result}");
                        }
                        else
                        {
                            Console.WriteLine("Невозможно преобразовать строку в дату.");
                        }


                        // Формат времени
                        string format1 = "HH:mm:ss";

                        // Попытка парсинга строки в объект TimeOnly
                        if (TimeOnly.TryParseExact(row[4].ToString(), format1, out TimeOnly result1))
                        {
                            Console.WriteLine($"Время: {result1}");
                        }
                        else
                        {
                            Console.WriteLine("Невозможно преобразовать строку в время.");
                        }


                        events.Add(

                            new Event
                            {
                                Date = result,
                                Time = result1,
                                Territory = row[8].ToString(),
                                EventType = eventtype,
                                Employee = employee
                            }

                            );

                    }


                    await _context.Events.AddRangeAsync(events);
                    await _context.Employees.AddAsync(employee);
                    await _context.SaveChangesAsync();
                }
                catch (Exception ex)
                {
                    // Логирование исключения
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }

        // Метод для получения уникальных значений столбца 
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
            return RedirectToPage("/Index");
        }
    }


}