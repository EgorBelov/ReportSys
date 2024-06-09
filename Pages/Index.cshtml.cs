using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.DAL.Entities;
using System.Data;
using System.Globalization;
using System.Linq;

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

            // Split the string into words, remove empty strings, and join back with a single space
            var words = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", words);
        }

        //public async Task<IActionResult> OnPostAsync()
        //{
        //    try
        //    {
        //        await LoadExcelFile();
        //        TempData["SuccessMessage"] = "File uploaded successfully.";
        //        return RedirectToPage("/PageUnavailability/Index");
        //    }
        //    catch (Exception ex)
        //    {
        //        TempData["ErrorMessage"] = $"Error processing file: {ex.Message}";
        //        return Page();
        //    }
        //}

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

            var positions = await _context.Positions.ToListAsync();
            var departments = await _context.Departments.ToListAsync();
            var eventTypes = await _context.EventTypes.ToListAsync();
            var workSchedules = await _context.WorkSchedules.ToListAsync();

            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                var employeesToAdd = new List<Employee>();
                var eventsToAdd = new List<Event>();

                Parallel.ForEach(uniqueEmployeeNames, employeeName =>
                {
                    string[] words = employeeName.Split(' ');

                    if (words.Length < 3)
                    {
                        throw new Exception($"Invalid employee name format: {employeeName}");
                    }

                    int id;
                    if (!int.TryParse(RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Карта №")).Trim(), out id))
                    {
                        throw new Exception($"Invalid ID for employee: {employeeName}");
                    }

                    string positionName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Должность")).Trim();
                    string divisionOrDepartmentName = RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Подразделение")).Trim();

                    var position = positions.FirstOrDefault(x => x.Name == positionName);
                    if (position == null)
                    {
                        throw new Exception($"Position not found: {positionName}");
                    }

                    var department = departments.FirstOrDefault(x => x.Name == divisionOrDepartmentName);
                    if (department == null)
                    {
                        throw new Exception($"Department not found: {divisionOrDepartmentName}");
                    }

                    var employee = new Employee
                    {
                        Id = id,
                        FirstName = words[0],
                        SecondName = words[1],
                        LastName = words[2]
                    };

                    lock (department.Employees)
                    {
                        department.Employees.Add(employee);
                    }

                    lock (position.Employees)
                    {
                        position.Employees.Add(employee);
                    }

                    lock (workSchedules)
                    {
                        if (workSchedules.Any())
                        {
                            workSchedules[0].Employees.Add(employee);
                        }
                    }

                    var needrows = GetRowsByColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName);

                    foreach (var row in needrows)
                    {
                        var eventtype = eventTypes.FirstOrDefault(x => x.Name == row[10].ToString());
                        if (eventtype == null)
                        {
                            throw new Exception($"Event type not found: {row[10].ToString()}");
                        }

                        if (!DateOnly.TryParseExact(row[3].ToString(), "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly dateResult))
                        {
                            throw new Exception($"Invalid date format for row: {row[3].ToString()}");
                        }

                        if (!TimeOnly.TryParseExact(row[4].ToString(), "H:mm:ss", out TimeOnly timeResult))
                        {
                            throw new Exception($"Invalid time format for row: {row[4].ToString()}");
                        }

                        lock (eventsToAdd)
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
                    }

                    lock (employeesToAdd)
                    {
                        employeesToAdd.Add(employee);
                    }
                });

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
                            .Where(value => !string.IsNullOrEmpty(value)) // Filter out empty values
                            .Distinct();
        }

        // Method to get the value of another column in the first row where the specified column value is found
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
            try
            {
                // Validate the uploaded Excel file data
                var validationErrors = await ValidateExcelFile();
                if (validationErrors.Any())
                {
                    TempData["ErrorMessage"] = string.Join(", ", validationErrors);
                    return Page();
                }

                // If validation passes, clear the old data and load new data
                await ClearOldData();
                await LoadExcelFile();

                TempData["SuccessMessage"] = "File uploaded and processed successfully.";
                return RedirectToPage("/PageUnavailability/Index");
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = $"Error processing file: {ex.Message}";
                return Page();
            }
        }

        private async Task<List<string>> ValidateExcelFile()
        {
            var errors = new List<string>();
            DataTable dataTable = new DataTable();

            using (var stream = new MemoryStream())
            {
                await Upload.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Use the first sheet

                    foreach (var firstRowCell in worksheet.Cells[4, 1, 4, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

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
                    errors.Add($"Invalid employee name format: {employeeName}");
                }

                if (!int.TryParse(RemoveExtraSpaces(GetOtherColumnValue(dataTable, "Сотрудник (Посетитель)", employeeName, "Карта №")).Trim(), out int id))
                {
                    errors.Add($"Invalid ID for employee: {employeeName}");
                }
            }

            return errors;
        }

        private async Task ClearOldData()
        {
            var employees = await _context.Employees.ToListAsync();
            var events = await _context.Events.ToListAsync();

            _context.Employees.RemoveRange(employees);
            _context.Events.RemoveRange(events);

            await _context.SaveChangesAsync();
        }

    }
}
