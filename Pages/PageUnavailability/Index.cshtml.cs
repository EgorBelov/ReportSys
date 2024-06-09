using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.DAL.Entities;
using System.Data;
using System.Globalization;

namespace ReportSys.Pages.PageUnavailability
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

        public async Task LoadExcelFile()
        {
            DataTable dataTable = new DataTable();

            // �������� ����������� ���� � �����
            using (var stream = new MemoryStream())
            {
                await Upload.CopyToAsync(stream);
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // ���������� ������ ����

                    // ��������� �������
                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

                    // ��������� ������
                    for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
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

            var groupedRows = GroupRowsByColumnValue(dataTable, "Таб№ ");
            

            foreach (var kvp in groupedRows)
            {
                var emp = await _context.Employees
                        .FirstOrDefaultAsync(x => x.Id.ToString() == kvp.Key.ToString());

                var unavs = new List<Unavailability>();
                foreach(var row in kvp.Value)
                {
                    var typeUnav = await _context.UnavailabilityTypes
                                .FirstOrDefaultAsync(x => x.Id.ToString() == row[6].ToString());


                    // ������ �������
                    string format1 = "H.mm.ss";

                    // ������� �������� ������ � ������ TimeOnly
                    //if (TimeOnly.TryParseExact(row[2].ToString(), format1, out TimeOnly result1))
                    //{
                    //    Console.WriteLine($"�����: {result1}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("���������� ������������� ������ � �����.");
                    //}

                    string[] formatsTime = { "H.mm.ss", "h:mm:ss tt", "HH:mm:ss", "h:mm tt" };
                    if (TimeOnly.TryParseExact(row[2].ToString(), formatsTime, CultureInfo.InvariantCulture, DateTimeStyles.None, out TimeOnly result1))
                    {
                        Console.WriteLine($"Время: {result1}");
                    }
                    else
                    {
                        Console.WriteLine("Невозможно распознать время.");
                    }
                    //if (TimeOnly.TryParse(row[2].ToString(), out TimeOnly result1))
                    //{
                    //    Console.WriteLine($"Время: {result1}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("Невозможно распознать время.");
                    //}





                    // ������� �������� ������ � ������ TimeOnly
                    //if (TimeOnly.TryParseExact(row[1].ToString(), format1, out TimeOnly result2))
                    //{
                    //    Console.WriteLine($"�����: {result2}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("���������� ������������� ������ � �����.");
                    //}
                    if (TimeOnly.TryParseExact(row[1].ToString(), formatsTime, CultureInfo.InvariantCulture, DateTimeStyles.None, out TimeOnly result2))
                    {
                        Console.WriteLine($"Время: {result2}");
                    }
                    else
                    {
                        Console.WriteLine("Невозможно распознать время.");
                    }
                    //if (TimeOnly.TryParse(row[1].ToString(), out TimeOnly result2))
                    //{
                    //    Console.WriteLine($"Время: {result2}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("Невозможно распознать время.");
                    //}




                    //// ������ ����
                    //string format = "d.M.yyyy";

                    //// ������� �������� ������ � ������ DateOnly
                    //if (DateOnly.TryParseExact(row[5].ToString(), format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly result))
                    //{
                    //    Console.WriteLine($"����: {result}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("���������� ������������� ������ � ����.");
                    //}


                    string[] formatsDate = { "d.M.yyyy", "M/d/yyyy", "MMMM d, yyyy", "dd.MM.yyyy", "yyyy-MM-dd" };

                    if (DateOnly.TryParse(row[5].ToString(), out DateOnly result))
                    {
                        Console.WriteLine($"Дата: {result.ToString("d", CultureInfo.InvariantCulture)}");
                    }
                    else
                    {
                        Console.WriteLine("Невозможно распознать дату.");
                    }
                    //if (DateOnly.TryParseExact(row[5].ToString(), formatsDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly result))
                    //{
                    //    Console.WriteLine($"Дата: {result.ToString("d", CultureInfo.InvariantCulture)}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine("Невозможно распознать дату.");
                    //}

                    unavs.Add(
                        new Unavailability
                        {
                            UnavailabilityFrom = result2,
                            UnavailabilityBefore = result1,
                            Reason = row[3].ToString(),
                            UnavailabilityType = typeUnav,
                            Date = result,
                            Employee = emp
                        }
                        );

                    
                }

                await _context.Unavailabilitys.AddRangeAsync(unavs);

            }


           
            await _context.SaveChangesAsync();
               
            
        }



        public static Dictionary<string, List<DataRow>> GroupRowsByColumnValue(DataTable dataTable, string groupColumn)
        {
            return dataTable.AsEnumerable()
                            .GroupBy(row => row.Field<string>(groupColumn))
                            .ToDictionary(
                                group => group.Key,
                                group => group.ToList()
                            );
        }


        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                await LoadExcelFile();
                TempData["SuccessMessage"] = "File uploaded successfully.";
                return RedirectToPage("/EntryAccess/Index");
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = $"Error processing file: {ex.Message}";
                return Page();
            }

            //await LoadExcelFile();

            ////foreach (DataRow row in data)
            ////{
            ////    var employee = new Employee
            ////    {
            ////        // ��������� �������� ������ ������� �� ������
            ////        Name = row["NameColumnName"].ToString(), // �������� �� �������� ��� �������
            ////        Position = row["PositionColumnName"].ToString() // �������� �� �������� ��� �������
            ////                                                        // �������� ������ �������� �� �������������
            ////    };

            ////    _context.Employees.Add(employee);
            ////}

            //await _context.SaveChangesAsync();

            //// Setting success message
            //TempData["SuccessMessage"] = "File uploaded successfully.";


            //return RedirectToPage("/EntryAccess/Index");
        }
    }
}
