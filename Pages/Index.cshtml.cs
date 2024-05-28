using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.DAL.Entities;
using System.Data;
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



        public async Task<DataTable> LoadExcelFile()
        {
            
            DataTable dataTable = new DataTable();


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

            

            return dataTable;
        }



        public async Task<IActionResult> OnPostAsync()
        {

            var data = LoadExcelFile();

            foreach (DataRow row in data)
            {
                var employee = new Employee
                {
                    // Заполняем свойства модели данными из строки
                    Name = row["NameColumnName"].ToString(), // Замените на реальное имя колонки
                    Position = row["PositionColumnName"].ToString() // Замените на реальное имя колонки
                                                                    // Добавьте другие свойства по необходимости
                };

                _context.Employees.Add(employee);
            }

            await _context.SaveChangesAsync();
            return RedirectToPage("/Index");
        }
    }


}