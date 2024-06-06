using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.DAL.Entities;
using ReportSys.Pages.Services;
using System.Globalization;
using System.Xml.Linq;

namespace ReportSys.Pages.PageAccess0
{
    public class IndexModel : ServicesPage
    {
        private readonly ReportSysContext _context;


        public string _id { get; set; }
        public string _name { get; set; }

        public IndexModel(ReportSysContext context)
        {
            _context = context; 
        }

        public async Task<IActionResult> OnGet()
        {
            _id = HttpContext.Session.GetString("EmployeeNumber");

            var employee = await _context.Employees
                .Include(e => e.Department)
                .FirstOrDefaultAsync(e => e.Id.ToString() == _id);

            _name = employee.FirstName + " " + employee.SecondName + " " + employee.LastName;

            return Page();
        }
        public async Task<IActionResult> OnPostAsync(DateOnly startDate, DateOnly endDate)
        {
            var employeeNumber = HttpContext.Session.GetString("EmployeeNumber");
           
            List<string> employeeNumbers = new List<string>();
            employeeNumbers.Add(employeeNumber);
            return await CreateXlsxFirst(_context, employeeNumbers, startDate, endDate);
        }
    }
}
