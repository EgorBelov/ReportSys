using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.Pages.Services;
using System.Globalization;

namespace ReportSys.Pages.PageAccess0
{
    public class IndexModel : ServicesPage
    {
        private readonly ReportSysContext _context;


        public string _id { get; set; }

        public IndexModel(ReportSysContext context)
        {
            _context = context; 
        }

        public void OnGet()
        {
            _id = HttpContext.Session.GetString("EmployeeNumber");
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
