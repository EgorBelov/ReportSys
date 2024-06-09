using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using ReportSys.DAL;
using ReportSys.Pages.Services;
using System.Threading.Tasks;

namespace ReportSys.Pages.EntryAccess
{
    public class IndexModel : ServicesPage
    {
        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public List<SelectListItem> AllEmployeeList { get; set; }


        [BindProperty]
        public int SelectedEmployeeId { get; set; }


        [BindProperty]
        public string EmployeeNumber { get; set; }

        public async Task<IActionResult> OnGetAsync()
        {


            await GetAllEmployeesAsync(_context);

            AllEmployeeList = AllEmployeesSL.ToList();

            return Page();

        }

        public async Task<IActionResult> OnPostAsync()
        {
            //if (string.IsNullOrEmpty(EmployeeNumber))
            //{
            //    ModelState.AddModelError(string.Empty, "Табельный номер обязателен.");
            //    return Page();
            //}

            var employee = await _context.Employees
                .Include(e => e.Position)
                .FirstOrDefaultAsync(e => e.Id.ToString() == SelectedEmployeeId.ToString());

            if (employee == null)
            {
                ModelState.AddModelError(string.Empty, "Табельный номер не найден.");
                return Page();
            }

            //HttpContext.Session.SetString("EmployeeNumber", EmployeeNumber);
            HttpContext.Session.SetString("EmployeeNumber", SelectedEmployeeId.ToString());

            switch (employee.Position.AccessLevel)
            {
                case 0:
                    return RedirectToPage("/PageAccess0/Index");
                case 1:
                    return RedirectToPage("/PageAccess1/Index");
                default:
                    ModelState.AddModelError(string.Empty, "Неизвестный доступ.");
                    return Page();
            }
        }
    }
}
