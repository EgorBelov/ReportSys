using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using ReportSys.DAL;

namespace ReportSys.Pages.EntryAccess
{
    public class IndexModel : PageModel
    {
        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public string EmployeeNumber { get; set; }

        public void OnGet()
        {
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (string.IsNullOrEmpty(EmployeeNumber))
            {
                ModelState.AddModelError(string.Empty, "Табельный номер обязателен.");
                return Page();
            }

            var employee = await _context.Employees
             .Include(e => e.Position) // Включаем связанную таблицу должностей
             .FirstOrDefaultAsync(e => e.Id.ToString() == EmployeeNumber);

            if (employee == null)
            {
                ModelState.AddModelError(string.Empty, "Табельный номер не найден.");
                return Page();
            }


            HttpContext.Session.SetString("EmployeeNumber", EmployeeNumber);

            switch (employee.Position.AccessLevel)
            {
                case 0:
                    return RedirectToPage("/PageAccess0/Index");
                case 1:
                    return RedirectToPage("/PageAccess1/Index");
                case 2:
                    return RedirectToPage("/PageAccess2/Index");
                default:
                    ModelState.AddModelError(string.Empty, "Неизвестный доступ.");
                    return Page();
            }
        }



    }
}
