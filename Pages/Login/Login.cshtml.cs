using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ReportSys.DAL;
using System.Security.Claims;

namespace ReportSys.Pages.Login
{
    public class LoginModel : PageModel
    {

        private readonly ReportSysContext _context;

        public LoginModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public string Username { get; set; }

        [BindProperty]
        public string Password { get; set; }

        public string ErrorMessage { get; set; }

        public async Task<IActionResult> OnPostAsync()
        {
            //var user = _context.AuthUsers.FirstOrDefault(u => u.Login == Username && u.Password == Password);
            //if (user == null)
            //{
            //    ErrorMessage = "Invalid username or password";
            //    return Page();
            //}

            //var claims = new List<Claim>
            //{
            //    new Claim(ClaimTypes.Name, user.Login),
            //    new Claim("AccessLevel", user.AccessLevel.ToString())
            //};

            //var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
            //var authProperties = new AuthenticationProperties { IsPersistent = true };

            //await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), authProperties);

            //string redirectPage = user.AccessLevel switch
            //{
            //    0 => "/PageAccess0/Index",
            //    1 => "/PageAccess1/Index",
            //    2 => "/PageAccess2/Index",
            //    _ => "/Login"
            //};

            return RedirectToPage();
        }
    }
}

