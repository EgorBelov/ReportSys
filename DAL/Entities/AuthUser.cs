using Microsoft.AspNetCore.Identity;

namespace ReportSys.DAL.Entities
{
    public class AuthUser 
    {
        public int Id { get; set; }
        
        public int AccessLevel { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public Employee Employee { get; set; }
        
        //public AuthUser(string email, string password)
        //{
        //    var passwordHasher = new PasswordHasher<AuthUser>();
        //    var hashedPassword = passwordHasher.HashPassword(this, password);

        //    base.Email = email;
        //    base.UserName = email;
        //    base.PasswordHash = hashedPassword;
        //}
        //public AuthUser()
        //{
        //}
    }
}
