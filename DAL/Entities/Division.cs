namespace ReportSys.DAL.Entities
{
    public class Division
    {
        public int Id { get; set; }
        public string? Name { get; set; }

        public List<Department> Departments { get; set; } = new List<Department>();

        public List<Employee> Employees { get; set; } = new List<Employee>();


    }
}
