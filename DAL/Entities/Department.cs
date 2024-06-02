namespace ReportSys.DAL.Entities
{
    public class Department
    {
        public int Id { get; set; }
        public string? Name { get; set; }

        public ICollection<Hierarchy> UpperDepartments { get; set; } = new List<Hierarchy>();
        public ICollection<Hierarchy> LowerDepartments { get; set; } = new List<Hierarchy>();

        public List<Employee> Employees { get; set; } = new List<Employee>();


    }
}
