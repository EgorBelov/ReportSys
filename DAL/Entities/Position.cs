namespace ReportSys.DAL.Entities
{
    public class Position
    {
        public int Id { get; set; }
        public string? Name { get; set; }

        public int  AccessLevel { get; set; }

        public List<Employee> Employees { get; set; } = new List<Employee>();

    }
}
