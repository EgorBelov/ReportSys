namespace ReportSys.DAL.Entities
{
    public class Department
    {
        public int Id { get; set; }
        public string? Name { get; set; }

        public int? DivisionId { get; set; }      // внешний ключ
        public Division? Division { get; set; }    // навигационное свойство

        public List<Employee> Employees { get; set; } = new List<Employee>();

    }
}
