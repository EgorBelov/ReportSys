namespace ReportSys.DAL.Entities
{
    public class Event
    {
        public int Id { get; set; }

        public string? Territory { get; set; }

        public DateOnly Date { get; set; }  
        public TimeOnly Time { get; set; }

        public int? EventTypeId { get; set; }      // внешний ключ
        public EventType? EventType { get; set; }    // навигационное свойство

        public int EmployeeId { get; set; }
        public Employee Employee { get; set; }
    }
}
