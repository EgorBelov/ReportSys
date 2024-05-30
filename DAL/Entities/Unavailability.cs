namespace ReportSys.DAL.Entities
{
    public class Unavailability
    {

        public int Id { get; set; }

        public string? Reason { get; set; }

        public DateOnly Date { get; set; }
        public TimeOnly UnavailabilityFrom { get; set; }
        public TimeOnly UnavailabilityBefore { get; set; }

        public int? UnavailabilityTypeId { get; set; }      // внешний ключ
        public UnavailabilityType? UnavailabilityType { get; set; }    // навигационное свойство

        public int EmployeeId { get; set; }
        public Employee Employee { get; set; }
    }
}
