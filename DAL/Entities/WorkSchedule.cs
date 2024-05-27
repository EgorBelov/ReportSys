namespace ReportSys.DAL.Entities
{
    public class WorkSchedule
    {
        public int Id { get; set; }

        public TimeOnly Arrival { get; set; }
        public TimeOnly Exit { get; set; }
        public TimeOnly LunchStart { get; set; }
        public TimeOnly LunchEnd { get; set; }
        public int EmployeeId { get; set; }
        public Employee Employee { get; set; }
    }
}
