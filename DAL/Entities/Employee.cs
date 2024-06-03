namespace ReportSys.DAL.Entities
{
    public class Employee
    {
        public int Id { get; set; }

        public string? FirstName { get; set; }
        public string? SecondName { get; set; }
        public string? LastName { get; set; }
        
        public int WorkScheduleId { get; set; }
        public WorkSchedule WorkSchedule { get; set; }

        public int PositionId { get; set; }
        public Position Position { get; set; }

        //public int? AuthUserId { get; set; }
        //public AuthUser? AuthUser { get; set; }

        public int DepartmentId { get; set; }
        public Department Department { get; set; }

   
        public List<Event> Events { get; set; } = new List<Event>();
        public List<Unavailability> Unavailabilitys { get; set; } = new List<Unavailability>();
    }
}
