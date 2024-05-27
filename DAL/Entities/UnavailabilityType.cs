namespace ReportSys.DAL.Entities
{
    public class UnavailabilityType
    {
        public int Id { get; set; }
        public string? Name { get; set; }

        public List<Unavailability> Unavailabilitys { get; set; } = new List<Unavailability>();
    }
}
