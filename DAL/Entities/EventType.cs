namespace ReportSys.DAL.Entities
{
    public class EventType
    {
        public int Id { get; set; }
        public string? Name { get; set; }

        public List<Event> Events { get; set; } = new List<Event>();
    }
}
