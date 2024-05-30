namespace ReportSys.DAL.Entities
{
    public class WorkSchedule
    {
        public int Id { get; set; }

        public TimeOnly Arrival { get; set; }
        public TimeOnly Exit { get; set; }
        public TimeOnly LunchStart { get; set; }
        public TimeOnly LunchEnd { get; set; }
        
        public List<Employee> Employees { get; set; } = new List<Employee>();


        // Метод для получения строки расписания
        public string GetScheduleString()
        {
            // Вычисляем продолжительность обеденного перерыва
            var lunchDuration = LunchEnd - LunchStart;

            // Формируем строку
            string schedule = $"с {Arrival:HH:mm} \nпо {Exit:HH:mm} \nобеденный перерыв \n{lunchDuration.TotalMinutes} мин";

            return schedule;
        }
    }
}
