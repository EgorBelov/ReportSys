namespace ReportSys.DAL.Entities
{
    public class Hierarchy
    {
        public int Id { get; set; }

       
        public int UpperDepartmentId { get; set; }
        public Department UpperDepartment { get; set; }

        public int LowerDepartmentId { get; set; }
        public Department LowerDepartment { get; set; }


    }
}
