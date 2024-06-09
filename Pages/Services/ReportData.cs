using OfficeOpenXml;

namespace ReportSys.Pages.Services
{
    public class ReportData
    {
        public int Row { get; set; }
        public int MinSumE { get; set; }
        public int MinSumS { get; set; }
        public int PlusSumE { get; set; }
        public int PlusSumS { get; set; }
        public TimeSpan MinTimeS { get; set; }
        public TimeSpan MinTimeE { get; set; }
        public TimeSpan PlusTimeS { get; set; }
        public TimeSpan PlusTimeE { get; set; }
        public ExcelWorksheet Worksheet { get; set; }
    }

}
