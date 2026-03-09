namespace ReportAPI.Models
{
    public class Report
    {
        public int ReportID { get; set; }
        public string ReportName { get; set; } = string.Empty;
        public string ReportQuery { get; set; } = string.Empty;
        public string? ReportConfig { get; set; }
        public string? GridState { get; set; }
        public string? ConnectionString { get; set; }
    }
}
