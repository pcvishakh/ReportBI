using ReportAPI.Models;

namespace ReportAPI.Repositories
{
    public interface IReportRepository
    {
        Task<IEnumerable<Report>> GetAllReportsAsync();
        Task<Report?> GetReportByIdAsync(int id);
        Task<IEnumerable<dynamic>> ExecuteReportQueryAsync(string query, string connectionString);
        Task UpdateGridStateAsync(int id, string gridState);
    }
}
