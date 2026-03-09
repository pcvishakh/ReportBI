using ReportAPI.Models;
using ReportAPI.Data;
using Dapper;
using Microsoft.Data.SqlClient;

namespace ReportAPI.Repositories
{
    public class ReportRepository : IReportRepository
    {
        private readonly DapperContext _context;

        public ReportRepository(DapperContext context)
        {
            _context = context;
        }

        public async Task<IEnumerable<Report>> GetAllReportsAsync()
        {
            var query = "SELECT ReportID, ReportName, ReportQuery, ReportConfig, GridState, ConnectionString FROM Reports";
            using var connection = _context.CreateConnection();
            var reports = await connection.QueryAsync<Report>(query);
            return reports;
        }

        public async Task<Report?> GetReportByIdAsync(int id)
        {
            var query = "SELECT ReportID, ReportName, ReportQuery, ReportConfig, GridState, ConnectionString FROM Reports WHERE ReportID = @Id";
            using var connection = _context.CreateConnection();
            var report = await connection.QuerySingleOrDefaultAsync<Report>(query, new { Id = id });
            return report;
        }

        public async Task<IEnumerable<dynamic>> ExecuteReportQueryAsync(string query, string connectionString)
        {
            using var connection = new SqlConnection(connectionString);
            var result = await connection.QueryAsync(query);
            return result;
        }

        public async Task UpdateGridStateAsync(int id, string gridState)
        {
            var query = "UPDATE Reports SET GridState = @GridState WHERE ReportID = @Id";
            using var connection = _context.CreateConnection();
            await connection.ExecuteAsync(query, new { GridState = gridState, Id = id });
        }
    }
}
