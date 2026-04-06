using Microsoft.AspNetCore.Mvc;
using ReportAPI.Models;
using ReportAPI.Repositories;
using ClosedXML.Excel;

namespace ReportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController : ControllerBase
    {
        private readonly IReportRepository _repository;

        public ReportsController(IReportRepository repository)
        {
            _repository = repository;
        }

        [HttpGet]
        public async Task<IActionResult> GetReports()
        {
            try
            {
                var reports = await _repository.GetAllReportsAsync();
                return Ok(reports);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetReport(int id)
        {
            try
            {
                var report = await _repository.GetReportByIdAsync(id);
                if (report == null)
                    return NotFound();

                return Ok(report);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("{id}/execute")]
        public async Task<IActionResult> ExecuteReport(int id)
        {
            try
            {
                var report = await _repository.GetReportByIdAsync(id);
                if (report == null)
                    return NotFound();

                if (string.IsNullOrWhiteSpace(report.ConnectionString) || string.IsNullOrWhiteSpace(report.ReportQuery))
                    return BadRequest("Report missing connection string or query.");

                var data = await _repository.ExecuteReportQueryAsync(report.ReportQuery, report.ConnectionString);

                return Ok(new
                {
                    Data = data,
                    GridState = report.GridState,
                    ColumnDefinitions = report.ColumnDefinitions
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost("{id}/state")]
        public async Task<IActionResult> UpdateState(int id, [FromBody] UpdateStateRequest request)
        {
            try
            {
                var report = await _repository.GetReportByIdAsync(id);
                if (report == null)
                    return NotFound();
                    
                await _repository.UpdateGridStateAsync(id, request.GridState);
                return Ok();
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
    [HttpGet("{id}/export")]
    public async Task<IActionResult> ExportReport(int id, [FromServices] ReportAPI.Services.ExportService exportService)
    {
        try
        {
            var report = await _repository.GetReportByIdAsync(id);
            if (report == null)
                return NotFound();

            if (string.IsNullOrWhiteSpace(report.ConnectionString) || string.IsNullOrWhiteSpace(report.ReportQuery))
                return BadRequest("Report missing connection string or query.");

            var data = await _repository.ExecuteReportQueryAsync(report.ReportQuery, report.ConnectionString);

            // Convert IEnumerable<dynamic> to IEnumerable<IDictionary<string, object>>
            var parsedData = data.Cast<IDictionary<string, object>>();

            IEnumerable<ColumnDefinition>? columnDefinitions = null;
            if (!string.IsNullOrWhiteSpace(report.ColumnDefinitions))
            {
                try
                {
                    columnDefinitions = System.Text.Json.JsonSerializer.Deserialize<IEnumerable<ColumnDefinition>>(
                        report.ColumnDefinitions, 
                        new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                }
                catch
                {
                    // Fallback to null if deserialization fails
                }
            }

            var excelBytes = exportService.GenerateExcel(parsedData, report.GridState, columnDefinitions);

            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{report.ReportName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }
        catch (Exception ex)
        {
            return StatusCode(500, ex.Message);
        }
    }

    [HttpGet("{id}/export-csv")]
    public async Task<IActionResult> ExportReportCsv(int id, [FromServices] ReportAPI.Services.ExportService exportService)
    {
        try
        {
            var report = await _repository.GetReportByIdAsync(id);
            if (report == null)
                return NotFound();

            if (string.IsNullOrWhiteSpace(report.ConnectionString) || string.IsNullOrWhiteSpace(report.ReportQuery))
                return BadRequest("Report missing connection string or query.");

            var data = await _repository.ExecuteReportQueryAsync(report.ReportQuery, report.ConnectionString);
            var parsedData = data.Cast<IDictionary<string, object>>();

            IEnumerable<ColumnDefinition>? columnDefinitions = null;
            if (!string.IsNullOrWhiteSpace(report.ColumnDefinitions))
            {
                try
                {
                    columnDefinitions = System.Text.Json.JsonSerializer.Deserialize<IEnumerable<ColumnDefinition>>(
                        report.ColumnDefinitions, 
                        new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                }
                catch
                {
                }
            }

            var csvBytes = exportService.GenerateCsv(parsedData, report.GridState, columnDefinitions);

            return File(csvBytes, "text/csv", $"{report.ReportName}_{DateTime.Now:yyyyMMddHHmmss}.csv");
        }
        catch (Exception ex)
        {
            return StatusCode(500, ex.Message);
        }
    }
}

public class UpdateStateRequest
{
    public string GridState { get; set; } = string.Empty;
}
}
