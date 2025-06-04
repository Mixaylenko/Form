using Microsoft.AspNetCore.Mvc;
using ServerForm.Models;
using ServerForm.Interfaces;
using System.Text;

namespace ServerForm.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController : ControllerBase
    {
        private readonly IWebHostEnvironment _env;
        private readonly IReportService _reportService;
        private readonly DatabaseContext _context;
        private readonly string _uploadsPath;

        public ReportsController(
            IWebHostEnvironment env,
            IReportService reportService,
            DatabaseContext context)
        {
            _reportService = reportService ?? throw new ArgumentNullException(nameof(reportService));
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _env = env ?? throw new ArgumentNullException(nameof(env));

            _uploadsPath = Path.Combine(_env.WebRootPath, "uploads");
            Directory.CreateDirectory(_uploadsPath);
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<ReportData>>> GetReports()
        {
            var reports = await _reportService.GetAllReportsAsync();
            return Ok(reports);
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<ReportData>> GetReport(int id)
        {
            try
            {
                var report = await _reportService.GetReportAsync(id);
                return Ok(report);
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Report with ID {id} not found.");
            }
        }

        [HttpPost]
        [DisableRequestSizeLimit]
        public async Task<ActionResult<ReportData>> CreateReport([FromForm] ReportModel model)
        {
            if (model?.File == null || model.File.Length == 0)
                return BadRequest("No file uploaded.");

            try
            {
                var report = await _reportService.CreateReportAsync(model.Name, model.File);
                return CreatedAtAction(
                    nameof(GetReport),
                    new { id = report.Id },
                    report);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateReport(int id, [FromBody] ReportData reportData)
        {
            if (id != reportData.Id)
                return BadRequest("ID mismatch");

            try
            {
                await _reportService.UpdateReportAsync(reportData);
                return NoContent();
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Report with ID {id} not found.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteReport(int id)
        {
            try
            {
                await _reportService.DeleteReportAsync(id);
                return NoContent();
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Report with ID {id} not found.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet("view/{id}")]
        public async Task<ActionResult> ViewExcel(int id)
        {
            try
            {
                var report = await _reportService.GetReportAsync(id);
                var content = _reportService.GetExcelContent(report);
                return Content(content, "text/plain", Encoding.UTF8);
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Report with ID {id} not found.");
            }
            catch (FileNotFoundException)
            {
                return NotFound("File not found.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet("Download/{id}")]
        public async Task<IActionResult> Download(int id)
        {
            try
            {
                var report = await _reportService.GetReportAsync(id);
                var fileStream = _reportService.GetExcelFileStream(report);

                return File(
                    fileStream,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    report.FileName);
            }
            catch (KeyNotFoundException)
            {
                return NotFound($"Report with ID {id} not found.");
            }
            catch (FileNotFoundException)
            {
                return NotFound("File not found.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        public class ReportModel
        {
            public string Name { get; set; }
            public IFormFile File { get; set; }
        }
    }
}