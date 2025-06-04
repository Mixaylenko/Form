using Microsoft.AspNetCore.Mvc;
using ServerForm.Interfaces;
using ServerForm.Models;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace ServerForm.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportsController : ControllerBase
    {
        private readonly IReportService _reportService;

        public ReportsController(IReportService reportService)
        {
            _reportService = reportService;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<ReportData>>> GetReports()
        {
            return Ok(await _reportService.GetAllReportsAsync());
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<ReportData>> GetReport(int id)
        {
            var report = await _reportService.GetReportAsync(id);
            return report != null ? Ok(report) : NotFound();
        }

        [HttpPost]
        public async Task<ActionResult<ReportData>> CreateReport(
            [FromForm] string name,
            IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File is required");

            var report = new ReportData { Name = name, FileName = file.FileName };
            using var stream = file.OpenReadStream();
            var created = await _reportService.CreateReportAsync(report, stream);
            return CreatedAtAction(nameof(GetReport), new { id = created.Id }, created);
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateReport(
            int id,
            [FromForm] string name,
            IFormFile file = null)
        {
            var report = new ReportData { Name = name };
            Stream stream = file != null ? file.OpenReadStream() : null;

            var updated = await _reportService.UpdateReportAsync(id, report, stream);
            return updated != null ? NoContent() : NotFound();
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteReport(int id)
        {
            await _reportService.DeleteReportAsync(id);
            return NoContent();
        }

        [HttpGet("{id}/download")]
        public async Task<IActionResult> DownloadReport(int id)
        {
            var report = await _reportService.GetReportAsync(id);
            if (report == null) return NotFound();

            var fileBytes = await System.IO.File.ReadAllBytesAsync(report.FilePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", report.FileName);
        }

        [HttpGet("{id}/convert-to-word")]
        public async Task<IActionResult> ConvertToWord(int id)
        {
            var wordBytes = await _reportService.ConvertToWordAsync(id);
            if (wordBytes == null) return NotFound();

            var report = await _reportService.GetReportAsync(id);
            var fileName = Path.GetFileNameWithoutExtension(report.FileName) + ".docx";
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
    }
}