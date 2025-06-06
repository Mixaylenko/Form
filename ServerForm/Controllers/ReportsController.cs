﻿using Microsoft.AspNetCore.Mvc;
using ServerForm.Interfaces;
using ServerForm.Models;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;
using System.Drawing.Imaging;
using System.Drawing;
using System.ComponentModel.DataAnnotations;

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

        [HttpPut("update/{id}")]
        public async Task<IActionResult> UpdateReport(int id, [FromForm] UpdateReportRequest request)
        {
            // Проверяем только обязательное поле Name
            if (string.IsNullOrEmpty(request.Name))
            {
                ModelState.AddModelError("Name", "Название отчета обязательно");
                return BadRequest(ModelState);
            }

            var reportData = new ReportData
            {
                Name = request.Name
            };

            Stream fileStream = null;
            string fileName = null;

            // Файл не обязателен - обрабатываем только если предоставлен
            if (request.File != null && request.File.Length > 0)
            {
                fileStream = request.File.OpenReadStream();
                fileName = request.File.FileName;
            }

            var updatedReport = await _reportService.UpdateReportAsync(
                id,
                reportData,
                fileStream,
                fileName
            );

            if (updatedReport == null)
                return NotFound();

            return Ok(updatedReport);
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteReport(int id)
        {
            await _reportService.DeleteReportAsync(id);
            return NoContent();
        }
        [HttpGet("{id}/preview")]
        public async Task<ActionResult<ReportPreview>> PreviewReport(int id)
        {
            var report = await _reportService.GetReportAsync(id);
            if (report == null) return NotFound();

            var ext = Path.GetExtension(report.FileName).ToLower();
            if (ext != ".xlsx" && ext != ".xls")
            {
                return BadRequest("Preview is available only for Excel files");
            }

            var worksheets = new List<WorksheetPreview>();

            using (var excelStream = new FileStream(report.FilePath, FileMode.Open, FileAccess.Read))
            using (var excelPackage = new ExcelPackage(excelStream))
            {
                foreach (var worksheet in excelPackage.Workbook.Worksheets.Where(w => w.Dimension != null))
                {
                    var wsPreview = new WorksheetPreview
                    {
                        Name = worksheet.Name,
                        TableData = new List<List<string>>(),
                        Images = new List<ImagePreview>()
                    };

                    // Чтение табличных данных
                    int rowCount = Math.Min(worksheet.Dimension.Rows, 100);
                    int colCount = Math.Min(worksheet.Dimension.Columns, 20);

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            rowData.Add(worksheet.Cells[row, col].Text);
                        }
                        wsPreview.TableData.Add(rowData);
                    }

                    worksheets.Add(wsPreview);
                }
            }

            return new ReportPreview
            {
                ReportId = id,
                ReportName = report.Name,
                FileName = report.FileName,
                Worksheets = worksheets
            };
        }

        // Вспомогательные классы
        public class ReportPreview
        {
            public int ReportId { get; set; }
            public string ReportName { get; set; }
            public string FileName { get; set; }
            public List<WorksheetPreview> Worksheets { get; set; }
        }

        public class WorksheetPreview
        {
            public string Name { get; set; }
            public List<List<string>> TableData { get; set; }
            public List<ImagePreview> Images { get; set; }
        }
        public class UpdateReportRequest
        {
            [Required(ErrorMessage = "Название отчета обязательно")]
            public string Name { get; set; }
            [Required(ErrorMessage = "Файл обязательно")]
            public IFormFile File { get; set; }
        }
        public class ImagePreview
        {
            public string Name { get; set; }
            public string Format { get; set; }
            public byte[] ImageData { get; set; }
        }

        // Сервис для рендеринга графиков (вынесен отдельно)
        public class ChartRenderer
        {
            public Bitmap RenderChart(ExcelChart chart)
            {
                // Реализация рендеринга графиков
                var bitmap = new Bitmap(800, 600);
                using var graphics = Graphics.FromImage(bitmap);
                graphics.Clear(Color.White);
                graphics.DrawString($"Chart: {chart.Name}",
                    new Font("Arial", 16),
                    Brushes.Black,
                    new PointF(50, 50));
                return bitmap;
            }
        }
        [HttpGet("download/{id}")]
        public async Task<IActionResult> DownloadReport(int id)
        {
            var report = await _reportService.GetReportAsync(id);
            if (report == null) return NotFound();

            var fileBytes = await System.IO.File.ReadAllBytesAsync(report.FilePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", report.FileName);
        }

        [HttpGet("convert-to-word/{id}")]
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