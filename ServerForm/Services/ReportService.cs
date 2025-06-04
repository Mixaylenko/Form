using System.Text;
using OfficeOpenXml;
using Microsoft.EntityFrameworkCore;
using ServerForm.Models;
using System.IO;
using ServerForm.Interfaces;
using OfficeOpenXml.Drawing;
using Microsoft.AspNetCore.Mvc;

namespace ServerForm.Services
{
    public class ReportService : IReportService
    {
        private readonly DatabaseContext _context;
        private readonly IWebHostEnvironment _hostingEnvironment;
        private readonly string _uploadsPath;

        static ReportService()
        {
            // Установка лицензии EPPlus (для версий 5+)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Альтернативный вариант для версий 4.x
            // ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public ReportService(DatabaseContext context, IWebHostEnvironment hostingEnvironment)
        {
            try
            {
                _context = context ?? throw new ArgumentNullException(nameof(context));
                _hostingEnvironment = hostingEnvironment ?? throw new ArgumentNullException(nameof(hostingEnvironment));
                _uploadsPath = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");

                if (!Directory.Exists(_uploadsPath))
                {
                    Directory.CreateDirectory(_uploadsPath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to initialize ReportService", ex);
            }
        }

        // Остальные методы класса остаются без изменений
        public async Task<ReportData> CreateReportAsync(string name, IFormFile file)
        {
            if (file == null || file.Length == 0)
                throw new ArgumentException("No file uploaded.");

            var uniqueFileName = $"{Guid.NewGuid()}{Path.GetExtension(file.FileName)}";
            var filePath = Path.Combine(_uploadsPath, uniqueFileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(fileStream);
            }

            var reportData = new ReportData
            {
                FileName = file.FileName,
                FilePath = filePath,
                Name = name
            };

            _context.ReportDatas.Add(reportData);
            await _context.SaveChangesAsync();

            return reportData;
        }

        public async Task DeleteReportAsync(int id)
        {
            var reportData = await _context.ReportDatas.FindAsync(id)
                ?? throw new KeyNotFoundException($"Report with ID {id} not found.");

            if (File.Exists(reportData.FilePath))
            {
                File.Delete(reportData.FilePath);
            }

            _context.ReportDatas.Remove(reportData);
            await _context.SaveChangesAsync();
        }

        public async Task<ReportData> GetReportAsync(int id)
        {
            return await _context.ReportDatas.FindAsync(id)
                ?? throw new KeyNotFoundException($"Report with ID {id} not found.");
        }

        public async Task<IEnumerable<ReportData>> GetAllReportsAsync()
        {
            return await _context.ReportDatas.AsNoTracking().ToListAsync();
        }

        public async Task<ReportData> UpdateReportAsync(ReportData reportData)
        {
            _context.Entry(reportData).State = EntityState.Modified;
            await _context.SaveChangesAsync();
            return reportData;
        }

        public string GetExcelContent(ReportData reportData)
        {
            if (string.IsNullOrEmpty(reportData.FilePath) || !File.Exists(reportData.FilePath))
                throw new FileNotFoundException("Report file not found.");

            var sb = new StringBuilder();
            var fileInfo = new FileInfo(reportData.FilePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet == null) continue;

                    sb.AppendLine($"=== Лист: {worksheet.Name} ===");

                    if (worksheet.Dimension == null)
                    {
                        sb.AppendLine("(пустой лист)");
                        continue;
                    }

                    // Заголовки
                    AppendExcelRow(sb, worksheet, 1);

                    // Данные
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        AppendExcelRow(sb, worksheet, row);
                    }

                    // Графики
                    AppendDrawingsInfo(sb, worksheet);

                    sb.AppendLine(new string('=', 50));
                }
            }

            return sb.ToString();
        }

        public FileStream GetExcelFileStream(ReportData reportData)
        {
            if (string.IsNullOrEmpty(reportData.FilePath) || !File.Exists(reportData.FilePath))
                throw new FileNotFoundException("Report file not found.");

            return new FileStream(reportData.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        }

        private void AppendExcelRow(StringBuilder sb, ExcelWorksheet worksheet, int row)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                sb.Append($"{worksheet.Cells[row, col].Text?.Trim()}\t");
            }
            sb.AppendLine();
        }

        private void AppendDrawingsInfo(StringBuilder sb, ExcelWorksheet worksheet)
        {
            if (worksheet.Drawings.Count == 0) return;

            sb.AppendLine($"\nГрафики на листе ({worksheet.Drawings.Count}):");

            foreach (ExcelDrawing drawing in worksheet.Drawings)
            {
                string drawingName = drawing.Name;
                string drawingType = drawing.GetType().Name;

                sb.AppendLine($"- {drawingName} ({drawingType})");
            }
        }
    }
}