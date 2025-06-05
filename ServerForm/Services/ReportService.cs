using ServerForm.Interfaces;
using ServerForm.Models;
using OfficeOpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml.Drawing.Chart;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OfficeOpenXml.Drawing;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ServerForm.Services
{
    public class ReportService : IReportService
    {
        private readonly DatabaseContext _context;
        private readonly IWebHostEnvironment _env;
        private readonly string _uploadsPath;

        public ReportService(DatabaseContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
            _uploadsPath = Path.Combine(_env.WebRootPath, "uploads");
            Directory.CreateDirectory(_uploadsPath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public async Task<ReportData> CreateReportAsync(ReportData report, Stream fileStream)
        {
            var extension = !string.IsNullOrEmpty(report.FileName) 
                ? Path.GetExtension(report.FileName) 
                : "";
            var fileName = $"{Guid.NewGuid()}{extension}";
            var filePath = Path.Combine(_uploadsPath, fileName);

            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                await fileStream.CopyToAsync(fs);
            }

            report.FilePath = filePath;
            report.FileName = fileName;

            _context.ReportDatas.Add(report);
            await _context.SaveChangesAsync();

            return report;
        }

        public async Task<IEnumerable<ReportData>> GetAllReportsAsync()
        {
            return await _context.ReportDatas.ToListAsync();
        }

        public async Task<ReportData> GetReportAsync(int id)
        {
            return await _context.ReportDatas.FindAsync(id);
        }

        public async Task<ReportData> UpdateReportAsync(int id, ReportData report, Stream fileStream = null, string originalFileName = null)
        {
            var existing = await _context.ReportDatas.FindAsync(id);
            if (existing == null) return null;

            // Всегда обновляем название отчета
            existing.Name = report.Name;

            if (fileStream != null)
            {
                // Удаляем старый файл
                if (!string.IsNullOrEmpty(existing.FilePath) && System.IO.File.Exists(existing.FilePath))
                {
                    System.IO.File.Delete(existing.FilePath);
                }

                // Сохраняем оригинальное расширение из нового файла
                var extension = !string.IsNullOrEmpty(originalFileName)
                    ? Path.GetExtension(originalFileName)
                    : "";

                var fileName = $"{Guid.NewGuid()}{extension}";
                var filePath = Path.Combine(_uploadsPath, fileName);

                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    await fileStream.CopyToAsync(fs);
                }

                // Обновляем путь и имя файла
                existing.FilePath = filePath;
                existing.FileName = fileName;
            }
            // Если файл не меняется, оставляем текущее имя и путь

            await _context.SaveChangesAsync();
            return existing;
        }

        public async Task DeleteReportAsync(int id)
        {
            var report = await _context.ReportDatas.FindAsync(id);
            if (report == null) return;

            if (!string.IsNullOrEmpty(report.FilePath) && System.IO.File.Exists(report.FilePath))
            {
                System.IO.File.Delete(report.FilePath);
            }

            _context.ReportDatas.Remove(report);
            await _context.SaveChangesAsync();
        }

        public async Task<byte[]> ConvertToWordAsync(int id)
        {
            var report = await GetReportAsync(id);
            if (report == null) return null;

            using var excelStream = new FileStream(report.FilePath, FileMode.Open, FileAccess.Read);
            using var excelPackage = new ExcelPackage(excelStream);
            using var wordStream = new MemoryStream();

            using (var wordDocument = WordprocessingDocument.Create(wordStream, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                foreach (var worksheet in excelPackage.Workbook.Worksheets.Where(w => w.Dimension != null))
                {
                    // Заголовок листа
                    var titlePara = new Paragraph(new Run(new Text(worksheet.Name)));
                    titlePara.ParagraphProperties = new ParagraphProperties(
                        new ParagraphStyleId() { Val = "Heading1" }
                    );
                    body.AppendChild(titlePara);

                    // Таблица
                    var table = new Table();
                    var tableProps = new TableProperties(
                        new TableBorders(
                            new TopBorder() { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                            new RightBorder() { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4 }
                        )
                    );
                    table.AppendChild(tableProps);

                    int rowCount = Math.Min(worksheet.Dimension.Rows, 100);
                    int colCount = Math.Min(worksheet.Dimension.Columns, 20);

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var tableRow = new TableRow();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cell = new TableCell();
                            cell.Append(new Paragraph(new Run(new Text(worksheet.Cells[row, col].Text))));
                            tableRow.Append(cell);
                        }
                        table.Append(tableRow);
                    }
                    body.Append(table);

                    // Изображения графиков
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            using var ms = new MemoryStream();
                            picture.Image.Save(ms, ImageFormat.Png);
                            await AddImageToDocumentAsync(mainPart, body, ms.ToArray());
                        }
                        else if (drawing is ExcelChart chart)
                        {
                            using var chartImage = ConvertChartToImage(chart);
                            using var imageStream = new MemoryStream();
                            chartImage.Save(imageStream, ImageFormat.Png);
                            await AddImageToDocumentAsync(mainPart, body, imageStream.ToArray());
                        }
                    }
                    // Разрыв страницы
                    body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }
                mainPart.Document.Save();
            }

            return wordStream.ToArray();
        }

        private Bitmap ConvertChartToImage(ExcelChart chart)
        {
            // В реальном проекте используйте библиотеку для рендеринга
            var bitmap = new Bitmap(800, 600);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.Clear(System.Drawing.Color.White);
            graphics.DrawString($"Chart: {chart.Name}",
                new System.Drawing.Font("Arial", 16),
                Brushes.Black,
                new PointF(50, 50));
            return bitmap;
        }

        private async Task AddImageToDocumentAsync(
            MainDocumentPart mainPart,
            Body body,
            byte[] imageBytes)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using var stream = new MemoryStream(imageBytes);
            imagePart.FeedData(stream);

            var imageId = mainPart.GetIdOfPart(imagePart);

            var element = new DocumentFormat.OpenXml.Office.Drawing.Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5000000L, Cy = 3000000L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = 1U, Name = "Chart" },
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip() { Embed = imageId }
                                ),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 5000000L, Cy = 3000000L }
                                    ),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                    )
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
            );

            body.Append(new Paragraph(new Run(element)));
        }
    }
}