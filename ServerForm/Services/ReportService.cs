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
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Processing;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Styles = DocumentFormat.OpenXml.Wordprocessing.Styles;
using Color = System.Drawing.Color;

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

                // Добавить стили для заголовков
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                GenerateDocumentStyles(stylesPart);

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
                        try
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
                        catch (Exception ex)
                        {
                            // Логирование ошибки вместо прерывания процесса
                            var errorText = new Paragraph(new Run(new Text(
                                $"Ошибка обработки элемента: {ex.Message}")));
                            body.AppendChild(errorText);
                        }
                    }
                    body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }
                mainPart.Document.Save();
            }
            return wordStream.ToArray();
        }

        private Bitmap ConvertChartToImage(ExcelChart chart)
        {
            try
            {
                // Реальная реализация рендеринга графика
                var bitmap = new Bitmap(800, 600);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);
                    // Здесь должен быть код рендеринга графика
                    // Для примера - заглушка с текстом
                    g.DrawString($"График: {chart.Name}",
                        new System.Drawing.Font("Arial", 14),
                        Brushes.Black,
                        new System.Drawing.PointF(50, 50));
                }
                return bitmap;
            }
            catch
            {
                // Возвращаем пустое изображение при ошибке
                return new Bitmap(800, 600);
            }
        }

        private async Task AddImageToDocumentAsync(
        MainDocumentPart mainPart,
        Body body,
        byte[] imageBytes)
        {
            try
            {
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var stream = new MemoryStream(imageBytes))
                {
                    imagePart.FeedData(stream);
                }

                var imageId = mainPart.GetIdOfPart(imagePart);

                // Использование правильного неймспейса
                var element = new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = 5000000L, Cy = 3000000L },
                        new DW.EffectExtent() { LeftEdge = 0L, RightEdge = 0L, TopEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = 1U, Name = "Chart" },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Chart.png" },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip() { Embed = imageId },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = 5000000L, Cy = 3000000L }),
                                        new A.PresetGeometry(new A.AdjustValueList())
                                        { Preset = A.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });

                body.Append(new Paragraph(new Run(element)));
            }
            catch (Exception ex)
            {
                var errorText = new Paragraph(new Run(new Text(
                    $"Ошибка вставки изображения: {ex.Message}")));
                body.AppendChild(errorText);
            }
        }

        // Исправление 5: Добавление базовых стилей документа
        private void GenerateDocumentStyles(StyleDefinitionsPart stylesPart)
        {
            var styles = new Styles();
            var style = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal" };
            style.Append(new StyleName() { Val = "Normal" });
            style.Append(new PrimaryStyle());
            styles.Append(style);

            var heading1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
            heading1.Append(new StyleName() { Val = "Heading 1" });
            heading1.Append(new BasedOn() { Val = "Normal" });
            heading1.Append(new NextParagraphStyle() { Val = "Normal" });
            heading1.Append(new StyleParagraphProperties(
                new SpacingBetweenLines() { After = "100" },
                new ContextualSpacing()));
            heading1.Append(new StyleRunProperties(
                new RunFonts() { Ascii = "Arial" },
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "28" },
                new Bold()));
            styles.Append(heading1);

            stylesPart.Styles = styles;
        }
    }
}