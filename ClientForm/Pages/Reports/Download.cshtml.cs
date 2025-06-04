using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing.Imaging;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using ClientForm.Models;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using Microsoft.Extensions.Configuration;
using ServerForm.Models;

namespace ClientForm.Pages.Reports
{
    public class DownloadModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        public DownloadModel(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<IActionResult> OnGetAsync(int id)
        {
            try
            {
                // 1. Получаем Excel файл с сервера
                var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/download/{id}");
                if (!response.IsSuccessStatusCode)
                    return NotFound();

                var report = await GetReportMetadataAsync(id);
                if (report == null)
                    return NotFound();

                // 2. Конвертируем Excel в Word
                using var excelStream = await response.Content.ReadAsStreamAsync();
                using var wordStream = new MemoryStream();

                await ConvertExcelToWordAsync(excelStream, wordStream);
                wordStream.Position = 0;

                // 3. Возвращаем Word документ
                return File(wordStream,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    $"{Path.GetFileNameWithoutExtension(report.FileName)}.docx");
            }
            catch (Exception ex)
            {
                return BadRequest($"Ошибка при конвертации: {ex.Message}");
            }
        }

        private async Task<ReportData> GetReportMetadataAsync(int id)
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/{id}");
            return response.IsSuccessStatusCode
                ? await response.Content.ReadFromJsonAsync<ReportData>()
                : null;
        }

        private async Task ConvertExcelToWordAsync(Stream excelStream, Stream wordStream)
        {
            using (var excelPackage = new ExcelPackage(excelStream))
            using (var wordDocument = WordprocessingDocument.Create(wordStream, WordprocessingDocumentType.Document, true))
            {
                // Основная часть документа
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                // Обрабатываем каждый лист
                foreach (var worksheet in excelPackage.Workbook.Worksheets.Where(w => w.Dimension != null))
                {
                    // Заголовок листа
                    AddWorksheetTitle(body, worksheet.Name);

                    // Таблица с данными
                    AddWorksheetTable(body, worksheet);

                    // Изображения и графики
                    await AddWorksheetDrawingsAsync(body, worksheet, mainPart);

                    // Разрыв страницы
                    body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }

                mainPart.Document.Save();
            }
        }

        private void AddWorksheetTitle(Body body, string title)
        {
            var paragraph = new Paragraph(
                new Run(
                    new Text(title),
                    new RunProperties(
                        new Bold(),
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "28" }
                    )
                )
            );
            body.AppendChild(paragraph);
        }

        private void AddWorksheetTable(Body body, ExcelWorksheet worksheet)
        {
            var table = new Table();

            // Стили таблицы
            var tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                    new RightBorder() { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4 }
                )
            );
            table.AppendChild(tableProperties);

            // Ограничим размер таблицы для безопасности
            int maxRows = Math.Min(worksheet.Dimension.Rows, 1000);
            int maxCols = Math.Min(worksheet.Dimension.Columns, 50);

            // Добавляем данные
            for (int row = 1; row <= maxRows; row++)
            {
                var tableRow = new TableRow();

                for (int col = 1; col <= maxCols; col++)
                {
                    var cell = new TableCell();
                    cell.AppendChild(new Paragraph(new Run(new Text(worksheet.Cells[row, col].Text))));
                    tableRow.AppendChild(cell);
                }

                table.AppendChild(tableRow);
            }

            body.AppendChild(table);
        }

        private async Task AddWorksheetDrawingsAsync(Body body, ExcelWorksheet worksheet, MainDocumentPart mainPart)
        {
            if (worksheet.Drawings.Count == 0) return;

            var drawingsParagraph = new Paragraph(new Run(new Text("Графики:")));
            body.AppendChild(drawingsParagraph);

            foreach (var drawing in worksheet.Drawings)
            {
                try
                {
                    if (drawing is ExcelPicture picture)
                    {
                        using var stream = new MemoryStream();
                        picture.Image.Save(stream, ImageFormat.Png);
                        await AddImageToDocumentAsync(mainPart, body, stream.ToArray(), "image/png");
                    }
                    else if (drawing is ExcelChart chart)
                    {
                        using var chartImage = await ConvertChartToImageAsync(chart);
                        using var imageStream = new MemoryStream();
                        chartImage.Save(imageStream, ImageFormat.Png);
                        await AddImageToDocumentAsync(mainPart, body, imageStream.ToArray(), "image/png");
                    }
                }
                catch (Exception ex)
                {
                    var errorParagraph = new Paragraph(new Run(new Text($"Ошибка обработки изображения: {ex.Message}")));
                    body.AppendChild(errorParagraph);
                }
            }
        }

        private async Task<Image> ConvertChartToImageAsync(ExcelChart chart)
        {
            using var bitmap = new Bitmap(800, 600);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(System.Drawing.Color.White);
                graphics.DrawString("Chart Placeholder",
                    new System.Drawing.Font("Arial", 20),
                    Brushes.Black,
                    new PointF(100, 100));
            }
            return new Bitmap(bitmap);
        }

        private async Task AddImageToDocumentAsync(
            MainDocumentPart mainPart,
            Body body,
            byte[] imageBytes,
            string contentType)
        {
            var imagePartType = contentType switch
            {
                "image/png" => ImagePartType.Png,
                "image/jpeg" => ImagePartType.Jpeg,
                "image/gif" => ImagePartType.Gif,
                _ => throw new NotSupportedException($"Unsupported image format: {contentType}")
            };

            var imagePart = mainPart.AddImagePart(imagePartType);
            await using (var stream = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(stream);
            }

            var imageId = mainPart.GetIdOfPart(imagePart);

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = 5000000L, Cy = 3000000L },
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties() { Id = 1U, Name = "Image" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Image" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip() { Embed = imageId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = 5000000L, Cy = 3000000L }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 });

            body.AppendChild(new Paragraph(new Run(drawing)));
        }
    }
}