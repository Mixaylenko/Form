using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using System.Drawing;
using System.Drawing.Imaging;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using ServerForm.Models;

namespace ClientForm.Pages.Reports
{
    public class DetailsModel : PageModel
    {
        private readonly IWebHostEnvironment _env;
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        public ReportData Report { get; set; }
        public List<WorksheetModel> Worksheets { get; set; } = new();

        public DetailsModel(
            IHttpClientFactory httpClientFactory,
            IConfiguration configuration,
            IWebHostEnvironment env)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"];
            _env = env;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public async Task<IActionResult> OnGetAsync(int id)
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/{id}");
            if (!response.IsSuccessStatusCode) return NotFound();

            Report = await response.Content.ReadFromJsonAsync<ReportData>();
            if (Report == null) return NotFound();

            var filePath = Path.Combine(_env.WebRootPath, "uploads", Report.FileName);

            // Проверка существования файла
            if (!System.IO.File.Exists(filePath))
            {
                ModelState.AddModelError(string.Empty, "Файл отчета не найден");
                return Page();
            }

            using var package = new ExcelPackage(new FileInfo(filePath));

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var model = new WorksheetModel
                {
                    Name = worksheet.Name,
                    TableData = new List<List<string>>(),
                    Images = new List<ImageModel>()
                };

                // Обработка табличных данных
                if (worksheet.Dimension != null)
                {
                    // Ограничим размер для безопасности
                    int maxRows = Math.Min(worksheet.Dimension.End.Row, 100);
                    int maxCols = Math.Min(worksheet.Dimension.End.Column, 20);

                    for (int row = 1; row <= maxRows; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= maxCols; col++)
                        {
                            rowData.Add(worksheet.Cells[row, col].Text);
                        }
                        model.TableData.Add(rowData);
                    }
                }

                // Обработка графиков - ИСПРАВЛЕННАЯ ЧАСТЬ
                foreach (var drawing in worksheet.Drawings)
                {
                    if (drawing is ExcelChart chart) // Исправлено здесь
                    {
                        using var image = ConvertChartToImage(chart);
                        using var ms = new MemoryStream();
                        image.Save(ms, ImageFormat.Png);
                        model.Images.Add(new ImageModel
                        {
                            Data = ms.ToArray(),
                            AltText = chart.Name
                        });
                    }
                    else if (drawing is ExcelPicture picture)
                    {
                        // Обработка обычных изображений
                        using var ms = new MemoryStream();
                        picture.Image.Save(ms, ImageFormat.Png);
                        model.Images.Add(new ImageModel
                        {
                            Data = ms.ToArray(),
                            AltText = picture.Name
                        });
                    }
                }

                Worksheets.Add(model);
            }

            return Page();
        }

        // Исправленная сигнатура метода
        private Bitmap ConvertChartToImage(ExcelChart chart) // Исправлено здесь
        {
            var bitmap = new Bitmap(600, 400);
            using var g = Graphics.FromImage(bitmap);
            g.Clear(Color.White);
            g.DrawString(chart.Name, new Font("Arial", 14), Brushes.Black, 10, 10);

            // Простая визуализация для примера
            g.DrawRectangle(Pens.Black, 50, 50, 500, 300);
            g.DrawString("Chart Placeholder", new Font("Arial", 20), Brushes.Blue, 100, 150);

            return bitmap;
        }

        public class WorksheetModel
        {
            public string Name { get; set; }
            public List<List<string>> TableData { get; set; }
            public List<ImageModel> Images { get; set; }
        }

        public class ImageModel
        {
            public byte[] Data { get; set; }
            public string AltText { get; set; }
        }
    }
}