using ClientForm.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace ClientForm.Pages.Reports
{
    public class DetailsModel : PageModel
    {
        public HttpClient _httpClient { get; }

        [BindProperty(SupportsGet = true)]
        public int Id { get; set; }

        public ReportData Report { get; set; }
        public bool IsExcel { get; set; }
        public List<ExcelWorksheetModel> Worksheets { get; set; } = new();

        public DetailsModel(IHttpClientFactory httpClientFactory)
        {
            _httpClient = httpClientFactory.CreateClient("ServerAPI");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<IActionResult> OnGetAsync()
        {
            var response = await _httpClient.GetAsync($"api/reports/{Id}");
            if (!response.IsSuccessStatusCode)
            {
                return NotFound();
            }

            Report = await response.Content.ReadFromJsonAsync<ReportData>();
            IsExcel = Report.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                      Report.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);

            if (IsExcel)
            {
                await LoadExcelData();
            }

            return Page();
        }

        public async Task<IActionResult> OnPostConvertToWordAsync()
        {
            var response = await _httpClient.GetAsync($"api/reports/download/{Id}");
            if (!response.IsSuccessStatusCode)
            {
                return NotFound();
            }

            using var stream = await response.Content.ReadAsStreamAsync();
            using var package = new ExcelPackage(stream);

            // Создаем временный файл Word
            var tempWordFile = Path.GetTempFileName() + ".docx";

            // Конвертируем первый лист Excel в Word
            if (package.Workbook.Worksheets.Count > 0)
            {
                ConvertExcelToWord(package.Workbook.Worksheets[0], tempWordFile);
            }

            // Возвращаем файл пользователю
            var fileStream = new FileStream(tempWordFile, FileMode.Open);
            return File(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       Path.GetFileNameWithoutExtension(Report.FileName) + ".docx");
        }

        private void ConvertExcelToWord(ExcelWorksheet worksheet, string wordFilePath)
        {
            // Создаем Word документ
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
            {
                // Добавляем главную часть документа
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Создаем таблицу в Word
                if (worksheet.Dimension != null)
                {
                    Table wordTable = new Table();

                    int rowCount = Math.Min(worksheet.Dimension.Rows, 100);
                    int colCount = Math.Min(worksheet.Dimension.Columns, 20);

                    for (int row = 1; row <= rowCount; row++)
                    {
                        TableRow wordRow = new TableRow();

                        for (int col = 1; col <= colCount; col++)
                        {
                            TableCell wordCell = new TableCell();
                            wordCell.Append(new Paragraph(new Run(new Text(worksheet.Cells[row, col].Text))));
                            wordCell.Append(new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                            wordRow.Append(wordCell);
                        }

                        wordTable.Append(wordRow);
                    }

                    body.Append(wordTable);
                }

                mainPart.Document.Save();
            }
        }

        private async Task LoadExcelData()
        {
            try
            {
                var response = await _httpClient.GetAsync($"api/reports/download/{Id}");
                if (!response.IsSuccessStatusCode) return;

                using var stream = await response.Content.ReadAsStreamAsync();
                using var package = new ExcelPackage(stream);

                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var model = new ExcelWorksheetModel
                    {
                        Name = worksheet.Name,
                        Images = new List<ExcelImageModel>(),
                        TableData = new List<List<string>>()
                    };

                    // Получаем данные таблицы
                    if (worksheet.Dimension != null)
                    {
                        int rowCount = Math.Min(worksheet.Dimension.Rows, 100);
                        int colCount = Math.Min(worksheet.Dimension.Columns, 20);

                        for (int row = 1; row <= rowCount; row++)
                        {
                            var currentRow = new List<string>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                currentRow.Add(worksheet.Cells[row, col].Text);
                            }
                            model.TableData.Add(currentRow);
                        }
                    }

                    // Получаем изображения
                    foreach (var drawing in worksheet.Drawings)
                    {
                        if (drawing is ExcelPicture picture)
                        {
                            using var ms = new MemoryStream();
                            picture.Image.Save(ms, picture.Image.RawFormat);
                            model.Images.Add(new ExcelImageModel
                            {
                                ImageData = ms.ToArray(),
                                Format = picture.Image.RawFormat.ToString(),
                                Name = picture.Name
                            });
                        }
                    }

                    Worksheets.Add(model);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Excel data: {ex.Message}");
            }
        }
    }

    public class ExcelWorksheetModel
    {
        public string Name { get; set; }
        public List<List<string>> TableData { get; set; }
        public List<ExcelImageModel> Images { get; set; }
    }

    public class ExcelImageModel
    {
        public byte[] ImageData { get; set; }
        public string Format { get; set; }
        public string Name { get; set; }
    }
}