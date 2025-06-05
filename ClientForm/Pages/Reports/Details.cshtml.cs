using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using System.Net.Http;
using System.Net.Http.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;

namespace ClientForm.Pages.Reports
{
    public class DetailsModel : PageModel
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IConfiguration _configuration;
        private readonly ILogger<DetailsModel> _logger;
        private HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        public DetailsModel(
            IHttpClientFactory httpClientFactory,
            IConfiguration configuration,
            ILogger<DetailsModel> logger)
        {
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;
            _logger = logger;
            _apiBaseUrl = _configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
        }

        public ReportData Report { get; set; }
        public bool IsExcel { get; set; }
        public List<WorksheetPreview> Worksheets { get; set; } = new();

        public async Task<IActionResult> OnGetAsync(int id)
        {
            try
            {
                _httpClient = _httpClientFactory.CreateClient();
                _httpClient.BaseAddress = new Uri(_apiBaseUrl);
                _httpClient.Timeout = TimeSpan.FromSeconds(30);

                // Загрузка основных данных отчета
                var reportResponse = await _httpClient.GetAsync($"/api/Reports/{id}");
                if (!reportResponse.IsSuccessStatusCode)
                {
                    _logger.LogError("Failed to get report. Status: {StatusCode}", reportResponse.StatusCode);
                    return NotFound();
                }

                Report = await reportResponse.Content.ReadFromJsonAsync<ReportData>();

                // Проверка типа файла
                var ext = Path.GetExtension(Report.FileName).ToLower();
                IsExcel = ext == ".xlsx" || ext == ".xls";

                if (IsExcel)
                {
                    // Загрузка данных для предпросмотра
                    var previewResponse = await _httpClient.GetAsync($"/api/Reports/{id}/preview");
                    if (previewResponse.IsSuccessStatusCode)
                    {
                        var preview = await previewResponse.Content.ReadFromJsonAsync<ReportPreview>();
                        Worksheets = preview.Worksheets;
                    }
                    else
                    {
                        _logger.LogWarning("Failed to get preview. Status: {StatusCode}", previewResponse.StatusCode);
                    }
                }

                return Page();
            }
            catch (HttpRequestException ex)
            {
                _logger.LogError(ex, "API request failed");
                return StatusCode(500, "Ошибка соединения с сервером");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error");
                return StatusCode(500, "Внутренняя ошибка сервера");
            }
        }

        public class ReportData
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string FileName { get; set; }
            public string FilePath { get; set; }
        }

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

        public class ImagePreview
        {
            public string Name { get; set; }
            public string Format { get; set; }
            public byte[] ImageData { get; set; }
        }
    }
}