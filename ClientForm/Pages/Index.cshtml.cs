using ClientForm.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using System.Net;

namespace ClientForm.Pages
{
    [Authorize]
    public class IndexModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(
            IHttpClientFactory httpClientFactory,
            IConfiguration configuration,
            ILogger<IndexModel> logger)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
            _logger = logger;
        }

        public List<ReportData> Reports { get; set; } = new List<ReportData>();

        [TempData]
        public string? ErrorMessage { get; set; }

        [TempData]
        public string? InfoMessage { get; set; } // Новое свойство для информационных сообщений

        [BindProperty]
        public UploadModel Upload { get; set; }

        public class UploadModel
        {
            [Required]
            public string FileName { get; set; }

            [Required]
            public IFormFile File { get; set; }
        }

        public async Task OnGetAsync()
        {
            await LoadReports();
        }

        private async Task LoadReports()
        {
            try
            {
                var isAdmin = User.IsInRole("Admin");
                var currentUsername = User.Identity?.Name;

                if (string.IsNullOrEmpty(currentUsername))
                {
                    ErrorMessage = "User not authenticated";
                    return;
                }

                var endpoint = isAdmin ? "all" : $"user/{WebUtility.UrlEncode(currentUsername)}";

                var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/{endpoint}");

                if (response.IsSuccessStatusCode)
                {
                    Reports = await response.Content.ReadFromJsonAsync<List<ReportData>>() ?? new List<ReportData>();
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    ErrorMessage = $"Error loading reports: {response.StatusCode}";
                    _logger.LogError("Failed to get reports. Status: {StatusCode}, Error: {Error}",
                        response.StatusCode, errorContent);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = "Error loading reports";
                _logger.LogError(ex, "Error getting reports");
            }
        }

        // Остальные методы остаются без изменений
        public async Task<IActionResult> OnPostUploadAsync()
        {
            if (!ModelState.IsValid)
            {
                await LoadReports();
                return Page();
            }

            try
            {
                using var content = new MultipartFormDataContent();
                var fileContent = new StreamContent(Upload.File.OpenReadStream());
                fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(Upload.File.ContentType);
                content.Add(fileContent, "file", Upload.File.FileName);
                content.Add(new StringContent(Upload.FileName), "fileName");

                var response = await _httpClient.PostAsync($"{_apiBaseUrl}/api/reports", content);

                if (response.IsSuccessStatusCode)
                {
                    InfoMessage = "Отчет успешно загружен";
                    return RedirectToPage();
                }

                ErrorMessage = await response.Content.ReadAsStringAsync();
                _logger.LogError("Failed to upload report: {ErrorMessage}", ErrorMessage);
            }
            catch (Exception ex)
            {
                ErrorMessage = "Ошибка при загрузке отчета";
                _logger.LogError(ex, "Error uploading report");
            }

            await LoadReports();
            return Page();
        }

        public async Task<IActionResult> OnPostDeleteAsync(int id)
        {
            try
            {
                var response = await _httpClient.DeleteAsync($"{_apiBaseUrl}/api/reports/{id}");

                if (response.IsSuccessStatusCode)
                {
                    InfoMessage = "Отчет успешно удален";
                }
                else
                {
                    ErrorMessage = await response.Content.ReadAsStringAsync();
                    _logger.LogError("Failed to delete report: {ErrorMessage}", ErrorMessage);
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = "Ошибка при удалении отчета";
                _logger.LogError(ex, "Error deleting report with id {Id}", id);
            }

            return RedirectToPage();
        }
    }
}