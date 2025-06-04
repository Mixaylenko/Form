using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using ClientForm.Models;
using ServerForm.Models; 

namespace ClientForm.Pages.Reports
{
    public class EditModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        [BindProperty(SupportsGet = true)]
        public int Id { get; set; }

        [BindProperty]
        public ReportInputModel Input { get; set; } = new();

        public string CurrentFileName { get; set; }

        public EditModel(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
        }

        public async Task<IActionResult> OnGetAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/{Id}");
                if (!response.IsSuccessStatusCode)
                {
                    return NotFound();
                }

                var report = await response.Content.ReadFromJsonAsync<ReportData>();
                if (report == null)
                {
                    return NotFound();
                }

                Input.Name = report.Name;
                CurrentFileName = report.FileName;

                return Page();
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, $"Ошибка при загрузке отчета: {ex.Message}");
                return Page();
            }
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                await LoadCurrentFileName();
                return Page();
            }

            try
            {
                // Создаем объект для обновления
                var updateData = new ReportData
                {
                    Id = Id,
                    Name = Input.Name,
                    FileName = CurrentFileName
                };

                var response = await _httpClient.PutAsJsonAsync(
                    $"{_apiBaseUrl}/api/reports/{Id}",
                    updateData);

                if (response.IsSuccessStatusCode)
                {
                    return RedirectToPage("Details", new { id = Id });
                }

                var errorMessage = await response.Content.ReadAsStringAsync();
                ModelState.AddModelError(string.Empty,
                    $"Ошибка при обновлении: {response.StatusCode} - {errorMessage}");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, $"Ошибка при сохранении: {ex.Message}");
            }

            await LoadCurrentFileName();
            return Page();
        }

        private async Task LoadCurrentFileName()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/{Id}");
                if (response.IsSuccessStatusCode)
                {
                    var report = await response.Content.ReadFromJsonAsync<ReportData>();
                    CurrentFileName = report?.FileName;
                }
            }
            catch
            {
                // Логируем ошибку, если необходимо
            }
        }
    }

}