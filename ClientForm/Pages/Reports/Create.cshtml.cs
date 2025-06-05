using ClientForm.Models; 
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Net.Http.Headers;

namespace ClientForm.Pages.Reports
{
    public class CreateModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        [BindProperty]
        public ReportInputModel Input { get; set; } = new ReportInputModel();

        public CreateModel(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
        }

        public IActionResult OnGet()
        {
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            try
            {
                using var content = new MultipartFormDataContent();
                content.Add(new StringContent(Input.Name), "Name");

                // Обработка файла
                if (Input.File != null)
                {
                    var fileContent = new StreamContent(Input.File.OpenReadStream());
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(Input.File.ContentType);
                    content.Add(fileContent, "File", Input.File.FileName);
                }

                var response = await _httpClient.PostAsync($"{_apiBaseUrl}/api/reports", content);

                if (response.IsSuccessStatusCode)
                {
                    var createdReport = await response.Content.ReadFromJsonAsync<ReportData>();
                    return RedirectToPage("./Details", new { id = createdReport?.Id });
                }

                var errorMessage = await response.Content.ReadAsStringAsync();
                ModelState.AddModelError(string.Empty, $"Ошибка при создании отчёта: {errorMessage}");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, $"Ошибка: {ex.Message}");
            }

            return Page();
        }
    }
}