using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using System.Net.Http;
using System.Net.Http.Json;
using System.Net.Http.Headers;
using Microsoft.Extensions.Configuration;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Http;

namespace ClientForm.Pages.Reports
{
    public class EditModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        [BindProperty(SupportsGet = true)]
        public int Id { get; set; }

        [BindProperty]
        public RIModel In { get; set; } = new();

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

                In.Name = report.Name;
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
            try
            {
                using var content = new MultipartFormDataContent();

                // Добавляем текстовые поля
                content.Add(new StringContent(In.Name), "name");

                // Добавляем файл только если он был предоставлен
                if (In.NewFile != null && In.NewFile.Length > 0)
                {
                    var fileContent = new StreamContent(In.NewFile.OpenReadStream());
                    fileContent.Headers.ContentType = new MediaTypeHeaderValue(In.NewFile.ContentType);
                    content.Add(fileContent, "file", In.NewFile.FileName);
                }

                var response = await _httpClient.PutAsync(
                    $"{_apiBaseUrl}/api/reports/{Id}",
                    content);

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
                ModelState.AddModelError(string.Empty, $"Ошибка при обновлении: {ex.Message}");
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
                // Игнорируем ошибки
            }
        }
    }

    public class RIModel
    {
        [Required(ErrorMessage = "Название отчета обязательно")]
        [StringLength(100, ErrorMessage = "Название слишком длинное")]
        public string Name { get; set; }

        [Display(Name = "Новый файл")]
        public IFormFile NewFile { get; set; }
    }

    public class ReportData
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
    }
}