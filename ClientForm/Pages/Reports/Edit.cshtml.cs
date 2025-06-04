using ClientForm.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;

namespace ClientForm.Pages.Reports
{
    public class EditModel : PageModel
    {
        private readonly HttpClient _httpClient;

        [BindProperty(SupportsGet = true)]
        public int Id { get; set; }

        [BindProperty]
        public ReportInputModel Input { get; set; } = new();

        public string CurrentFileName { get; set; } = string.Empty;

        public EditModel(IHttpClientFactory httpClientFactory)
        {
            _httpClient = httpClientFactory.CreateClient("ServerAPI");
        }

        public async Task<IActionResult> OnGetAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"api/reports/{Id}");
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
                return Page();
            }

            try
            {
                using var content = new MultipartFormDataContent();
                content.Add(new StringContent(Id.ToString()), "id");
                content.Add(new StringContent(Input.Name), "name");

                if (Input.File != null)
                {
                    await using var fileStream = Input.File.OpenReadStream();
                    content.Add(new StreamContent(fileStream), "file", Input.File.FileName);
                }

                var response = await _httpClient.PutAsync($"api/reports/{Id}", content);

                if (response.IsSuccessStatusCode)
                {
                    return RedirectToPage("Details", new { id = Id });
                }

                var errorMessage = await response.Content.ReadAsStringAsync();
                ModelState.AddModelError(string.Empty,
                    $"Ошибка при обновлении отчета: {response.StatusCode} - {errorMessage}");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, $"Ошибка при сохранении: {ex.Message}");
            }

            // При ошибке нужно заново загрузить CurrentFileName
            var reportResponse = await _httpClient.GetAsync($"api/reports/{Id}");
            if (reportResponse.IsSuccessStatusCode)
            {
                var report = await reportResponse.Content.ReadFromJsonAsync<ReportData>();
                CurrentFileName = report?.FileName ?? string.Empty;
            }

            return Page();
        }
    }
}