using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ServerForm.Models;

namespace ClientForm.Pages.Reports
{
    public class DeleteModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        public ReportData Report { get; set; }

        public DeleteModel(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"];
        }

        public async Task<IActionResult> OnGetAsync(int id)
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports/{id}");
            if (!response.IsSuccessStatusCode) return NotFound();

            Report = await response.Content.ReadFromJsonAsync<ReportData>();
            return Page();
        }

        public async Task<IActionResult> OnPostAsync(int id)
        {
            await _httpClient.DeleteAsync($"{_apiBaseUrl}/api/reports/{id}");
            return RedirectToPage("Index");
        }
    }
}