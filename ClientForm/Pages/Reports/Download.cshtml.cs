using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace ClientForm.Pages.Reports
{
    public class DownloadModel : PageModel
    {
        private readonly HttpClient _httpClient;

        public DownloadModel(IHttpClientFactory httpClientFactory)
        {
            _httpClient = httpClientFactory.CreateClient("ServerAPI");
        }

        public async Task<IActionResult> OnGetAsync(int id)
        {
            var response = await _httpClient.GetAsync($"api/reports/download/{id}");
            if (!response.IsSuccessStatusCode)
                return NotFound();

            var fileBytes = await response.Content.ReadAsByteArrayAsync();
            return File(fileBytes, "application/octet-stream", "report.xlsx");
        }
    }
}