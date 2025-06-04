using Microsoft.AspNetCore.Mvc.RazorPages;
using ServerForm.Models;

namespace ClientForm.Pages.Reports
{
    public class IndexModel : PageModel
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiBaseUrl;

        public List<ReportData> Reports { get; set; } = new();

        public IndexModel(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClient = httpClientFactory.CreateClient();
            _apiBaseUrl = configuration["ApiBaseUrl"];
        }

        public async Task OnGetAsync()
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/reports");
            if (response.IsSuccessStatusCode)
            {
                Reports = await response.Content.ReadFromJsonAsync<List<ReportData>>();
            }
        }
    }
}