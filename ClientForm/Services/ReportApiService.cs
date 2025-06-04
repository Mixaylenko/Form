// D:\prog\Form\ClientForm\Services\ReportApiService.cs
using ServerForm.Models;
using System.Net.Http.Headers;

namespace ClientForm.Services
{
    public interface IReportApiService
    {
        Task<IEnumerable<ReportData>> GetAllReportsAsync();
        Task<ReportData> GetReportAsync(int id);
        Task<ReportData> CreateReportAsync(string name, IFormFile file);
        Task UpdateReportAsync(ReportData reportData);
        Task DeleteReportAsync(int id);
        Task<string> GetExcelContentAsync(int id);
        Task<Stream> GetExcelFileStreamAsync(int id);
    }

    public class ReportApiService : ApiService, IReportApiService
    {
        public ReportApiService(HttpClient httpClient, IConfiguration configuration)
            : base(httpClient, configuration)
        {
        }

        public async Task<IEnumerable<ReportData>> GetAllReportsAsync()
        {
            return await GetAsync<IEnumerable<ReportData>>("/api/Reports");
        }

        public async Task<ReportData> GetReportAsync(int id)
        {
            return await GetAsync<ReportData>($"/api/Reports/{id}");
        }

        public async Task<ReportData> CreateReportAsync(string name, IFormFile file)
        {
            using var content = new MultipartFormDataContent();

            content.Add(new StringContent(name), "Name");

            var fileContent = new StreamContent(file.OpenReadStream())
            {
                Headers = { ContentType = new MediaTypeHeaderValue(file.ContentType) }
            };
            content.Add(fileContent, "File", file.FileName);

            var response = await PostFormDataAsync("/api/Reports", content);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadFromJsonAsync<ReportData>();
        }

        public async Task UpdateReportAsync(ReportData reportData)
        {
            await PutAsync($"/api/Reports/{reportData.Id}", reportData);
        }

        public async Task DeleteReportAsync(int id)
        {
            await DeleteAsync($"/api/Reports/{id}");
        }

        public async Task<string> GetExcelContentAsync(int id)
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/Reports/view/{id}");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public async Task<Stream> GetExcelFileStreamAsync(int id)
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}/api/Reports/Download/{id}");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStreamAsync();
        }
    }
}