// D:\prog\Form\ClientForm\Services\ApiService.cs
using System.Net.Http.Json;
using System.Text;

namespace ClientForm.Services
{
    public class ApiService
    {
        public readonly HttpClient _httpClient;
        public readonly string _apiBaseUrl;

        public ApiService(HttpClient httpClient, IConfiguration configuration)
        {
            _httpClient = httpClient;
            _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
        }

        protected async Task<T> GetAsync<T>(string endpoint)
        {
            var response = await _httpClient.GetAsync($"{_apiBaseUrl}{endpoint}");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<T>();
        }

        protected async Task<TResponse> PostAsync<TRequest, TResponse>(string endpoint, TRequest data)
        {
            var response = await _httpClient.PostAsJsonAsync($"{_apiBaseUrl}{endpoint}", data);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<TResponse>();
        }

        protected async Task PutAsync<T>(string endpoint, T data)
        {
            var response = await _httpClient.PutAsJsonAsync($"{_apiBaseUrl}{endpoint}", data);
            response.EnsureSuccessStatusCode();
        }

        protected async Task DeleteAsync(string endpoint)
        {
            var response = await _httpClient.DeleteAsync($"{_apiBaseUrl}{endpoint}");
            response.EnsureSuccessStatusCode();
        }

        protected async Task<HttpResponseMessage> PostFormDataAsync(string endpoint, MultipartFormDataContent content)
        {
            return await _httpClient.PostAsync($"{_apiBaseUrl}{endpoint}", content);
        }
    }
}