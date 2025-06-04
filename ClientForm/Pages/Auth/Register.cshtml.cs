using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using Microsoft.AspNetCore.Authorization;

[AllowAnonymous]
public class RegisterModel : PageModel
{
    private readonly HttpClient _httpClient;
    private readonly string _apiBaseUrl;
    private readonly ILogger<RegisterModel> _logger;

    public RegisterModel(
        IHttpClientFactory httpClientFactory,
        IConfiguration configuration,
        ILogger<RegisterModel> logger)
    {
        _httpClient = httpClientFactory.CreateClient();
        _apiBaseUrl = configuration["ApiBaseUrl"];
        _logger = logger;

        // ��������� HttpClient
        _httpClient.Timeout = TimeSpan.FromSeconds(30);
    }

    [BindProperty]
    public InputModel Input { get; set; }

    [TempData]
    public string Message { get; set; }

    public class InputModel
    {
        [Required(ErrorMessage = "��� �����������")]
        [StringLength(50, ErrorMessage = "��� �� ������ ��������� 50 ��������")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Email ����������")]
        [EmailAddress(ErrorMessage = "������������ ������ email")]
        public string Email { get; set; }

        [Required(ErrorMessage = "������ ����������")]
        [MinLength(6, ErrorMessage = "������ ������ ��������� ������� 6 ��������")]
        [DataType(DataType.Password)]
        public string Password { get; set; }
    }

    public void OnGet()
    {
        // ������������� ������ ��� GET-�������
        Input ??= new InputModel();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid)
        {
            _logger.LogWarning("���������� ������ ��� �����������");
            return Page();
        }

        try
        {
            _logger.LogInformation($"������� ����������� ������������: {Input.Email}");

            var response = await _httpClient.PostAsJsonAsync(
                $"{_apiBaseUrl}/api/user/register",
                new
                {
                    Input.Name,
                    Input.Email,
                    Input.Password
                });

            if (response.IsSuccessStatusCode)
            {
                _logger.LogInformation($"�������� �����������: {Input.Email}");
                Message = "����������� ������ �������! �������������, ��������� ���� ������.";
                return RedirectToPage("/Auth/Login");
            }
            else
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                _logger.LogError($"������ �����������: {errorContent}");

                // ��������� ������ ������ �����
                if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                {
                    ModelState.AddModelError(string.Empty, "������������ � ����� email ��� ����������");
                }
                else
                {
                    ModelState.AddModelError(string.Empty, $"������ �������: {errorContent}");
                }

                return Page();
            }
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "������ ����������� � API");
            ModelState.AddModelError(string.Empty, "������ ����������� � �������. ���������� �����.");
            return Page();
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "������� ��� �����������");
            ModelState.AddModelError(string.Empty, "����� �������� ������ �������. ���������� �����.");
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "����������� ������ ��� �����������");
            ModelState.AddModelError(string.Empty, "��������� �������������� ������. ���������� �����.");
            return Page();
        }
    }
}