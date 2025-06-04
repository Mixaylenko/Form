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

        // Настройка HttpClient
        _httpClient.Timeout = TimeSpan.FromSeconds(30);
    }

    [BindProperty]
    public InputModel Input { get; set; }

    [TempData]
    public string Message { get; set; }

    public class InputModel
    {
        [Required(ErrorMessage = "Имя обязательно")]
        [StringLength(50, ErrorMessage = "Имя не должно превышать 50 символов")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Email обязателен")]
        [EmailAddress(ErrorMessage = "Некорректный формат email")]
        public string Email { get; set; }

        [Required(ErrorMessage = "Пароль обязателен")]
        [MinLength(6, ErrorMessage = "Пароль должен содержать минимум 6 символов")]
        [DataType(DataType.Password)]
        public string Password { get; set; }
    }

    public void OnGet()
    {
        // Инициализация модели при GET-запросе
        Input ??= new InputModel();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid)
        {
            _logger.LogWarning("Невалидные данные при регистрации");
            return Page();
        }

        try
        {
            _logger.LogInformation($"Попытка регистрации пользователя: {Input.Email}");

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
                _logger.LogInformation($"Успешная регистрация: {Input.Email}");
                Message = "Регистрация прошла успешно! Авторизуйтесь, используя свои данные.";
                return RedirectToPage("/Auth/Login");
            }
            else
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                _logger.LogError($"Ошибка регистрации: {errorContent}");

                // Обработка разных статус кодов
                if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                {
                    ModelState.AddModelError(string.Empty, "Пользователь с таким email уже существует");
                }
                else
                {
                    ModelState.AddModelError(string.Empty, $"Ошибка сервера: {errorContent}");
                }

                return Page();
            }
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Ошибка подключения к API");
            ModelState.AddModelError(string.Empty, "Ошибка подключения к серверу. Попробуйте позже.");
            return Page();
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "Таймаут при регистрации");
            ModelState.AddModelError(string.Empty, "Время ожидания ответа истекло. Попробуйте позже.");
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Неожиданная ошибка при регистрации");
            ModelState.AddModelError(string.Empty, "Произошла непредвиденная ошибка. Попробуйте позже.");
            return Page();
        }
    }
}