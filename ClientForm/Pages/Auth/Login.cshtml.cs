using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authentication.Cookies;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;

[AllowAnonymous]
public class LoginModel : PageModel
{
    private readonly HttpClient _httpClient;
    private readonly string _apiBaseUrl;
    private readonly ILogger<LoginModel> _logger;

    public LoginModel(
        IHttpClientFactory httpClientFactory,
        IConfiguration configuration,
        ILogger<LoginModel> logger)
    {
        _httpClient = httpClientFactory.CreateClient();
        _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
        _logger = logger;
        _httpClient.Timeout = TimeSpan.FromSeconds(30);
    }

    [BindProperty]
    public InputModel Input { get; set; }

    [TempData]
    public string ErrorMessage { get; set; }

    public string ReturnUrl { get; set; }

    public class InputModel
    {
        [Required(ErrorMessage = "Email обязателен")]
        [EmailAddress(ErrorMessage = "Некорректный формат email")]
        public string Email { get; set; }

        [Required(ErrorMessage = "Пароль обязателен")]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        [Display(Name = "Запомнить меня")]
        public bool RememberMe { get; set; }
    }

    public void OnGet(string returnUrl = null)
    {
        if (!string.IsNullOrEmpty(ErrorMessage))
        {
            ModelState.AddModelError(string.Empty, ErrorMessage);
        }

        // Очищаем существующий внешний cookie для обеспечения чистого процесса входа
        ReturnUrl = returnUrl;
    }

    public async Task<IActionResult> OnPostAsync(string returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/Report/Index");

        if (!ModelState.IsValid)
        {
            return Page();
        }

        try
        {
            var response = await _httpClient.PostAsJsonAsync(
                $"{_apiBaseUrl}/api/user/login",
                new
                {
                    Email = Input.Email,
                    Password = Input.Password
                });

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadFromJsonAsync<LoginResponse>();

                var claims = new List<Claim>
                {
                    new Claim(ClaimTypes.NameIdentifier, result.Id.ToString()),
                    new Claim(ClaimTypes.Name, result.Name),
                    new Claim(ClaimTypes.Email, Input.Email),
                    new Claim(ClaimTypes.Role, result.Role)
                };

                var claimsIdentity = new ClaimsIdentity(
                    claims, CookieAuthenticationDefaults.AuthenticationScheme);

                var authProperties = new AuthenticationProperties
                {
                    IsPersistent = Input.RememberMe,
                    ExpiresUtc = Input.RememberMe ? DateTimeOffset.UtcNow.AddDays(30) : null,
                    AllowRefresh = true
                };

                await HttpContext.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    new ClaimsPrincipal(claimsIdentity),
                    authProperties);

                _logger.LogInformation("User {Email} logged in.", Input.Email);
                return LocalRedirect(ReturnUrl);
            }

            ErrorMessage = await response.Content.ReadAsStringAsync();
            ModelState.AddModelError(string.Empty, ErrorMessage);
            return Page();
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "API request failed");
            ErrorMessage = "Ошибка соединения с сервером";
            ModelState.AddModelError(string.Empty, ErrorMessage);
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Login error");
            ErrorMessage = "Ошибка при входе в систему";
            ModelState.AddModelError(string.Empty, ErrorMessage);
            return Page();
        }
    }

    private class LoginResponse
    {
        public string Message { get; set; }
        public int Id { get; set; }
        public string Name { get; set; }
        public string Role { get; set; }
    }
}