using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;

[Authorize]
public class LogoutModel : PageModel
{
    private readonly ILogger<LogoutModel> _logger;

    public LogoutModel(ILogger<LogoutModel> logger)
    {
        _logger = logger;
    }

    public async Task<IActionResult> OnPostAsync()
    {
        await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        _logger.LogInformation("User logged out.");
        return RedirectToPage("/Auth/Login");
    }

    public async Task<IActionResult> OnGetAsync()
    {
        await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
        _logger.LogInformation("User logged out.");
        return RedirectToPage("/Auth/Login");
    }
}