using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;

namespace ClientForm.Pages.Reports
{
    public class DownloadModel : PageModel
    {
        private readonly string _apiBaseUrl;

        public DownloadModel(IConfiguration configuration)
        {
            _apiBaseUrl = configuration["ApiBaseUrl"]
                ?? throw new ArgumentNullException("ApiBaseUrl");
        }

        public IActionResult OnGet(int id, [FromQuery] string format)
        {
            var url = format == "word"
                ? $"{_apiBaseUrl}/api/reports/convert-to-word/{id}"
                : $"{_apiBaseUrl}/api/reports/download/{id}";

            return Redirect(url);
        }
    }
}