// D:\prog\Form\ClientForm\Pages\Report\Create.cshtml.cs
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ClientForm.Services;
using Microsoft.AspNetCore.Http;

namespace ClientForm.Pages.Report
{
    public class CreateModel : PageModel
    {
        private readonly IReportApiService _reportService;
        private readonly ILogger<CreateModel> _logger;

        public CreateModel(
            IReportApiService reportService,
            ILogger<CreateModel> logger)
        {
            _reportService = reportService;
            _logger = logger;
        }

        [BindProperty]
        public string Name { get; set; }

        [BindProperty]
        public IFormFile File { get; set; }

        public void OnGet()
        {
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid || File == null)
            {
                return Page();
            }

            try
            {
                await _reportService.CreateReportAsync(Name, File);
                return RedirectToPage("Index");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating report");
                ModelState.AddModelError(string.Empty, "Error creating report");
                return Page();
            }
        }
    }
}