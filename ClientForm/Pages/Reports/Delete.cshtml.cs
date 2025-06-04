// D:\prog\Form\ClientForm\Pages\Reports\Delete.cshtml.cs
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ClientForm.Services;
using Microsoft.AspNetCore.Authorization;
using ClientForm.Models;

namespace ClientForm.Pages.Reports
{
    [Authorize]
    public class DeleteModel : PageModel
    {
        private readonly IReportApiService _reportService;
        private readonly string _apiBaseUrl;
        private readonly ILogger<DeleteModel> _logger;

        public DeleteModel(
            IReportApiService reportService,
            IConfiguration configuration,
            ILogger<DeleteModel> logger)
        {
            _apiBaseUrl = configuration["ApiBaseUrl"] ?? throw new ArgumentNullException("ApiBaseUrl");
            _reportService = reportService;
            _logger = logger;
        }

        [BindProperty]
        public ReportData Report { get; set; }

        public async Task<IActionResult> OnGetAsync(int id)
        {
            try
            {
                var serverReport = await _reportService.GetReportAsync(id);
                if (serverReport == null)
                {
                    return NotFound();
                }

                // �������������� ServerForm.Models.ReportData � ClientForm.Models.ReportData
                Report = new ClientForm.Models.ReportData
                {
                    Id = serverReport.Id,
                    Name = serverReport.Name,
                    FileName = serverReport.FileName,
                    // �������� ��� ����������� ��������
                };

                return Page();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading report for deletion");
                return RedirectToPage("/Error");
            }
        }

        public async Task<IActionResult> OnPostAsync(int id)
        {
            try
            {
                await _reportService.DeleteReportAsync(id);
                return RedirectToPage("./Index");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting report");
                return RedirectToPage("/Error");
            }
        }
    }
}