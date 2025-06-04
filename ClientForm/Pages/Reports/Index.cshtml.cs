using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ClientForm.Services;
using ServerForm.Models;
using ClientForm.Pages.Reports; // Добавьте это

namespace ClientForm.Pages.Reports
{
    public class IndexModel : PageModel
    {
        private readonly IReportApiService _reportService;
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(
            IReportApiService reportService,
            ILogger<IndexModel> logger)
        {
            _reportService = reportService;
            _logger = logger;
        }

        public IEnumerable<ReportData> Reports { get; set; }

        public async Task OnGetAsync()
        {
            try
            {
                var serverReports = await _reportService.GetAllReportsAsync();
                Reports = serverReports.Select(r => new ReportData
                {
                    Id = r.Id,
                    Name = r.Name,
                    FileName = r.FileName,
                    FilePath = r.FilePath
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading reports");
                Reports = Enumerable.Empty<ReportData>();
            }
        }

    }
}