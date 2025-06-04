using ServerForm.Models;
using ServerForm.Services;

namespace ServerForm.Interfaces
{
    public interface IReportService
    {
        Task<ReportData> CreateReportAsync(string name, IFormFile file);
        Task DeleteReportAsync(int id);
        Task<ReportData> GetReportAsync(int id);
        Task<IEnumerable<ReportData>> GetAllReportsAsync();
        Task<ReportData> UpdateReportAsync(ReportData reportData);
        string GetExcelContent(ReportData reportData);
        FileStream GetExcelFileStream(ReportData reportData);
    }
}