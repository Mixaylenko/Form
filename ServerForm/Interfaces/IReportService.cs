using ServerForm.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ServerForm.Interfaces
{
    public interface IReportService
    {
        Task<ReportData> CreateReportAsync(ReportData report, Stream fileStream);
        Task<IEnumerable<ReportData>> GetAllReportsAsync();
        Task<ReportData> GetReportAsync(int id);
        Task<ReportData> UpdateReportAsync(int id, ReportData report, Stream fileStream = null, string originalFileName = null);
        Task DeleteReportAsync(int id);
        Task<byte[]> ConvertToWordAsync(int id);
    }
}