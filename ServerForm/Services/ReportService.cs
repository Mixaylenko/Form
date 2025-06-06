using ServerForm.Interfaces;
using ServerForm.Models;
using OfficeOpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml.Drawing.Chart;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OfficeOpenXml.Drawing;
using System.IO;
using Microsoft.Office.Interop.Excel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Processing;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Styles = DocumentFormat.OpenXml.Wordprocessing.Styles;
using Color = System.Drawing.Color;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ServerForm.Services
{
    public class ReportService : IReportService
    {
        private readonly DatabaseContext _context;
        private readonly IWebHostEnvironment _env;
        private readonly string _uploadsPath;

        public ReportService(DatabaseContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
            _uploadsPath = Path.Combine(_env.WebRootPath, "uploads");
            Directory.CreateDirectory(_uploadsPath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public async Task<ReportData> CreateReportAsync(ReportData report, Stream fileStream)
        {
            var extension = !string.IsNullOrEmpty(report.FileName) 
                ? Path.GetExtension(report.FileName) 
                : "";
            var fileName = $"{Guid.NewGuid()}{extension}";
            var filePath = Path.Combine(_uploadsPath, fileName);

            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                await fileStream.CopyToAsync(fs);
            }

            report.FilePath = filePath;
            report.FileName = fileName;

            _context.ReportDatas.Add(report);
            await _context.SaveChangesAsync();

            return report;
        }

        public async Task<IEnumerable<ReportData>> GetAllReportsAsync()
        {
            return await _context.ReportDatas.ToListAsync();
        }

        public async Task<ReportData> GetReportAsync(int id)
        {
            return await _context.ReportDatas.FindAsync(id);
        }

        public async Task<ReportData> UpdateReportAsync(int id, ReportData report, Stream fileStream = null, string originalFileName = null)
        {
            var existing = await _context.ReportDatas.FindAsync(id);
            if (existing == null) return null;

            // Всегда обновляем название отчета
            existing.Name = report.Name;

            if (fileStream != null)
            {
                // Удаляем старый файл
                if (!string.IsNullOrEmpty(existing.FilePath) && System.IO.File.Exists(existing.FilePath))
                {
                    System.IO.File.Delete(existing.FilePath);
                }

                // Сохраняем оригинальное расширение из нового файла
                var extension = !string.IsNullOrEmpty(originalFileName)
                    ? Path.GetExtension(originalFileName)
                    : "";

                var fileName = $"{Guid.NewGuid()}{extension}";
                var filePath = Path.Combine(_uploadsPath, fileName);

                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    await fileStream.CopyToAsync(fs);
                }

                // Обновляем путь и имя файла
                existing.FilePath = filePath;
                existing.FileName = fileName;
            }
            // Если файл не меняется, оставляем текущее имя и путь

            await _context.SaveChangesAsync();
            return existing;
        }

        public async Task DeleteReportAsync(int id)
        {
            var report = await _context.ReportDatas.FindAsync(id);
            if (report == null) return;

            if (!string.IsNullOrEmpty(report.FilePath) && System.IO.File.Exists(report.FilePath))
            {
                System.IO.File.Delete(report.FilePath);
            }

            _context.ReportDatas.Remove(report);
            await _context.SaveChangesAsync();
        }

        public async Task<byte[]> ConvertToWordAsync(int id)
        {
            var report = await GetReportAsync(id);
            if (report == null) return null;

            // Создаем временную копию файла для работы
            var tempExcelFile = Path.GetTempFileName() + Path.GetExtension(report.FileName);
            System.IO.File.Copy(report.FilePath, tempExcelFile, true);

            var tempWordFile = Path.GetTempFileName() + ".docx";
            var pythonScript = GeneratePythonScript(tempExcelFile, tempWordFile);

            var scriptPath = Path.GetTempFileName() + ".py";
            await System.IO.File.WriteAllTextAsync(scriptPath, pythonScript);

            try
            {
                var process = new Process();
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = $"\"{scriptPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                process.Start();
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    var error = await process.StandardError.ReadToEndAsync();
                    throw new Exception($"Python error: {error}");
                }

                return await System.IO.File.ReadAllBytesAsync(tempWordFile);
            }
            finally
            {
                // Гарантированное освобождение ресурсов
                await Task.Delay(1000); // Даем время на завершение процессов
                SafeDelete(tempExcelFile);
                SafeDelete(tempWordFile);
                SafeDelete(scriptPath);

                // Удаление временных файлов Excel (типа ~$...)
                var tempDir = Path.GetDirectoryName(tempExcelFile);
                var lockFiles = Directory.GetFiles(tempDir, "~$" + Path.GetFileName(tempExcelFile));
                foreach (var lockFile in lockFiles) SafeDelete(lockFile);
            }
        }

        // Добавьте эту вспомогательную функцию в класс
        private void SafeDelete(string path)
        {
            try
            {

                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }
            }
            catch
            {
                // Логирование ошибки при необходимости
            }
        }

        private string GeneratePythonScript(string excelPath, string wordPath)
        {
            return $@"
import win32com.client as win32
import os
import sys
import pythoncom

def convert_excel_to_word(excel_path, word_path):
    pythoncom.CoInitialize()  # Инициализация COM для потока
    excel = None
    word = None
    
    try:
        excel = win32.Dispatch('Excel.Application', pythoncom.CoInitialize())
        excel.Visible = False
        excel.DisplayAlerts = False  # Отключаем предупреждения
        
        # Открываем книгу в режиме read-only
        wb = excel.Workbooks.Open(
            os.path.abspath(excel_path),
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True
        )
        
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Add()
        
        # Обработка всех листов
        for sheet in wb.Sheets:
            # Добавляем название листа как заголовок
            header = doc.Content
            header.InsertAfter(f'Лист: {{sheet.Name}}\\n\\n')
            
            # Копируем данные
            sheet.UsedRange.Copy()
            doc.Content.PasteExcelTable(False, False, False)
            doc.Content.InsertAfter('\\n\\n')
        
        doc.SaveAs(os.path.abspath(word_path))
        
    except Exception as e:
        print(f""Ошибка конвертации: {{str(e)}}"", file=sys.stderr)
        raise
    finally:
        # Гарантированное освобождение ресурсов в правильном порядке
        try:
            if 'doc' in locals() and doc:
                doc.Close(SaveChanges=False)
        except: pass
        
        try:
            if 'wb' in locals() and wb:
                wb.Close(SaveChanges=False)
        except: pass
        
        try:
            if excel:
                excel.DisplayAlerts = True
                excel.Quit()
        except: pass
        
        try:
            if word:
                word.Quit()
        except: pass
        
        # Принудительное освобождение COM-объектов
        for obj in [excel, word, wb, doc, sheet]:
            try:
                while obj and win32._ole32_.CoReleaseServerProcess() == 0:
                    pass
            except: pass
        
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    excel_path = r'{excelPath.Replace("\\", "\\\\")}'
    word_path = r'{wordPath.Replace("\\", "\\\\")}'
    convert_excel_to_word(excel_path, word_path)
";
        }

    }
}