using ExcelTool.Models;

namespace ExcelTool.Services
{
    public interface IExcelUploadService
    {
        Task<ExcelUploadResult> UploadAsync(string tableName, IFormFile excelFile);
    }
}

