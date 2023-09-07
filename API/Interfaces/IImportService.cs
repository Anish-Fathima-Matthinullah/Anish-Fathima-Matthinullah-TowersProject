using OfficeOpenXml;

namespace API.Interfaces
{
    public interface IImportService
    {
        Task<int> AddToImportTable(IFormFile formFile);
        Task<bool> ImportToDatabase(ExcelWorksheet worksheet, int importId);
    }
}