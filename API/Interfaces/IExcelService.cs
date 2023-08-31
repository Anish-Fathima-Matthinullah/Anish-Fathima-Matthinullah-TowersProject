using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Interfaces
{
    public interface IExcelService
    {
       Task<bool> SaveAllAsync();
       Task<int> AddToImportTable(IFormFile formFile, string? fileName, CancellationToken cancellationToken);
    //    Task<string> SelcetTablebasedOnColor(ExcelStyle rng, ExcelWorksheet worksheet, int row);
       Task<bool> SelectTablebasedOnColor(ExcelStyle rng, ExcelWorksheet worksheet, int row, int importId);
    }
}