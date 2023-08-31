using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Interfaces
{
    public interface IExcelService
    {
       Task<bool> SaveAllAsync();
       Task<bool> AddToImportTable(IFormFile formFile, string? fileName, CancellationToken cancellationToken);
       string SelcetTablebasedOnColor(ExcelStyle rng, ExcelWorksheet worksheet, int row);
    }
}