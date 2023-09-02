using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Interfaces
{
    public interface IExcelService
    {
       Task<int> AddToImportTable(IFormFile formFile, string? fileName, CancellationToken cancellationToken);
    //    Task<bool> SelectTablebasedOnColor2(ExcelStyle rng, ExcelWorksheet worksheet, int row, int importId);
       Task<bool> SelectTablebasedOnColor(ExcelWorksheet worksheet, int importId);
    }
}