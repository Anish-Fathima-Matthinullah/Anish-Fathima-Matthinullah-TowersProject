using API.Dtos;
using API.Models;

namespace API.Interfaces
{
    public interface IExportService
    {
        ExcelListDetailsDto ExportData();
        List<ClientDataDto> GetData();
    }
}