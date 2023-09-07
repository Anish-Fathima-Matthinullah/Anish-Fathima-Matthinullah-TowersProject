using API.Dtos;
using API.Models;
using OfficeOpenXml;

namespace API.Interfaces
{
    public interface IStyleService
    {
        void SetWorksheetStyle(ExcelWorksheet workSheet);
        void SetDataFormatStyle(ExcelWorksheet workSheet, ExcelListDetailsDto? excelListDetails);
    }
}