
using API.Models;

namespace API.Dtos
{
    public class ExcelListDetailsDto
    {
        public List<ExcelDataDto>? ExcelList { get; set; }
        public List<ExcelData>? TableData { get; set; }
    }
}