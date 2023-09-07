using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using API.Dtos;
using API.Interfaces;

namespace API.Controllers
{
    public class ExcelController : BaseApiController
    {
        private readonly IExportService _exportService;
        private readonly IImportService _importService;
        private readonly IStyleService _styleService;
        private readonly IUpdateService _updateService;
        public ExcelController(IExportService exportService, IImportService importService, IStyleService styleService, IUpdateService updateService)
        {
            _exportService = exportService;
            _importService = importService;
            _styleService = styleService;
            _updateService = updateService;
        }

        [HttpGet]
        public List<ClientDataDto> GetData()
        {
            return _exportService.GetData();
        }

        [HttpGet("export")]
        public Task<IActionResult> Export()
        {
            const string fileFormat = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            string excelName = $"Sample-{DateTime.Now:dd-MM-yyyy}.xlsx";
            var stream = new MemoryStream();

            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sample");

                _styleService.SetWorksheetStyle(workSheet);

                ExcelListDetailsDto excelListDetails = _exportService.ExportData();

                workSheet.Cells[3, 1].LoadFromCollection(excelListDetails.ExcelList, false);

                _styleService.SetDataFormatStyle(workSheet, excelListDetails);

                package.Save();
            }

            stream.Position = 0;
            return Task.FromResult<IActionResult>(File(stream, fileFormat, excelName));
        }

        [HttpPost("import")]
        public async Task<IActionResult> ExcelUpload(IFormFile formFile)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return BadRequest("File is empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("File not supported");
            }

            var importId = await _importService.AddToImportTable(formFile);
            if (!(importId > 0))
            {
                return BadRequest("Import Table failed to Update");
            }

            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    await _importService.ImportToDatabase(worksheet, importId);
                }
            }
            return Ok();
        }

        [HttpPut("update")]
        public async Task<IActionResult> UpdateTable(ClientDataDto dataToUpdate)
        {
            if(dataToUpdate is null) return BadRequest("Invalid data update request");
            
            await _updateService.updateTable(dataToUpdate);
            return Ok();
        }
    }
}
