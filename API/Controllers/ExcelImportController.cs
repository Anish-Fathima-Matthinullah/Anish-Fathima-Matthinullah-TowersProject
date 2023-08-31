using API.Interfaces;
using API.Modals;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Controllers
{
    public class ExcelImportController : BaseApiController
    {
        private readonly IExcelService _excelService;

        public ExcelImportController(IExcelService excelService)
        {
            _excelService = excelService;
        }  

        [HttpPost]
        public async Task<ActionResult<List<ImportHistory>>> ExcelUpload(IFormFile formFile, string? fileName, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)  
            {  
                return BadRequest("File is empty");  
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))  
            {  
                return BadRequest("File not supported"); 
            } 

            if(!await _excelService.AddToImportTable(formFile, fileName, cancellationToken)) 
            {
                return BadRequest("Import Table failed to Update");
            }

            using (var stream = new MemoryStream())  
            {
                await formFile.CopyToAsync(stream, cancellationToken);  
        
                using (var package = new ExcelPackage(stream)) 
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 3; row <= rowCount; row++)  
                    {
                        ExcelStyle rng = worksheet.Cells[row, 1].Style;
                        string color = rng.Fill.BackgroundColor.LookupColor();

                        var tablename = _excelService.SelcetTablebasedOnColor(rng, worksheet, row);

                        if(tablename == "") return BadRequest("Invalid color found");


                    }
                }
            }
            return Ok();
        }
    }
}