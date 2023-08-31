using Microsoft.AspNetCore.Mvc;
using API.Modals;
using OfficeOpenXml;
using API.Data;
using OfficeOpenXml.Style;
using Microsoft.EntityFrameworkCore;

namespace API.Controllers
{
    public class ExcelDataController : BaseApiController
    {
        private readonly DataContext _context;
        public ExcelDataController(DataContext context)
        {
            _context = context;
        }

        [HttpPost]
        public async Task<ActionResult<List<ExcelData>>> ExcelUpload(IFormFile formFile, CancellationToken cancellationToken, string fileName)
        {
            if (formFile == null || formFile.Length <= 0)  
            {  
                return BadRequest("File is empty");  
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))  
            {  
                return BadRequest("File not supported"); 
            } 

            var importFile = new ImportHistory{
                FileName = formFile.FileName,
                Name = fileName == "" ? formFile.FileName : fileName
            };

            _context.ImportHistory?.Add(importFile);
            await _context.SaveChangesAsync(cancellationToken);

            var excelList = new List<ImportHistory?>();
  
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
                        string color = rng.Fill.BackgroundColor.LookupColor().Replace("#", "");

                        ExcelData data = new ExcelData()
                        {  
                            ActivityId = Convert.ToString(worksheet.Cells[row, 1].Value),  
                            ActivityName = Convert.ToString(worksheet.Cells[row, 2].Value),
                            BlProjectStart = Convert.ToDateTime(worksheet.Cells[row, 3].Value),
                            BlProjectFinish = Convert.ToDateTime(worksheet.Cells[row, 4].Value),
                            ActualStart = Convert.ToDateTime(worksheet.Cells[row, 5].Value),
                            ActualFinish = Convert.ToDateTime(worksheet.Cells[row, 6].Value),
                            ActivityComplete = Convert.ToDecimal(worksheet.Cells[row, 7].Value) * 100,
                            MaterialCostComplete = Convert.ToDecimal(worksheet.Cells[row, 8].Value) * 100,
                            LaborCostComplete = Convert.ToDecimal(worksheet.Cells[row, 9].Value) * 100,
                            NonLaborCostComplete = Convert.ToDecimal(worksheet.Cells[row, 10].Value) * 100,
                            ImportHistoryId = Convert.ToInt32(importFile.Id),
                            Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
                        };



                        excelList.Add(importFile);
                        _context.Buildings?.Add(data);
                    }  
                }  
            } 
            await _context.SaveChangesAsync(cancellationToken);
            return Ok(excelList);
                   
        }

        [HttpGet]
        public async Task<IEnumerable<ImportHistory>> GetData()
        {
            return await _context.ImportHistory?
                            .Include(b => b.ExcelDatas)
                            .ToListAsync();
        }
                
    }
}
