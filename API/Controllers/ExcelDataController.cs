using Microsoft.AspNetCore.Mvc;
using API.Modals;
using OfficeOpenXml;
using API.Data;
using OfficeOpenXml.Style;
using Microsoft.EntityFrameworkCore;
using System.Drawing;
using AutoMapper;
using API.Dtos;
using AutoMapper.QueryableExtensions;

namespace API.Controllers
{
    public class ExcelDataController : BaseApiController
    {
        private readonly DataContext _context;
        private readonly IMapper _mapper;
        public ExcelDataController(DataContext context, IMapper mapper)
        {
            _mapper = mapper;
            _context = context;
        }

        [HttpPost]
        public async Task<ActionResult<List<string>>> ExcelUpload(IFormFile formFile, CancellationToken cancellationToken, string fileName)
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
            List<string> styleInfo = new();
  
            using (var stream = new MemoryStream())  
            {  
                await formFile.CopyToAsync(stream, cancellationToken);  
        
                using (var package = new ExcelPackage(stream))  
                {  
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  
                    var rowCount = worksheet.Dimension.Rows;

                    styleInfo.Add(
                        " headerHeight = " + Convert.ToString(worksheet.Row(1).Height) +
                        " headerFontSize = " + Convert.ToString(worksheet.Row(1).Style.Font.Size) +
                        " headerFont = " + Convert.ToString(worksheet.Row(1).Style.Font.Family) + 
                        " headerFontColor = " + Convert.ToString(worksheet.Row(1).Style.Font.Color) +
                        " columnHeight = " + Convert.ToString(worksheet.Row(2).Height) +
                        " columnFontSize = " + Convert.ToString(worksheet.Row(2).Style.Font.Size) +
                        " columnFont = " + Convert.ToString(worksheet.Row(2).Style.Font.Family) + 
                        " columnFontColor = " + Convert.ToString(worksheet.Row(2).Style.Font.Color) + 
                        " column1Width = " + Convert.ToString(worksheet.Column(1).Width)+
                        " column2Width = " + Convert.ToString(worksheet.Column(2).Width)+
                        " column3Width = " + Convert.ToString(worksheet.Column(3).Width)+
                        " column4Width = " + Convert.ToString(worksheet.Column(4).Width)+
                        " column5Width = " + Convert.ToString(worksheet.Column(5).Width)+
                        " column6Width = " + Convert.ToString(worksheet.Column(6).Width)+
                        " column7Width = " + Convert.ToString(worksheet.Column(7).Width)+
                        " column8Width = " + Convert.ToString(worksheet.Column(8).Width)+
                        " column9Width = " + Convert.ToString(worksheet.Column(9).Width)+
                        " column10Width = " + Convert.ToString(worksheet.Column(10).Width)
                    );
                    
                    for (int row = 3; row <= rowCount; row++)  
                    {    
                        ExcelStyle rng = worksheet.Cells[row, 1].Style;
                        string color = rng.Fill.BackgroundColor.LookupColor();

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
        public async Task<ActionResult<IEnumerable<ImportHistory>>> GetData()
        {
            if (_context.ImportHistory is null)
            {
                return BadRequest("No records to display");
            }
            
            return await _context.ImportHistory
                            .Include(b => b.ExcelDatas)
                            .ToListAsync();
        }

        [HttpGet("exportv2")]  
        public async Task<IActionResult> ExportV2(CancellationToken cancellationToken)  
        {
            var stream = new MemoryStream();  
  
            using (var package = new ExcelPackage(stream))  
            {  
                var workSheet = package.Workbook.Worksheets.Add("Sample");
                workSheet.TabColor = Color.Black;

                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Height = 27;
                workSheet.Row(1).Style.Font.Size = 20;
                workSheet.Row(1).Style.Font.Name = "Calibri";
                workSheet.Row(1).Style.Font.Color.SetColor(1,0,32,96);
                workSheet.Row(1).Style.Font.Bold = true;

                workSheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Row(2).Height = 94.5;
                workSheet.Row(2).Style.Font.Size = 14;
                workSheet.Row(2).Style.Font.Name = "Calibri";
                workSheet.Row(2).Style.Font.Color.SetColor(1,0,32,96);
                workSheet.Row(2).Style.Font.Bold = true;
                workSheet.Row(2).Style.WrapText = true;
                
                ExcelRange header = workSheet.Cells["A1:J1"];
                ExcelRange columnHeader = workSheet.Cells["A2:J2"];
                header.Merge = true;
                header.Value = "Time Schedule Follow Up";
                // header.Style.Font.SetFromFont("", 20, true);
                // header.Style.Font.Color.SetColor(0,0,0,128);
                // header.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                
                // columnHeader.LoadFromText("Activity ID, Activity Name, BL Project Start, BL Project Finish, Actual Start, Actual Finish, Activity % Complete, Material Cost % Complete, Labor Cost % Complete, Non Labor Cost % Complete");
                // workSheet.Cells.LoadFromCollection(list, true);
                workSheet.Column(1).Width = 68;
                workSheet.Column(2).Width = 62;
                workSheet.Column(3).Width = 22;
                workSheet.Column(4).Width = 14;
                workSheet.Column(5).Width = 13;
                workSheet.Column(6).Width = 11;
                workSheet.Column(7).Width = 13;
                workSheet.Column(8).Width = 11;
                workSheet.Column(9).Width = 12;   
                workSheet.Column(10).Width = 16;
                workSheet.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thick);

                if (_context.Buildings is null) return BadRequest("No records to display");
                
                // IEnumerable<ExcelDataDto> databaseItems = await _context.Buildings
                //                                     .ProjectTo<ExcelDataDto>(_mapper.ConfigurationProvider)
                //                                     .ToListAsync(cancellationToken: cancellationToken);
                
                // if(databaseItems == null) return NotFound();
                // workSheet.Cells[2,1].LoadFromCollection(databaseItems, true); 

                IEnumerable<ExcelData> databaseItems = await _context.Buildings.ToListAsync(cancellationToken: cancellationToken);
                IEnumerable<ExcelDataDto> items = _mapper.Map<IEnumerable<ExcelDataDto>>(databaseItems);
                workSheet.Cells[2,1].LoadFromCollection(items, true); 
                columnHeader.LoadFromText("Activity ID, Activity Name, BL Project Start, BL Project Finish, Actual Start, Actual Finish, Activity % Complete, Material Cost % Complete, Labor Cost % Complete, Non Labor Cost % Complete");
                package.Save();  
            } 
            stream.Position = 0;  
            string excelName = $"UserList-{DateTime.Now:dd-MM-yyyy}.xlsx";
            return File(stream, "application/octet-stream", excelName);  
        }     
    }
}
