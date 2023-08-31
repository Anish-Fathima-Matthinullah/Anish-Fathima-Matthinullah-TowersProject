using API.Data;
using API.Interfaces;
using API.Modals;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Services
{
    public class ExcelService : IExcelService
    {
        private readonly DataContext _context;
        public ExcelService(DataContext context)
        {
            _context = context;
        }

        public async Task<bool> SaveAllAsync()
        {
            return await _context.SaveChangesAsync() > 0;
        }

        public async Task<int> AddToImportTable(IFormFile formFile, string? fileName, CancellationToken cancellationToken)
        {
            var importFile = new ImportHistory
            {
                FileName = formFile.FileName,
                Name = fileName == null ? formFile.FileName : fileName
            };

            _context.ImportHistory?.Add(importFile);
            await _context.SaveChangesAsync(cancellationToken);
            return importFile.Id;
        }

        // public async Task<string> SelcetTablebasedOnColor(ExcelStyle rng, ExcelWorksheet worksheet, int row)
        // {
        //     var color = rng.Fill.BackgroundColor.LookupColor();
        //     var tableName = "";
        //     switch(color)
        //     {
        //         case "#FFFFF2CC":
        //             tableName = "Building";
        //             Building building = new()
        //                 {  
        //                     ActivityId = Convert.ToString(worksheet.Cells[row, 1].Value),  
        //                     ActivityName = Convert.ToString(worksheet.Cells[row, 2].Value),
        //                     BlProjectStart = Convert.ToDateTime(worksheet.Cells[row, 3].Value),
        //                     BlProjectFinish = Convert.ToDateTime(worksheet.Cells[row, 4].Value),
        //                     ActualStart = Convert.ToDateTime(worksheet.Cells[row, 5].Value),
        //                     ActualFinish = Convert.ToDateTime(worksheet.Cells[row, 6].Value),
        //                     ActivityComplete = Convert.ToDecimal(worksheet.Cells[row, 7].Value) * 100,
        //                     MaterialCostComplete = Convert.ToDecimal(worksheet.Cells[row, 8].Value) * 100,
        //                     LaborCostComplete = Convert.ToDecimal(worksheet.Cells[row, 9].Value) * 100,
        //                     NonLaborCostComplete = Convert.ToDecimal(worksheet.Cells[row, 10].Value) * 100,
        //                     ImportHistoryId = Convert.ToInt32(1),
        //                     Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
        //                 };
        //             _context.Building?.Add(building);
        //             await _context.SaveChangesAsync();
        //             break;

        //         case "#FFFBE5D6":
        //             tableName = "Tower";
        //             break;
        //         case "#FFD9D9D9":
        //             tableName = "Milestone";
        //             break;
        //         case "#FF000000":
        //             tableName = "Activity";
        //             break;
        //         case "#FF7DFFB8":
        //             tableName = "Area";
        //             break;
        //         case "#FFE2F0D9":
        //             tableName = "Region";
        //             break;
        //         case "#FFDEEBF7":
        //             tableName = "Floor";
        //             break;
        //         case "#FFDCC5ED":
        //             tableName = "Work";
        //             break;
        //         default: 
        //             break;
        //     }
        //     return tableName;
        // }

        // public void CreateModel(string tablename)
        // {
        //     switch(tablename)
        //     {
        //         case "Building":

        //         break;
        //     }
        // }

        public async Task<bool> SelectTablebasedOnColor(ExcelStyle rng, ExcelWorksheet worksheet, int row, int importId)
        {
            var color = rng.Fill.BackgroundColor.LookupColor();
            int id = 0;
            ParentInfo parentInfo = new();
            switch(color)
            {
                case "#FF000000":
                    Activity activity = new()
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
                            HeaderId = id,
                            Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
                        };
                        _context.Activity?.Add(activity);
                    break;
                
                default: 
                    Header headerRow = new()
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
                            ImportHistoryId = Convert.ToInt32(importId),
                            Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
                        };
                    
                    if(headerRow.ActivityId != null)
                    {
                        var trimLength = headerRow.ActivityId.Length - headerRow.ActivityId.TrimStart(' ').Length;
                        if(trimLength == 0)
                        {
                            parentInfo.Length = trimLength;
                        }
                        else if(trimLength - parentInfo.Length == 2)
                        {
                            parentInfo.Length = trimLength;
                            parentInfo.Id = id;
                            headerRow.ParentId = parentInfo.Id;
                        } 
                        else if(trimLength == parentInfo.Length)
                        {
                            headerRow.ParentId = parentInfo.Id;
                        }
                    }
                    
                    _context.Header?.Add(headerRow);
                    id = headerRow.Id;
                    break;
            }
            return await _context.SaveChangesAsync() > 0;
        }
    }
}