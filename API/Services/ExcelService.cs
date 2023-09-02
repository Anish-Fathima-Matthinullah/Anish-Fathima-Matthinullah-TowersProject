using API.Data;
using API.Interfaces;
using API.Modals;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Services
{
    public class ExcelService : IExcelService
    {
        private readonly DataContext _context;
        private readonly int importId;

        private static Dictionary<string, int> colors = new(){
            { "#FFFFF2CC", 10},
            { "#FFFBE5D6", 9},
            { "#FFD9D9D9", 8},
            { "#FF000000", 0},
            { "#FF7DFFB8", 8},
            { "#FFE2F0D9", 7},
            { "#FFDEEBF7", 6},
            { "#FFDCC5ED", 5}
        };

        public ExcelService(DataContext context)
        {
            _context = context;
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

        

        // public async Task<bool> SelectTablebasedOnColor2(ExcelStyle rng, ExcelWorksheet worksheet, int row, int importId)
        // {
        //     var color = rng.Fill.BackgroundColor.LookupColor();
        //     Guid id = Guid.NewGuid();
        //     ParentInfo parentInfo = new();
        //     switch(color)
        //     {
        //         case "#FF000000":
        //             Activity activity = new()
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
        //                     HeaderId = id,
        //                     Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
        //                 };
        //                 _context.Activity?.Add(activity);
        //             break;
                
        //         default: 
        //             Header headerRow = new()
        //                 {  
        //                     Id = Guid.NewGuid(),
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
        //                     ImportHistoryId = Convert.ToInt32(importId),
        //                     Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
        //                 };
                    
        //             if(headerRow.ActivityId != null)
        //             {
        //                 var trimLength = headerRow.ActivityId.Length - headerRow.ActivityId.TrimStart(' ').Length;
        //                 if(trimLength == 0)
        //                 {
        //                     parentInfo.Length = trimLength;
        //                 }
        //                 else if(trimLength - parentInfo.Length == 2)
        //                 {
        //                     parentInfo.Length = trimLength;
        //                     // parentInfo.Id = id;
        //                     headerRow.ParentId = parentInfo.Id;
        //                 } 
        //                 else if(trimLength == parentInfo.Length)
        //                 {
        //                     headerRow.ParentId = parentInfo.Id;
        //                 }
        //             }
                    
        //             _context.Header?.Add(headerRow);
        //             id = headerRow.Id;
        //             break;
        //     }
        //     return await _context.SaveChangesAsync() > 0;
        // }

        public async Task<bool> SelectTablebasedOnColor(ExcelWorksheet worksheet, int importId)
        {
            var rowCount = worksheet.Dimension.Rows;
            Stack<HeaderItem> stack = new();
            Guid id = Guid.NewGuid();

            for (int row = 3; row <= rowCount; row++)  
            {
                Guid rowId = Guid.NewGuid();
                Guid? parentId = new();
                ExcelStyle rng = worksheet.Cells[row, 1].Style;
                var color = rng.Fill.BackgroundColor.LookupColor();

                if(stack.Count == 0)
                {
                    parentId = null;
                    stack.Push(new HeaderItem {
                        Color = color,
                        Id = rowId
                    });
                }
                else {
                    HeaderItem topHeaderItem = stack.Peek();

                    colors.TryGetValue(color, out int currentColorValue);
                    colors.TryGetValue(topHeaderItem.Color, out int topColorValue);

                    if(currentColorValue < topColorValue) {
                        parentId = topHeaderItem.Id;
                        stack.Push(new HeaderItem{
                            Color = color,
                            Id = rowId
                        });
                    }
                    else {
                        colors.TryGetValue(stack.Peek().Color, out int currTopColorValue);
                        while(currTopColorValue <= currentColorValue)
                        {
                            stack.Pop();
                            colors.TryGetValue(stack.Peek().Color, out currTopColorValue);
                        }

                        HeaderItem headerItem = stack.Peek();
                        parentId = headerItem.Id;
                        stack.Push(new HeaderItem {
                            Color = color,
                            Id = rowId
                        });
                    }
                }
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
                                HeaderId = (Guid)parentId,
                                Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
                            };
                            _context.Activity?.Add(activity);
                        break;
                    
                    default: 
                        Header headerRow = new()
                            {  
                                Identifier = rowId,
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
                                ParentId = parentId,
                                Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
                            };
                        
                        _context.Header?.Add(headerRow);
                        break;
                    }
            }
            
            return await _context.SaveChangesAsync() > 0;
        }
    }
}