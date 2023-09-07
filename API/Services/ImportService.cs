using API.Data;
using API.Interfaces;
using API.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace API.Services
{
    public class ImportService : IImportService
    {
        private readonly DataContext _context;
        private static readonly Dictionary<string, int> colors = new(){
            { "#FFFFF2CC", 10},
            { "#FFFBE5D6", 9},
            { "#FFD9D9D9", 8},
            { "#FFFFFFFF", 0},
            { "#FF7DFFB8", 8},
            { "#FFE2F0D9", 7},
            { "#FFDEEBF7", 6},
            { "#FFDCC5ED", 5}
        };
        private static readonly string BLACK_HEX_CODE = "#FF000000";
        private static readonly string WHITE_HEX_CODE = "#FFFFFFFF";

        public ImportService(DataContext context)
        {
            _context = context;
        }

        public async Task<int> AddToImportTable(IFormFile formFile)
        {
            var importFile = new ImportHistory
            {
                FileName = formFile.FileName,
                Name = formFile.FileName
            };

            _context.ImportHistory?.Add(importFile);
            await _context.SaveChangesAsync();
            return importFile.Id;
        }

        public async Task<bool> ImportToDatabase(ExcelWorksheet worksheet, int importId)
        {
            await deleteDataTable();

            var rowCount = worksheet.Dimension.Rows;
            Stack<HeaderItem> stack = new();

            for (int row = 3; row <= rowCount; row++)
            {
                Guid rowId = Guid.NewGuid();
                ExcelStyle rng = worksheet.Cells[row, 1].Style;
                var color = rng.Fill.BackgroundColor.LookupColor();
                color = color == BLACK_HEX_CODE ? WHITE_HEX_CODE : color;
                string style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold;

                Guid? parentId = getParent(stack, rowId, color);
                ExcelData data = ConvertRowToDataTable(rowId, worksheet, row, parentId, importId, style);
                
                _context.TimeSchedule?.Add(data);
            }

            return await _context.SaveChangesAsync() > 0;
        }

        private ExcelData ConvertRowToDataTable(Guid rowId, ExcelWorksheet worksheet, int row, Guid? parentId, int importId, string style)
        {
            return new ExcelData() {
                Identifier = rowId,
                ActivityId = worksheet.Cells[row, 1].Text,
                ActivityName = worksheet.Cells[row, 2].Text,
                BlProjectStart = worksheet.Cells[row, 3].Value == null ? null : Convert.ToDateTime(worksheet.Cells[row, 3].Value),
                BlProjectFinish = worksheet.Cells[row, 4].Value == null ? null : Convert.ToDateTime(worksheet.Cells[row, 4].Value),
                ActualStart = worksheet.Cells[row, 5].Value == null ? null : Convert.ToDateTime(worksheet.Cells[row, 5].Value),
                ActualFinish = worksheet.Cells[row, 6].Value == null ? null : Convert.ToDateTime(worksheet.Cells[row, 6].Value),
                ActivityComplete = worksheet.Cells[row, 7].Value == null ? null : Convert.ToDecimal(worksheet.Cells[row, 7].Value),
                MaterialCostComplete = worksheet.Cells[row, 8].Value == null ? null : Convert.ToDecimal(worksheet.Cells[row, 8].Value),
                LaborCostComplete = worksheet.Cells[row, 9].Value == null ? null : Convert.ToDecimal(worksheet.Cells[row, 9].Value),
                NonLaborCostComplete = worksheet.Cells[row, 10].Value == null ? null : Convert.ToDecimal(worksheet.Cells[row, 10].Value),
                ImportHistoryId = importId,
                ParentId = parentId,
                Style = style
            };
        }

        private Guid? getParent(Stack<HeaderItem> stack, Guid rowId, string color) {
            Guid? parentId = new();
            
            if (stack.Count == 0)
            {
                parentId = null;
                stack.Push(new HeaderItem
                {
                    Color = color,
                    Id = rowId
                });
            }
            else
            {
                HeaderItem topHeaderItem = stack.Peek();
                colors.TryGetValue(color, out int currentColorValue);
                colors.TryGetValue(topHeaderItem.Color, out int topColorValue);

                if (currentColorValue < topColorValue)
                {
                    parentId = topHeaderItem.Id;
                    stack.Push(new HeaderItem
                    {
                        Color = color,
                        Id = rowId
                    });
                }
                else
                {
                    colors.TryGetValue(stack.Peek().Color, out int currTopColorValue);
                    while (currTopColorValue <= currentColorValue)
                    {
                        stack.Pop();
                        colors.TryGetValue(stack.Peek().Color, out currTopColorValue);
                    }

                    HeaderItem headerItem = stack.Peek();
                    parentId = headerItem.Id;
                    stack.Push(new HeaderItem
                    {
                        Color = color,
                        Id = rowId
                    });
                }
            }

            return parentId;
        }

        private async Task deleteDataTable() {
            if(_context.TimeSchedule is null) return;
            await _context.TimeSchedule.ExecuteDeleteAsync();
        }
        
        class HeaderItem {
            public Guid Id { get; set; }
            public string? Color { get; set; } = "";
        }
    }
}