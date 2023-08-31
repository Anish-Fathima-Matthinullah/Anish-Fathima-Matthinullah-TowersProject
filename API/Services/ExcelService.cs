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
        public ExcelService(DataContext context)
        {
            _context = context;
        }

        public async Task<bool> SaveAllAsync()
        {
            return await _context.SaveChangesAsync() > 0;
        }

        public async Task<bool> AddToImportTable(IFormFile formFile, string? fileName, CancellationToken cancellationToken)
        {
            var importFile = new ImportHistory
            {
                FileName = formFile.FileName,
                Name = fileName == null ? formFile.FileName : fileName
            };

            _context.ImportHistory?.Add(importFile);
            return await _context.SaveChangesAsync(cancellationToken) > 0;
        }

        public string SelcetTablebasedOnColor(ExcelStyle rng, ExcelWorksheet worksheet, int row)
        {
            var color = rng.Fill.BackgroundColor.LookupColor();
            var tableName = "";
            switch(color)
            {
                case "#FFFFF2CC":
                    tableName = "Building";
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
                            ImportHistoryId = Convert.ToInt32(1),
                            Style = color + ", " + rng.Font.Size + ", " + rng.Font.Bold
                        };
                    break;
                case "#FFFBE5D6":
                    tableName = "Tower";
                    break;
                case "#FFD9D9D9":
                    tableName = "Milestone";
                    break;
                case "#FF000000":
                    tableName = "Activity";
                    break;
                case "#FF7DFFB8":
                    tableName = "Area";
                    break;
                case "#FFE2F0D9":
                    tableName = "Region";
                    break;
                case "#FFDEEBF7":
                    tableName = "Floor";
                    break;
                case "#FFDCC5ED":
                    tableName = "Work";
                    break;
                default: 
                    break;
            }
            return tableName;
        }

        // public void CreateModel(string tablename)
        // {
        //     switch(tablename)
        //     {
        //         case "Building":

        //         break;
        //     }
        // }
    }
}