using System.Drawing;
using API.Dtos;
using API.Interfaces;
using API.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace API.Services
{
    public class StyleService : IStyleService
    {
        private static readonly string TITLE = "Time Schedule Follow Up";
        private static readonly string FONT_FAMILY = "Calibri";
        private static readonly string COLUMN_HEADERS = "Activity ID, Activity Name, BL Project Start, BL Project Finish, Actual Start, Actual Finish, Activity % Complete, Material Cost % Complete, Labor Cost % Complete, Non Labor Cost % Complete";
        private static readonly string DEFAULT_STYLE = "#FFFFFFFF, 11, false";

        public void SetWorksheetStyle(ExcelWorksheet workSheet)
        {
            workSheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells.Style.Font.Name = FONT_FAMILY;

            workSheet.Rows.Height = 17;
            workSheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            workSheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

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

            SetHeaderStyle(workSheet);
            SetColumnHeaderStyle(workSheet);
        }
        
        public void SetDataFormatStyle(ExcelWorksheet workSheet, ExcelListDetailsDto? excelListDetails)
        {
            SetPercentFormat(workSheet);
            if(excelListDetails is null)
                return;
                
            SetDateFormat(workSheet, excelListDetails.ExcelList);

            List<ExcelData>? rowItems = excelListDetails.TableData;
            if (rowItems != null)
            {
                int i = 3;
                foreach (ExcelData item in rowItems)
                {
                    if (item.Style == null) item.Style = DEFAULT_STYLE;
                    string[] styleInfo = item.Style.Split(", ");

                    var query = from cell in workSheet.Cells["A:J"]
                                where cell.Value?.ToString()?.Contains(item.ActivityId ?? "") == true
                                select cell;

                    foreach (var cell in query)
                    {
                        cell.Style.Font.Size = Convert.ToInt16(styleInfo[1]);
                        cell.Style.Font.Bold = Convert.ToBoolean(styleInfo[2]);
                    }

                    Color colFromHex = ColorTranslator.FromHtml(styleInfo[0]);
                    workSheet.Cells[i, 1, i, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[i, 1, i, 10].Style.Fill.BackgroundColor.SetColor(colFromHex);

                    var rowCount = workSheet.Dimension.Rows;
                    var borderCells = from cell in workSheet.Cells[1, 1, rowCount, 10]
                                      select cell;

                    foreach (var cell in borderCells)
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells[1, 1, rowCount, 10].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    workSheet.Cells["A1:J1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    workSheet.Cells["A2:J2"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    i++;
                }
            }
        }

        private void SetPercentFormat(ExcelWorksheet workSheet)
        {
            workSheet.Cells["G:J"].Style.Numberformat.Format = "0%";
        }

        private void SetDateFormat(ExcelWorksheet workSheet, List<ExcelDataDto>? data)
        {
            if(data is null) return;
            var indexs = FindDateTimeColumns(data.FirstOrDefault()).ToList();
            indexs.ForEach(i => {
                workSheet.Column(++i).Style.Numberformat.Format = "d-mmm-yy"; // assign datetime style
            });

            // find datetime columns
            static IEnumerable<int> FindDateTimeColumns(ExcelDataDto? obj)
            {
                if(obj is null) return Enumerable.Empty<int>();
                var indexs = obj.GetType()
                    .GetProperties()
                    .Select((item, index) => new { Item = item, Index = index })
                    .Where(o => o.Item.PropertyType == typeof(DateTime?))
                    .Select(o => o.Index);

                return indexs;
            }
        }

        private void SetHeaderStyle(ExcelWorksheet workSheet)
        {
            ExcelRange header = workSheet.Cells["A1:J1"];
            header.Merge = true;
            header.Value = TITLE;

            workSheet.Row(1).Height = 27;
            workSheet.Row(1).Style.Font.Size = 20;
            workSheet.Row(1).Style.Font.Color.SetColor(1, 0, 32, 96);
            workSheet.Row(1).Style.Font.Bold = true;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private void SetColumnHeaderStyle(ExcelWorksheet workSheet)
        {
            ExcelRange columnHeader = workSheet.Cells["A2:J2"];
            columnHeader.LoadFromText(COLUMN_HEADERS);

            workSheet.Row(2).Height = 94.5;
            workSheet.Row(2).Style.Font.Size = 14;
            workSheet.Row(2).Style.Font.Color.SetColor(1, 0, 32, 96);
            workSheet.Row(2).Style.Font.Bold = true;
            workSheet.Row(2).Style.WrapText = true;
            workSheet.Row(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            workSheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
    }
}