namespace API.Modals
{
    public class ImportHistory
    {
        public int Id { get; set; }
        public string? FileName { get; set; } = "";
        public string? Name { get; set; } = "";
        public List<ExcelData>? ExcelDatas { get; set; } = null;
    }
}