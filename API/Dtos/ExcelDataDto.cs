namespace API.Dtos
{
    public class ExcelDataDto
    {
        public string? ActivityId { get; set; } = "";
        public string? ActivityName { get; set; } = "";
        public DateTime? BlProjectStart { get; set; }
        public DateTime? BlProjectFinish { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualFinish { get; set; }
        public decimal? ActivityComplete { get; set; }
        public decimal? MaterialCostComplete { get; set; }
        public decimal? LaborCostComplete { get; set; }
        public decimal? NonLaborCostComplete { get; set; }
    }
}