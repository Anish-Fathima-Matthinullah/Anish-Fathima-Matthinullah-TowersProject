namespace API.Modals
{
    public class Area
    {
        public int Id { get; set; }
        public string? ActivityId { get; set; } = "";
        public string? ActivityName { get; set; } = "";
        public DateTime BlProjectStart { get; set; }
        public DateTime BlProjectFinish { get; set; }
        public DateTime ActualStart { get; set; }
        public DateTime ActualFinish { get; set; }
        public decimal ActivityComplete { get; set; }
        public decimal MaterialCostComplete { get; set; }
        public decimal LaborCostComplete { get; set; }
        public decimal NonLaborCostComplete { get; set; }
        public int TowerId { get; set; }
        public string? Style { get; set; } = "";
        public List<Region> Regions { get; set; } = new();
    }
}