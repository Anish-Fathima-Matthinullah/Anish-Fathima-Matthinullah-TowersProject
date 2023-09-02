using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;

namespace API.Modals
{
    public class ExcelData
    {
        public int Id { get; set; }
        [DisplayName("Activity ID")]
        public string? ActivityId { get; set; } = "";
        [DisplayName("Activity Name")]
        public string? ActivityName { get; set; } = "";
        [DisplayName("BL Project Start")]
        public DateTime? BlProjectStart { get; set; }
        [DisplayName("BL Project Finish")]
        public DateTime? BlProjectFinish { get; set; }
        [DisplayName("Actual Start")]
        public DateTime? ActualStart { get; set; }
        [DisplayName("Actual Finish")]
        public DateTime? ActualFinish { get; set; }
        [DisplayName("Activity % Complete")]
        public decimal? ActivityComplete { get; set; }
        [DisplayName("Material Cost % Complete")]
        public decimal? MaterialCostComplete { get; set; }
        [DisplayName("Labor Cost % Complete")]
        public decimal? LaborCostComplete { get; set; }
        [DisplayName("Non Labor Cost % Complete")]
        public decimal? NonLaborCostComplete { get; set; }
        public int? ImportHistoryId { get; set; }
        public string? Style { get; set; } = "";
    }
}