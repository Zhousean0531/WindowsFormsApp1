using System.Collections.Generic;
using System;
namespace WindowsFormsApp1.Data_Access.Page2
{
    public class P2Batch
    {
        public DateTime? ProductionDate { get; set; }
        public DateTime? TestDate { get; set; }
        public string WorkOrder { get; set; }
        public string MaterialNo { get; set; }
        public string BatchNo{get;set;}
        public string ReportNo { get; set; }
        public string ProductSize { get; set; }
        public string ProductDisplay { get; set; }
        public string FilterSize { get; set; }
        public string Material { get; set; }
        public decimal? TargetGsm { get; set; }
        public string Glue { get; set; }
        public decimal? Speed { get; set; }
        public decimal? UpperTemp { get; set; }
        public decimal? LowerTemp { get; set; }
        public decimal? Pressure { get; set; }
        public decimal? WindSpeed { get; set; }
        public string CarbonLine { get; set; }
        public string Username { get; set; }

        public List<P2GasTest> GasTests { get; set; } = new List<P2GasTest>();
    }
    public class P2GasTest
    {
        public string ReportNo { get; set; }
        public string GasName { get; set; }
        public decimal? Concentration { get; set; }
        public decimal? Background { get; set; }

        public List<P2Sample> Samples { get; set; } = new List<P2Sample>();
    }
    public class P2Sample
    {
        public decimal? Weight { get; set; }
        public decimal? Thickness { get; set; }
        public decimal? PressureDrop { get; set; }
        public bool IsSelected { get; set; }

        // 可動態 11~40~N 筆效率
        public List<decimal> Efficiencies { get; set; } = new List<decimal>();

        // Excel 匯入時可只提供指定分鐘，例如 0 與 10 分鐘
        public Dictionary<int, decimal> EfficiencyPoints { get; set; } = new Dictionary<int, decimal>();
    }
}
