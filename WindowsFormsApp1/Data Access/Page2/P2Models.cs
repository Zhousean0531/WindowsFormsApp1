using System.Collections.Generic;
using System;
namespace WindowsFormsApp1.Data_Access.Page2
{
    public class P2Batch
    {
        public string ProductionDate { get; set; }
        public string TestDate { get; set; }
        public string WorkOrder { get; set; }
        public string MaterialNo { get; set; }
        public string ReportNo { get; set; }
        public string ProductSize { get; set; }
        public string ProductDisplay { get; set; }
        public string FilterSize { get; set; }
        public string Material { get; set; }
        public string MaterialBatchNo { get; set; }
        public double? TargetGsm { get; set; }
        public double? Glue { get; set; }
        public double? Speed { get; set; }
        public double? UpperTemp { get; set; }
        public double? LowerTemp { get; set; }
        public double? Pressure { get; set; }
        public double? WindSpeed { get; set; }
        public string CarbonLine { get; set; }
        public string Username { get; set; }

        public List<P2GasTest> GasTests { get; set; } = new List<P2GasTest>();
    }

    public class P2GasTest
    {
        public string GasName { get; set; }
        public double? Concentration { get; set; }
        public double? Background { get; set; }

        public List<P2Sample> Samples { get; set; } = new List<P2Sample>();
    }

    public class P2Sample
    {
        public double? Weight { get; set; }
        public double? PressureDrop { get; set; }
        public bool IsSelected { get; set; }

        // 可動態 11~40~N 筆效率
        public List<double> Efficiencies { get; set; } = new List<double>();
    }
}