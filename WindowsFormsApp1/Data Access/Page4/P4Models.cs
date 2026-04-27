using System.Collections.Generic;

namespace WindowsFormsApp1.Data_Access.Page4
{
    public class P4Batch
    {
        public string ReportNo { get; set; }
        public string Material { get; set; }
        public string MaterialNo { get; set; }

        public string ArrivalDate { get; set; }
        public string TestingDate { get; set; }
        public double Density { get; set; }
        public string QtyText { get; set; }
        public string UserName { get; set; }

        // ⭐ 核心
        public string SupplierLot { get; set; }
        public string FactoryLot { get; set; }

        public List<P4Row> Rows { get; set; } = new List<P4Row>();

        public decimal? Moisture { get; set; }
        public decimal? Butane { get; set; }
        public decimal? Ash { get; set; }

        public Dictionary<string, double> ParticleSizePercentages
            = new Dictionary<string, double>();

        public List<P4EfficiencyGroup> EfficiencyGroups
            = new List<P4EfficiencyGroup>();
    }

    public class P4Row
    {
        public string LotNo { get; set; }
        public string LotFull { get; set; }

        public double Weight { get; set; }
        public double Density { get; set; }
        public double VocIn { get; set; }
        public double VocOut { get; set; }
        public double DeltaP { get; set; }

        public string Outgassing { get; set; }
        public bool IsSelected { get; set; }
    }

    public class P4EfficiencyGroup
    {
        public string GasName { get; set; }
        public double? Concentration { get; set; }

        public double? Eff0 { get; set; }   // 初始
        public double? Eff10 { get; set; }  // 40分鐘

        public List<double> Efficiencies11 { get; set; } = new List<double>();
    }
}