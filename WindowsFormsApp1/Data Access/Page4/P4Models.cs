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

        public string QtyText { get; set; }
        public string UserName { get; set; }

        public List<string> LotNos { get; set; } = new List<string>();
        public List<string> LotFulls { get; set; } = new List<string>();

        public List<double> Weights { get; set; } = new List<double>();
        public List<double> Densities { get; set; } = new List<double>();

        public List<double> VocIns { get; set; } = new List<double>();
        public List<double> VocOuts { get; set; } = new List<double>();

        public List<double> DeltaPs { get; set; } = new List<double>();
        public List<string> OutgassingList { get; set; } = new List<string>();

        public int SelectedIndex { get; set; }

        public Dictionary<string, double> ParticleSizePercentages
            = new Dictionary<string, double>();

        public List<EfficiencyGroup> EfficiencyGroups
            = new List<EfficiencyGroup>();
    }
    public class P4Lot
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
    public class P4Particle
    {
        public string SizeName { get; set; }
        public double Percentage { get; set; }
    }
    public class P4EfficiencyGroup
    {
        public string GasName { get; set; }
        public double? Concentration { get; set; }
        public double? Eff0 { get; set; }
        public double? Eff10 { get; set; }
        public List<double> Efficiencies { get; set; } = new List<double>();
    }
}