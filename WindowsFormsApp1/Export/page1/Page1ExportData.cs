using System;
using System.Collections.Generic;

namespace WindowsFormsApp1.Export.Page1
{
    public class Page1ExportData
    {
        public string UserName { get; set; }
        public string ReportNo { get; set; }
        public string FilterRawQtyWeight { get; set; }
        public string FilterRawQuantity { get; set; }
        public string QtyText { get; set; }
        public string Material { get; set; }
        public DateTime ArrivalDate { get; set; }
        public string MaterialNo { get; set; }
        public DateTime TestingDate { get; set; }
        public List<string> LotFulls { get; set; }
        public List<double> Densities { get; set; }
        public List<double> DeltaPs { get; set; }
        public List<double> VocIns { get; set; }
        public List<double> VocOuts { get; set; }
        public Dictionary<string, double> ParticleSizes { get; set; }
        public List<string> OutgassingList { get; set; }
        public List<string> MeshSummaries { get; set; }

        public int SelectedIndex { get; set; }
        public double Eff0 { get; set; }
        public double Eff10 { get; set; }
        public List<double> Efficiencies11 { get; set; }

        // 新增：第二欄百分比（對應 Excel B7,B8...）
        public Dictionary<string, double> ParticleSizePercentages { get; set; }
    }
}
