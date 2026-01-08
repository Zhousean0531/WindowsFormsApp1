using System.Collections.Generic;
using static EfficiencyHelper;

public class Page4ExportData
{
    // ===== 基本資訊 =====
    public string ReportNo { get; set; }
    public string Material { get; set; }
    public string ArrivalDate { get; set; }
    public string TestingDate { get; set; }
    public string Lot { get; set; }
    public string QtyText { get; set; }
    public List<string> LotFulls { get; set; }
    public List<double> Weights { get; set; }
    public List<double> PressureDrops { get; set; }
    public string MaterialNo { get; set; }
    public List<double> Densities { get; set; }
    public List<double> DeltaPs { get; set; }
    public List<double> VocIns { get; set; }
    public List<double> VocOuts { get; set; }
    public List<string> LotNos { get; set; }
    public List<string> OutgassingList { get; set; }
    public List<string> MeshSummaries { get; set; }
    public int SelectedIndex { get; set; }
    public List<EfficiencyGroup> EfficiencyGroups { get; set; }
}
