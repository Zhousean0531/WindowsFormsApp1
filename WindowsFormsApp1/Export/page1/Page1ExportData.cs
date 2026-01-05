using System.Collections.Generic;

public class Page1ExportData
{
    public string ReportNo { get; set; }
    public string FilterRawQtyWeight { get; set; }
    public string FilterRawQuantity { get; set; }
    public string QtyText { get; set; }
    public string Material { get; set; }
    public string ArrivalDate { get; set; }
    public string MaterialNo { get; set; }
    public string TestingDate { get; set; }
    public List<string> LotFulls { get; set; }
    public List<double> Densities { get; set; }
    public List<double> DeltaPs { get; set; }
    public List<double> VocIns { get; set; }
    public List<double> VocOuts { get; set; }
    public Dictionary<string, double> ParticleSizes { get; set; }
    public List<string> OutgassingList { get; set; }
    public List<string> MeshSummaries { get; set; }

    public int SelectedIndex { get; set; }
    public string Eff0 { get; set; }
    public string Eff10 { get; set; }
}
