using System.Collections.Generic;
using System.Windows.Forms;

public class Page2ExportData
{
    public string UserName { get; set; }
    public string ReportNo { get; set; }
    public List<string> TypeMaterialDisplays { get; set; }
    public string TestDate { get; set; }
    public string OrderDisplay { get; set; }
    public string CarbonOrder { get; set; }
    public string ProductName { get; set; }
    public string ProductNo { get; set; }
    public string ProductDisplay { get; set; }
    public string FilterSize { get; set; }
    public string Wind { get; set; }
    public string ProductType { get; set; }
    public string Gsm { get; set; }
    public string ProductionDate { get; set; }
    public List<double> TestWeights { get; set; }
    public List<double> PressureDrops { get; set; }
    public string Pressure { get; set; }
    public string CarbonInfo { get; set; }
    public int SelectedIndex { get; set; }
    public string TestGas { get; set; }
    public string Gile { get; set; }
    public string Speed { get; set; }
    public string Upper { get; set; }
    public string Lower { get; set; }
    public List<EfficiencyGroup> EfficiencyGroups { get; set; }
    // 11 點效率結果
}
