using System.Collections.Generic;

public class Page3ExportData
{
    public string TestingDate { get; set; }
    public string ArrivalDate { get; set; }
    public string Material { get; set; }
    public List<string> LotNos { get; set; }
    public List<double> Weights { get; set; }
    public List<double> DeltaPs { get; set; }
    public List<double> VocIns { get; set; }
    public List<double> VocOuts { get; set; }
    public List<string> OutMinusIn { get; set; }
}
