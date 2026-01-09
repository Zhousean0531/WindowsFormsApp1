using System.Collections.Generic;

public class EfficiencyGroup
{
    public string GasName { get; set; }
    public EfficiencyResult Result { get; set; }
    public double Concentration { get; set; }
    public string Eff0 { get; set; }
    public string Eff10 { get; set; }
    public List<double> Readings11 { get; set; }
    public List<double> Efficiencies11 { get; set; }

}
