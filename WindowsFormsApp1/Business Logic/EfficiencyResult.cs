using System.Collections.Generic;

public class EfficiencyResult
{
    public double Concentration { get; set; }
    public double Background { get; set; }
    public List<double> Readings { get; set; } = new List<double>();     // 11 筆讀值
    public List<double> Efficiencies { get; set; } = new List<double>(); // 11 筆效率
    public string ConcentrationText => Concentration > 0 ? Concentration.ToString("F0") : "";
    public double? Eff0 =>
        Efficiencies.Count > 0 ? (double?)Efficiencies[0] : null;
    public double? Eff10 =>
        Efficiencies.Count > 10 ? (double?)Efficiencies[10] : null;
}