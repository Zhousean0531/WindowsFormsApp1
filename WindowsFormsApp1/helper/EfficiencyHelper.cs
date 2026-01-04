using System;
using System.Collections.Generic;
using System.Linq;

public static class EfficiencyHelper
{
    // 只負責：字串 → 11 筆數值
    public static List<double> ParseReadings(string raw)
    {
        return (raw ?? "")
            .Split(new[] { '/', ',', '\n', '\r', ' ' },
                   StringSplitOptions.RemoveEmptyEntries)
            .Select(s => double.TryParse(s, out var v) ? v : double.NaN)
            .Where(v => !double.IsNaN(v))
            .Take(11)
            .ToList();
    }

    // 只負責：濃度 + 背景 + 11 筆 → 計算結果
    public static EfficiencyResult Compute11Points(
        double concentration,
        double background,
        List<double> readings11)
    {
        if (concentration <= 0)
            throw new ArgumentException("Concentration must be > 0");

        if (readings11 == null || readings11.Count < 11)
            throw new ArgumentException("Readings must contain 11 values");

        var result = new EfficiencyResult
        {
            Concentration = concentration,
            Background = background,
            Readings = readings11
        };

        foreach (var r in readings11)
        {
            double eff = (concentration + background - r) / concentration * 100.0;
            result.Efficiencies.Add(Math.Round(eff, 1));
        }

        return result;
    }
}
