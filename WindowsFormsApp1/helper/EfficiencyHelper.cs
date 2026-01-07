using System;
using System.Collections.Generic;
using System.Linq;

public static class EfficiencyHelper
{
    // 只負責：字串 → 11 筆數值
    public static List<double?> ParseReadings(string tbVal)
    {
        return (tbVal ?? "")
            .Split(new[] { '/', ',', '\n', '\r', ' ' },
                   StringSplitOptions.RemoveEmptyEntries)
            .Select(s =>
            {
                if (double.TryParse(s, out var v))
                    return (double?)v;
                return null; // 無效值視為空
            })
            .ToList();
    }

    // 只負責：濃度 + 背景 + 11 筆 → 計算結果
    public static EfficiencyResult Compute11Points(
    double concentration,
    double background,
    List<double?> readings)
    {
        if (concentration <= 0)
            throw new ArgumentException("濃度需大於0");

        if (readings == null || readings.Count == 0)
            throw new ArgumentException("請確認讀值");

        var result = new EfficiencyResult
        {
            Concentration = concentration,
            Background = background
        };

        int count = 0;

        foreach (var r in readings)
        {
            if (r == null)
                break; // ⭐ 遇到空值，停止計算

            double eff =
                (concentration + background - r.Value)
                / concentration * 100.0;

            result.Readings.Add(r.Value);
            result.Efficiencies.Add(Math.Round(eff, 1));
            count++;
        }
        return result;
    }

}
