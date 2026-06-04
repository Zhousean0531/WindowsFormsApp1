using System;
using System.Collections.Generic;
using System.Linq;

public static class EfficiencyHelper
{
    public static bool TryValidateInputs(
        string label,
        string concentrationText,
        string backgroundText,
        string readingText,
        int minimumReadingCount,
        out double concentration,
        out double background,
        out List<double?> readings,
        out string errorMessage)
    {
        concentration = 0;
        background = 0;
        readings = new List<double?>();
        errorMessage = "";

        string name = string.IsNullOrWhiteSpace(label) ? "效率" : label.Trim();

        if (!double.TryParse((concentrationText ?? "").Trim(), out concentration) || concentration <= 0)
        {
            errorMessage = $"{name} 濃度需為數字且大於 0";
            return false;
        }

        if (!double.TryParse((backgroundText ?? "").Trim(), out background))
        {
            errorMessage = $"{name} 背景值需為數字";
            return false;
        }

        var parts = SplitReadingTokens(readingText);

        if (parts.Count == 0)
        {
            errorMessage = $"{name} 讀值不可空白";
            return false;
        }

        foreach (string part in parts)
        {
            if (!double.TryParse(part, out double value))
            {
                errorMessage = $"{name} 讀值只能輸入數字，錯誤值：{part}";
                return false;
            }

            readings.Add(value);
        }

        if (minimumReadingCount > 0 && readings.Count < minimumReadingCount)
        {
            errorMessage = $"{name} 讀值需至少 {minimumReadingCount} 筆";
            return false;
        }

        return true;
    }

    public static List<double?> ParseReadings(string tbVal)
    {
        return SplitReadingTokens(tbVal)
            .Select(s =>
            {
                if (double.TryParse(s, out var v))
                    return (double?)v;
                return null; // 無效值視為空
            })
            .ToList();
    }

    private static List<string> SplitReadingTokens(string text)
    {
        return (text ?? "")
            .Split(new[] { '/', '／', ',', '，', '\n', '\r', ' ', '\t' },
                   StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList();
    }

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
