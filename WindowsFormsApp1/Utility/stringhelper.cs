using System;
using System.Collections.Generic;
using System.Linq;

public static class ParseHelper
{
    public static List<string> SplitStr(string raw)
    {
        var sep = new[] { '/', ',', '、', '|' };
        return (raw ?? "")
            .Split(sep, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim())
            .ToList();
    }

    public static List<double> SplitDouble(string raw)
    {
        return SplitStr(raw)
            .Select(s => double.TryParse(s, out var v) ? v : 0)
            .ToList();
    }

    public static bool TrySplitDouble(string raw, out List<double> values)
    {
        values = new List<double>();

        foreach (string token in SplitStr(raw))
        {
            if (!double.TryParse(token, out double value))
                return false;

            values.Add(value);
        }

        return true;
    }
}
