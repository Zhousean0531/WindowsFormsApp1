using System;
using System.Collections.Generic;
using System.Linq;

public static class StringUtil
{
    public static List<string> Split(string raw)
    {
        var sep = new[] { '/', ',', '、', '|', ' ', '\r', '\n', ';' };
        return (raw ?? "")
            .Split(sep, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim())
            .ToList();
    }

    public static List<double> SplitDouble(string raw)
    {
        return Split(raw)
            .Select(s => double.TryParse(s, out var v) ? v : double.NaN)
            .Where(v => !double.IsNaN(v))
            .ToList();
    }
}
