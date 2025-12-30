using System;
using System.Collections.Generic;

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
}
