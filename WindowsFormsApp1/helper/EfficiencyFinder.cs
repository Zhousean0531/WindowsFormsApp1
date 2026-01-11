using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

public static class EfficiencyFinder
{
    public static Dictionary<string, string> FindMinEfficiencyByCarbonLot(
        IXLWorksheet ws,
        string carbonLot,
        List<string> targetTypes)
    {
        var result = new Dictionary<string, string>
        {
            { "MA", "N/A" },
            { "MB", "N/A" },
            { "MC", "N/A" }
        };

        if (string.IsNullOrWhiteSpace(carbonLot))
            return result;

        string lotKey = carbonLot.Trim();

        foreach (var r in ws.RowsUsed())
        {

            string excelLot = r.Cell("E").GetString().Trim();
            string testItem = r.Cell("K").GetString().Trim();
            string efficiencyStr = r.Cell("P").GetString().Trim();

            // 用 Contains 比對原料批號
            if (!excelLot.ToUpper().Contains(lotKey.ToUpper()))
                continue;

            if (!double.TryParse(efficiencyStr, out double efficiency))
                continue;

            string type = null;

            // ===== 效率分類規則 =====
            if ((testItem == "SO2" || testItem == "H2S") && targetTypes.Contains("MA"))
                type = "MA";
            else if (testItem == "NH3" && targetTypes.Contains("MB"))
                type = "MB";
            else if ((testItem == "Toluene" || testItem == "Acetone" || testItem == "IPA") && targetTypes.Contains("MC"))
                type = "MC";

            if (type == null)
                continue;

            // ===== 取最小效率 =====
            if (result[type] == "N/A")
            {
                result[type] = efficiency.ToString("0.###");
            }
            else if (double.TryParse(result[type], out double existVal) && efficiency < existVal)
            {
                result[type] = efficiency.ToString("0.###");
            }
        }

        return result;
    }
}
