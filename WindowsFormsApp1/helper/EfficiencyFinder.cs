using ClosedXML.Excel;
using System;
using System.Collections.Generic;

public static class EfficiencyFinder
{
    public static Dictionary<string, string> FindEfficiencyByTypeWithMin(IXLWorksheet ws, string carbonLot, List<string> targetTypes)
    {
        // 建立結果字典，先預設目標種類都為 "N/A"
        var result = new Dictionary<string, string>
        {
            { "MA", "N/A" },
            { "MB", "N/A" },
            { "MC", "N/A" }
        };

        foreach (var r in ws.RowsUsed())
        {
            int i, j, k;
            if (ws.Name == "濾網")
            {
                i = 21; j = 31; k = 32;
            }
            else
            {
                i = 5; j = 10; k = 15;
            }
            string existingLot = r.Cell(i).GetString().Trim();
            existingLot = existingLot.Length >= 14 ? existingLot.Substring(0, 14) : existingLot;
            string testItem = r.Cell(j).GetString().Trim();
            string efficiencyStr = r.Cell(k).GetString().Trim();

            if (existingLot.Equals(carbonLot.Trim(), StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(efficiencyStr, out double efficiency))
            {
                string type = null;
                if ((testItem == "SO2" || testItem == "H2S") && targetTypes.Contains("MA")) type = "MA";
                else if (testItem == "NH3" && targetTypes.Contains("MB")) type = "MB";
                else if ((testItem == "Toluene" || testItem == "Acetone" || testItem == "IPA") && targetTypes.Contains("MC")) type = "MC";

                if (type != null)
                {
                    if (result[type] == "N/A")
                        result[type] = efficiency.ToString("0.###");
                    else if (double.TryParse(result[type], out double existVal) && efficiency < existVal)
                        result[type] = efficiency.ToString("0.###");
                }
            }
        }
        return result;
    }
}
