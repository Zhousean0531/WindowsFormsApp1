using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public static class Page5MasterExporter
{
    public static void Export(Page5ExportData d, string rawEfficiencyText)
    {
        if (d == null || d.Rows == null || d.Rows.Count == 0)
            return;
        string sourcePath = Path.Combine(
            Application.StartupPath,
            "總表.xlsx"
        );

        var wb = new XLWorkbook(sourcePath);
        var ws = wb.Worksheet("濾筒");

        // 調整：欄位索引向右移 +2（與 Page5LookupHelper 使用欄位對齊）
        int row = (ws.Column(38).CellsUsed()
                     .LastOrDefault()?.Address.RowNumber ?? 3) + 1;
        int? effCol = ResolveEfficiencyColumn(d.FilterType);

        foreach (var r in d.Rows)
        {
            ws.Cell(row, 21).Value = d.TestDate;
            ws.Cell(row, 22).Value = d.ReportNo;
            ws.Cell(row, 23).Value = d.CylinderNo;
            ws.Cell(row, 24).Value = d.CarbonLot;
            ws.Cell(row, 25).Value = d.Customer;
            ws.Cell(row, 26).Value = "CYL";
            ws.Cell(row, 27).Value = d.FilterType;
            ws.Cell(row, 28).Value = d.ReCylinderNo;
            ws.Cell(row, 29).Value = r.SN;
            ws.Cell(row, 30).Value = r.Weight;
            ws.Cell(row, 35).Value = "N/A";
            ws.Cell(row, 36).Value = "N/A";
            ws.Cell(row, 37).Value = "N/A";
            ws.Cell(row, 53).Value = "130";
            ws.Cell(row, 55).Value = d.UserName;
            ws.Cell(row, 56).Value = DateTime.Now.ToString("yyyyMMddHHmmss");
            if (r.ControlValues != null)
            {
                foreach (var kv in r.ControlValues)
                {
                    if (kv.Key > 0)
                        ws.Cell(row, kv.Key).Value = kv.Value;
                }
            }
            if (effCol.HasValue && !string.IsNullOrWhiteSpace(rawEfficiencyText))
            {
                ws.Cell(row, effCol.Value).Value = rawEfficiencyText;
            }
            // MaterialInfo
            var materialInfo = MaterialMasterHelper.Get(r.SN);
            if (materialInfo != null)
            {
                ws.Cell(row, 31).Value = materialInfo.MaterialNo;
                ws.Cell(row, 32).Value = materialInfo.MaterialName;
                ws.Cell(row, 33).Value = materialInfo.InUnit;
                ws.Cell(row, 34).Value = materialInfo.SampleQty;
                ws.Cell(row, 35).Value = materialInfo.InspectUnit;
                ws.Cell(row, 36).Value = materialInfo.Spec;
            }
            row++;
        }
        wb.Save();
        wb.Dispose();
    }
    private static int? ResolveEfficiencyColumn(string filterType)
    {
        if (string.IsNullOrWhiteSpace(filterType))
            return null;

        filterType = filterType.Trim().ToUpper();

        if (filterType.Contains("MA")) return 35;
        if (filterType.Contains("MB")) return 36;
        if (filterType.Contains("MC")) return 37;

        return null;
    }

}
