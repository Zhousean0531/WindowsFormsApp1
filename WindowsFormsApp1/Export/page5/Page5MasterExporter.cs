using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public static class Page5MasterExporter
{
    public static void Export(Page5ExportData d)
    {
        if (d == null || d.Rows == null || d.Rows.Count == 0)
            return;

        // ===== 跟 Page1 一樣的總表來源（桌面）=====
        string sourcePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "總表.xlsx"
        );

        var wb = new XLWorkbook(sourcePath);
        var ws = wb.Worksheet("濾筒");

        // 調整：欄位索引向右移 +2（與 Page5LookupHelper 使用欄位對齊）
        int row = (ws.Column(38).CellsUsed()
                     .LastOrDefault()?.Address.RowNumber ?? 3) + 1;

        var targetTypes = (d.FilterType ?? "")
            .Split('+')
            .Select(s => s.Trim())
            .ToList();

        foreach (var r in d.Rows)
        {
            ws.Cell(row, 21).Value = d.TestDate;
            ws.Cell(row, 22).Value = d.ReportNo;
            ws.Cell(row, 23).Value = d.CylinderNo;
            ws.Cell(row, 24).Value = d.CylinderNo;
            ws.Cell(row, 25).Value = d.Customer;
            ws.Cell(row, 26).Value = "CYL";
            ws.Cell(row, 27).Value = d.FilterType;
            ws.Cell(row, 28).Value = d.ReCylinderNo;
            ws.Cell(row, 29).Value = r.SN;
            ws.Cell(row, 30).Value = r.Weight;
            ws.Cell(row, 57).Value = d.UserName;

            if (r.ControlValues != null)
            {
                foreach (var kv in r.ControlValues)
                {
                    if (kv.Key > 0)
                        ws.Cell(row, kv.Key).Value = kv.Value;
                }
            }

            var eff = EfficiencyFinder
                .FindMinEfficiencyByCarbonLot(ws, d.CarbonLot, targetTypes);

            if (eff.ContainsKey("MA")) ws.Cell(row, 34).Value = eff["MA"];
            if (eff.ContainsKey("MB")) ws.Cell(row, 35).Value = eff["MB"];
            if (eff.ContainsKey("MC")) ws.Cell(row, 36).Value = eff["MC"];

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
        MessageBox.Show("匯入完成！");
    }
}
