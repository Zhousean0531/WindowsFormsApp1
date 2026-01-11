using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public static class Page5MasterExporter
{
    public static void Export(Page5ExportData d)
    {
        var wb = new XLWorkbook(@"C:\Users\User\Desktop\總表.xlsx");
        var ws = wb.Worksheet("濾筒");

        var targetTypes = d.FilterType.Split('+')
                                      .Select(s => s.Trim())
                                      .ToList();
        Dictionary<string, string> eff = null;

        if (!string.IsNullOrWhiteSpace(d.CarbonLot))
        {
            eff = EfficiencyFinder
                .FindMinEfficiencyByCarbonLot(ws, d.CarbonLot, targetTypes);
        }

        int row = (ws.Column(38).CellsUsed()
                     .LastOrDefault()?.Address.RowNumber ?? 3) + 1;

        foreach (var r in d.Rows)
        {
            MessageBox.Show(
    $"即將寫入 row={row}\nCylinderNo={d.CylinderNo}"
);
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
            ws.Cell(row, 31).Value = "V";
            ws.Cell(row, 32).Value = "V";
            ws.Cell(row, 33).Value = "V";
            ws.Cell(row, 34).Value = "V";
            ws.Cell(row, 53).Value = "130";
            ws.Cell(row, 55).Value = d.UserName;

            foreach (var kv in r.ControlValues)
            {
                ws.Cell(row, kv.Key).Value = kv.Value;
            }
            if (eff != null)
            {
                ws.Cell(row, 35).Value = eff["MA"];
                ws.Cell(row, 36).Value = eff["MB"];
                ws.Cell(row, 37).Value = eff["MC"];
            }
            row++;
        }

        wb.Save();
        MessageBox.Show("匯入完成！");
    }
}
