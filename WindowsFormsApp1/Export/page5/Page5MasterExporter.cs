using ClosedXML.Excel;
using System.Linq;
using System.Windows.Forms;

public static class Page5MasterExporter
{
    public static void Export(Page5ExportData d)
    {
        var wb = new XLWorkbook(@"C:\Users\User\Desktop\總表.xlsx");
        var ws = wb.Worksheet("濾筒");

        int row = (ws.Column(38).CellsUsed()
                     .LastOrDefault()?.Address.RowNumber ?? 3) + 1;

        var targetTypes = d.FilterType.Split('+')
                                      .Select(s => s.Trim())
                                      .ToList();

        foreach (var r in d.Rows)
        {
            ws.Cell(row, 19).Value = d.TestDate;
            ws.Cell(row, 20).Value = d.ReportNo;
            ws.Cell(row, 21).Value = d.CylinderNo;
            ws.Cell(row, 22).Value = d.Customer;
            ws.Cell(row, 23).Value = "CYL";
            ws.Cell(row, 24).Value = d.FilterType;
            ws.Cell(row, 25).Value = d.ReCylinderNo;

            ws.Cell(row, 26).Value = r.SN;
            ws.Cell(row, 27).Value = r.Weight;

            foreach (var kv in r.ControlValues)
            {
                ws.Cell(row, kv.Key).Value = kv.Value;
            }

            var eff = EfficiencyFinder
                .FindEfficiencyByTypeWithMin(ws, d.CarbonLot, targetTypes);

            if (eff.ContainsKey("MA")) ws.Cell(row, 32).Value = eff["MA"];
            if (eff.ContainsKey("MB")) ws.Cell(row, 33).Value = eff["MB"];
            if (eff.ContainsKey("MC")) ws.Cell(row, 34).Value = eff["MC"];

            row++;
        }

        wb.Save();
        MessageBox.Show("TabPage5 匯入完成！");
    }
}
