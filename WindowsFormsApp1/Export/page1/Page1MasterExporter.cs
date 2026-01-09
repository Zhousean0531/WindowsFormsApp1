using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

public static class Page1MasterExporter
{
    public static void Export(Page1ExportData d)
    {
        string filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "總表.xlsx"
        );

        using (var wb = new XLWorkbook(filePath))
        {
            var ws = wb.Worksheet("濾網");
            int row = (ws.Column(2).CellsUsed().LastOrDefault()?.Address.RowNumber ?? 4) + 1;
            for (int i = 0; i < d.LotFulls.Count; i++)
            {
                ws.Cell(row, 1).Value = d.TestingDate;
                ws.Cell(row, 2).Value = d.ArrivalDate;
                ws.Cell(row, 3).Value = d.Material;
                ws.Cell(row, 4).Value = d.LotFulls[i];
                ws.Cell(row, 5).Value = d.QtyText;
                ws.Cell(row, 8).Value = d.Densities[i];
                ws.Cell(row, 11).Value = d.VocIns[i];
                ws.Cell(row, 12).Value = d.VocOuts[i];
                ws.Cell(row, 14).Value = d.DeltaPs[i];
                ws.Cell(row, 18).Value = d.MeshSummaries[i];
                ws.Cell(row, 19).Value = d.UserName;
                if (i == d.SelectedIndex)
                    ws.Cell(row, 15).Value = d.Eff0;
                row++;
            }

            wb.Save();
        }
    }
}
