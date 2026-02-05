using ClosedXML.Excel;
using System;
using System.Linq;
using System.Windows.Forms;
using System.IO;

public static class Page4MasterExporter
{
    public static void Export(Page4ExportData d)
    {
        string filePath = Path.Combine(
            Application.StartupPath,
            "總表.xlsx"
        );
        using (var wb = new XLWorkbook(filePath))
        {
            var ws = wb.Worksheet("濾筒");
            int row = (ws.Column(2).CellsUsed().LastOrDefault()?.Address.RowNumber ?? 4) + 1;
            foreach (var eff in d.EfficiencyGroups)
            {
                for (int i = 0; i < d.LotFulls.Count; i++)
                {
                    ws.Cell(row, 1).Value = d.TestingDate;
                    ws.Cell(row, 2).Value = d.ArrivalDate;
                    ws.Cell(row, 3).Value = d.Material;
                    ws.Cell(row, 4).Value = d.LotNos[i];
                    ws.Cell(row, 5).Value = d.LotFulls[i];
                    ws.Cell(row, 6).Value = d.QtyText;
                    ws.Cell(row, 8).Value = d.Weights[i];
                    ws.Cell(row, 9).Value = d.Densities[i];
                    ws.Cell(row, 11).Value = eff.GasName;
                    ws.Cell(row, 12).Value = d.VocIns[i];
                    ws.Cell(row, 13).Value = d.VocOuts[i];
                    ws.Cell(row, 14).Value = d.OutgassingList[i];
                    ws.Cell(row, 15).Value = d.DeltaPs[i];
                    ws.Cell(row, 19).Value = d.MeshSummary;
                    ws.Cell(row, 20).Value = d.UserName;
                    if (i == d.SelectedIndex)
                    {
                        ws.Cell(row, 16).Value = eff.Eff0;
                        ws.Cell(row, 17).Value = eff.Eff10;
                        ws.Cell(row, 18).Value = eff.Concentration;
                    }
                    row++;
                }
            }
            wb.Save();
            MasterTableHelper.CopyToOneDrive(filePath);
        }
    }

}
