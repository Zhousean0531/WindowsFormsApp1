using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
public static class Page3MasterExporter
{
    public static void Export(Page3ExportData d)
    {
        string filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "總表.xlsx"
        );

        using (var wb = new XLWorkbook(filePath))
        {
            var ws = wb.Worksheet("濾網");
            if (ws == null)
            {
                MessageBox.Show("找不到「濾網」工作表");
                return;
            }

            int row = (ws.Column(2).CellsUsed()
                        .LastOrDefault()?.Address.RowNumber ?? 4) + 1;

            for (int i = 0; i < d.LotNos.Count; i++)
            {
                ws.Cell(row, 1).Value = d.TestingDate;
                ws.Cell(row, 2).Value = d.ArrivalDate;
                ws.Cell(row, 3).Value = d.Material;
                ws.Cell(row, 4).Value = d.LotNos[i];

                ws.Cell(row, 8).Value = d.Weights[i];
                ws.Cell(row, 9).Value = d.DeltaPs[i];
                ws.Cell(row, 10).Value = d.VocIns[i];
                ws.Cell(row, 11).Value = d.VocOuts[i];
                ws.Cell(row, 12).Value = d.OutMinusIn[i];

                row++;
            }

            wb.Save();
        }
    }
}