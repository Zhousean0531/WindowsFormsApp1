using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public static class Page2MasterExporter
{
    public static void Export(Page2ExportData d)
    {
        string filePath = Path.Combine(
            Application.StartupPath,
            "總表.xlsx"
        );

        if (!File.Exists(filePath))
        {
            MessageBox.Show("找不到 總表.xlsx");
            return;
        }

        if (d == null)
            return;

        // 防護 null：使用 null-coalescing 取得安全的長度
        int n = Math.Min(d.TestWeights?.Count ?? 0, d.PressureDrops?.Count ?? 0);
        if (n == 0)
        {
            MessageBox.Show("沒有可用的測試資料可供匯出");
            return;
        }

        if (d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有效率群組資料可供匯出");
            return;
        }

        using (var wb = new XLWorkbook(filePath))
        {
            var ws = wb.Worksheet("濾網");
            if (ws == null)
            {
                MessageBox.Show("總表中找不到「濾網」工作表");
                return;
            }

            var used = ws.Column(20)
                         .Cells(c => c.Address.RowNumber >= 5 && !c.IsEmpty());

            int rowIdx = (used.LastOrDefault()?.Address.RowNumber ?? 4) + 1;

            foreach (var g in d.EfficiencyGroups)
            {
                if (g == null) continue; // 防護空物件
                for (int i = 0; i < n; i++)
                {
                    ws.Cell(rowIdx, 20).Value = d.ProductionDate; //T
                    ws.Cell(rowIdx, 21).Value = d.TestDate; //U
                    ws.Cell(rowIdx, 22).Value = d.OrderDisplay; //V
                    ws.Cell(rowIdx, 23).Value = d.ProductType; //W  
                    ws.Cell(rowIdx, 24).Style.Alignment.WrapText = true;
                    ws.Cell(rowIdx, 24).Value = g.TypeMaterialDisplay ?? ""; //X
                    ws.Cell(rowIdx, 25).Value = d.Gsm;//Y
                    ws.Cell(rowIdx, 26).Value = d.Gile;//Z
                    ws.Cell(rowIdx, 27).Value = d.Speed;//AA
                    ws.Cell(rowIdx, 28).Value = d.Upper;//AB
                    ws.Cell(rowIdx, 29).Value = d.Lower;//AC
                    ws.Cell(rowIdx, 30).Value = d.Pressure;//AD
                    ws.Cell(rowIdx, 31).Value = d.Wind;//AE
                    ws.Cell(rowIdx, 32).Value = d.TestWeights[i];//AF
                    ws.Cell(rowIdx, 33).Value = d.PressureDrops[i];//AG
                    ws.Cell(rowIdx, 38).Value = d.CarbonInfo;//AL
                    ws.Cell(rowIdx, 39).Value = d.UserName;//AM

                    if (i == d.SelectedIndex)
                    {
                        ws.Cell(rowIdx, 34).Value = g.GasName;//AH
                        ws.Cell(rowIdx, 35).Value = g.Concentration;//AI
                        ws.Cell(rowIdx, 36).Value = g.Eff0;//AJ
                        ws.Cell(rowIdx, 37).Value = g.Eff10;//AK
                    }
                    else
                    {
                        ws.Cell(rowIdx, 34).Value = "";
                        ws.Cell(rowIdx, 35).Value = "";
                        ws.Cell(rowIdx, 36).Value = "";
                    }
                    rowIdx++;
                }
            }
            wb.Save();
            MasterTableHelper.CopyToOneDrive(filePath);
        }
    }
}
