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
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "總表.xlsx"
        );

        if (!File.Exists(filePath))
        {
            MessageBox.Show("找不到 總表.xlsx");
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

            // 以第 5 列開始，找最後一筆
            var used = ws.Column(14)
                         .Cells(c => c.Address.RowNumber >= 5 && !c.IsEmpty());

            int rowIdx = (used.LastOrDefault()?.Address.RowNumber ?? 4) + 1;
            // ⭐ 必須先檢查
            if (d == null)
            {
                return;
            }
            int n = Math.Min(
                d.TestWeights.Count,
                d.PressureDrops.Count
            );
            foreach (var g in d.EfficiencyGroups)
            {
                for (int i = 0; i < n; i++)
                {
                    // ───── 共通欄位（每一列都要寫） ─────
                    ws.Cell(rowIdx, 19).Value = d.ProductionDate;
                    ws.Cell(rowIdx, 20).Value = d.TestDate;
                    ws.Cell(rowIdx, 21).Value = d.CarbonOrder;
                    ws.Cell(rowIdx, 22).Value = d.Gsm;
                    ws.Cell(rowIdx, 23).Value = d.Gile;
                    ws.Cell(rowIdx, 24).Value = d.Speed;
                    ws.Cell(rowIdx, 25).Value = d.Upper;
                    ws.Cell(rowIdx, 26).Value = d.Lower;
                    ws.Cell(rowIdx, 27).Value = d.Pressure;
                    ws.Cell(rowIdx, 28).Value = d.Wind;

                    ws.Cell(rowIdx, 29).Value = d.TestWeights[i];
                    ws.Cell(rowIdx, 30).Value = d.PressureDrops[i];
                    ws.Cell(rowIdx, 34).Value = d.CarbonInfo;

                    // ───── 只有選定壓損那一列才寫效率 ─────
                    if (i == d.SelectedIndex)
                    {
                        ws.Cell(rowIdx, 31).Value = g.GasName;
                        ws.Cell(rowIdx, 32).Value = g.Eff0;
                        ws.Cell(rowIdx, 33).Value = g.Eff10;
                    }
                    else
                    {
                        ws.Cell(rowIdx, 31).Value = "";
                        ws.Cell(rowIdx, 32).Value = "";
                        ws.Cell(rowIdx, 33).Value = "";
                    }

                    rowIdx++; // ⭐ rowIdx 只在這裡 +1
                }
            }
            wb.Save();
        }
    }
}
