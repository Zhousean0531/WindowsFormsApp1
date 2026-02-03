using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

public static class Page4HelperExporter
{
    public static void Export(string helperSavePath, Page4ExportData d)
    {
        if (d == null) return;

        // =====================================================
        // (A) 確保 Helper 檔存在（不存在就從範本複製）
        // =====================================================
        string templatePath = Path.Combine(
            Application.StartupPath,
            "Helper_Template.xlsm"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Helper_Template.xlsm");
            return;
        }

        if (!File.Exists(helperSavePath))
        {
            File.Copy(templatePath, helperSavePath, true);
        }

        Excel.Application app = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        try
        {
            app = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            wb = app.Workbooks.Open(helperSavePath);
            ws = (Excel.Worksheet)wb.Worksheets["濾筒工作表"];
            // ↑ 工作表名稱請確認與 Helper_Template 一致

            // =====================================================
            // (B) 找下一列（以第 1 欄為基準）
            // =====================================================
            int startRow = 2; // 從第 2 列開始寫入
            int n = new[]
            {
                d.LotNos.Count,
                d.Weights.Count,
                d.Densities.Count,
                d.VocIns.Count,
                d.VocOuts.Count,
                d.OutgassingList.Count,
                d.DeltaPs.Count
            }.Min();

            for (int i = 0; i < n; i++)
            {
                int row = startRow + i;

                ws.Cells[row, 1].Value = d.TestingDate;
                ws.Cells[row, 2].Value = d.ArrivalDate;
                ws.Cells[row, 3].Value = d.Material;
                ws.Cells[row, 4].Value = d.LotNos[i];
                ws.Cells[row, 5].Value = d.LotFulls[i];
                ws.Cells[row, 8].Value = d.Weights[i];
                ws.Cells[row, 9].Value = d.Densities[i];
                ws.Cells[row, 11].Value = d.VocIns[i];
                ws.Cells[row, 12].Value = d.VocOuts[i];
                ws.Cells[row, 13].Value = d.OutgassingList[i];
                ws.Cells[row, 14].Value = d.DeltaPs[i];
                System.Threading.Thread.Sleep(90);
                Application.DoEvents();
            }

            // =====================================================
            // (D) Mesh（結構化資料）
            // 每一個粒徑各佔一列（跟 Page1 一樣的概念）
            // =====================================================
            if (d.ParticleSizePercentages != null &&
                d.ParticleSizePercentages.Count > 0)
            {
                int meshRow = 8;

                foreach (var kv in d.ParticleSizePercentages)
                {
                    ws.Cells[meshRow, 1].Value = kv.Key;          // 粒徑
                    ws.Cells[meshRow, 2].Value = kv.Value/100; // 百分比
                    ws.Cells[meshRow, 2].NumberFormat = "0.0%";
                    meshRow++;
                }
            }
            // (E) Efficiency（含標題）
            if (d.EfficiencyGroups != null && d.EfficiencyGroups.Count > 0)
            {
                int defaultCol = 21; // U

                // ★ 這裡假設「樣品名稱、壓損」取 SelectedIndex 那一筆
                int idx = d.SelectedIndex;

                for (int gIdx = 0; gIdx < d.EfficiencyGroups.Count; gIdx++)
                {
                    var g = d.EfficiencyGroups[gIdx];

                    if (g.Efficiencies11 == null || g.Efficiencies11.Count == 0)
                        continue;

                    // ───── 決定欄位 ─────
                    int col;
                    if (string.Equals(g.GasName, "Acetone", StringComparison.OrdinalIgnoreCase))
                        col = 22; // V
                    else if (string.Equals(g.GasName, "IPA", StringComparison.OrdinalIgnoreCase))
                        col = 23; // W
                    else
                        col = defaultCol; // U
                    string testDateMMDD =
                        DateTime.TryParse(d.TestingDate, out var td)
                            ? td.ToString("MM.dd")
                            : d.TestingDate;

                    string arrivalDateMMDD =
                        DateTime.TryParse(d.ArrivalDate, out var ad)
                            ? ad.ToString("MM.dd")
                            : d.ArrivalDate;

                    // 取樣品名稱右邊兩個字（防呆）
                    string lotName = d.LotFulls[idx] ?? "";
                    string lotShort =
                        lotName.Length >= 2
                            ? lotName.Substring(lotName.Length - 2)
                            : lotName;

                    // ───── 組標題字串（第一列）─────
                    string headerText =
                        $"{testDateMMDD} " + $"{d.Material}" +$"#{lotShort}" +$"({d.DeltaPs[idx]}Pa)" +$"-{arrivalDateMMDD}Arrival_" +$"{g.GasName}";

                    ws.Cells[startRow-1, col].Value = headerText;

                    // ───── 寫入所有效率值（從下一列開始）─────
                    for (int i = 0; i < g.Efficiencies11.Count; i++)
                    {
                        ws.Cells[startRow + 1 + i, col].Value = g.Efficiencies11[i];
                    }
                }
            }
            wb.Save();
        }
        finally
        {
            wb?.Close(false);
            app?.Quit();

            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null) Marshal.ReleaseComObject(wb);
            if (app != null) Marshal.ReleaseComObject(app);
        }
    }
}
