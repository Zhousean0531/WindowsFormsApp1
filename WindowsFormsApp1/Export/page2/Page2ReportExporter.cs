using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static SignatureHelper;
using Excel = Microsoft.Office.Interop.Excel;

public static class Page2ReportExporter
{
    public static void Export(Page2ExportData d)
    {
        if (d == null || d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有可匯出的效率資料");
            return;
        }

        if (d.PressureDrops == null || d.TestWeights == null)
        {
            MessageBox.Show("壓損或重量資料為空");
            return;
        }

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_SemiFinished_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_SemiFinished_Template.xlsx");
            return;
        }

        // ─────────────────────────────
        // 逐一氣體輸出
        // ─────────────────────────────
        foreach (var g in d.EfficiencyGroups)
        {
            if (g.Efficiencies11 == null || g.Efficiencies11.Count == 0)
            {
                MessageBox.Show($"氣體 {g.GasName} 的效率資料為空");
                continue;
            }

            string savePath;
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                sfd.FileName = $"{d.ReportNo}_{g.GasName}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    continue;

                savePath = sfd.FileName;
            }

            File.Copy(templatePath, savePath, true);

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(savePath);
                ws = (Excel.Worksheet)wb.Sheets["濾網半成品報告"];

                int idx = d.SelectedIndex;
                if (idx < 0 || idx >= g.Efficiencies11.Count)
                {
                    MessageBox.Show($"氣體 {g.GasName} 的效率資料與壓損索引不符");
                    continue;
                }

                // ─────────────────────────────
                // (A) 基本資料
                // ─────────────────────────────
                ws.Range["C6"].Value = d.ReportNo;
                ws.Range["F7"].Value = d.ProductDisplay;
                ws.Range["F8"].Value = d.FilterSize;
                ws.Range["F9"].Value = d.OrderDisplay;
                ws.Range["H6"].Value = d.TestDate;

                ws.Range["F11"].Value = g.GasName;

                if (idx >= 0 && idx < d.TestWeights.Count)
                    ws.Range["H10"].Value = d.TestWeights[idx];

                ws.Range["F12"].Value = g.Concentration;
                ws.Range["F13"].Value = d.Wind;

                ws.Range["F14"].Value = d.PressureDrops[idx];
                ws.Range["F16"].Value = g.Efficiencies11[idx];

                // ─────────────────────────────
                // (B) 11 點效率明細（從 M2 開始）
                // ─────────────────────────────
                int startRow = 2;
                for (int i = 0; i < g.Efficiencies11.Count; i++)
                {
                    ws.Cells[startRow + i, 13].Value = g.Efficiencies11[i];
                }

                // ─────────────────────────────
                // (C) 簽名
                // ─────────────────────────────
                ExcelSignatureHelper.TryAddSignature(ws, "H28");

                wb.Save();
            }
            finally
            {
                if (wb != null) wb.Close(false);
                if (app != null) app.Quit();

                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }
    }
}
