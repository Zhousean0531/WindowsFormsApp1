using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static SignatureHelper;
using Excel = Microsoft.Office.Interop.Excel;

public static class Page2ReportExporter
{
    // =====================================================
    // 對外唯一入口（跟 Page1 一樣）
    // =====================================================
    public static void Export(Page2ExportData d)
    {
        if (d == null) return;

        // ───── 選擇 Report 儲存路徑 ─────
        string reportSavePath;
        using (var sfd = new SaveFileDialog())
        {
            DateTime testDt = DateTime.Parse(d.TestDate);

            sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
            sfd.FileName =
                $"{d.ReportNo}{d.ProductName}_{d.ProductType}_{d.Gsm}_{d.OrderDisplay}({testDt:MMdd}生產).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            reportSavePath = sfd.FileName;
        }

        // ───── 選擇 Helper 儲存路徑 ─────
        string helperSavePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Macro (*.xlsm)|*.xlsm";
            sfd.FileName = "Helper.xlsm";
            sfd.Title = "請選擇 Helper 檔案存放位置";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            helperSavePath = sfd.FileName;
        }

        // ───── 匯出報告檔 ─────
        Export_Report(reportSavePath, d);

        // ───── 匯出 Helper ─────
        Page2HelperExporter.Export(helperSavePath, d);
    }
    private static void Export_Report(string savePath, Page2ExportData d)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_SemiFinished_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_SemiFinished_Template.xlsx");
            return;
        }

        foreach (var g in d.EfficiencyGroups)
        {
            if (g.Efficiencies11 == null || g.Efficiencies11.Count == 0)
                continue;

            File.Copy(templatePath, savePath, true);

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

                wb = app.Workbooks.Open(savePath);
                ws = (Excel.Worksheet)wb.Sheets["濾網半成品報告"];

                int idx = d.SelectedIndex;

                DateTime testDt = DateTime.Parse(d.TestDate);
                DateTime prodDt = DateTime.Parse(d.ProductionDate);

                string l1Text =
                    $"{testDt:MM.dd} {d.ProductType} {d.TestWeights[idx]}gsm " +
                    $"({d.PressureDrops[idx]}Pa) -{prodDt:MMdd}生產";

                // ───── 基本資料 ─────
                ws.Range["C6"].Value = d.ReportNo;
                ws.Range["F7"].Value = d.ProductDisplay;
                ws.Range["F8"].Value = d.FilterSize;
                ws.Range["F9"].Value = d.OrderDisplay;
                ws.Range["H6"].Value = d.TestDate;
                ws.Range["F11"].Value = g.GasName;

                ws.Range["H10"].Value = d.TestWeights[idx];
                ws.Range["F12"].Value = g.Concentration;
                ws.Range["F13"].Value = d.Wind;
                ws.Range["F14"].Value = d.PressureDrops[idx];
                ws.Range["F16"].Value = g.Eff0;

                // ───── 11 點效率 ─────
                ws.Range["L1"].Value = l1Text;

                int startRow = 2;
                for (int i = 0; i < g.Efficiencies11.Count; i++)
                {
                    ws.Cells[startRow + i, 13].Value = g.Efficiencies11[i];
                }

                System.Threading.Thread.Sleep(100);
                ExcelSignatureHelper.TryAddSignature(ws, "H28");

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
    public static class Page2HelperExporter
    {
        public static void Export(string helperSavePath, Page2ExportData d)
        {
            if (d == null || d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
                return;

            // ───── Helper Template 一定來自程式路徑 ─────
            string templatePath = Path.Combine(
                Application.StartupPath,
                "Helper_Template.xlsm"
            );

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("找不到 Helper_Template.xlsm");
                return;
            }
                File.Copy(templatePath, helperSavePath, true);
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
                ws = (Excel.Worksheet)wb.Worksheets["濾網半成品工作表"];

                // ───── 找起始列（A 欄）─────
                int startRow =
                    ws.Cells[ws.Rows.Count, 1]
                      .End(Excel.XlDirection.xlUp)
                      .Row + 1;

                int idx = d.SelectedIndex;
                int n = new[]
                    {
                        d.PressureDrops?.Count ?? 0,
                        d.TestWeights?.Count ?? 0
                    }.Min();
                if (n <= 0)
                    return;
                var eff = d.EfficiencyGroups.First();
                DateTime testDt = DateTime.Parse(d.TestDate);
                DateTime prodDt = DateTime.Parse(d.ProductionDate);

                double weight = d.TestWeights[idx];
                double dp = d.PressureDrops[idx];
                // ───── 逐筆壓損展開成多列 ─────
                for (int i = 0; i < n; i++)
                {
                    int row = startRow + i;

                    // 每一列都有的欄位（對應你的截圖）
                    ws.Cells[row, 1].Value = d.ProductionDate; // A
                    ws.Cells[row, 3].Value = d.OrderDisplay;   // C
                    ws.Cells[row, 4].Value = d.ProductType;    // D
                    ws.Cells[row, 5].Value = eff.TypeMaterialDisplay; // E
                    ws.Cells[row, 6].Value = d.Gsm;             // F
                    ws.Cells[row, 7].Value = d.Gile;            // G
                    ws.Cells[row, 8].Value = d.Speed;           // H
                    ws.Cells[row, 9].Value = d.Pressure;        // I
                    ws.Cells[row, 10].Value = d.Wind;            // J
                    ws.Cells[row, 11].Value = d.TestWeights[i];  // K
                    ws.Cells[row, 12].Value = d.PressureDrops[i];// L
                    ws.Cells[row, 18].Value = d.CarbonInfo;      // R
                    ws.Cells[1, 21].value = $"{testDt:MM.dd} {d.ProductType} {weight}gsm ({dp}Pa)-{prodDt:MMdd}生產";
                    // ★ 只有「選中的壓損列」才寫的欄位
                    if (i == idx)
                    {
                        ws.Cells[row, 2].Value = d.TestDate;        // B
                        ws.Cells[row, 13].Value = eff.GasName;      // M
                        ws.Cells[row, 14].Value = eff.Concentration;// N
                        ws.Cells[row, 15].Value = eff.Eff0;         // O
                        ws.Cells[row, 16].Value = eff.Eff10;        //P
                        
                    }
                }

                if (eff.Efficiencies11 != null && eff.Efficiencies11.Count > 0)
                {
                    int startRowEff = 3;
                    int colEff = 21; // U
                    int count = Math.Min(11, eff.Efficiencies11.Count);

                    for (int i = 0; i < count; i++)
                    {
                        ws.Cells[startRowEff + i, colEff].Value = eff.Efficiencies11[i];
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
}
