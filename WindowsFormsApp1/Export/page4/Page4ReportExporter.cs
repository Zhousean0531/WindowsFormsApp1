using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static SignatureHelper;

public static class Page4ReportExporter
{
    public static void Export(Page4ExportData d)
    {
        if (d == null || d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有可匯出的效率資料");
            return;
        }

        // =====================================================
        // (A) 先選擇 Helper 存檔路徑（只問一次）
        // =====================================================
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

        // =====================================================
        // (B) 報告範本檢查
        // =====================================================
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_RawMaterial_Template.xlsx");
            return;
        }

        // =====================================================
        // (C) 多氣體 → 多份報告
        // =====================================================
        foreach (var g in d.EfficiencyGroups)
        {
            if (g.Efficiencies11 == null || g.Efficiencies11.Count == 0)
                continue;

            DateTime arrivalDt = DateTime.Parse(d.ArrivalDate);
            string reportSavePath;

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                sfd.FileName =
                    $"{d.ReportNo}_{d.Material}({arrivalDt:MMdd}到廠)_{g.GasName}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    continue;

                reportSavePath = sfd.FileName;
            }

            // ───── 複製範本 ─────
            File.Copy(templatePath, reportSavePath, true);

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

                wb = app.Workbooks.Open(reportSavePath);
                ws = (Excel.Worksheet)wb.Worksheets["濾網原料報告"];

                int idx = d.SelectedIndex;

                // ───── 表頭 ─────
                ws.Range["C4"].Value = d.ReportNo;
                ws.Range["C5"].Value = d.ArrivalDate;
                ws.Range["C6"].Value = d.TestingDate;
                ws.Range["E4"].Value = d.MaterialNo;
                ws.Range["E5"].Value = d.Material;
                ws.Range["E6"].Value = d.QtyText;

                // ───── 明細 ─────
                const int COL_FIRST = 3;
                const int ROW_LOT = 10;
                const int ROW_DEN = 13;
                const int ROW_DP = 14;
                const int ROW_VIN = 15;
                const int ROW_VOUT = 16;
                const int ROW_OUTG = 17;
                const int ROW_EFF = 18;

                for (int i = 0; i < d.LotFulls.Count; i++)
                {
                    int col = COL_FIRST + i;

                    ws.Cells[ROW_LOT, col].Value = d.LotFulls[i];
                    ws.Cells[ROW_DEN, col].Value = d.Densities[i];
                    ws.Cells[ROW_DP, col].Value = d.DeltaPs[i];
                    ws.Cells[ROW_VIN, col].Value = d.VocIns[i];
                    ws.Cells[ROW_VOUT, col].Value = d.VocOuts[i];
                    ws.Cells[ROW_OUTG, col].Value = d.OutgassingList[i];

                    ws.Cells[ROW_EFF, col].Value =
                        (i == idx) ? g.Eff0 : "N.D.";
                }

                // ───── 簽名 ─────
                ExcelSignatureHelper.TryAddSignature(ws, "E25");

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

        // =====================================================
        // (D) 最後再匯出一次 Helper（一次）
        // =====================================================
        Page4HelperExporter.Export(helperSavePath, d);

        // =====================================================
        // (E) 南京特規（如需要）
        // =====================================================
        Page4ReportExporterForNanJing.Export(d);
    }
}
