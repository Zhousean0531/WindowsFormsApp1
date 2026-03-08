using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page1;
using static SignatureHelper;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelReportUtil
{
    // =====================================================
    // 對外唯一入口
    // =====================================================
    public static void ExportPage1(
        string reportSavePath,
        string helperSavePath,
        P1Batch batch)
    {
        if (batch == null) return;

        Export_QC_RawMaterial_Report(reportSavePath, batch);
        Export_Helper(helperSavePath, batch);
    }

    // =====================================================
    // (A) 匯出 QC_RawMaterial_Template.xlsx（報告用）
    // =====================================================
    private static void Export_QC_RawMaterial_Report(
        string savePath,
        P1Batch batch)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_RawMaterial_Template.xlsx");
            return;
        }

        File.Copy(templatePath, savePath, true);

        Excel.Application app = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        try
        {
            app = new Excel.Application
            {
                DisplayAlerts = false,
                Visible = false
            };

            wb = app.Workbooks.Open(savePath);
            ws = (Excel.Worksheet)wb.Sheets[1];

            // ===== 表頭 =====
            ws.Range["C4"].Value = batch.ReportNo;
            ws.Range["E4"].Value = batch.MaterialNo;
            ws.Range["E5"].Value = batch.Material;
            ws.Range["C5"].Value = batch.ArrivalDate;
            ws.Range["C6"].Value = batch.TestingDate;
            ws.Range["E6"].Value = batch.QtyText;
            ws.Range["C13:E13"].Value2 = "N/A";

            // ===== 明細 =====
            const int COL_FIRST = 3; // C
            const int ROW_LOTNO = 10;
            const int ROW_DENS = 13;
            const int ROW_DP = 14;
            const int ROW_VIN = 15;
            const int ROW_VOUT = 16;
            const int ROW_OUTG = 17;
            const int ROW_EFF = 18;

            int max = batch.Samples.Count;

            for (int i = 0; i < max; i++)
            {
                int col = COL_FIRST + i;
                var sample = batch.Samples[i];
                var selectedSample = batch.Samples.FirstOrDefault(s => s.IsSelected);
                double? eff0 = selectedSample?.Efficiencies?.FirstOrDefault();
                ws.Cells[ROW_LOTNO, col].Value = sample.LotFull;
                ws.Cells[ROW_DENS, col].Value = sample.Density;
                ws.Cells[ROW_DP, col].Value = sample.DeltaP;
                ws.Cells[ROW_VIN, col].Value = sample.VocIn;
                ws.Cells[ROW_VOUT, col].Value = sample.VocOut;
                ws.Cells[ROW_OUTG, col].Value = sample.Outgassing;

                ws.Cells[ROW_EFF, col].Value =
                    sample.IsSelected
                    ? (eff0.HasValue ? (object)eff0.Value : "N.D.")
                    : "N.D.";

                System.Threading.Thread.Sleep(20);
                Application.DoEvents();
            }

            // ===== 簽名 =====
            ExcelSignatureHelper.TryAddSignature(ws, "E25");

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

    // =====================================================
    // (B) 匯出 Helper_Template.xlsm（彙整用）
    // =====================================================
    private static void Export_Helper(
        string helperSavePath,
        P1Batch batch)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "Helper_Template.xlsm");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Helper_Template.xlsm");
            return;
        }
        var selectedSample = batch.Samples.FirstOrDefault(s => s.IsSelected);
        var eff0 = selectedSample?.Efficiencies?.FirstOrDefault();
        var eff10 = selectedSample?.Efficiencies?.ElementAtOrDefault(10);
        File.Copy(templatePath, helperSavePath, true);
        
        HelperWorkbookInterop.Append(
            helperSavePath,
            "濾網工作表",
            (wshelper, startRow) =>
            {
                int n = batch.Samples.Count;

                // ===== 每一筆資料 → 一列 =====
                for (int i = 0; i < n; i++)
                {
                    int row = startRow + i;
                    var sample = batch.Samples[i];
                    wshelper.Cells[row, 1].Value = batch.TestingDate;
                    wshelper.Cells[row, 2].Value = batch.Material;
                    wshelper.Cells[row, 3].Value = "";
                    wshelper.Cells[row, 4].Value = sample.LotFull;
                    wshelper.Cells[row, 5].Value = sample.VocIn;
                    wshelper.Cells[row, 6].Value = sample.VocOut;
                    wshelper.Cells[row, 7].Value = sample.Outgassing;
                    wshelper.Cells[row, 8].Value = sample.DeltaP;

                    if (sample.IsSelected)
                    {
                        wshelper.Cells[row, 9].Value =
                            eff0.HasValue ? (object)eff0.Value : "N.D.";
                        wshelper.Cells[row, 10].Value =
                            eff10.HasValue ? (object)eff10.Value : "N.D.";
                    }
                    else
                    {
                        wshelper.Cells[row, 9].Value = "N.D.";
                        wshelper.Cells[row, 10].Value = "N.D.";
                    }
                }

                // ===== Mesh（粒徑）=====
                if (batch.ParticleSizePercentages != null)
                {
                    int r = 7;
                    foreach (var kv in batch.ParticleSizePercentages)
                    {
                        if (kv.Key.Contains("總重"))
                            continue;

                        wshelper.Cells[r, 1].Value = kv.Key;
                        wshelper.Cells[r, 2].Value = kv.Value / 100;
                        wshelper.Cells[r, 2].NumberFormat = "0.0%";
                        r++;
                    }
                }
                // ===== 效率 11 點 =====
                    int startEffRow = 3;
                    int col = 19;

                    var effList = selectedSample?.Efficiencies;

                    if (effList != null)
                    {
                        for (int i = 0; i < effList.Count && i < 11; i++)
                        {
                            wshelper.Cells[startEffRow + i, col].Value = effList[i];
                        }
                    }
            }
        );
    }
}