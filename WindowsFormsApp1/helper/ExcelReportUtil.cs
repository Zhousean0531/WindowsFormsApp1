using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsFormsApp1.Export.Page1;
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
        Page1ExportData d)
    {
        if (d == null) return;

        Export_QC_RawMaterial_Report(reportSavePath, d);
        Export_Helper(helperSavePath, d);
    }

    // =====================================================
    // (A) 匯出 QC_RawMaterial_Template.xlsx（報告用）
    // =====================================================
    private static void Export_QC_RawMaterial_Report(
        string savePath,
        Page1ExportData d)
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
            ws.Range["C4"].Value = d.ReportNo;
            ws.Range["E4"].Value = d.MaterialNo;
            ws.Range["E5"].Value = d.Material;
            ws.Range["C5"].Value = d.ArrivalDate;
            ws.Range["C6"].Value = d.TestingDate;
            ws.Range["E6"].Value = d.QtyText;
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

            int max = new[]
            {
                d.LotFulls.Count,
                d.Densities.Count,
                d.DeltaPs.Count,
                d.VocIns.Count,
                d.VocOuts.Count,
                d.OutgassingList.Count
            }.Min();

            for (int i = 0; i < max; i++)
            {
                int col = COL_FIRST + i;

                ws.Cells[ROW_LOTNO, col].Value = d.LotFulls[i];
                ws.Cells[ROW_DENS, col].Value = d.Densities[i];
                ws.Cells[ROW_DP, col].Value = d.DeltaPs[i];
                ws.Cells[ROW_VIN, col].Value = d.VocIns[i];
                ws.Cells[ROW_VOUT, col].Value = d.VocOuts[i];
                ws.Cells[ROW_OUTG, col].Value = d.OutgassingList[i];

                ws.Cells[ROW_EFF, col].Value =
                    (i == d.SelectedIndex) ? (object)d.Eff0 : "N.D.";

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
        Page1ExportData d)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "Helper_Template.xlsm");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Helper_Template.xlsm");
            return;
        }

        // 永遠重新複製 template
        File.Copy(templatePath, helperSavePath, true);

        HelperWorkbookInterop.Append(
            helperSavePath,
            "濾網工作表",
            (wshelper, startRow) =>
            {
                int n = d.LotFulls.Count;

                // ===== 每一筆資料 → 一列（關鍵）=====
                for (int i = 0; i < n; i++)
                {
                    int row = startRow + i;

                    wshelper.Cells[row, 1].Value = d.TestingDate;
                    wshelper.Cells[row, 2].Value = d.Material;
                    wshelper.Cells[row, 3].Value = "";
                    wshelper.Cells[row, 4].Value = d.LotFulls[i];
                    wshelper.Cells[row, 5].Value = d.VocIns[i];
                    wshelper.Cells[row, 6].Value = d.VocOuts[i];
                    wshelper.Cells[row, 7].Value = d.OutgassingList[i];
                    wshelper.Cells[row, 8].Value = d.DeltaPs[i];

                    // ★ 只有 SelectedIndex 那一列才有效率
                    if (i == d.SelectedIndex)
                    {
                        wshelper.Cells[row, 9].Value = d.Eff0;
                        wshelper.Cells[row, 10].Value = d.Eff10;
                    }
                    else
                    {
                        wshelper.Cells[row, 9].Value = "N.D.";
                        wshelper.Cells[row, 10].Value = "N.D.";
                    }
                }

                // ===== Mesh（結構化百分比，固定區塊）=====
                if (d.ParticleSizePercentages != null)
                {
                    int r = 7;
                    foreach (var kv in d.ParticleSizePercentages)
                    {
                        if (kv.Key.Contains("總重"))
                            continue;

                        wshelper.Cells[r, 1].Value = kv.Key;
                        wshelper.Cells[r, 2].Value = kv.Value/100;
                        wshelper.Cells[r, 2].NumberFormat = "0.0%";
                        r++;
                    }
                }

                // ===== 效率 11 點（S3 起，固定區塊）=====
                if (d.Efficiencies11 != null)
                {
                    int startEffRow = 3;
                    int col = 19; // S

                    for (int i = 0; i < d.Efficiencies11.Count && i < 11; i++)
                    {
                        wshelper.Cells[startEffRow + i, col].Value =
                            d.Efficiencies11[i];
                    }
                }
            }
        );
    }
}
