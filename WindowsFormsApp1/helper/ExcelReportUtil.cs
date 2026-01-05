using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelReportUtil
{
    public static void ExportPage1(string savePath, Page1ExportData d)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到公版檔：" + templatePath);
            return;
        }

        File.Copy(templatePath, savePath, true);

        var app = new Excel.Application();
        var wb = app.Workbooks.Open(savePath);
        var ws = (Excel.Worksheet)wb.Sheets[1];

        try
        {
            // ─────────────────────────────
            // (A) 一定要寫的資料（不管有沒有簽名）
            // ─────────────────────────────
            ws.Range["C4"].Value = d.ReportNo;
            ws.Range["E4"].Value = d.MaterialNo;
            ws.Range["E5"].Value = d.Material;
            ws.Range["C5"].Value = d.ArrivalDate;
            ws.Range["C6"].Value = d.TestingDate;
            ws.Range["E6"].Value = d.QtyText;

            const int COL_FIRST = 3;
            const int ROW_LOTNO = 10;
            const int ROW_DENS = 13;
            const int ROW_DP = 14;
            const int ROW_VIN = 15;
            const int ROW_VOUT = 16;
            const int ROW_OUTG = 17;
            const int ROW_EFF = 18;

            for (int i = 0; i < d.LotFulls.Count; i++)
            {
                int col = COL_FIRST + i;
                ws.Cells[ROW_LOTNO, col].Value = d.LotFulls[i];
                ws.Cells[ROW_DENS, col].Value = d.Densities[i];
                ws.Cells[ROW_DP, col].Value = d.DeltaPs[i];
                ws.Cells[ROW_VIN, col].Value = d.VocIns[i];
                ws.Cells[ROW_VOUT, col].Value = d.VocOuts[i];
                ws.Cells[ROW_OUTG, col].Value = d.OutgassingList[i];
                ws.Cells[ROW_EFF, col].Value =
                    (i == d.SelectedIndex) ? d.Eff0 : "N.D.";
            }

            // ─────────────────────────────
            // (B) 簽名：有就加，沒有就略過
            // ─────────────────────────────
            string sigPath = SignatureHelper.FindSignaturePath();
            if (!string.IsNullOrEmpty(sigPath) && File.Exists(sigPath))
            {
                Excel.Range target = ws.Range["E25"];

                ws.Shapes.AddPicture(
                    sigPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)target.Left,
                    (float)target.Top,
                    -1,
                    -1
                );
            }
            else
            {
                MessageBox.Show(
                    $"找不到使用者 {Environment.UserName} 的簽名檔，將略過簽名。",
                    "簽名檔不存在",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }

            wb.Save();
        }
        finally
        {
            wb.Close();
            app.Quit();
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);
        }
    }
}

