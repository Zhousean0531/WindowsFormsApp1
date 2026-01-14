using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using static SignatureHelper;

public static class Page6ReportExporter
{
    public static void Export(Page6ExportData d)
    {
        if (d == null || d.DataGrid == null || d.DataGrid.Rows.Count == 0)
        {
            MessageBox.Show("目前沒有資料可匯出");
            return;
        }

        string templatePath = Path.Combine(
            Application.StartupPath,
            "RawMaterial_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到物料報告範本");
            return;
        }

        string savePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
            sfd.FileName =
                $"{d.ReportNo}({d.TestDate:MMdd}到廠).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

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
            ws = (Excel.Worksheet)wb.Sheets[1]; // 或指定名稱

            // ===== 表頭 =====
            ws.Range["C3"].Value =
                $"抽檢日期 Testing Date: {d.TestDate:yyyy.MM.dd}";

            ws.Range["I3"].Value =
                $"報告編號 Report No: {d.ReportNo}";

            // ===== DGV 資料（B6 起）=====
            int startRow = 6;
            int startCol = 2; // B

            for (int r = 0; r <12; r++)
            {
                var row = d.DataGrid.Rows[r];
                if (row.IsNewRow) continue;

                for (int c = 0; c < 12; c++)
                {
                    ws.Cells[startRow + r, startCol + c].Value =
                        row.Cells[c].Value?.ToString();
                }
            }

            // ===== 簽名（跟 Page2 一樣）=====
            ExcelSignatureHelper.TryAddSignature(ws, "L18");

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
