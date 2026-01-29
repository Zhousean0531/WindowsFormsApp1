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
            int maxRow = Math.Min(
                12,
                d.DataGrid.Rows.Count - 1   // 排除最後的新資料列
            );

            for (int r = 0; r < maxRow; r++)
            {
                var row = d.DataGrid.Rows[r];
                int excelRow = startRow + r;

                // B 欄 ← DGV[0]
                ws.Cells[excelRow, 2].Value =
                    row.Cells[0].Value?.ToString();

                // C 欄 ← DGV[1]
                ws.Cells[excelRow, 3].Value =
                    row.Cells[1].Value?.ToString();

                // D 欄 ← DGV[2] + DGV[3]
                ws.Cells[excelRow, 4].Value =
                    $"{row.Cells[2].Value} {row.Cells[3].Value}".Trim();

                // E 欄 ← DGV[4] + DGV[5]
                ws.Cells[excelRow, 5].Value =
                    $"{row.Cells[4].Value} {row.Cells[5].Value}".Trim();

                // F 欄 ← DGV[6]
                ws.Cells[excelRow, 6].Value =
                    row.Cells[6].Value?.ToString();

                // G 欄 ← DGV[7]
                ws.Cells[excelRow, 7].Value =
                    row.Cells[7].Value?.ToString();

                // 如果後面還有欄位（8 之後）
                int excelCol = 8; // H
                for (int c = 8; c < row.Cells.Count; c++)
                {
                    ws.Cells[excelRow, excelCol].Value =
                        row.Cells[c].Value?.ToString();
                    excelCol++;
                }

                System.Threading.Thread.Sleep(20);
                Application.DoEvents();
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
