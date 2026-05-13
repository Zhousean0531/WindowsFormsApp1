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

            ws.Range["G3"].Value =
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

                // B ← 日期
                ws.Cells[excelRow, 2].Value = row.Cells[0].Value?.ToString();

                // C ← 品名
                ws.Cells[excelRow, 3].Value = row.Cells[1].Value?.ToString();

                // D ← 進貨量 + 單位
                ws.Cells[excelRow, 4].Value =
                    $"{row.Cells[2].Value} {row.Cells[3].Value}".Trim();

                // E ← 抽樣數 + 單位
                ws.Cells[excelRow, 5].Value =
                    $"{row.Cells[4].Value} {row.Cells[5].Value}".Trim();

                // F ← 外觀
                ws.Cells[excelRow, 6].Value =
                    row.Cells[6].Value?.ToString();

                // G ← 規格
                ws.Cells[excelRow, 7].Value =
                    row.Cells[7].Value?.ToString();

                // H ← ⭐測量值
                ws.Cells[excelRow, 9].Value =
                    row.Cells[8].Value?.ToString();

                // I ← 判定
                ws.Cells[excelRow, 10].Value =
                    row.Cells[9].Value?.ToString();

                // J ← 備註
                ws.Cells[excelRow, 11].Value =
                    row.Cells[10].Value?.ToString();

                System.Threading.Thread.Sleep(90);
                Application.DoEvents();
            }
            // =====簽名=====
            ExcelSignatureHelper.TryAddSignature(ws, "J18");

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
