using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

public static class HelperWorkbookInterop
{
    public static void Append(
        string helperPath,
        string sheetName,
        Action<Excel.Worksheet, int> writeRow
    )
    {
        if (!File.Exists(helperPath))
        {
            MessageBox.Show($"找不到 {Path.GetFileName(helperPath)}");
            return;
        }

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

            wb = app.Workbooks.Open(
                helperPath,
                ReadOnly: false
            );

            ws = wb.Worksheets[sheetName] as Excel.Worksheet;
            if (ws == null)
            {
                MessageBox.Show($"找不到工作表：{sheetName}");
                return;
            }
            int startRow = 2;

            // 用「穩定的一欄」找最後一筆（通常是 A 欄）
            int lastRow = ws.Cells[
                ws.Rows.Count, 3
            ].End[Excel.XlDirection.xlUp].Row;

            int rowIdx = Math.Max(lastRow + 1, startRow);

            writeRow(ws, rowIdx);

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
