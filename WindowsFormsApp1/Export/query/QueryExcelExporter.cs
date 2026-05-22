using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

public static class QueryExcelExporter
{
    public static void Export(DataTable table)
    {
        if (table == null || table.Rows.Count == 0)
        {
            MessageBox.Show("目前沒有可匯出的查詢資料");
            return;
        }

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = "QC查詢結果_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("查詢結果");
                ws.Cell(1, 1).InsertTable(table, "QueryResult", true);
                ws.Columns().AdjustToContents();

                string dir = Path.GetDirectoryName(sfd.FileName);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                wb.SaveAs(sfd.FileName);
            }

            MessageBox.Show("匯出完成");
        }
    }
}
