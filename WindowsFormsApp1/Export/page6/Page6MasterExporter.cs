using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

public static class Page6MasterExporter
{
    public static void Export(Page6ExportData d)
    {
        if (d == null || d.DataGrid == null || d.DataGrid.Rows.Count == 0)
            return;

        string filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "總表.xlsx"
        );

        using (var wb = new XLWorkbook(filePath))
        {
            // ⚠️ 工作表名稱請依你實際總表調整
            var ws = wb.Worksheet("物料");

            // ===== 找最後一列（照 Page1）=====
            int row = (ws.Column(1).CellsUsed().LastOrDefault()?.Address.RowNumber ?? 1) + 1;

            foreach (var dgvRow in d.DataGrid.Rows.Cast<System.Windows.Forms.DataGridViewRow>())
            {
                if (dgvRow.IsNewRow) continue;
                ws.Cell(row, 1).Value = dgvRow.Cells[0].Value?.ToString();
                ws.Cell(row, 2).Value = d.TestDate.ToString("yyyy.MM.dd");
                ws.Cell(row, 3).Value = d.ReportNo;
                ws.Cell(row, 4).Value = dgvRow.Cells[1].Value?.ToString();
                ws.Cell(row, 5).Value = $"{dgvRow.Cells[2].Value}{dgvRow.Cells[3].Value}";
                ws.Cell(row, 6).Value = $"{dgvRow.Cells[4].Value}{dgvRow.Cells[5].Value}";
                ws.Cell(row, 7).Value = dgvRow.Cells[7].Value?.ToString();
                ws.Cell(row, 8).Value = dgvRow.Cells[8].Value?.ToString();
                ws.Cell(row, 9).Value = dgvRow.Cells[10].Value?.ToString();
                row++;
            }

            wb.Save();
        }
    }
}
