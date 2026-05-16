using System.Windows.Forms;

namespace WindowsFormsApp1.Data_Access.Page6
{
    public static class P6Mapper
    {
        public static P6Batch FromExportData(Page6ExportData d, string userName)
        {
            var batch = new P6Batch
            {
                ReportNo = d.ReportNo,
                TestDate = d.TestDate,
                UserName = userName,
                SuppliedNO = d.SuppliedNO
            };

            foreach (DataGridViewRow row in d.DataGrid.Rows)
            {
                if (row.IsNewRow) continue;

                var item = new P6Item
                {
                    Col1 = row.Cells[0].Value?.ToString(),
                    Col2 = row.Cells[1].Value?.ToString(),

                    Spec1 = row.Cells[2].Value?.ToString(),
                    Spec2 = row.Cells[3].Value?.ToString(),

                    Range1 = row.Cells[4].Value?.ToString(),
                    Range2 = row.Cells[5].Value?.ToString(),

                    Result = row.Cells[6].Value?.ToString(),
                    Judgment = row.Cells[7].Value?.ToString(),

                    Extra1 = row.Cells.Count > 8 ? row.Cells[8].Value?.ToString() : "",
                    Extra2 = row.Cells.Count > 9 ? row.Cells[9].Value?.ToString() : "",
                    Extra3 = row.Cells.Count > 10 ? row.Cells[10].Value?.ToString() : ""
                };

                batch.Items.Add(item);
            }

            return batch;
        }
    }
}