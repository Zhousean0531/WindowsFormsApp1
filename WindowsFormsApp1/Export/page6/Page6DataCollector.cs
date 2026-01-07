using System.Windows.Forms;

public static class Page6DataCollector
{
    public static Page6ExportData Collect(TabPage tab)
    {
        var dgv = ControlHelper.Find<DataGridView>(tab, "RawMaterialdgv");
        if (dgv == null || dgv.Rows.Count == 0)
        {
            MessageBox.Show("目前沒有資料可匯出");
            return null;
        }

        var reportNoTB =
            ControlHelper.Find<TextBox>(tab, "MaterialReportNOTB");

        var testDatePicker =
            ControlHelper.Find<DateTimePicker>(tab, "MaterialTestDateBox");

        if (reportNoTB == null || testDatePicker == null)
        {
            MessageBox.Show("找不到報告編號或測試日期欄位");
            return null;
        }

        return new Page6ExportData
        {
            ReportNo = reportNoTB.Text.Trim(),
            TestDate = testDatePicker.Value,
            DataGrid = dgv
        };
    }
}
