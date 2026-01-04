using System.Windows.Forms;

public static class Page1ReportExporter
{
    public static void Export(Page1ExportData d)
    {
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{d.ReportNo}_{d.Material}.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK) return;

            ExcelReportUtil.ExportPage1(sfd.FileName, d);
        }
    }
}
