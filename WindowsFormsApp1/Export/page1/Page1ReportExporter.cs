using System.Windows.Forms;
using System;

public static class Page1ReportExporter
{
    public static void Export(Page1ExportData d)
    {
        DateTime ArrivalDt = DateTime.Parse(d.ArrivalDate);
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{d.ReportNo}_{d.Material}({ArrivalDt:MMdd}到廠).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK) return;

            ExcelReportUtil.ExportPage1(sfd.FileName, d);
        }
    }
}
