using System;
using System.Windows.Forms;

public static class Page1ReportExporter
{
    public static void Export(Page1ExportData d)
    {
        DateTime arrivalDt = DateTime.Parse(d.ArrivalDate);

        // ───── 選擇「報告檔」存檔路徑 ─────
        string reportSavePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{d.ReportNo}_{d.Material}({arrivalDt:MMdd}到廠).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            reportSavePath = sfd.FileName;
        }

        // ───── 選擇「Helper 檔」存檔路徑（跟報告一樣）─────
        string helperSavePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Macro (*.xlsm)|*.xlsm";
            sfd.FileName = "Helper.xlsm";
            sfd.Title = "請選擇 Helper 檔案存放位置";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            helperSavePath = sfd.FileName;
        }

        // ───── 交給匯出工具處理 ─────
        ExcelReportUtil.ExportPage1(
            reportSavePath,
            helperSavePath,
            d
        );
    }
}
