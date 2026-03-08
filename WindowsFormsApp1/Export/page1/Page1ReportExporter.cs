using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page1;

public static class Page1ReportExporter
{
    public static void Export(P1Batch batch)
    {
        DateTime arrivalDt = batch.ArrivalDate;

        // ───── 選擇報告存檔路徑 ─────
        string reportSavePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{batch.ReportNo}_{batch.Material}({arrivalDt:MMdd}到廠).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            reportSavePath = sfd.FileName;
        }

        // ───── 選擇 Helper 存檔路徑 ─────
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

        // ───── 呼叫原本的 Excel 工具 ─────
        ExcelReportUtil.ExportPage1(
            reportSavePath,
            helperSavePath,
            batch
        );
    }
}