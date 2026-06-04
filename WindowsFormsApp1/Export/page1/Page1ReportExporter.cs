using System;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page1;

public static class Page1ReportExporter
{
    public static bool Export(P1Batch batch)
    {
        var staged = ExportStaged(batch);
        if (!staged.Success)
            return false;

        ReportExportStaging.Commit(staged.Files);
        return true;
    }

    public static ReportExportResult ExportStaged(P1Batch batch)
    {
        DateTime arrivalDt = batch.ArrivalDate;

        // ───── 選擇報告存檔路徑 ─────
        string reportSavePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{batch.ReportNo}_{batch.Material}({arrivalDt:MMdd}到廠).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return ReportExportResult.Failed();

            reportSavePath = sfd.FileName;
        }

        string helperSavePath;
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Macro (*.xlsm)|*.xlsm";
            sfd.FileName = "Helper.xlsm";
            sfd.Title = "請選擇 Helper 檔案存放位置";

            if (sfd.ShowDialog() != DialogResult.OK)
                return ReportExportResult.Failed();

            helperSavePath = sfd.FileName;
        }

        string tempReportPath = ReportExportStaging.CreateTempPath(reportSavePath);
        string tempHelperPath = ReportExportStaging.CreateTempPath(helperSavePath);

        // ───── 呼叫原本的 Excel 工具 ─────
        bool ok = ExcelReportUtil.ExportPage1(
            tempReportPath,
            tempHelperPath,
            batch
        );

        if (!ok)
            return ReportExportResult.Failed();

        var result = new ReportExportResult { Success = true };
        result.Files.Add(new ReportOutputFile { TempPath = tempReportPath, FinalPath = reportSavePath });
        result.Files.Add(new ReportOutputFile { TempPath = tempHelperPath, FinalPath = helperSavePath });
        return result;
    }
}
