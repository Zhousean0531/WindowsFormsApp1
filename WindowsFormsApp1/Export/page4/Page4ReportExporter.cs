using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WindowsFormsApp1.Data_Access.Page4;
using static SignatureHelper;

public static class Page4ReportExporter
{
    public static bool Export(P4Batch d)
    {
        var staged = ExportStaged(d);
        if (!staged.Success)
            return false;

        ReportExportStaging.Commit(staged.Files);
        return true;
    }

    public static ReportExportResult ExportStaged(P4Batch d)
    {
        if (d == null || d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有效率資料");
            return ReportExportResult.Failed();
        }

        string helperSavePath;

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Macro (*.xlsm)|*.xlsm";
            sfd.FileName = "Helper.xlsm";

            if (sfd.ShowDialog() != DialogResult.OK)
                return ReportExportResult.Failed();

            helperSavePath = sfd.FileName;
        }

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到報告範本");
            return ReportExportResult.Failed();
        }

        var result = new ReportExportResult { Success = true };

        foreach (var g in d.EfficiencyGroups)
        {
            DateTime arrivalDt = ParseDateOrToday(d.ArrivalDate);
            string reportNo = g.ReportNo ?? d.ReportNo;
            string savePath;

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel (*.xlsx)|*.xlsx";
                sfd.FileName =
                    $"{reportNo}_{d.Material}({arrivalDt:MMdd}到廠)_{g.GasName}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                {
                    ReportExportStaging.Cleanup(result.Files);
                    return ReportExportResult.Failed();
                }

                savePath = sfd.FileName;
            }

            string tempPath = ReportExportStaging.CreateTempPath(savePath);

            if (!ExportSingleReport(tempPath, templatePath, d, g))
            {
                ReportExportStaging.Cleanup(result.Files);
                return ReportExportResult.Failed();
            }

            result.Files.Add(new ReportOutputFile
            {
                TempPath = tempPath,
                FinalPath = savePath
            });
        }

        string tempHelperPath = ReportExportStaging.CreateTempPath(helperSavePath);
        Page4HelperExporter.Export(tempHelperPath, d);

        if (!File.Exists(tempHelperPath))
        {
            ReportExportStaging.Cleanup(result.Files);
            return ReportExportResult.Failed();
        }

        result.Files.Add(new ReportOutputFile
        {
            TempPath = tempHelperPath,
            FinalPath = helperSavePath
        });

        var nanJing = Page4ReportExporterForNanJing.ExportStaged(d);
        if (!nanJing.Success)
        {
            ReportExportStaging.Cleanup(result.Files);
            return ReportExportResult.Failed();
        }

        result.Files.AddRange(nanJing.Files);
        return result;
    }

    private static bool ExportSingleReport(
        string savePath,
        string templatePath,
        P4Batch d,
        P4EfficiencyGroup g
    )
    {
        File.Copy(templatePath, savePath, true);

        Excel.Application app = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        try
        {
            app = ExcelInteropHelper.CreateApplication();
            wb = ExcelInteropHelper.OpenWorkbook(app, savePath);
            ws = (Excel.Worksheet)wb.Worksheets["濾網原料報告"];

            ws.Range["C4"].Value = g.ReportNo ?? d.ReportNo;
            ws.Range["C5"].Value = d.ArrivalDate;
            ws.Range["C6"].Value = d.TestingDate;
            ws.Range["E4"].Value = d.MaterialNo;
            ws.Range["E5"].Value = d.Material;
            ws.Range["E6"].Value = d.QtyText;

            const int COL_FIRST = 3;
            const int ROW_LOT = 10;
            const int ROW_DEN = 13;
            const int ROW_DP = 14;
            const int ROW_VIN = 15;
            const int ROW_VOUT = 16;
            const int ROW_OUTG = 17;
            const int ROW_EFF = 18;
            const int ROW_MOISTURE = 19;
            const int ROW_BUTANE = 20;
            const int ROW_ASH = 21;

            for (int i = 0; i < d.Rows.Count; i++)
            {
                int col = COL_FIRST + i;
                var r = d.Rows[i];

                ws.Cells[ROW_LOT, col].Value = r.LotFull;
                ws.Cells[ROW_DEN, col].Value = r.Density;
                ws.Cells[ROW_DP, col].Value = r.DeltaP;
                ws.Cells[ROW_VIN, col].Value = r.VocIn;
                ws.Cells[ROW_VOUT, col].Value = r.VocOut;
                ws.Cells[ROW_OUTG, col].Value = r.Outgassing;

                ws.Cells[ROW_EFF, col].Value =
                    r.IsSelected
                        ? (g.Eff0?.ToString("F1") ?? "N.D.")
                        : "N.D.";

                string moistureText = d.Moisture;
                string butaneText = d.Butane;
                string ashText = d.Ash;

                ws.Cells[ROW_MOISTURE, col].Value =
                    r.IsSelected
                        ? moistureText
                        : (string.IsNullOrWhiteSpace(moistureText) ? "" : "N.D.");

                ws.Cells[ROW_BUTANE, col].Value =
                    r.IsSelected
                        ? butaneText
                        : (string.IsNullOrWhiteSpace(butaneText) ? "" : "N.D.");

                ws.Cells[ROW_ASH, col].Value =
                    r.IsSelected
                        ? ashText
                        : (string.IsNullOrWhiteSpace(ashText) ? "" : "N.D.");
            }

            ExcelSignatureHelper.TryAddSignature(ws, "E25");
            ExcelInteropHelper.Save(wb);
            return true;
        }
        finally
        {
            ExcelInteropHelper.CloseWorkbook(wb, false);
            ExcelInteropHelper.Quit(app);

            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null) Marshal.ReleaseComObject(wb);
            if (app != null) Marshal.ReleaseComObject(app);
        }
    }

    private static DateTime ParseDateOrToday(string value)
    {
        if (DateTime.TryParse(value, out DateTime result))
            return result;

        return DateTime.Today;
    }
}
