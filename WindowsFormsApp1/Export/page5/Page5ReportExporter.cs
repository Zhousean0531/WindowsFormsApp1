using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;

public static class Page5ReportExporter
{
    public static bool Export(
        Page5ExportData d,
        string cylTypeText,
        string rawEfficiencyText,
        string fileNameOverride = null
    )

    {
        var staged = ExportStaged(d, cylTypeText, rawEfficiencyText, fileNameOverride);
        if (!staged.Success)
            return false;

        ReportExportStaging.Commit(staged.Files);
        return true;
    }

    public static ReportExportResult ExportStaged(
        Page5ExportData d,
        string cylTypeText,
        string rawEfficiencyText,
        string fileNameOverride = null
    )

    {
        if (d == null || d.Rows == null || d.Rows.Count == 0)
            return ReportExportResult.Failed();

        string templatePath = Path.Combine(
            Application.StartupPath,
            "Report.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Report.xlsx");
            return ReportExportResult.Failed();
        }

        string reportType = ResolveFileNameReportType(d);

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = BuildFileName(d, reportType, fileNameOverride);

            if (sfd.ShowDialog() != DialogResult.OK)
                return ReportExportResult.Failed();

            string tempPath = ReportExportStaging.CreateTempPath(sfd.FileName);
            File.Copy(templatePath, tempPath, true);

            using (var wb = new XLWorkbook(tempPath))
            {
                var ws = wb.Worksheet(1);
                // ===== 1. 儀器校正資訊 =====
                WriteInstrumentInfo(ws, cylTypeText);
                // ===== 2. 檢驗資料 =====
                int row = 4; // Report 起始列
                int? effCol = ResolveEfficiencyColumn(cylTypeText);

                foreach (var r in d.Rows)
                {
                    // ===== 基本資料 =====
                    ws.Cell(row, "A").Value = ToReportText(d.TestDate);
                    ws.Cell(row, "B").Value = ToReportText(d.ReportNo);
                    ws.Cell(row, "C").Value = ToReportText(d.CylinderNo);
                    ws.Cell(row, "D").Value = ToReportText(d.Customer);
                    ws.Cell(row, "E").Value = "CYL";
                    ws.Cell(row, "F").Value = ToReportText(d.FilterType);
                    ws.Cell(row, "G").Value = ToReportText(d.ReCylinderNo);
                    ws.Cell(row, "H").Value = ToReportText(r.SN);
                    ws.Cell(row, "I").Value = ToReportText(r.Weight);
                    ws.Cell(row, "J").Value = "V";
                    ws.Cell(row, "K").Value = "V";
                    ws.Cell(row, "L").Value = "V";
                    ws.Cell(row, "M").Value = "V";
                    ws.Cell(row, "N").Value = "N/A";
                    ws.Cell(row, "O").Value = "N/A";
                    ws.Cell(row, "P").Value = "N/A";
                    if (effCol.HasValue)
                    {
                        ws.Cell(row, effCol.Value).Value = ToReportText(rawEfficiencyText);
                    }
                    ws.Cell(row, "AI").Value = "130";
                    if (r.ControlValues != null)
                    {
                        foreach (var kv in r.ControlValues)
                        {
                            if (ReportColumnMap.TryGetValue(kv.Key, out int reportCol))
                            {
                                ws.Cell(row, reportCol).Value = ToReportText(kv.Value);
                            }
                        }
                    }
                    row++;
                }

                wb.Save();
            }

            return ReportExportResult.FromFile(tempPath, sfd.FileName);
        }
    }

    // ===== 檔名處理 =====
    private static string BuildFileName(
        Page5ExportData d,
        string reportType,
        string fileNameOverride
    )
    {
        string fileName;

        if (!string.IsNullOrWhiteSpace(fileNameOverride))
        {
            fileName = fileNameOverride.Trim();
        }
        else
        {
            fileName = $"{d.ReportNo}({d.CylinderNo}-{reportType})";
        }

        foreach (char c in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(c.ToString(), "");
        }

        if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            fileName += ".xlsx";

        return fileName;
    }
    private static readonly Dictionary<int, int> ReportColumnMap =
        new Dictionary<int, int>
    {
        { 38, 20 }, // T
        { 39, 21 }, // U
        { 40, 22 }, // V
        { 41, 24 }, // X
        { 42, 25 }, // Y
        { 43, 26 }, // Z
        { 44, 27 }, // AA
        { 45, 28 }, // AB
        { 46, 29 }, // AC
        { 47, 30 }, // AD
        { 48, 31 }, // AE
        { 49, 32 }, // AF
        { 50, 33 }, // AG
        { 54, 36 }  // AJ
    };
    // ===== 一般欄位空值轉 N/A =====
    private static string ToReportText(string value)
    {
        return string.IsNullOrWhiteSpace(value)
            ? "N/A"
            : value.Trim();
    }

    // ===== 報告類型 =====
    private static string ResolveFileNameReportType(Page5ExportData d)
    {
        if (d == null)
            return "";

        string rawMaterialType = (d.RawMaterialType ?? "").Trim().ToUpper();
        string materialNo = (d.MaterialNo ?? "").Trim().ToUpper();

        if (rawMaterialType == "MA")
        {
            if (materialNo.Contains("11A0C00Y000002"))
                return "ACID";

            if (materialNo.Contains("11D0S00Y000002"))
                return "DMS";
        }

        return ResolveReportType(d.RawMaterialType);
    }

    private static string ResolveReportType(string filterType)
    {
        if (string.IsNullOrWhiteSpace(filterType))
            return "";

        filterType = filterType.Trim().ToUpper();

        if (filterType == "MA") return "DMS";
        if (filterType == "MB") return "BASE";
        if (filterType == "MC") return "TOC";

        return "";
    }

    // ===== 初始效率欄位 =====
    private static int? ResolveEfficiencyColumn(string cylType)
    {
        if (string.IsNullOrWhiteSpace(cylType))
            return null;

        cylType = cylType.Trim().ToUpper();

        if (cylType.Contains("MA")) return 14; // N
        if (cylType.Contains("MB")) return 15; // O
        if (cylType.Contains("MC")) return 16; // P

        return null;
    }
    private static void WriteInstrumentInfo(IXLWorksheet ws, string cylTypeText)
    {
        // 先全部寫 N/A，避免模板殘留舊資料
        WriteInstrumentNA(ws, "Q4", "Q5", "Q6");    // MA
        WriteInstrumentNA(ws, "R4", "R5", "R6");    // MB
        WriteInstrumentNA(ws, "S4", "S5", "S6");    // MC
        WriteInstrumentNA(ws, "W4", "W5", "W6");    // Met One / GT-324
        WriteInstrumentNA(ws, "AH4", "AH5", "AH6"); // FT3+
        WriteInstrumentNA(ws, "AK4", "AK5", "AK6"); // PressureDrop

        List<string> instrumentNos = BuildPage5InstrumentNos(cylTypeText);

        Dictionary<string, InstrumentReportInfo> instruments =
            InstrumentRepository.GetByInstrumentNos(instrumentNos);

        string type = "";
        if (!string.IsNullOrWhiteSpace(cylTypeText))
            type = cylTypeText.Trim().ToUpper();

        // MA → Thermo Fisher / 43i
        if (type.Contains("MA"))
        {
            WriteInstrumentByNo(ws, instruments, "QAD-API-05", "Q4", "Q5", "Q6");
        }

        // MB → Teledyne / T201
        if (type.Contains("MB"))
        {
            WriteInstrumentByNo(ws, instruments, "QAD-API-01", "R4", "R5", "R6");
        }

        // MC → Honeywell / RAE3000
        if (type.Contains("MC"))
        {
            WriteInstrumentByNo(ws, instruments, "QAD-PID-01", "S4", "S5", "S6");
        }

        // Page5 固定顯示
        WriteInstrumentByNo(ws, instruments, "QAD-GT-03", "W4", "W5", "W6");
        WriteInstrumentByNo(ws, instruments, "QAD-MIT-03", "AH4", "AH5", "AH6");
        WriteInstrumentByNo(ws, instruments, "QAD-IAQ-05", "AK4", "AK5", "AK6");
    }

    private static List<string> BuildPage5InstrumentNos(string cylTypeText)
    {
        List<string> nos = new List<string>();

        string type = "";
        if (!string.IsNullOrWhiteSpace(cylTypeText))
            type = cylTypeText.Trim().ToUpper();

        if (type.Contains("MA"))
            nos.Add("QAD-API-05");

        if (type.Contains("MB"))
            nos.Add("QAD-API-01");

        if (type.Contains("MC"))
            nos.Add("QAD-PID-01");

        // Page5 固定顯示
        nos.Add("QAD-GT-03");
        nos.Add("QAD-MIT-03");
        nos.Add("QAD-IAQ-05");
        return nos;
    }

    private static void WriteInstrumentByNo(
        IXLWorksheet ws,
        Dictionary<string, InstrumentReportInfo> instruments,
        string instrumentNo,
        string nameCell,
        string calibrationDateCell,
        string expireDateCell
    )
    {
        InstrumentReportInfo info;

        if (instruments == null || !instruments.TryGetValue(instrumentNo, out info))
        {
            WriteInstrumentNA(ws, nameCell, calibrationDateCell, expireDateCell);
            return;
        }

        ws.Cell(nameCell).Value = ToReportText(info.InstrumentName);
        ws.Cell(calibrationDateCell).Value = info.CalibrationDate.ToString("yyyy.MM.dd");
        ws.Cell(expireDateCell).Value = info.ExpireDate.ToString("yyyy.MM.dd");
    }

    private static void WriteInstrumentNA(
        IXLWorksheet ws,
        string nameCell,
        string calibrationDateCell,
        string expireDateCell
    )
    {
        ws.Cell(nameCell).Value = "N/A";
        ws.Cell(calibrationDateCell).Value = "N/A";
        ws.Cell(expireDateCell).Value = "N/A";
    }
}
