using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;

public static class Page5ReportExporter
{
    public static void Export(
        Page5ExportData d,
        string cylTypeText,
        string rawEfficiencyText,
        List<int> instrumentIds,
        string fileNameOverride = null
    )
    {
        if (d == null || d.Rows == null || d.Rows.Count == 0)
            return;

        string templatePath = Path.Combine(
            Application.StartupPath,
            "Report.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Report.xlsx");
            return;
        }

        string reportType = ResolveReportType(d.FilterType);

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = BuildFileName(d, reportType, fileNameOverride);

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            File.Copy(templatePath, sfd.FileName, true);

            using (var wb = new XLWorkbook(sfd.FileName))
            {
                var ws = wb.Worksheet(1);

                // ===== 1. 儀器校正資訊 =====
                var calibInfos = InstrumentRepository.GetByIds(instrumentIds);
                WriteCalibrationInfo(ws, calibInfos);

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
                    ws.Cell(row, "AI").Value = "N/A";
                    ws.Cell(row, "AJ").Value = "N/A";
                    ws.Cell(row, "AK").Value = "130";
                    ws.Cell(row, "AL").Value = "N/A";
                    if (r.ControlValues != null)
                    {
                        foreach (var kv in r.ControlValues)
                        {
                            if (ReportColumnMap.TryGetValue(kv.Key, out int reportCol))
                            {
                                ws.Cell(row, reportCol).Value = kv.Value;
                            }
                        }
                    }
                    row++;
                }

                wb.Save();
            }

            MessageBox.Show("匯出完成！");
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
        { 54, 38 }  // AL
    };
    // ===== 一般欄位空值轉 N/A =====
    private static string ToReportText(string value)
    {
        return string.IsNullOrWhiteSpace(value)
            ? "N/A"
            : value.Trim();
    }

    // ===== 報告類型 =====
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

    // ===== 儀器校正欄位對照 =====
    private static readonly Dictionary<string, Tuple<string, string, string>>
        CalibrationCellMap =
        new Dictionary<string, Tuple<string, string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            // SO2
            { "SO2 ANALYZER 43I", Tuple.Create("Q4", "Q5", "Q6") },

            // NH3
            { "NH3 ANALYZER T201", Tuple.Create("R4", "R5", "R6") },

            // VOC
            { "PORTABLE HANDHELD VOC MONITOR PPBRAE 3000", Tuple.Create("S4", "S5", "S6") },

            // Particle Counter
            { "HANDHELD PARTICLE COUNTER GT 324", Tuple.Create("W4", "W5", "W6") },

            // MiTAP
            { "MITAP SFT3", Tuple.Create("AH4", "AH5", "AH6") },

            // Pressure Drop
            { "UNIVERSAL IAQ INSTRUMENT TESTO 440DP", Tuple.Create("AM4", "AM5", "AM6") },
        };

    // ===== 寫入儀器校正資訊 =====
    private static void WriteCalibrationInfo(
        IXLWorksheet ws,
        List<CalibrationInfo> infos
    )
    {
        if (infos == null || infos.Count == 0)
            return;

        foreach (var info in infos)
        {
            if (info == null)
                continue;

            string instrumentName = info.InstrumentName?.Trim();

            if (string.IsNullOrWhiteSpace(instrumentName))
                continue;

            string key = NormalizeInstrumentName(instrumentName);

            if (!CalibrationCellMap.TryGetValue(key, out var cellSet))
                continue;

            ws.Cell(cellSet.Item1).Value = instrumentName;
            ws.Cell(cellSet.Item2).Value = info.CalibrationDate.ToString("yyyy.MM.dd");
            ws.Cell(cellSet.Item3).Value = info.ExpireDate.ToString("yyyy.MM.dd");
        }
    }

    // ===== 儀器名稱正規化 =====
    private static string NormalizeInstrumentName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "";

        string s = name.Trim().ToUpper();

        s = s.Replace("/", " ");
        s = s.Replace("\\", " ");
        s = s.Replace("-", " ");
        s = s.Replace("_", " ");
        s = s.Replace(".", " ");

        while (s.Contains("  "))
        {
            s = s.Replace("  ", " ");
        }

        return s.Trim();
    }
}