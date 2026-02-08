using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

public static class Page5ReportExporter
{
    public static void Export(
        Page5ExportData d,
        string cylTypeText,          // MA / MB / MC
        string rawEfficiencyText,    // CYLRawEffTB
        List<int> columnsToCheck
    )
    {
        if (d == null || d.Rows == null || d.Rows.Count == 0)
            return;

        string templatePath = Path.Combine(
            Application.StartupPath,
            "CYLReport.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 CYLReport.xlsx");
            return;
        }

        string reportType = ResolveReportType(d.FilterType);

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{d.ReportNo}({d.CylinderNo}-{reportType}).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            File.Copy(templatePath, sfd.FileName, true);

            using (var wb = new XLWorkbook(sfd.FileName))
            {
                var ws = wb.Worksheet(1);

                // ===== 1️⃣ 儀器校正 =====
                var calibInfos =
                    CalibrationHelper.GetCalibrationInfosByColumns(columnsToCheck);

                WriteCalibrationInfo(ws, calibInfos, d.FilterType);

                // ===== 2️⃣ 檢驗資料 =====
                int row = 4; // Report 起始列
                int? effCol = ResolveEfficiencyColumn(cylTypeText);

                foreach (var r in d.Rows)
                {
                    ws.Cell(row, "A").Value = d.TestDate;
                    ws.Cell(row, "B").Value = d.ReportNo;
                    ws.Cell(row, "C").Value = d.CylinderNo;
                    ws.Cell(row, "D").Value = d.Customer;
                    ws.Cell(row, "E").Value = "濾筒";
                    ws.Cell(row, "F").Value = d.FilterType;
                    ws.Cell(row, "G").Value = d.ReCylinderNo;
                    ws.Cell(row, "H").Value = r.SN;
                    ws.Cell(row, "I").Value = r.Weight;

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

                    if (effCol.HasValue && !string.IsNullOrWhiteSpace(rawEfficiencyText))
                    {
                        ws.Cell(row, effCol.Value).Value = rawEfficiencyText;
                    }

                    row++;
                }

                wb.Save();
            }

            MessageBox.Show("匯出完成！");
        }
    }

    // =====================================================
    // 📌 工具區（Report 所需的「全部工具」都在下面）
    // =====================================================

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

    // ===== Master → Report 欄位對照 =====
    private static readonly Dictionary<int, int> ReportColumnMap =
        new Dictionary<int, int>
    {
        { 38, 20 }, { 39, 21 }, { 40, 22 },
        { 41, 24 }, { 42, 25 }, { 43, 26 },
        { 44, 27 }, { 45, 28 }, { 46, 29 },
        { 47, 30 }, { 48, 31 }, { 49, 32 },
        { 50, 33 }, { 54, 38 }
    };

    // ===== 儀器 → Report 位置 =====
    private static readonly Dictionary<string, Tuple<string, string, string>>
        CalibrationCellMap =
        new Dictionary<string, Tuple<string, string, string>>(StringComparer.OrdinalIgnoreCase)
    {
        { "SO2 Analyzer/43i",  Tuple.Create("Q4", "Q5", "Q6") },
        { "NH3 Analyzer/T201", Tuple.Create("R4", "R5", "R6") },
        { "Portable Handheld  VOC Monitor/ ppbRAE 3000",Tuple.Create("S4", "S5", "S6") },
        { "Handheld Particle Counter/GT-521S", Tuple.Create("W4", "W5", "W6") },
        { "MiTAP/FT3+", Tuple.Create("AH4", "AH5", "AH6") },
        { "Universal IAQ instrument/testo 400", Tuple.Create("AM4", "AM5", "AM6") },
    };

    // ===== MA / MB / MC → 儀器 =====
    private static readonly Dictionary<string, string>
        EfficiencyInstrumentMap =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "MA", "SO2 Analyzer/43i" },
        { "MB", "NH3 Analyzer/T201" },
        { "MC", "Portable Handheld  VOC Monitor/ ppbRAE 3000" }
    };

    // ===== 決定本次 Report 要寫哪些儀器 =====
    private static HashSet<string> ResolveInstrumentsToWrite(string filterType)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Handheld Particle Counter/GT-521S",
            "MiTAP/FT3+",
            "Universal IAQ instrument/testo 400"
        };

        if (!string.IsNullOrWhiteSpace(filterType) &&
            EfficiencyInstrumentMap.TryGetValue(filterType.Trim().ToUpper(), out string effInst))
        {
            result.Add(effInst);
        }

        return result;
    }

    // ===== 寫入儀器校正 =====
    private static void WriteCalibrationInfo(
        IXLWorksheet ws,
        List<CalibrationInfo> infos,
        string filterType
    )
    {
        if (infos == null || infos.Count == 0)
            return;

        var instrumentsToWrite = ResolveInstrumentsToWrite(filterType);

        foreach (var info in infos)
        {
            if (!instrumentsToWrite.Contains(info.InstrumentName))
                continue;

            if (!CalibrationCellMap.TryGetValue(info.InstrumentName, out var cellSet))
                continue;

            ws.Cell(cellSet.Item1).Value = info.InstrumentName;
            ws.Cell(cellSet.Item2).Value = info.CalibrationDate.ToString("yyyy.MM.dd");
            ws.Cell(cellSet.Item3).Value = info.ExpireDate.ToString("yyyy.MM.dd");
        }
    }
}
