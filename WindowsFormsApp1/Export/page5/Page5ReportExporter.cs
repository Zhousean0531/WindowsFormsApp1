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
        List<int> instrumentIds
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
            sfd.FileName = $"{d.ReportNo}({d.CylinderNo}-{reportType}).xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            File.Copy(templatePath, sfd.FileName, true);

            using (var wb = new XLWorkbook(sfd.FileName))
            {
                var ws = wb.Worksheet(1);
                var calibInfos = InstrumentRepository.GetByIds(instrumentIds);
                WriteCalibrationInfo(ws, calibInfos);
                // ===== 2. 檢驗資料 =====
                int row = 4; // Report 起始列
                int? effCol = ResolveEfficiencyColumn(cylTypeText);

                foreach (var r in d.Rows)
                {
                    ws.Cell(row, "A").Value = d.TestDate;
                    ws.Cell(row, "B").Value = d.ReportNo;
                    ws.Cell(row, "C").Value = d.CylinderNo;
                    ws.Cell(row, "D").Value = d.Customer;
                    ws.Cell(row, "E").Value = "CYL";
                    ws.Cell(row, "F").Value = d.FilterType;
                    ws.Cell(row, "G").Value = d.ReCylinderNo;
                    ws.Cell(row, "H").Value = r.SN;
                    ws.Cell(row, "I").Value = r.Weight;
                    ws.Cell(row, "J").Value = "V";
                    ws.Cell(row, "K").Value = "V";
                    ws.Cell(row, "L").Value = "V";
                    ws.Cell(row, "M").Value = "V";
                    ws.Cell(row, "N").Value = "N/A";
                    ws.Cell(row, "O").Value = "N/A";
                    ws.Cell(row, "P").Value = "N/A";
                    ws.Cell(row, "AK").Value = "130";
                    if (effCol.HasValue && !string.IsNullOrWhiteSpace(rawEfficiencyText))
                    {
                        ws.Cell(row, effCol.Value).Value = rawEfficiencyText.Trim();
                    }
                    // ===== Master → Report 欄位對照 =====
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
    // 工具區
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

    private static readonly Dictionary<string, Tuple<string, string, string>>
        CalibrationCellMap =
        new Dictionary<string, Tuple<string, string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            { "SO2 Analyzer/43i", Tuple.Create("Q4", "Q5", "Q6") },
            { "NH3 Analyzer/T201", Tuple.Create("R4", "R5", "R6") },
            { "Portable Handheld  VOC Monitor/ ppbRAE 3000", Tuple.Create("S4", "S5", "S6") },
            { "Handheld Particle Counter/GT-324", Tuple.Create("W4", "W5", "W6") },
            { "MiTAP/SFT3", Tuple.Create("AH4", "AH5", "AH6") },
            { "Universal IAQ instrument/testo 440dp", Tuple.Create("AM4", "AM5", "AM6") },
        };

    // ===== 寫入儀器校正 =====
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

            if (!CalibrationCellMap.TryGetValue(instrumentName, out var cellSet))
                continue;

            ws.Cell(cellSet.Item1).Value = instrumentName;
            ws.Cell(cellSet.Item2).Value = info.CalibrationDate.ToString("yyyy.MM.dd");
            ws.Cell(cellSet.Item3).Value = info.ExpireDate.ToString("yyyy.MM.dd");
        }
    }
}