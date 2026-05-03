using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;

public static class Page3ReportExporter
{
    public static void Export(Page3ExportData d)
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

        List<int> instrumentIds = new List<int>
        {
            1, 2, 3, 5, 7, 9
        };

        using (SaveFileDialog sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = BuildFileName(d);

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            File.Copy(templatePath, sfd.FileName, true);

            using (XLWorkbook wb = new XLWorkbook(sfd.FileName))
            {
                IXLWorksheet ws = wb.Worksheet(1);

                // ===== 1. 儀器校正 =====
                List<CalibrationInfo> calibInfos =
                    InstrumentRepository.GetByIds(instrumentIds);

                WriteCalibrationInfo(ws, calibInfos);

                // ===== 2. 檢驗資料 =====
                int row = 4;

                foreach (Page3RowData r in d.Rows)
                {
                    ws.Cell(row, "A").Value = d.TestingDate;
                    ws.Cell(row, "B").Value = d.FilterReportNo;
                    ws.Cell(row, "C").Value = d.PackageNo;
                    ws.Cell(row, "D").Value = d.Customer;
                    ws.Cell(row, "E").Value = d.Model;
                    ws.Cell(row, "F").Value = ResolveFilterTypeText(d);
                    ws.Cell(row, "G").Value = d.ReFilterNo;

                    ws.Cell(row, "H").Value = r.SN;
                    ws.Cell(row, "I").Value = r.Weight;
                    ws.Cell(row, "J").Value = r.Length;
                    ws.Cell(row, "K").Value = r.Width;
                    ws.Cell(row, "L").Value = r.Height;
                    ws.Cell(row, "M").Value = r.Diagonal;

                    // MA / MB / MC 效率
                    ws.Cell(row, "N").Value = string.IsNullOrWhiteSpace(d.MA) ? "N/A" : d.MA;
                    ws.Cell(row, "O").Value = string.IsNullOrWhiteSpace(d.MB) ? "N/A" : d.MB;
                    ws.Cell(row, "P").Value = string.IsNullOrWhiteSpace(d.MC) ? "N/A" : d.MC;

                    if (r.ControlValues != null)
                    {
                        foreach (KeyValuePair<int, string> kv in r.ControlValues)
                        {
                            int reportCol;

                            if (ReportColumnMap.TryGetValue(kv.Key, out reportCol))
                            {
                                ws.Cell(row, reportCol).Value = kv.Value;
                            }
                        }
                    }

                    row++;
                }

                wb.Save();
            }
        }
    }

    private static string BuildFileName(Page3ExportData d)
    {
        string reportNo = SafeFileName(d.FilterReportNo);
        string materialNo = SafeFileName(d.FilterMaterialNo);
        string packageNo = SafeFileName(d.PackageNo);
        string workOrder = SafeFileName(d.WorkOrder);

        string inside;

        if (string.IsNullOrWhiteSpace(workOrder) || workOrder == "-")
        {
            inside = materialNo + "_" + packageNo;
        }
        else
        {
            inside = materialNo + "_" + packageNo + "_" + workOrder;
        }

        return reportNo + "(" + inside + ").xlsx";
    }

    private static string SafeFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "";

        char[] invalidChars = Path.GetInvalidFileNameChars();

        foreach (char c in invalidChars)
        {
            value = value.Replace(c.ToString(), "");
        }

        return value.Trim();
    }

    // ===== ControlValues key → Report 欄位對照 =====
    private static readonly Dictionary<int, int> ReportColumnMap =
        new Dictionary<int, int>
    {
        { 38, 20 }, // ParticleIn       T
        { 39, 21 }, // ParticleOut      U
        { 40, 22 }, // ParticleDiff     V

        { 41, 24 }, // IPAIn            X
        { 42, 25 }, // IPAOut           Y
        { 43, 26 }, // IPADiff          Z

        { 44, 27 }, // AcetoneIn        AA
        { 45, 28 }, // AcetoneOut       AB
        { 46, 29 }, // AcetoneDiff      AC

        { 47, 30 }, // NonTargetIn      AD
        { 48, 31 }, // NonTargetOut     AE
        { 49, 32 }, // NonTargetDiff    AF

        { 50, 33 }, // TotalDiff        AG

        { 53, 37 }, // PressureDropSpec AK
        { 54, 38 }  // PressureDrop     AL
    };

    // ===== 儀器 → Report 位置 =====
    private static readonly Dictionary<string, Tuple<string, string, string>>
        CalibrationCellMap =
        new Dictionary<string, Tuple<string, string, string>>(StringComparer.OrdinalIgnoreCase)
    {
        { "SO2 Analyzer/43i", Tuple.Create("Q4", "Q5", "Q6") },
        { "NH3 Analyzer/T201", Tuple.Create("R4", "R5", "R6") },
        { "Portable Handheld  VOC Monitor/ ppbRAE 3000", Tuple.Create("S4", "S5", "S6") },
        { "Handheld Particle Counter/GT-324", Tuple.Create("W4", "W5", "W6") },
        { "MiTAP/SFT3", Tuple.Create("AH4", "AH5", "AH6") },
        { "Universal IAQ instrument/testo 440dp", Tuple.Create("AM4", "AM5", "AM6") }
    };

    private static void WriteCalibrationInfo(
        IXLWorksheet ws,
        List<CalibrationInfo> infos
    )
    {
        if (infos == null || infos.Count == 0)
            return;

        foreach (CalibrationInfo info in infos)
        {
            Tuple<string, string, string> cellSet;

            if (!CalibrationCellMap.TryGetValue(info.InstrumentName, out cellSet))
                continue;

            ws.Cell(cellSet.Item1).Value = info.InstrumentName;
            ws.Cell(cellSet.Item2).Value = info.CalibrationDate.ToString("yyyy.MM.dd");
            ws.Cell(cellSet.Item3).Value = info.ExpireDate.ToString("yyyy.MM.dd");
        }
    }
    private static string ResolveFilterTypeText(Page3ExportData d)
    {
        List<string> types = new List<string>();

        if (!string.IsNullOrWhiteSpace(d.MA))
            types.Add("MA");

        if (!string.IsNullOrWhiteSpace(d.MB))
            types.Add("MB");

        if (!string.IsNullOrWhiteSpace(d.MC))
            types.Add("MC");

        return string.Join("/", types);
    }
}