using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;

public static class Page3ReportExporter
{
    public static void Export(
        Page3ExportData d,
        Page3PressureMode pressureMode,
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

        using (SaveFileDialog sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = BuildFileName(d, fileNameOverride);

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            File.Copy(templatePath, sfd.FileName, true);

            using (XLWorkbook wb = new XLWorkbook(sfd.FileName))
            {
                IXLWorksheet ws = wb.Worksheet(1);
                // ===== 1. 儀器校正資訊 =====
                WriteInstrumentInfo(ws, d);
                // ===== 2. 檢驗資料 =====
                int row = 4;

                foreach (Page3RowData r in d.Rows)
                {
                    // ===== 基本資料 =====
                    ws.Cell(row, "A").Value = ToReportText(d.TestingDate);
                    ws.Cell(row, "B").Value = ToReportText(d.FilterReportNo);
                    ws.Cell(row, "C").Value = ToReportText(d.PackageNo);
                    ws.Cell(row, "D").Value = ToReportText(d.Customer);
                    ws.Cell(row, "E").Value = ToReportText(d.Model);
                    ws.Cell(row, "F").Value = ResolveFilterTypeText(d);
                    ws.Cell(row, "G").Value = ToReportText(d.ReFilterNo);

                    ws.Cell(row, "H").Value = ToReportText(r.SN);
                    ws.Cell(row, "I").Value = ToReportText(r.Weight);
                    ws.Cell(row, "J").Value = "V";
                    ws.Cell(row, "K").Value = "V";
                    ws.Cell(row, "L").Value = "V";
                    ws.Cell(row, "M").Value = "V";

                    // ===== MA / MB / MC 初始效率 =====
                    ws.Cell(row, "N").Value = ToReportText(d.MA);
                    ws.Cell(row, "O").Value = ToReportText(d.MB);
                    ws.Cell(row, "P").Value = ToReportText(d.MC);

                    // ===== 檢測數據 =====
                    if (r.ControlValues != null)
                    {
                        foreach (KeyValuePair<int, string> kv in r.ControlValues)
                        {
                            int reportCol;

                            if (ReportColumnMap.TryGetValue(kv.Key, out reportCol))
                            {
                                ws.Cell(row, reportCol).Value = ToReportText(kv.Value);
                            }
                        }
                    }
                    WritePressureDrop(ws, row, r, pressureMode);

                    row++;
                }

                wb.Save();
            }
        }
    }

    // ===== 檔名處理 =====
    private static string BuildFileName(
        Page3ExportData d,
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

            fileName = reportNo + "(" + inside + ")";
        }

        fileName = SafeFileName(fileName);

        if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            fileName += ".xlsx";

        return fileName;
    }

    private static string SafeFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "";

        foreach (char c in Path.GetInvalidFileNameChars())
        {
            value = value.Replace(c.ToString(), "");
        }

        return value.Trim();
    }

    private static string ToReportText(string value)
    {
        return string.IsNullOrWhiteSpace(value)
            ? "N/A"
            : value.Trim();
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
    };
    private static void WritePressureDrop(
    IXLWorksheet ws,
    int row,
    Page3RowData r,
    Page3PressureMode pressureMode
)
    {
        string pressureSpec = GetControlText(r, 53);
        string pressureValue = GetControlText(r, 54);

        // 先全部寫 N/A，避免模板殘留舊值
        ws.Cell(row, "AI").Value = "N/A"; // 單片壓損規格
        ws.Cell(row, "AJ").Value = "N/A"; // 單片壓損數據
        ws.Cell(row, "AK").Value = "N/A"; // 整組壓損規格
        ws.Cell(row, "AL").Value = "N/A"; // 整組壓損數據

        if (pressureMode == Page3PressureMode.Single)
        {
            // 單片
            ws.Cell(row, "AI").Value = ToReportText(pressureSpec);
            ws.Cell(row, "AJ").Value = ToReportText(pressureValue);
        }
        else
        {
            // 整組
            ws.Cell(row, "AK").Value = ToReportText(pressureSpec);
            ws.Cell(row, "AL").Value = ToReportText(pressureValue);
        }
    }

    private static string GetControlText(Page3RowData r, int key)
    {
        if (r == null || r.ControlValues == null)
            return "";

        string value;

        if (r.ControlValues.TryGetValue(key, out value))
            return value;

        return "";
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

        if (types.Count == 0)
            return "N/A";

        return string.Join("/", types);
    }
    private static void WriteInstrumentInfo(IXLWorksheet ws, Page3ExportData d)
    {
        // 先全部寫 N/A，避免模板殘留舊資料
        WriteInstrumentNA(ws, "Q4", "Q5", "Q6");   // MA
        WriteInstrumentNA(ws, "R4", "R5", "R6");   // MB
        WriteInstrumentNA(ws, "S4", "S5", "S6");   // MC
        WriteInstrumentNA(ws, "W4", "W5", "W6");   // Particle
        WriteInstrumentNA(ws, "AH4", "AH5", "AH6"); // FT3+
        WriteInstrumentNA(ws, "AM4", "AM5", "AM6"); //PressureDrop

        List<string> instrumentNos = BuildPage3InstrumentNos(d);

        Dictionary<string, InstrumentReportInfo> instruments =
            InstrumentRepository.GetByInstrumentNos(instrumentNos);

        // MA → Thermo Fisher / 43i
        if (!string.IsNullOrWhiteSpace(d.MA))
        {
            WriteInstrumentByNo(ws, instruments, "QAD-API-05", "Q4", "Q5", "Q6");
        }

        // MB → Teledyne / T201
        if (!string.IsNullOrWhiteSpace(d.MB))
        {
            WriteInstrumentByNo(ws, instruments, "QAD-API-01", "R4", "R5", "R6");
        }

        // MC → Honeywell / RAE3000
        if (!string.IsNullOrWhiteSpace(d.MC))
        {
            WriteInstrumentByNo(ws, instruments, "QAD-PID-01", "S4", "S5", "S6");
        }

        // Page3 固定顯示
        WriteInstrumentByNo(ws, instruments, "QAD-GT-03", "W4", "W5", "W6");
        WriteInstrumentByNo(ws, instruments, "QAD-MIT-02", "AH4", "AH5", "AH6");
        WriteInstrumentByNo(ws, instruments, "QAD-IAQ-05", "AM4", "AM5", "AM6");
    }

    private static List<string> BuildPage3InstrumentNos(Page3ExportData d)
    {
        List<string> nos = new List<string>();

        if (!string.IsNullOrWhiteSpace(d.MA))
            nos.Add("QAD-API-05");

        if (!string.IsNullOrWhiteSpace(d.MB))
            nos.Add("QAD-API-01");

        if (!string.IsNullOrWhiteSpace(d.MC))
            nos.Add("QAD-PID-01");

        // Page3 固定顯示
        nos.Add("QAD-GT-03");
        nos.Add("QAD-MIT-02");
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
    public enum Page3PressureMode
    {
        Set,    // 整組
        Single  // 單片
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