using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

public static class Page5ReportExporter
{
    // Master 欄位 index → Report 欄位 index
    private static readonly Dictionary<int, int> ReportColumnMap =
        new Dictionary<int, int>
    {
    { 38, 20 },  // Particle_In    
    { 39, 21 },  // Particle_Out   
    { 40, 22 },  // Particle Diff  

    { 41, 24 },  // IPA_In
    { 42, 25 },  // IPA_Out
    { 43, 26 }, // IPA Diff

    { 44, 27 }, // Acetone_In
    { 45, 28 }, // Acetone_Out
    { 46, 29 }, // Acetone Diff

    { 47, 30 }, // Nontarget_In
    { 48, 31 }, // Nontarget_Out
    { 49, 32 }, // Nontarget Diff

    { 50, 33 }, // TVOC out-in
    { 54, 38 }  // Pressure_Drop
    };

    public static void Export(
        Page5ExportData d,
        string cylTypeText,
        string rawEfficiencyText
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

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = $"{d.ReportNo}_{d.CylinderNo}.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            File.Copy(templatePath, sfd.FileName, true);

            using (var wb = new XLWorkbook(sfd.FileName))
            {
                var ws = wb.Worksheet(1);

                int row = 4; // 假設第 1 列是表頭

                // 👉 效率欄位只判斷一次
                int? effCol = ResolveEfficiencyColumn(cylTypeText);

                foreach (var r in d.Rows)
                {
                    ws.Cell(row, "A").Value = d.TestDate;
                    ws.Cell(row, "B").Value = d.ReportNo;
                    ws.Cell(row, "C").Value = d.CylinderNo;
                    ws.Cell(row, "D").Value = d.Customer;
                    ws.Cell(row, "E").Value = "濾筒";
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
                    // ===== 效率 =====
                    if (effCol.HasValue && !string.IsNullOrWhiteSpace(rawEfficiencyText))
                    {
                        ws.Cell(row, effCol.Value).Value = rawEfficiencyText;
                    }

                    row++;
                }

                wb.Save();
            }

            MessageBox.Show("Report 匯出完成！");
        }
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
}
