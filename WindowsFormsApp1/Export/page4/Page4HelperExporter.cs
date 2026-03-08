using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WindowsFormsApp1.Data_Access.Page4;

public static class Page4HelperExporter
{
    public static void Export(string helperSavePath, P4Batch d)
    {
        if (d == null) return;

        string templatePath = Path.Combine(
            Application.StartupPath,
            "Helper_Template.xlsm"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Helper_Template.xlsm");
            return;
        }

        if (!File.Exists(helperSavePath))
        {
            File.Copy(templatePath, helperSavePath, true);
        }

        Excel.Application app = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        try
        {
            app = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            wb = app.Workbooks.Open(helperSavePath);
            ws = (Excel.Worksheet)wb.Worksheets["濾筒工作表"];

            int startRow = 2;
            int row = startRow;

            int n = d.LotFulls.Count;

            // ────────────────
            // Lot / 測試資料
            // ────────────────
            for (int i = 0; i < n; i++)
            {
                ws.Cells[row, 1].Value = d.TestingDate;
                ws.Cells[row, 2].Value = d.ArrivalDate;
                ws.Cells[row, 3].Value = d.Material;

                ws.Cells[row, 4].Value = d.LotNos[i];
                ws.Cells[row, 5].Value = d.LotFulls[i];

                ws.Cells[row, 8].Value = d.Weights[i];
                ws.Cells[row, 9].Value = d.Densities[i];

                ws.Cells[row, 11].Value = d.VocIns[i];
                ws.Cells[row, 12].Value = d.VocOuts[i];

                ws.Cells[row, 13].Value = d.OutgassingList[i];
                ws.Cells[row, 14].Value = d.DeltaPs[i];

                row++;
            }

            // ────────────────
            // Particle Size
            // ────────────────
            if (d.ParticleSizePercentages != null &&
                d.ParticleSizePercentages.Count > 0)
            {
                int meshRow = 8;

                foreach (var p in d.ParticleSizePercentages)
                {
                    ws.Cells[meshRow, 1].Value = p.Key;
                    ws.Cells[meshRow, 2].Value = p.Value / 100;
                    ws.Cells[meshRow, 2].NumberFormat = "0.0%";

                    meshRow++;
                }
            }

            // ────────────────
            // Efficiency
            // ────────────────
            if (d.EfficiencyGroups != null && d.EfficiencyGroups.Count > 0)
            {
                int defaultCol = 21; // U
                int idx = d.SelectedIndex;

                for (int gIdx = 0; gIdx < d.EfficiencyGroups.Count; gIdx++)
                {
                    var g = d.EfficiencyGroups[gIdx];

                    if (g.Efficiencies11 == null ||
                        g.Efficiencies11.Count == 0)
                        continue;

                    int col = defaultCol;

                    if (string.Equals(g.GasName, "Acetone", StringComparison.OrdinalIgnoreCase))
                        col = 22;

                    if (string.Equals(g.GasName, "IPA", StringComparison.OrdinalIgnoreCase))
                        col = 23;

                    string testDate =
                        DateTime.Parse(d.TestingDate).ToString("MM.dd");

                    string arrivalDate =
                        DateTime.Parse(d.ArrivalDate).ToString("MM.dd");

                    string lotFull = d.LotFulls[idx];

                    string header =
                        $"{testDate} {d.Material}#{lotFull}({d.DeltaPs[idx]}Pa)-{arrivalDate}Arrival_{g.GasName}";

                    ws.Cells[startRow - 1, col].Value = header;

                    for (int i = 0; i < g.Efficiencies11.Count; i++)
                    {
                        ws.Cells[startRow + 1 + i, col].Value =
                            g.Efficiencies11[i];
                    }
                }
            }

            wb.Save();
        }
        finally
        {
            wb?.Close(false);
            app?.Quit();

            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null) Marshal.ReleaseComObject(wb);
            if (app != null) Marshal.ReleaseComObject(app);
        }
    }
}