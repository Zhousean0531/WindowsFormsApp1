using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page4;
using Excel = Microsoft.Office.Interop.Excel;

public static class Page4HelperExporter
{
    public static void Export(string helperSavePath, P4Batch d)
    {
        if (d == null || d.Rows == null || d.Rows.Count == 0)
            return;

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
            app = new Excel.Application();
            app.Visible = false;
            app.DisplayAlerts = false;

            wb = app.Workbooks.Open(helperSavePath);
            ws = (Excel.Worksheet)wb.Worksheets["濾筒工作表"];
            int startRow = 2;
            int row = startRow;

            // ===== Lot / 測試資料（改成 Rows）=====
            foreach (var r in d.Rows)
            {
                ws.Cells[row, 1].Value = d.TestingDate;
                ws.Cells[row, 2].Value = d.ArrivalDate;
                ws.Cells[row, 3].Value = d.Material;

                ws.Cells[row, 4].Value = r.LotNo;
                ws.Cells[row, 5].Value = r.LotFull;

                ws.Cells[row, 8].Value = r.Weight;
                ws.Cells[row, 9].Value = r.Density;
                ws.Cells[row, 10].value = d.Moisture;
                ws.Cells[row, 11].Value = r.VocIn;
                ws.Cells[row, 12].Value = r.VocOut;

                ws.Cells[row, 13].Value = r.Outgassing;
                ws.Cells[row, 14].Value = r.DeltaP;

                row++;
            }

            //Particle Size
            if (d.ParticleSizePercentages != null &&
                d.ParticleSizePercentages.Count > 0)
            {
                int meshRow = 8;

                foreach (var p in d.ParticleSizePercentages)
                {
                    ws.Cells[meshRow-1, 1].Value = p.Key;
                    ws.Cells[meshRow, 2].Value = p.Value / 100;
                    ws.Cells[meshRow, 2].NumberFormat = "0.0%";

                    meshRow++;
                }
            }

            // ===== Efficiency
            if (d.EfficiencyGroups != null && d.EfficiencyGroups.Count > 0)
            {
                int defaultCol = 27;

                //找選中的 row
                var selected = d.Rows.FirstOrDefault(x => x.IsSelected);
                if (selected == null) return;

                foreach (var g in d.EfficiencyGroups)
                {
                    if (g.Efficiencies11 == null ||
                        g.Efficiencies11.Count == 0)
                        continue;

                    int col = defaultCol;

                    if (string.Equals(g.GasName, "Acetone", StringComparison.OrdinalIgnoreCase))
                        col = 28;

                    if (string.Equals(g.GasName, "IPA", StringComparison.OrdinalIgnoreCase))
                        col = 29;

                    string testDate =
                        DateTime.Parse(d.TestingDate).ToString("MM.dd");

                    string arrivalDate =
                        DateTime.Parse(d.ArrivalDate).ToString("MM.dd");

                    //用 selected
                    string materialLot = selected.LotFull;

                    if (!string.IsNullOrEmpty(materialLot))
                    {
                        var parts = materialLot.Split('#');
                        materialLot = parts.Length > 1 ? parts[parts.Length - 1] : materialLot;
                    }
                    int selectedRow = startRow + d.Rows.IndexOf(selected);
                    int summaryCol = 15;

                    if (string.Equals(g.GasName, "Acetone", StringComparison.OrdinalIgnoreCase))
                        summaryCol = 17;

                    if (string.Equals(g.GasName, "IPA", StringComparison.OrdinalIgnoreCase))
                        summaryCol = 19;

                    ws.Cells[selectedRow, summaryCol].Value =
                        g.Eff0.HasValue ? (object)g.Eff0.Value : "N.D.";

                    ws.Cells[selectedRow, summaryCol + 1].Value =
                        g.Eff10.HasValue ? (object)g.Eff10.Value : "N.D.";

                    string header =
                        $"{testDate} {d.Material}#{materialLot}({selected.DeltaP}Pa)-{arrivalDate}Arrival_{g.GasName}";

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
            if (wb != null) wb.Close(false);
            if (app != null) app.Quit();

            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null) Marshal.ReleaseComObject(wb);
            if (app != null) Marshal.ReleaseComObject(app);
        }
    }
}