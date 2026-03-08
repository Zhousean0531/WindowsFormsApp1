using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WindowsFormsApp1.Data_Access.Page4;
using static SignatureHelper;

public static class Page4ReportExporter
{
    public static void Export(P4Batch d)
    {
        if (d == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有效率資料");
            return;
        }

        string helperSavePath;

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Macro (*.xlsm)|*.xlsm";
            sfd.FileName = "Helper.xlsm";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            helperSavePath = sfd.FileName;
        }

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到報告範本");
            return;
        }

        foreach (var g in d.EfficiencyGroups)
        {
            DateTime arrivalDt = DateTime.Parse(d.ArrivalDate);

            string savePath;

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel (*.xlsx)|*.xlsx";

                sfd.FileName =
                    $"{d.ReportNo}_{d.Material}({arrivalDt:MMdd}到廠)_{g.GasName}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    continue;

                savePath = sfd.FileName;
            }

            File.Copy(templatePath, savePath, true);

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

                wb = app.Workbooks.Open(savePath);
                ws = (Excel.Worksheet)wb.Worksheets["濾網原料報告"];

                int idx = d.SelectedIndex;

                ws.Range["C4"].Value = d.ReportNo;
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

                for (int i = 0; i < d.LotFulls.Count; i++)
                {
                    int col = COL_FIRST + i;

                    ws.Cells[ROW_LOT, col].Value = d.LotFulls[i];
                    ws.Cells[ROW_DEN, col].Value = d.Densities[i];
                    ws.Cells[ROW_DP, col].Value = d.DeltaPs[i];
                    ws.Cells[ROW_VIN, col].Value = d.VocIns[i];
                    ws.Cells[ROW_VOUT, col].Value = d.VocOuts[i];
                    ws.Cells[ROW_OUTG, col].Value = d.OutgassingList[i];

                    ws.Cells[ROW_EFF, col].Value =
                        i == d.SelectedIndex
                        ? (g.Eff0?.ToString("F1") ?? "N.D.")
                        : "N.D.";
                }

                ExcelSignatureHelper.TryAddSignature(ws, "E25");

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

        Page4HelperExporter.Export(helperSavePath, d);

        Page4ReportExporterForNanJing.Export(d);
    }
}