using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WindowsFormsApp1.Data_Access.Page4;
using static SignatureHelper;

public static class Page4ReportExporterForNanJing
{
    public static void Export(P4Batch d)
    {
        if (d.Material != "SG017_A" && d.Material != "SG017_D")
            return;

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterialForNanJing_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到南京範本");
            return;
        }

        DateTime arrivalDt = DateTime.Parse(d.ArrivalDate);

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";

            sfd.FileName =
                $"{d.ReportNo}_{d.Material}({arrivalDt:MMdd}到廠)_加測.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

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

                wb = app.Workbooks.Open(templatePath);
                ws = (Excel.Worksheet)wb.Worksheets[1];

                var ipa = d.EfficiencyGroups.FirstOrDefault(x => x.GasName == "IPA");
                var acetone = d.EfficiencyGroups.FirstOrDefault(x => x.GasName == "Acetone");

                if (ipa?.Efficiencies11 != null)
                {
                    for (int i = 0; i < ipa?.Efficiencies11.Count && i < 41; i++)
                    {
                        ws.Cells[2 + i, 11].Value = ipa.Efficiencies11[i];
                    }
                }

                if (acetone?.Efficiencies11 != null)
                {
                    for (int i = 0; i < acetone.Efficiencies11.Count && i < 41; i++)
                    {
                        ws.Cells[2 + i, 12].Value = acetone.Efficiencies11[i];
                    }
                }

                ws.Range["B6"].Value =
                    $"檢測日期：{d.TestingDate}\nTesting Date";

                ws.Range["B6"].WrapText = true;

                int idx = d.SelectedIndex;

                ws.Range["F8"].Value = d.LotFulls[idx];

                ExcelSignatureHelper.TryAddSignature(ws, "G26");

                wb.SaveAs(sfd.FileName);
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
}