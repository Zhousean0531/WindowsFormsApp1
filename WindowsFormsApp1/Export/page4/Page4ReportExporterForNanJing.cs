using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static SignatureHelper;
using Excel = Microsoft.Office.Interop.Excel;

public static class Page4ReportExporterForNanJing
{
    public static void Export(Page4ExportData d)
    {
        // ───── 條件判斷 ─────
        if (d.Material != "SG017_A" && d.Material != "SG017_D")
            return;

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterialForNanJing_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到加測範本");
            return;
        }

        DateTime arrivalDt = DateTime.Parse(d.ArrivalDate);
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";
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

                // ───── 抓效率群組 ─────
                var ipaGroup = d.EfficiencyGroups
                    .FirstOrDefault(g => g.GasName == "IPA");

                var acetoneGroup = d.EfficiencyGroups
                    .FirstOrDefault(g => g.GasName == "Acetone");

                if (ipaGroup?.Efficiencies11 != null)
                {
                    int count = ipaGroup.Efficiencies11.Count;
                    int rows = Math.Min(count, 41);
                    if (count <= 40)
                    {
                        MessageBox.Show("請確認IPA數值是否滿足");
                    }
                    for (int i = 0; i < rows; i++)
                    {
                        ws.Cells[2 + i, 11].Value = ipaGroup.Efficiencies11[i]; // K
                    }
                }


                if (acetoneGroup?.Efficiencies11 != null)
                {
                    int count = acetoneGroup.Efficiencies11.Count;
                    int rows = Math.Min(count, 41);
                    if (count <= 40)
                    {
                        MessageBox.Show("請確認Acetone數值是否滿足");
                    }
                    for (int i = 0; i < rows; i++)
                    {
                        ws.Cells[2 + i, 12].Value = acetoneGroup.Efficiencies11[i]; // L
                    }
                }


                // ───── B6：檢測日期（含換行） ─────
                ws.Range["B6"].Value =
                    $"檢測日期：{d.TestingDate}\nTesting Date";
                ws.Range["B6"].WrapText = true;

                // ───── F8：Lot ─────
                ws.Range["F8"].Value = d.Lot;

                // ───── 初始 / 40 分鐘效率 ─────
                if (ipaGroup?.Efficiencies11?.Count >= 11)
                {
                    ws.Range["F14"].Value = ipaGroup.Efficiencies11[0];
                    ws.Range["F15"].Value = ipaGroup.Efficiencies11[40];
                }

                if (acetoneGroup?.Efficiencies11?.Count >= 11)
                {
                    ws.Range["G14"].Value = acetoneGroup.Efficiencies11[0];
                    ws.Range["G15"].Value = acetoneGroup.Efficiencies11[40];
                }

                // ───── 插入簽名 ─────
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
