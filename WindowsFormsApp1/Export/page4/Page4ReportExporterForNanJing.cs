using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WindowsFormsApp1.Data_Access.Page4;
using static SignatureHelper;

public static class Page4ReportExporterForNanJing
{
    public static bool Export(P4Batch d)
    {
        if (d.Material != "SG017_A" && d.Material != "SG017_D")
            return true;

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterialForNanJing_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到南京範本");
            return false;
        }

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";
            sfd.FileName = d.ReportNo + "_" + d.Material + "_加測.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return false;

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                app = ExcelInteropHelper.CreateApplication();
                wb = ExcelInteropHelper.OpenWorkbook(app, templatePath);
                ws = (Excel.Worksheet)wb.Worksheets[1];

                // ⭐ 統一抓一次（重點修正）
                P4EfficiencyGroup ipa = null;
                P4EfficiencyGroup acetone = null;

                foreach (var g in d.EfficiencyGroups)
                {
                    if (g.GasName == "IPA")
                        ipa = g;
                    else if (g.GasName == "Acetone")
                        acetone = g;
                }

                // ===== 曲線（完全照你原本）=====
                if (ipa != null)
                {
                    for (int i = 0; i < ipa.Efficiencies11.Count && i < 41; i++)
                        ws.Cells[2 + i, 11].Value = ipa.Efficiencies11[i];
                }

                if (acetone != null)
                {
                    for (int i = 0; i < acetone.Efficiencies11.Count && i < 41; i++)
                        ws.Cells[2 + i, 12].Value = acetone.Efficiencies11[i];
                }

                // ===== 日期 =====
                ws.Range["B6"].Value = "檢測日期：" + d.TestingDate;

                // ===== 批號 =====
                var selected = d.Rows.FirstOrDefault(x => x.IsSelected);

                if (selected != null)
                {
                    ws.Range["F8"].Value = selected.LotFull;
                }

                // ===== F14~G15（摘要）=====
                ws.Range["F14"].Value = ipa?.Efficiencies11.FirstOrDefault().ToString("F1") ?? "N.D.";
                ws.Range["F15"].Value = ipa?.Efficiencies11.LastOrDefault().ToString("F1") ?? "N.D.";

                ws.Range["G14"].Value = acetone?.Efficiencies11.FirstOrDefault().ToString("F1") ?? "N.D.";
                ws.Range["G15"].Value = acetone?.Efficiencies11.LastOrDefault().ToString("F1") ?? "N.D.";

                ExcelSignatureHelper.TryAddSignature(ws, "G26");

                ExcelInteropHelper.SaveAs(wb, sfd.FileName);
            }
            finally
            {
                ExcelInteropHelper.CloseWorkbook(wb, false);
                ExcelInteropHelper.Quit(app);

                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }

        return true;
    }
}
