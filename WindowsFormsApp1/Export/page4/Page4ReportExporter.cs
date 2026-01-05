using ClosedXML.Excel;
using Google.Protobuf.WellKnownTypes;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class Page4ReportExporter
{
    public static void Export(Page4ExportData d)
    {
        if (d == null || d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有可匯出的效率資料");
            return;
        }

        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到原料模板");
            return;
        }

        foreach (var g in d.EfficiencyGroups)
        {
            string savePath;
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                sfd.FileName = $"{d.ReportNo}_{g.GasName}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    continue;

                savePath = sfd.FileName;
            }

            File.Copy(templatePath, savePath, true);

            using (var wb = new XLWorkbook(savePath))
            {
                var ws = wb.Worksheet("濾網原料報告");

                int idx = d.SelectedIndex;
                if (idx < 0 || idx >= g.Efficiencies11.Count)
                {
                    MessageBox.Show($"氣體 {g.GasName} 的效率資料與壓損索引不符");
                    continue;
                }

                // 基本資料
                ws.Cell("C4").Value = d.ReportNo;
                ws.Cell("C5").Value = d.ArrivalDate;
                ws.Cell("C6").Value = d.TestingDate;
                ws.Cell("E5").Value = d.Material;
                ws.Cell("E6").Value = d.QtyText;

                // 批次資料
                const int COL_FIRST = 3;
                for (int i = 0; i < d.LotFulls.Count; i++)
                {
                    int col = COL_FIRST + i;
                    ws.Cell(10, col).Value = d.LotFulls[i];
                    ws.Cell(13, col).Value = d.Densities[i];
                    ws.Cell(14, col).Value = d.DeltaPs[i];
                    ws.Cell(15, col).Value = d.VocIns[i];
                    ws.Cell(16, col).Value = d.VocOuts[i];
                    ws.Cell(17, col).Value = d.OutgassingList[i];

                    ws.Cell(18, col).Value =
                        (i == idx) ? g.Eff0 : "N.D.";
                }
                wb.Save();
            }
        }
    }
}
