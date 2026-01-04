using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public static class Page2ReportExporter
{
    public static void Export(Page2ExportData d)
    {
        if (d == null)
        {
            return;
        }
        if (d.EfficiencyGroups == null)
        {
            MessageBox.Show("EfficiencyGroups 是 null");
            return;
        }

        if (d.PressureDrops == null)
        {
            MessageBox.Show("PressureDrops 是 null");
            return;
        }

        foreach (var g in d.EfficiencyGroups)
        {
            if (g.Efficiencies11 == null)
            {
                MessageBox.Show($"氣體 {g.GasName} 的 Efficiencies11 是 null");
                return;
            }
        }
        if (d == null || d.EfficiencyGroups == null || d.EfficiencyGroups.Count == 0)
        {
            MessageBox.Show("沒有可匯出的效率資料");
            return;
        }

        // ─────────────────────────────
        // (A) 固定 QC_semi 模板路徑
        // ─────────────────────────────
        string templatePath = Path.Combine(
            Application.StartupPath,   // 程式所在資料夾
            "QC_SemiFinished_Template.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_semi.xlsx 模板");
            return;
        }

        // ─────────────────────────────
        // (B) 逐一氣體輸出
        // ─────────────────────────────
        foreach (var g in d.EfficiencyGroups)
        {
            string savePath;
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                sfd.FileName = $"{d.ReportNo}_{g.GasName}.xlsx";

                if (sfd.ShowDialog() != DialogResult.OK)
                    continue; // 多氣體時，略過這一顆

                savePath = sfd.FileName;
            }

            File.Copy(templatePath, savePath, true);

            using (var wb = new XLWorkbook(savePath))
            {
                // ⚠️ 請確認 QC_semi.xlsx 的工作表名稱
                
                var ws = wb.Worksheet("濾網半成品報告");

                int idx = d.SelectedIndex;

                if (idx < 0 || idx >= g.Efficiencies11.Count)
                {
                    MessageBox.Show($"氣體 {g.GasName} 的效率資料與壓損索引不符");
                    continue;
                }

                // ─────────────────────────────
                // (C) 基本資料（依 QC_semi 樣板）
                // ─────────────────────────────
                ws.Cell("C6").Value = d.ReportNo; //報告編號
                ws.Cell("F7").Value = d.ProductDisplay;//產品名稱
                ws.Cell("F8").Value = d.FilterSize;//尺寸
                ws.Cell("F9").Value = d.OrderDisplay; //產品編號
                ws.Cell("H6").Value = d.TestDate; //檢測日期
                ws.Cell("F11").Value = g.GasName; //測試氣體
                if (idx >= 0 && idx < d.TestWeights.Count)//測試品重量
                {
                    ws.Cell("H10").Value = d.TestWeights[idx];
                }
                ws.Cell("F12").Value = g.Concentration;//背景濃度
                ws.Cell("F13").Value = d.Wind;//測試風速
                // (D) 壓損 & 對應效率
                // ─────────────────────────────
                ws.Cell("F14").Value = d.PressureDrops[idx]; // 樣品壓損
                ws.Cell("F16").Value = g.Efficiencies11[idx]; // 樣品效率
                // ─────────────────────────────
                // (E) 11 點效率明細
                // 假設從第 20 列開始
                // ─────────────────────────────
                int startRow = 1;
                for (int i = 0; i < g.Efficiencies11.Count; i++)
                {
                    ws.Cell(startRow + i, 11).Value = g.Efficiencies11[i];
                }

                wb.Save();
            }
        }
    }
}
