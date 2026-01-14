using ClosedXML.Excel;
using System;
using System.IO;
using System.Windows.Forms;

public static class Page5ReportExporter
{
    public static void Export(
        Page5ExportData data,
        Page5LookupResult lookupResult,
        string efficiencyType   // "MA" / "MB" / "MC"
    )
    {
        // ===== 1️⃣ 檢查模板是否存在 =====
        string templatePath = Path.Combine(
            Application.StartupPath,
            "CYLReport.xlsx"
        );

        if (!File.Exists(templatePath))
        {
            MessageBox.Show(
                "找不到報告模板 CYLReport.xlsx，請確認檔案是否放在程式目錄。",
                "錯誤",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return;
        }

        // ===== 2️⃣ 組預設檔名 =====
        string defaultFileName =
            $"{data.ReportNo}_{data.CylinderNo}-{efficiencyType}.xlsx";

        // ===== 3️⃣ 讓使用者選擇存檔位置 =====
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
            sfd.FileName = defaultFileName;

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                // ===== 4️⃣ 複製模板成新檔 =====
                File.Copy(templatePath, sfd.FileName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "建立報告檔案時發生錯誤：\n" + ex.Message,
                    "錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}
