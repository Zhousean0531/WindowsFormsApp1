using System;
using System.Collections.Generic;
using System.Windows.Forms;

public static class CalibrationHelper
{
    public static bool CheckCalibrationByColumns(List<int> columns)
    {
        try
        {
            string filePath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "總表.xlsx"
            );

            if (!System.IO.File.Exists(filePath))
            {
                MessageBox.Show("找不到總表.xlsx，無法檢查儀器校正。");
                return true; // 不擋匯出
            }

            using (var wb = new ClosedXML.Excel.XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("儀器校正");
                if (ws == null)
                {
                    MessageBox.Show("找不到「儀器校正」工作表。");
                    return true; // 不擋匯出
                }

                DateTime today = DateTime.Today;
                var warnList = new List<string>();

                foreach (int col in columns)
                {
                    string instName = ws.Cell(1, col).GetString().Trim(); // 儀器名稱
                    string exp = ws.Cell(3, col).GetString().Trim(); // 有效日期（第 3 列）

                    if (string.IsNullOrWhiteSpace(instName))
                        continue;
                    if (!DateTime.TryParse(exp, out DateTime expireDate))
                        continue;

                    double daysDiff = (expireDate - today).TotalDays;
                    // 正負 14 天內都提醒
                    // 只在「有效日期距離今天的差異在 ±14 天內」才提醒
                    if ((daysDiff > 0 && daysDiff <= 14) || daysDiff <= 0)
                    {
                        string desc;

                        if (daysDiff > 0) // 即將到期（未過期）
                        {
                            desc = $"{instName}：{expireDate:yyyy.MM.dd} 即將到期，剩 {Math.Ceiling(daysDiff)} 天";
                        }
                        else if (daysDiff < 0) // 已過期但不超過14天
                        {
                            desc = $"{instName}：{expireDate:yyyy.MM.dd} 已過期 {Math.Abs(Math.Floor(daysDiff))} 天";
                        }
                        else // daysDiff == 0
                        {
                            return true;
                        }

                        warnList.Add(desc);
                    }
                }

                if (warnList.Count > 0)
                {
                    MessageBox.Show(
                        "【儀器校正提醒（有效日期 ±14 天）】\r\n\r\n" +
                        string.Join("\r\n", warnList),
                        "校正有效日期提醒",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("校正檢查錯誤：" + ex.Message);
        }
        return true;
    }
}
