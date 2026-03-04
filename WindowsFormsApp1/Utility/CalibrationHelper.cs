using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

public static class CalibrationHelper
{
    public static bool CheckCalibrationByColumns(List<int> columns)
    {
        try
        {
            string filePath = Path.Combine(
            Application.StartupPath,
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
    public static List<CalibrationInfo> GetCalibrationInfosByColumns(List<int> columns)
    {
        var result = new List<CalibrationInfo>();

        try
        {
            string filePath = Path.Combine(
                Application.StartupPath,
                "總表.xlsx"
            );

            if (!File.Exists(filePath))
                return result;

            using (var wb = new ClosedXML.Excel.XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("儀器校正");
                if (ws == null)
                    return result;

                foreach (int col in columns)
                {
                    // 第 1 列：儀器名稱
                    string instName = ws.Cell(1, col).GetString().Trim();
                    if (string.IsNullOrWhiteSpace(instName))
                        continue;

                    // 第 2 列：校正日期
                    string calText = ws.Cell(2, col).GetString().Trim();
                    DateTime calDate;
                    if (!DateTime.TryParse(calText, out calDate))
                        continue;

                    // 第 3 列：有效日期
                    string expText = ws.Cell(3, col).GetString().Trim();
                    DateTime expDate;
                    if (!DateTime.TryParse(expText, out expDate))
                        continue;

                    result.Add(new CalibrationInfo
                    {
                        InstrumentName = instName,
                        CalibrationDate = calDate,
                        ExpireDate = expDate
                    });
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "讀取儀器校正資料時發生錯誤：\r\n" + ex.Message,
                "校正資料錯誤",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }

        return result;
    }

}
public class CalibrationInfo
{
    public string InstrumentName { get; set; }
    public DateTime CalibrationDate { get; set; }
    public DateTime ExpireDate { get; set; }
}

