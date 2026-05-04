using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WindowsFormsApp1;
using WindowsFormsApp1.Data_Access;

public static class CalibrationHelper
{
    public static bool CheckByInstrumentIds(List<int> ids)
    {
        List<CalibrationInfo> infos = InstrumentRepository.GetByIds(ids);

        return CheckAndHandleExpired(infos);
    }

    public static List<CalibrationInfo> GetCalibrationInfos(List<int> ids)
    {
        return InstrumentRepository.GetByIds(ids);
    }

    public static bool CheckAndHandleExpired(List<CalibrationInfo> infos)
    {
        if (infos == null || infos.Count == 0)
            return true;

        List<string> warningList = new List<string>();
        List<string> expiredWarningList = new List<string>();
        List<string> blockList = new List<string>();

        DateTime today = DateTime.Today;

        foreach (CalibrationInfo info in infos)
        {
            if (info == null)
                continue;

            DateTime expireDate = info.ExpireDate.Date;
            double daysDiff = (expireDate - today).TotalDays;

            // ==============================
            // 已過期超過 30 天：停止匯出
            // ==============================
            if (daysDiff < -30)
            {
                blockList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 已過期 " +
                    Math.Abs(Math.Floor(daysDiff)).ToString("0") +
                    " 天，超過 30 天，不可匯出"
                );
            }
            // ==============================
            // 已過期 1～30 天：提醒，但可以匯出
            // ==============================
            else if (daysDiff < 0)
            {
                expiredWarningList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 已過期 " +
                    Math.Abs(Math.Floor(daysDiff)).ToString("0") +
                    " 天"
                );
            }
            // ==============================
            // 今天到期：提醒，但可以匯出
            // ==============================
            else if (daysDiff == 0)
            {
                expiredWarningList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 今天到期"
                );
            }
            // ==============================
            // 30 天內到期：提醒，但可以匯出
            // ==============================
            else if (daysDiff > 0 && daysDiff <= 30)
            {
                warningList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 即將到期，剩 " +
                    Math.Ceiling(daysDiff).ToString("0") +
                    " 天"
                );
            }
            // ==============================
            // 超過 30 天後才到期：不提醒
            // ==============================
            else
            {
                continue;
            }
        }

        // =====================================================
        // 已過期超過 30 天：停止流程
        // =====================================================
        if (blockList.Count > 0)
        {
            string message = "";

            message += "【儀器校正已過期超過 30 天】\r\n";
            message += string.Join("\r\n", blockList);
            message += "\r\n\r\n";
            message += "請先更新儀器校正日期與有效日期。\r\n";
            message += "系統將停止匯出 DB 與匯出報告流程。";

            MessageBox.Show(
                message,
                "儀器校正有效日期逾期",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );

            using (CalibrationManagerForm form = new CalibrationManagerForm())
            {
                form.ShowDialog();
            }

            return false;
        }

        // =====================================================
        // 30 天內到期、今天到期、或已過期 30 天內：提醒但繼續
        // =====================================================
        if (warningList.Count > 0 || expiredWarningList.Count > 0)
        {
            string message = "";

            if (warningList.Count > 0)
            {
                message += "【儀器校正提醒（30 天內到期）】\r\n";
                message += string.Join("\r\n", warningList);
                message += "\r\n\r\n";
            }

            if (expiredWarningList.Count > 0)
            {
                message += "【儀器校正已到期提醒（尚未超過 30 天）】\r\n";
                message += string.Join("\r\n", expiredWarningList);
                message += "\r\n\r\n";
            }

            message += "儀器尚未超過停止匯出的期限，可以繼續匯出。\r\n";
            message += "是否要現在重新填寫校正日期與有效日期？";

            DialogResult result = MessageBox.Show(
                message,
                "儀器校正有效日期提醒",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result == DialogResult.Yes)
            {
                using (CalibrationManagerForm form = new CalibrationManagerForm())
                {
                    form.ShowDialog();
                }
            }

            return true;
        }

        return true;
    }
}