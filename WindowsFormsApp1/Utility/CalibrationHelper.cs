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

        List<string> expiredList = new List<string>();
        List<string> warningList = new List<string>();

        DateTime today = DateTime.Today;

        foreach (CalibrationInfo info in infos)
        {
            DateTime expireDate = info.ExpireDate.Date;
            double daysDiff = (expireDate - today).TotalDays;

            if (daysDiff < 0)
            {
                expiredList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 已過期 " +
                    Math.Abs(Math.Floor(daysDiff)).ToString("0") +
                    " 天"
                );
            }
            else if (daysDiff == 0)
            {
                expiredList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 今天到期"
                );
            }
            else if (daysDiff > 0 && daysDiff <= 14)
            {
                warningList.Add(
                    info.InstrumentName + "：" +
                    expireDate.ToString("yyyy.MM.dd") +
                    " 即將到期，剩 " +
                    Math.Ceiling(daysDiff).ToString("0") +
                    " 天"
                );
            }
        }

        if (expiredList.Count > 0 || warningList.Count > 0)
        {
            string message = "";

            if (expiredList.Count > 0)
            {
                message += "【儀器已過期 / 今天到期】\r\n";
                message += string.Join("\r\n", expiredList);
                message += "\r\n\r\n";
            }

            if (warningList.Count > 0)
            {
                message += "【儀器校正提醒（14 天內）】\r\n";
                message += string.Join("\r\n", warningList);
                message += "\r\n\r\n";
            }

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
        }

        // 不管使用者是否更新，都繼續流程
        return true;
    }
}