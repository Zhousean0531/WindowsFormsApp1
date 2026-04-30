using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Data.SqlClient;
using WindowsFormsApp1.Data_Access;
public static class CalibrationHelper
{
    public static List<CalibrationInfo> GetCalibrationInfos(List<int> ids)
    {
        return InstrumentRepository.GetByIds(ids);
    }

    public static bool CheckAndHandleExpired(List<CalibrationInfo> infos)
    {
        foreach (var info in infos)
        {
            if (info.ExpireDate < DateTime.Today)
            {
                MessageBox.Show(
                    $"{info.InstrumentName} 已過期 ({info.ExpireDate:yyyy.MM.dd})",
                    "儀器過期",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                return false; // 👉 直接擋掉
            }
        }

        return true;
    }
}