using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace WindowsFormsApp1.Data_Access
{
    public class InstrumentReportInfo
    {
        public string InstrumentNo { get; set; }
        public string InstrumentName { get; set; }
        public DateTime CalibrationDate { get; set; }
        public DateTime ExpireDate { get; set; }
    }

    public static class InstrumentRepository
    {
        private static string ResolveReportInstrumentNameByNo(string instrumentNo)
        {
            if (string.IsNullOrWhiteSpace(instrumentNo))
                return "";

            instrumentNo = instrumentNo.Trim().ToUpper();

            switch (instrumentNo)
            {
                case "QAD-API-01":
                    return "NH3 Analyzer/T201";

                case "QAD-API-05":
                    return "SO2 Analyzer/43i";

                case "QAD-PID-01":
                    return "Portable Handheld  VOC Monitor/ ppbRAE 3000";

                case "QAD-GT-03":
                    return "Particle Counter/Lasair Pro 310";

                case "QAD-MIT-02":
                    return "MiTAP/FT3+";

                case "QAD-MIT-03":
                    return "MiTAP/SFT3";

                case "QAD-GT-02":
                    return "Handheld Particle Counter/GT-324";

                case "QAD-IAQ-05":
                    return "Universal IAQ instrument/\r\ntesto 440dp";

                default:
                    return "";
            }
        }
        private static string ResolveDbInstrumentNameByNo(string instrumentNo)
        {
            if (string.IsNullOrWhiteSpace(instrumentNo))
                return "";

            instrumentNo = instrumentNo.Trim().ToUpper();

            switch (instrumentNo)
            {
                case "QAD-API-01":
                    return "Teledyne / T201";

                case "QAD-API-05":
                    return "Thermo Fisher / 43i";

                case "QAD-PID-01":
                    return "Honeywell / RAE3000";

                case "QAD-GT-03":
                    return "PMS / Lasair® Pro";

                case "QAD-MIT-02":
                    return "創控 / FT3+";

                case "QAD-MIT-03":
                    return "創控 / SFT3";

                case "QAD-GT-02":
                    return "Met One / GT-324";

                case "QAD-IAQ-05":
                    return "Testo / 440 dP";

                default:
                    return "";
            }
        }
        public static Dictionary<string, InstrumentReportInfo> GetByInstrumentNos(
    List<string> instrumentNos
)
        {
            Dictionary<string, InstrumentReportInfo> result =
                new Dictionary<string, InstrumentReportInfo>(StringComparer.OrdinalIgnoreCase);

            if (instrumentNos == null || instrumentNos.Count == 0)
                return result;

            foreach (string instrumentNo in instrumentNos)
            {
                if (string.IsNullOrWhiteSpace(instrumentNo))
                    continue;

                string no = instrumentNo.Trim();

                string dbInstrumentName = ResolveDbInstrumentNameByNo(no);
                if (string.IsNullOrWhiteSpace(dbInstrumentName))
                    continue;

                string reportInstrumentName = ResolveReportInstrumentNameByNo(no);
                if (string.IsNullOrWhiteSpace(reportInstrumentName))
                    continue;

                InstrumentReportInfo info = GetByInstrumentName(dbInstrumentName);
                if (info == null)
                    continue;

                info.InstrumentNo = no;

                // 很重要：報告要顯示指定名稱，不是 DB 名稱
                info.InstrumentName = reportInstrumentName;

                if (!result.ContainsKey(no))
                    result.Add(no, info);
            }

            return result;
        }
        // ===== 匯入 Excel 時使用 =====
        public static void UpsertByInstrumentName(
            string instrumentName,
            DateTime calibrationDate,
            DateTime expireDate,
            string updatedBy
        )
        {
            if (string.IsNullOrWhiteSpace(instrumentName))
                return;

            instrumentName = instrumentName.Trim();

            if (string.IsNullOrWhiteSpace(updatedBy))
                updatedBy = "SYSTEM";

            using (SqlConnection conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                using (SqlTransaction tran = conn.BeginTransaction(IsolationLevel.Serializable))
                {
                    try
                    {
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tran;

                            cmd.CommandText = @"
UPDATE [dbo].[Instruments]
SET
    [CalibrationDate] = @CalibrationDate,
    [ExpireDate] = @ExpireDate,
    [UpdatedAt] = GETDATE(),
    [UpdatedBy] = @UpdatedBy
WHERE LTRIM(RTRIM([InstrumentName])) = @InstrumentName;

IF @@ROWCOUNT = 0
BEGIN
    INSERT INTO [dbo].[Instruments]
    (
        [InstrumentName],
        [CalibrationDate],
        [ExpireDate],
        [CreatedAt],
        [UpdatedAt],
        [UpdatedBy]
    )
    VALUES
    (
        @InstrumentName,
        @CalibrationDate,
        @ExpireDate,
        GETDATE(),
        GETDATE(),
        @UpdatedBy
    );
END
";

                            cmd.Parameters.Add("@InstrumentName", SqlDbType.NVarChar, 200).Value = instrumentName;
                            cmd.Parameters.Add("@CalibrationDate", SqlDbType.DateTime).Value = calibrationDate;
                            cmd.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = expireDate;
                            cmd.Parameters.Add("@UpdatedBy", SqlDbType.NVarChar, 50).Value = updatedBy;

                            cmd.ExecuteNonQuery();
                        }

                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }
        }

        // ===== 報告匯出時使用：一次查多台儀器 =====
        public static Dictionary<string, InstrumentReportInfo> GetByInstrumentNames(
            List<string> instrumentNames
        )
        {
            Dictionary<string, InstrumentReportInfo> result =
                new Dictionary<string, InstrumentReportInfo>(StringComparer.OrdinalIgnoreCase);

            if (instrumentNames == null || instrumentNames.Count == 0)
                return result;

            List<string> cleanNames = new List<string>();

            foreach (string name in instrumentNames)
            {
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                string cleanName = name.Trim();

                if (!cleanNames.Contains(cleanName))
                    cleanNames.Add(cleanName);
            }

            if (cleanNames.Count == 0)
                return result;

            using (SqlConnection conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                List<string> paramNames = new List<string>();

                for (int i = 0; i < cleanNames.Count; i++)
                {
                    paramNames.Add("@Name" + i);
                }

                string sql = @"
SELECT
    [InstrumentName],
    [CalibrationDate],
    [ExpireDate]
FROM [dbo].[Instruments]
WHERE LTRIM(RTRIM([InstrumentName])) IN (" + string.Join(",", paramNames) + @");
";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    for (int i = 0; i < cleanNames.Count; i++)
                    {
                        cmd.Parameters.Add("@Name" + i, SqlDbType.NVarChar, 200).Value = cleanNames[i];
                    }

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            InstrumentReportInfo info = new InstrumentReportInfo();

                            info.InstrumentName = reader["InstrumentName"].ToString().Trim();

                            if (reader["CalibrationDate"] != DBNull.Value)
                                info.CalibrationDate = Convert.ToDateTime(reader["CalibrationDate"]);

                            if (reader["ExpireDate"] != DBNull.Value)
                                info.ExpireDate = Convert.ToDateTime(reader["ExpireDate"]);

                            if (!result.ContainsKey(info.InstrumentName))
                            {
                                result.Add(info.InstrumentName, info);
                            }
                        }
                    }
                }
            }

            return result;
        }

        // ===== 報告匯出時使用：查單台儀器 =====
        public static InstrumentReportInfo GetByInstrumentName(string instrumentName)
        {
            if (string.IsNullOrWhiteSpace(instrumentName))
                return null;

            instrumentName = instrumentName.Trim();

            using (SqlConnection conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                string sql = @"
SELECT TOP 1
    [InstrumentName],
    [CalibrationDate],
    [ExpireDate]
FROM [dbo].[Instruments]
WHERE LTRIM(RTRIM([InstrumentName])) = @InstrumentName;
";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@InstrumentName", SqlDbType.NVarChar, 200).Value = instrumentName;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (!reader.Read())
                            return null;

                        InstrumentReportInfo info = new InstrumentReportInfo();

                        info.InstrumentName = reader["InstrumentName"].ToString().Trim();

                        if (reader["CalibrationDate"] != DBNull.Value)
                            info.CalibrationDate = Convert.ToDateTime(reader["CalibrationDate"]);

                        if (reader["ExpireDate"] != DBNull.Value)
                            info.ExpireDate = Convert.ToDateTime(reader["ExpireDate"]);
                        return info;
                    }
                }
            }
        }
    }
}