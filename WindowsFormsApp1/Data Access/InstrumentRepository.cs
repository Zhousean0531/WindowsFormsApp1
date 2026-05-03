using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace WindowsFormsApp1.Data_Access
{
    public static class InstrumentRepository
    {
        public static List<CalibrationInfo> GetByIds(List<int> ids)
        {
            var list = new List<CalibrationInfo>();

            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                string sql = $@"
                    SELECT Id, InstrumentName, CalibrationDate, ExpireDate
                    FROM Instruments
                    WHERE Id IN ({string.Join(",", ids)})";

                using (var cmd = new SqlCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new CalibrationInfo
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            InstrumentName = reader["InstrumentName"].ToString(),
                            CalibrationDate = Convert.ToDateTime(reader["CalibrationDate"]),
                            ExpireDate = Convert.ToDateTime(reader["ExpireDate"])
                        });
                    }
                }
            }

            return list;
        }
        public static List<CalibrationInfo> GetAll()
        {
            var list = new List<CalibrationInfo>();

            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                string sql = @"
            SELECT Id, InstrumentName, CalibrationDate, ExpireDate
            FROM Instruments
            ORDER BY Id";

                using (var cmd = new SqlCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new CalibrationInfo
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            InstrumentName = reader["InstrumentName"].ToString(),
                            CalibrationDate = reader["CalibrationDate"] == DBNull.Value
                                ? DateTime.Today
                                : Convert.ToDateTime(reader["CalibrationDate"]),
                            ExpireDate = reader["ExpireDate"] == DBNull.Value
                                ? DateTime.Today
                                : Convert.ToDateTime(reader["ExpireDate"])
                        });
                    }
                }
            }

            return list;
        }
        public static void Update(CalibrationInfo info)
        {
            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                string sql = @"
            UPDATE Instruments
            SET CalibrationDate = @cal,
                ExpireDate = @exp,
                UpdatedAt = GETDATE(),
                UpdatedBy = @user
            WHERE Id = @id";

                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@cal", System.Data.SqlDbType.DateTime).Value = info.CalibrationDate;
                    cmd.Parameters.Add("@exp", System.Data.SqlDbType.DateTime).Value = info.ExpireDate;
                    cmd.Parameters.Add("@user", System.Data.SqlDbType.NVarChar, 100).Value = Environment.UserName;
                    cmd.Parameters.Add("@id", System.Data.SqlDbType.Int).Value = info.Id;

                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}