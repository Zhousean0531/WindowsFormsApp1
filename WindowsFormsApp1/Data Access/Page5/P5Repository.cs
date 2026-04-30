using System;
using System.Data.SqlClient;

namespace WindowsFormsApp1.Data_Access.Page5
{
    internal class P5Repository
    {
        public static void Insert(Page5ExportData data, string rawEfficiencyText)
        {
            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    try
                    {
                        // ===== 1. 寫入 Batch =====
                        string batchSql = @"
                            INSERT INTO P5_Batch
                            (
                                ReportNo,
                                CylinderNo,
                                Customer,
                                FilterType,
                                TestDate,
                                ReCylinderNo,
                                CarbonLot,
                                UserName,
                                CreatedAt
                            )
                            OUTPUT INSERTED.Id
                            VALUES
                            (
                                @ReportNo,
                                @CylinderNo,
                                @Customer,
                                @FilterType,
                                @TestDate,
                                @ReCylinderNo,
                                @CarbonLot,
                                @UserName,
                                GETDATE()
                            )";

                        int batchId;

                        using (var cmd = new SqlCommand(batchSql, conn, tran))
                        {
                            cmd.Parameters.AddWithValue("@ReportNo", DbValue(data.ReportNo));
                            cmd.Parameters.AddWithValue("@CylinderNo", DbValue(data.CylinderNo));
                            cmd.Parameters.AddWithValue("@Customer", DbValue(data.Customer));
                            cmd.Parameters.AddWithValue("@FilterType", DbValue(data.FilterType));
                            cmd.Parameters.AddWithValue("@TestDate", DbValue(data.TestDate));
                            cmd.Parameters.AddWithValue("@ReCylinderNo", DbValue(data.ReCylinderNo));
                            cmd.Parameters.AddWithValue("@CarbonLot", DbValue(data.CarbonLot));
                            cmd.Parameters.AddWithValue("@UserName", DbValue(data.UserName));

                            batchId = Convert.ToInt32(cmd.ExecuteScalar());
                        }

                        // ===== 2. 寫入 Row =====
                        for (int i = 0; i < data.Rows.Count; i++)
                        {
                            var row = data.Rows[i];

                            string rowSql = @"
                                INSERT INTO P5_Row
                                (
                                    BatchId,
                                    RowNo,
                                    SN,
                                    Weight,
                                    Efficiency,

                                    ParticleIn,
                                    ParticleOut,
                                    ParticleDiff,

                                    IPAIn,
                                    IPAOut,
                                    IPADiff,

                                    AcetoneIn,
                                    AcetoneOut,
                                    AcetoneDiff,

                                    NonTargetIn,
                                    NonTargetOut,
                                    NonTargetDiff,

                                    TotalDiff,
                                    PressureDrop
                                )
                                VALUES
                                (
                                    @BatchId,
                                    @RowNo,
                                    @SN,
                                    @Weight,
                                    @Efficiency,

                                    @ParticleIn,
                                    @ParticleOut,
                                    @ParticleDiff,

                                    @IPAIn,
                                    @IPAOut,
                                    @IPADiff,

                                    @AcetoneIn,
                                    @AcetoneOut,
                                    @AcetoneDiff,

                                    @NonTargetIn,
                                    @NonTargetOut,
                                    @NonTargetDiff,

                                    @TotalDiff,
                                    @PressureDrop
                                )";

                            using (var cmd = new SqlCommand(rowSql, conn, tran))
                            {
                                cmd.Parameters.AddWithValue("@BatchId", batchId);
                                cmd.Parameters.AddWithValue("@RowNo", i + 1);

                                cmd.Parameters.AddWithValue("@SN", DbValue(row.SN));
                                cmd.Parameters.AddWithValue("@Weight", DbValue(row.Weight));
                                cmd.Parameters.AddWithValue("@Efficiency", DbValue(rawEfficiencyText));

                                cmd.Parameters.AddWithValue("@ParticleIn", Ctrl(row, 38));
                                cmd.Parameters.AddWithValue("@ParticleOut", Ctrl(row, 39));
                                cmd.Parameters.AddWithValue("@ParticleDiff", Ctrl(row, 40));

                                cmd.Parameters.AddWithValue("@IPAIn", Ctrl(row, 41));
                                cmd.Parameters.AddWithValue("@IPAOut", Ctrl(row, 42));
                                cmd.Parameters.AddWithValue("@IPADiff", Ctrl(row, 43));

                                cmd.Parameters.AddWithValue("@AcetoneIn", Ctrl(row, 44));
                                cmd.Parameters.AddWithValue("@AcetoneOut", Ctrl(row, 45));
                                cmd.Parameters.AddWithValue("@AcetoneDiff", Ctrl(row, 46));

                                cmd.Parameters.AddWithValue("@NonTargetIn", Ctrl(row, 47));
                                cmd.Parameters.AddWithValue("@NonTargetOut", Ctrl(row, 48));
                                cmd.Parameters.AddWithValue("@NonTargetDiff", Ctrl(row, 49));

                                cmd.Parameters.AddWithValue("@TotalDiff", Ctrl(row, 50));
                                cmd.Parameters.AddWithValue("@PressureDrop", Ctrl(row, 54));

                                cmd.ExecuteNonQuery();
                            }
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

        private static object DbValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return DBNull.Value;

            return value;
        }

        private static object Ctrl(Page5RowData row, int key)
        {
            if (row.ControlValues != null &&
                row.ControlValues.TryGetValue(key, out string value))
            {
                return DbValue(value);
            }

            return DBNull.Value;
        }
    }
}