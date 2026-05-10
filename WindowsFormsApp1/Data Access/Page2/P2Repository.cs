using System;
using System.Data.SqlClient;

namespace WindowsFormsApp1.Data_Access.Page2
{
    internal static class P2Repository
    {
        public static void Insert(P2Batch batch)
        {
            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    long batchId = InsertBatch(conn, tran, batch);

                    foreach (var gas in batch.GasTests)
                    {
                        long gasId = InsertGasTest(conn, tran, batchId, gas);

                        foreach (var sample in gas.Samples)
                        {
                            long sampleId = InsertSample(conn, tran, gasId, sample);
                            InsertEfficiencies(conn, tran, sampleId, sample);
                        }
                    }

                    tran.Commit();
                }
            }
        }

        private static long InsertBatch(SqlConnection conn, SqlTransaction tran, P2Batch batch)
        {
            string sql = @"
                INSERT INTO P2_Batch
                (ProductionDate, TestDate, WorkOrder, Material, MaterialNo,
                 TargetGsm, Glue, Speed, UpperTemp, LowerTemp,
                 Pressure, WindSpeed, CarbonLine,
                 ReportNo,  FilterSize,
                 CreatedAt, Username)
                VALUES
                (@ProductionDate, @TestDate, @WorkOrder, @Material, @MaterialNo,
                 @TargetGsm, @Glue, @Speed, @UpperTemp, @LowerTemp,
                 @Pressure, @WindSpeed, @CarbonLine,
                 @ReportNo, @FilterSize,
                 @CreatedAt, @Username);
                SELECT SCOPE_IDENTITY();
                ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@ProductionDate", batch.ProductionDate);
                cmd.Parameters.AddWithValue("@TestDate", batch.TestDate);
                cmd.Parameters.AddWithValue("@WorkOrder", batch.WorkOrder);
                cmd.Parameters.AddWithValue("@Material", batch.Material);
                cmd.Parameters.AddWithValue("@MaterialNo", batch.MaterialNo);
                cmd.Parameters.AddWithValue("@TargetGsm", batch.TargetGsm);
                cmd.Parameters.AddWithValue("@Glue", batch.Glue);
                cmd.Parameters.AddWithValue("@Speed", batch.Speed);
                cmd.Parameters.AddWithValue("@UpperTemp", batch.UpperTemp);
                cmd.Parameters.AddWithValue("@LowerTemp", batch.LowerTemp);
                cmd.Parameters.AddWithValue("@Pressure", batch.Pressure);
                cmd.Parameters.AddWithValue("@WindSpeed", batch.WindSpeed);
                cmd.Parameters.AddWithValue("@CarbonLine", batch.CarbonLine);
                cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now);
                cmd.Parameters.AddWithValue("@Username", batch.Username);
                cmd.Parameters.AddWithValue("@ReportNo", (object)batch.ReportNo ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@FilterSize", (object)batch.FilterSize ?? DBNull.Value);
                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }

        private static long InsertGasTest(SqlConnection conn, SqlTransaction tran, long batchId, P2GasTest gas)
        {
            string sql = @"
                INSERT INTO P2_GasTest
                (BatchId, GasName, Concentration, Background)
                VALUES
                (@BatchId, @GasName, @Concentration, @Background);
                SELECT SCOPE_IDENTITY();
            ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);
                cmd.Parameters.AddWithValue("@GasName", gas.GasName);
                cmd.Parameters.AddWithValue("@Concentration", gas.Concentration);
                cmd.Parameters.AddWithValue("@Background", gas.Background);

                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }

        private static long InsertSample(SqlConnection conn, SqlTransaction tran, long gasId, P2Sample sample)
        {
            string sql = @"
                INSERT INTO P2_Sample
                (GasTestId, Weight, PressureDrop, IsSelected)
                VALUES
                (@GasTestId, @Weight, @PressureDrop, @IsSelected);
                SELECT SCOPE_IDENTITY();
            ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@GasTestId", gasId);
                cmd.Parameters.AddWithValue("@Weight", sample.Weight);
                cmd.Parameters.AddWithValue("@PressureDrop", sample.PressureDrop);
                cmd.Parameters.AddWithValue("@IsSelected", sample.IsSelected);

                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }

        private static void InsertEfficiencies(SqlConnection conn, SqlTransaction tran, long sampleId, P2Sample sample)
        {
            if (sample.Efficiencies == null || sample.Efficiencies.Count == 0)
                return;

            string sql = @"
                INSERT INTO P2_Efficiency
                (SampleId, SequenceIndex, EfficiencyValue)
                VALUES
                (@SampleId, @SequenceIndex, @EfficiencyValue);
            ";

            for (int i = 0; i < sample.Efficiencies.Count; i++)
            {
                using (var cmd = new SqlCommand(sql, conn, tran))
                {
                    cmd.Parameters.AddWithValue("@SampleId", sampleId);
                    cmd.Parameters.AddWithValue("@SequenceIndex", i);
                    cmd.Parameters.AddWithValue("@EfficiencyValue", sample.Efficiencies[i]);

                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}