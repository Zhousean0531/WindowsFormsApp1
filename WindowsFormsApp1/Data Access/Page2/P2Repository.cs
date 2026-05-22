using System;
using System.Data.SqlClient;
using System.Linq;

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
                (ProductionDate, TestDate, WorkOrder, Material, MaterialNo,BatchNo,
                 TargetGsm, Glue, Speed, UpperTemp, LowerTemp,
                 Pressure, WindSpeed, CarbonLine,
                 ReportNo,  FilterSize,
                 CreatedAt, Username)
                VALUES
                (@ProductionDate, @TestDate, @WorkOrder, @Material, @MaterialNo,@BatchNo,
                 @TargetGsm, @Glue, @Speed, @UpperTemp, @LowerTemp,
                 @Pressure, @WindSpeed, @CarbonLine,
                 @ReportNo, @FilterSize,
                 @CreatedAt, @Username);
                SELECT SCOPE_IDENTITY();
                ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@ProductionDate", DbValue(batch.ProductionDate));
                cmd.Parameters.AddWithValue("@TestDate", DbValue(batch.TestDate));
                cmd.Parameters.AddWithValue("@WorkOrder", DbValue(batch.WorkOrder));
                cmd.Parameters.AddWithValue("@Material", DbValue(batch.Material));
                cmd.Parameters.AddWithValue("@MaterialNo", DbValue(batch.MaterialNo));
                cmd.Parameters.AddWithValue("@BatchNo", DbValue(batch.BatchNo));
                cmd.Parameters.AddWithValue("@TargetGsm", DbValue(batch.TargetGsm));
                cmd.Parameters.AddWithValue("@Glue", DbValue(batch.Glue));
                cmd.Parameters.AddWithValue("@Speed", DbValue(batch.Speed));
                cmd.Parameters.AddWithValue("@UpperTemp", DbValue(batch.UpperTemp));
                cmd.Parameters.AddWithValue("@LowerTemp", DbValue(batch.LowerTemp));
                cmd.Parameters.AddWithValue("@Pressure", DbValue(batch.Pressure));
                cmd.Parameters.AddWithValue("@WindSpeed", DbValue(batch.WindSpeed));
                cmd.Parameters.AddWithValue("@CarbonLine", DbValue(batch.CarbonLine));
                cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now);
                cmd.Parameters.AddWithValue("@Username", DbValue(batch.Username));
                cmd.Parameters.AddWithValue("@ReportNo", DbValue(batch.ReportNo));
                cmd.Parameters.AddWithValue("@FilterSize", DbValue(batch.FilterSize));
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
                cmd.Parameters.AddWithValue("@GasName", DbValue(gas.GasName));
                cmd.Parameters.AddWithValue("@Concentration", DbValue(gas.Concentration));
                cmd.Parameters.AddWithValue("@Background", DbValue(gas.Background));

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
                cmd.Parameters.AddWithValue("@Weight", DbValue(sample.Weight));
                cmd.Parameters.AddWithValue("@PressureDrop", DbValue(sample.PressureDrop));
                cmd.Parameters.AddWithValue("@IsSelected", sample.IsSelected);

                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }

        private static void InsertEfficiencies(SqlConnection conn, SqlTransaction tran, long sampleId, P2Sample sample)
        {
            bool hasPointEfficiencies =
                sample.EfficiencyPoints != null && sample.EfficiencyPoints.Count > 0;

            bool hasListEfficiencies =
                sample.Efficiencies != null && sample.Efficiencies.Count > 0;

            if (!hasPointEfficiencies && !hasListEfficiencies)
                return;

            string sql = @"
                INSERT INTO P2_Efficiency
                (SampleId, SequenceIndex, EfficiencyValue)
                VALUES
                (@SampleId, @SequenceIndex, @EfficiencyValue);
            ";

            if (hasPointEfficiencies)
            {
                foreach (var kv in sample.EfficiencyPoints.OrderBy(x => x.Key))
                {
                    using (var cmd = new SqlCommand(sql, conn, tran))
                    {
                        cmd.Parameters.AddWithValue("@SampleId", sampleId);
                        cmd.Parameters.AddWithValue("@SequenceIndex", kv.Key);
                        cmd.Parameters.AddWithValue("@EfficiencyValue", kv.Value);

                        cmd.ExecuteNonQuery();
                    }
                }

                return;
            }

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

        private static object DbValue(object value)
        {
            if (value == null)
                return DBNull.Value;

            return value;
        }
    }
}
