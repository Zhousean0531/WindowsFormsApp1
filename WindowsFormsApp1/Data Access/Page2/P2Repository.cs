using System;
using System.Data.SQLite;

namespace WindowsFormsApp1.Data_Access.Page2
{
    internal static class P2Repository
    {
        public static void Insert(P2Batch batch)
        {
            using (var conn = new SQLiteConnection(DbBootstrap.ConnStr))
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    long batchId = InsertBatch(conn, batch);

                    foreach (var gas in batch.GasTests)
                    {
                        long gasId = InsertGasTest(conn, batchId, gas);

                        foreach (var sample in gas.Samples)
                        {
                            long sampleId = InsertSample(conn, gasId, sample);

                            InsertEfficiencies(conn, sampleId, sample);
                        }
                    }

                    tran.Commit();
                }
            }
        }
        private static long InsertBatch(SQLiteConnection conn, P2Batch batch)
        {
            string sql = @"
                INSERT INTO P2_Batch
                (ProductionDate, TestDate, WorkOrder, Material, MaterialBatchNo,MaterialNo,
                 TargetGsm, Glue, Speed, UpperTemp, LowerTemp,
                 Pressure, WindSpeed, CarbonLine, CreatedAt, Username)
                VALUES
                (@ProductionDate, @TestDate, @WorkOrder, @Material, @MaterialBatchNo,@MaterialNo,
                 @TargetGsm, @Glue, @Speed, @UpperTemp, @LowerTemp,
                 @Pressure, @WindSpeed, @CarbonLine, @CreatedAt, @Username);
                SELECT last_insert_rowid();
            ";
            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@ProductionDate", batch.ProductionDate);
                cmd.Parameters.AddWithValue("@TestDate", batch.TestDate);
                cmd.Parameters.AddWithValue("@WorkOrder", batch.WorkOrder);
                cmd.Parameters.AddWithValue("@Material", batch.Material);
                cmd.Parameters.AddWithValue("@MaterialBatchNo", batch.MaterialBatchNo);
                cmd.Parameters.AddWithValue("@MaterialNo", batch.MaterialNo);
                cmd.Parameters.AddWithValue("@TargetGsm", batch.TargetGsm);
                cmd.Parameters.AddWithValue("@Glue", batch.Glue);
                cmd.Parameters.AddWithValue("@Speed", batch.Speed);
                cmd.Parameters.AddWithValue("@UpperTemp", batch.UpperTemp);
                cmd.Parameters.AddWithValue("@LowerTemp", batch.LowerTemp);
                cmd.Parameters.AddWithValue("@Pressure", batch.Pressure);
                cmd.Parameters.AddWithValue("@WindSpeed", batch.WindSpeed);
                cmd.Parameters.AddWithValue("@CarbonLine", batch.CarbonLine);
                cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Username", batch.Username);
                return (long)cmd.ExecuteScalar();
            }
        }

        private static long InsertGasTest(SQLiteConnection conn, long batchId, P2GasTest gas)
        {
            string sql = @"
                INSERT INTO P2_GasTest
                (BatchId, GasName, Concentration, Background)
                VALUES
                (@BatchId, @GasName, @Concentration, @Background);
                SELECT last_insert_rowid();
            ";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);
                cmd.Parameters.AddWithValue("@GasName", gas.GasName);
                cmd.Parameters.AddWithValue("@Concentration", gas.Concentration);
                cmd.Parameters.AddWithValue("@Background", gas.Background);

                return (long)cmd.ExecuteScalar();
            }
        }

        private static long InsertSample(SQLiteConnection conn, long gasId, P2Sample sample)
        {
            string sql = @"
                INSERT INTO P2_Sample
                (GasTestId, Weight, PressureDrop, IsSelected)
                VALUES
                (@GasTestId, @Weight, @PressureDrop, @IsSelected);
                SELECT last_insert_rowid();
            ";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@GasTestId", gasId);
                cmd.Parameters.AddWithValue("@Weight", sample.Weight);
                cmd.Parameters.AddWithValue("@PressureDrop", sample.PressureDrop);
                cmd.Parameters.AddWithValue("@IsSelected", sample.IsSelected ? 1 : 0);

                return (long)cmd.ExecuteScalar();
            }
        }

        private static void InsertEfficiencies(SQLiteConnection conn, long sampleId, P2Sample sample)
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
                using (var cmd = new SQLiteCommand(sql, conn))
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
