using System;
using System.Data.SQLite;

namespace WindowsFormsApp1.Data_Access.Page1
{
    internal static class P1Repository
    {
        public static void Insert(P1Batch batch)
        {
            using (var conn = new SQLiteConnection(DbBootstrap.ConnStr))
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    long batchId = InsertBatch(conn, batch);

                    foreach (var sample in batch.Samples)
                    {
                        long sampleId = InsertSample(conn, batchId, sample);
                        InsertEfficiencies(conn, sampleId, sample);
                    }

                    tran.Commit();
                }
            }
        }

        private static long InsertBatch(SQLiteConnection conn, P1Batch batch)
        {
            string sql = @"
                INSERT INTO P1_Batch
                (ArrivalDate, TestDate, ReportNo, MaterialType,
                 Quantity, ParticleAnalysis, CreatedAt, Username)
                VALUES
                (@ArrivalDate, @TestDate, @ReportNo, @MaterialType,
                 @Quantity, @ParticleAnalysis, @CreatedAt, @Username);
                SELECT last_insert_rowid();
            ";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@ArrivalDate", batch.ArrivalDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@TestDate", batch.TestingDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@ReportNo", batch.ReportNo);
                cmd.Parameters.AddWithValue("@MaterialType", batch.Material);
                cmd.Parameters.AddWithValue("@Quantity", batch.QtyText);
                cmd.Parameters.AddWithValue("@ParticleAnalysis", string.Join(",", batch.ParticleSizePercentages));
                cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Username", batch.Username);

                return (long)cmd.ExecuteScalar();
            }
        }

        private static long InsertSample(SQLiteConnection conn, long batchId, P1Sample sample)
        {
            string sql = @"
                INSERT INTO P1_Sample
                (GasTestId, InBatchNo, InternalBatchNo, Weight,
                 Density, PressureDrop, VOCIn, VOCOut, VOCOutgassing)
                VALUES
                (@GasTestId, @InBatchNo, @InternalBatchNo, @Weight,
                 @Density, @PressureDrop, @VOCIn, @VOCOut, @VOCOutgassing);
                SELECT last_insert_rowid();
            ";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@GasTestId", batchId);
                cmd.Parameters.AddWithValue("@InBatchNo", sample.LotFull);
                cmd.Parameters.AddWithValue("@InternalBatchNo", "");
                cmd.Parameters.AddWithValue("@Weight", sample.Weight);
                cmd.Parameters.AddWithValue("@Density", sample.Density);
                cmd.Parameters.AddWithValue("@PressureDrop", sample.DeltaP);
                cmd.Parameters.AddWithValue("@VOCIn", sample.VocIn);
                cmd.Parameters.AddWithValue("@VOCOut", sample.VocOut);
                cmd.Parameters.AddWithValue("@VOCOutgassing", sample.Outgassing);

                return (long)cmd.ExecuteScalar();
            }
        }

        private static void InsertEfficiencies(SQLiteConnection conn, long sampleId, P1Sample sample)
        {
            if (sample.Efficiencies == null || sample.Efficiencies.Count == 0)
                return;

            string sql = @"
                INSERT INTO P1_Efficiency
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