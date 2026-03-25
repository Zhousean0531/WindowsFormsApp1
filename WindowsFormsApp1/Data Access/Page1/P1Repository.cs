using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1.Data_Access.Page1
{
    internal static class P1Repository
    {
        public static void Insert(P1Batch batch)
        {
            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    long batchId = InsertBatch(conn, tran, batch);

                    foreach (var sample in batch.Samples)
                    {
                        long sampleId = InsertSample(conn, tran, batchId, sample);
                        InsertEfficiencies(conn, tran, sampleId, sample);
                    }

                    tran.Commit();
                }
            }
        }
        private static long InsertBatch(SqlConnection conn, SqlTransaction tran, P1Batch batch)
        {
            string sql = @"
               INSERT INTO P1_Batch
                (ArrivalDate, TestingDate, ReportNo, Material,
                 QtyText, ParticleAnalysis, Concentration, Background, CreatedAt, Username)
                VALUES
                (@ArrivalDate, @TestingDate, @ReportNo, @Material,
                 @QtyText, @ParticleAnalysis, @Concentration, @Background, @CreatedAt, @Username);
            ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.Add("@ArrivalDate", SqlDbType.DateTime).Value = batch.ArrivalDate;
                cmd.Parameters.Add("@TestingDate", SqlDbType.DateTime).Value = batch.TestingDate;
                cmd.Parameters.Add("@ParticleAnalysis", SqlDbType.NVarChar).Value =
                    string.Join(",",
                        batch.ParticleSizePercentages
                            .Select(kv => $"{kv.Key}:{kv.Value}")
                    );
                cmd.Parameters.Add("@ReportNo", SqlDbType.NVarChar).Value =
                    (object)batch.ReportNo ?? DBNull.Value;

                cmd.Parameters.Add("@Material", SqlDbType.NVarChar).Value =
                    (object)batch.Material ?? DBNull.Value;
                cmd.Parameters.Add("@Concentration", SqlDbType.Decimal).Value =
                    (object)batch.Concentration ?? DBNull.Value;

                cmd.Parameters.Add("@Background", SqlDbType.Decimal).Value =
                    (object)batch.Background ?? DBNull.Value;
                cmd.Parameters.Add("@QtyText", SqlDbType.NVarChar).Value =
                    (object)batch.QtyText ?? DBNull.Value;

                cmd.Parameters.Add("@CreatedAt", SqlDbType.DateTime).Value = DateTime.Now;

                cmd.Parameters.Add("@Username", SqlDbType.NVarChar).Value =
                    (object)batch.Username ?? DBNull.Value;

                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }
        private static long InsertSample(SqlConnection conn, SqlTransaction tran, long batchId, P1Sample sample)
        {
            string sql = @"
                INSERT INTO P1_Sample
                (BatchId, InBatchNo, InternalBatchNo, Weight,
                 Density, PressureDrop, VOCIn, VOCOut, VOCOutgassing)
                VALUES
                (@BatchId, @InBatchNo, @InternalBatchNo, @Weight,
                 @Density, @PressureDrop, @VOCIn, @VOCOut, @VOCOutgassing);
                SELECT SCOPE_IDENTITY();
            ";
            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.Add("@BatchId", SqlDbType.BigInt).Value = batchId;

                cmd.Parameters.Add("@InBatchNo", SqlDbType.NVarChar).Value =
                    (object)sample.LotFull ?? DBNull.Value;

                cmd.Parameters.Add("@InternalBatchNo", SqlDbType.NVarChar).Value =
                    DBNull.Value;

                cmd.Parameters.Add("@Weight", SqlDbType.Decimal).Value =
                    (object)sample.Weight ?? DBNull.Value;

                cmd.Parameters.Add("@Density", SqlDbType.Decimal).Value =
                    (object)sample.Density ?? DBNull.Value;

                cmd.Parameters.Add("@PressureDrop", SqlDbType.Decimal).Value =
                    (object)sample.DeltaP ?? DBNull.Value;

                cmd.Parameters.Add("@VOCIn", SqlDbType.Decimal).Value =
                    (object)sample.VocIn ?? DBNull.Value;

                cmd.Parameters.Add("@VOCOut", SqlDbType.Decimal).Value =
                    (object)sample.VocOut ?? DBNull.Value;

                cmd.Parameters.Add("@VOCOutgassing", SqlDbType.Decimal).Value =
                    (object)sample.Outgassing ?? DBNull.Value;

                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }
        private static void InsertEfficiencies(SqlConnection conn, SqlTransaction tran, long sampleId, P1Sample sample)
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
                using (var cmd = new SqlCommand(sql, conn, tran))
                {
                    cmd.Parameters.Add("@SampleId", SqlDbType.BigInt).Value = sampleId;
                    cmd.Parameters.Add("@SequenceIndex", SqlDbType.Int).Value = i;

                    cmd.Parameters.Add("@EfficiencyValue", SqlDbType.Decimal).Value =
                        sample.Efficiencies[i].HasValue
                            ? Math.Round(sample.Efficiencies[i].Value, 1)
                            : (object)DBNull.Value;

                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}