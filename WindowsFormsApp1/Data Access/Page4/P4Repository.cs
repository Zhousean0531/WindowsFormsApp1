using System;
using System.Data.SQLite;

namespace WindowsFormsApp1.Data_Access.Page4
{
    public static class P4Repository
    {
        public static void Insert(P4Batch batch)
        {
            using (var conn = new SQLiteConnection(DbBootstrap.ConnStr))
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    long batchId = InsertBatch(conn, batch);

                    InsertLots(conn, batchId, batch);
                    InsertParticles(conn, batchId, batch);
                    InsertEfficiencies(conn, batchId, batch);

                    tran.Commit();
                }
            }
        }

        private static long InsertBatch(SQLiteConnection conn, P4Batch batch)
        {
            string sql = @"
            INSERT INTO P4_Batch
            (ReportNo,Material,MaterialNo,ArrivalDate,TestingDate,QtyText,CreatedAt,Username)
            VALUES
            (@ReportNo,@Material,@MaterialNo,@ArrivalDate,@TestingDate,@QtyText,@CreatedAt,@Username);
            SELECT last_insert_rowid();
            ";

            using (var cmd = new SQLiteCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@ReportNo", batch.ReportNo);
                cmd.Parameters.AddWithValue("@Material", batch.Material);
                cmd.Parameters.AddWithValue("@MaterialNo", batch.MaterialNo);
                cmd.Parameters.AddWithValue("@ArrivalDate", batch.ArrivalDate);
                cmd.Parameters.AddWithValue("@TestingDate", batch.TestingDate);
                cmd.Parameters.AddWithValue("@QtyText", batch.QtyText);
                cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@Username", batch.UserName);

                return (long)cmd.ExecuteScalar();
            }
        }

        private static void InsertLots(SQLiteConnection conn, long batchId, P4Batch batch)
        {
            string sql = @"
            INSERT INTO P4_Lot
            (BatchId,LotNo,LotFull,Weight,Density,VocIn,VocOut,DeltaP,Outgassing,IsSelected)
            VALUES
            (@BatchId,@LotNo,@LotFull,@Weight,@Density,@VocIn,@VocOut,@DeltaP,@Outgassing,@IsSelected);
            ";

            for (int i = 0; i < batch.Weights.Count; i++)
            {
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@BatchId", batchId);
                    cmd.Parameters.AddWithValue("@LotNo", batch.LotNos[i]);
                    cmd.Parameters.AddWithValue("@LotFull", batch.LotFulls[i]);

                    cmd.Parameters.AddWithValue("@Weight", batch.Weights[i]);
                    cmd.Parameters.AddWithValue("@Density", batch.Densities[i]);

                    cmd.Parameters.AddWithValue("@VocIn", batch.VocIns[i]);
                    cmd.Parameters.AddWithValue("@VocOut", batch.VocOuts[i]);

                    cmd.Parameters.AddWithValue("@DeltaP", batch.DeltaPs[i]);
                    cmd.Parameters.AddWithValue("@Outgassing", batch.OutgassingList[i]);

                    cmd.Parameters.AddWithValue("@IsSelected",
                        i == batch.SelectedIndex ? 1 : 0);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        private static void InsertParticles(SQLiteConnection conn, long batchId, P4Batch batch)
        {
            string sql = @"
            INSERT INTO P4_Particle
            (BatchId,SizeName,Percentage)
            VALUES
            (@BatchId,@SizeName,@Percentage);
            ";

            foreach (var p in batch.ParticleSizePercentages)
            {
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@BatchId", batchId);
                    cmd.Parameters.AddWithValue("@SizeName", p.Key);
                    cmd.Parameters.AddWithValue("@Percentage", p.Value);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        private static void InsertEfficiencies(SQLiteConnection conn, long batchId, P4Batch batch)
        {
            string sql = @"
            INSERT INTO P4_Efficiency
            (BatchId,GasName,Concentration,SequenceIndex,EfficiencyValue)
            VALUES
            (@BatchId,@GasName,@Concentration,@SequenceIndex,@EfficiencyValue);
            ";

            foreach (var g in batch.EfficiencyGroups)
            {
                for (int i = 0; i < g.Efficiencies11.Count; i++)
                {
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@BatchId", batchId);
                        cmd.Parameters.AddWithValue("@GasName", g.GasName);
                        cmd.Parameters.AddWithValue("@Concentration", g.Concentration);
                        cmd.Parameters.AddWithValue("@SequenceIndex", i);
                        cmd.Parameters.AddWithValue("@EfficiencyValue", g.Efficiencies11[i]);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}