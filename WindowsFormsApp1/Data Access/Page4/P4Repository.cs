using System;
using System.Data.SqlClient;
using WindowsFormsApp1.Data_Access.Page4;

public static class P4Repository
{
    public static void Insert(P4Batch batch)
    {
        using (var conn = DbBootstrap.GetConnection())
        {
            conn.Open();
            using (var tran = conn.BeginTransaction())
            {
                long batchId = InsertBatch(conn, tran, batch);
                InsertLots(conn, tran, batchId, batch);
                InsertParticles(conn, tran, batchId, batch);
                InsertEfficiencies(conn, tran, batchId, batch);
                tran.Commit();
            }
        }
    }

    private static long InsertBatch(SqlConnection conn, SqlTransaction tran, P4Batch b)
    {
        string sql = @"
    INSERT INTO P4_Batch
    (ReportNo, Material, MaterialNo, ArrivalDate, TestingDate, QtyText, CreatedAt, Username,
     Moisture, Butane, Ash)
    VALUES
    (@ReportNo, @Material, @MaterialNo, @ArrivalDate, @TestingDate, @QtyText, @CreatedAt, @Username,
     @Moisture, @Butane, @Ash);
    SELECT SCOPE_IDENTITY();";

        using (var cmd = new SqlCommand(sql, conn, tran))
        {
            cmd.Parameters.AddWithValue("@ReportNo", b.ReportNo);
            cmd.Parameters.AddWithValue("@Material", b.Material);
            cmd.Parameters.AddWithValue("@MaterialNo", b.MaterialNo);
            cmd.Parameters.AddWithValue("@ArrivalDate", b.ArrivalDate);
            cmd.Parameters.AddWithValue("@TestingDate", b.TestingDate);
            cmd.Parameters.AddWithValue("@QtyText", b.QtyText);
            cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now);
            cmd.Parameters.AddWithValue("@Username", b.UserName);

            cmd.Parameters.AddWithValue("@Moisture", (object)b.Moisture ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Butane", (object)b.Butane ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Ash", (object)b.Ash ?? DBNull.Value);

            return Convert.ToInt64(cmd.ExecuteScalar());
        }
    }
    private static void InsertLots(SqlConnection conn, SqlTransaction tran, long batchId, P4Batch b)
    {
        string sql = @"
        INSERT INTO P4_Lot
        (BatchId, LotNo, LotFull, Weight, Density, VocIn, VocOut, DeltaP, Outgassing, IsSelected)
        VALUES
        (@BatchId, @LotNo, @LotFull, @Weight, @Density, @VocIn, @VocOut, @DeltaP, @Outgassing, @IsSelected);";

        foreach (var r in b.Rows)
        {
            using (var cmd = new SqlCommand(sql, conn, tran)) 
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);
                cmd.Parameters.AddWithValue("@LotNo", r.LotNo);
                cmd.Parameters.AddWithValue("@LotFull", r.LotFull);
                cmd.Parameters.AddWithValue("@Weight", r.Weight);
                cmd.Parameters.AddWithValue("@Density", r.Density);
                cmd.Parameters.AddWithValue("@VocIn", r.VocIn);
                cmd.Parameters.AddWithValue("@VocOut", r.VocOut);
                cmd.Parameters.AddWithValue("@DeltaP", r.DeltaP);
                cmd.Parameters.AddWithValue("@Outgassing", r.Outgassing);
                cmd.Parameters.AddWithValue("@IsSelected", r.IsSelected);

                cmd.ExecuteNonQuery();
            }
        }
    }
    private static void InsertParticles(SqlConnection conn, SqlTransaction tran, long batchId, P4Batch b)
    {
        string sql = @"
    INSERT INTO P4_Particle
    (BatchId, SizeName, Percentage)
    VALUES
    (@BatchId, @SizeName, @Percentage);";

        foreach (var p in b.ParticleSizePercentages)
        {
            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);
                cmd.Parameters.AddWithValue("@SizeName", p.Key);
                cmd.Parameters.AddWithValue("@Percentage", p.Value);

                cmd.ExecuteNonQuery();
            }
        }
    }
    private static void InsertEfficiencies(SqlConnection conn, SqlTransaction tran, long batchId, P4Batch b)
    {
        string sql = @"
    INSERT INTO P4_Efficiency
    (BatchId, GasName, Concentration, SequenceIndex, EfficiencyValue)
    VALUES
    (@BatchId, @GasName, @Concentration, @Index, @Value);";

        foreach (var g in b.EfficiencyGroups)
        {
            for (int i = 0; i < g.Efficiencies11.Count; i++)
            {
                using (var cmd = new SqlCommand(sql, conn, tran))
                {
                    cmd.Parameters.AddWithValue("@BatchId", batchId);
                    cmd.Parameters.AddWithValue("@GasName", g.GasName);
                    cmd.Parameters.AddWithValue("@Concentration", (object)g.Concentration ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Index", i);
                    cmd.Parameters.AddWithValue("@Value", g.Efficiencies11[i]);

                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}