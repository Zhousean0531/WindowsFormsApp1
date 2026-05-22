using System;
using System.Data.SqlClient;
using System.Linq;
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
            cmd.Parameters.AddWithValue("@ReportNo", DbValue(b.ReportNo));
            cmd.Parameters.AddWithValue("@Material", DbValue(b.Material));
            cmd.Parameters.AddWithValue("@MaterialNo", DbValue(b.MaterialNo));
            cmd.Parameters.AddWithValue("@ArrivalDate", DbValue(b.ArrivalDate));
            cmd.Parameters.AddWithValue("@TestingDate", DbValue(b.TestingDate));
            cmd.Parameters.AddWithValue("@QtyText", DbValue(b.QtyText));
            cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now);
            cmd.Parameters.AddWithValue("@Username", DbValue(b.UserName));

            cmd.Parameters.AddWithValue("@Moisture", DbValue(b.Moisture));
            cmd.Parameters.AddWithValue("@Butane", DbValue(b.Butane));
            cmd.Parameters.AddWithValue("@Ash", DbValue(b.Ash));

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
                cmd.Parameters.AddWithValue("@LotNo", DbValue(r.LotNo));
                cmd.Parameters.AddWithValue("@LotFull", DbValue(r.LotFull));
                cmd.Parameters.AddWithValue("@Weight", r.Weight);
                cmd.Parameters.AddWithValue("@Density", r.Density);
                cmd.Parameters.AddWithValue("@VocIn", r.VocIn);
                cmd.Parameters.AddWithValue("@VocOut", r.VocOut);
                cmd.Parameters.AddWithValue("@DeltaP", r.DeltaP);
                cmd.Parameters.AddWithValue("@Outgassing", DbValue(r.Outgassing));
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
                cmd.Parameters.AddWithValue("@SizeName", DbValue(p.Key));
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
            if (g.EfficiencyPoints != null && g.EfficiencyPoints.Count > 0)
            {
                foreach (var kv in g.EfficiencyPoints.OrderBy(x => x.Key))
                {
                    using (var cmd = new SqlCommand(sql, conn, tran))
                    {
                        cmd.Parameters.AddWithValue("@BatchId", batchId);
                        cmd.Parameters.AddWithValue("@GasName", DbValue(g.GasName));
                        cmd.Parameters.AddWithValue("@Concentration", DbValue(g.Concentration));
                        cmd.Parameters.AddWithValue("@Index", kv.Key);
                        cmd.Parameters.AddWithValue("@Value", kv.Value);

                        cmd.ExecuteNonQuery();
                    }
                }

                continue;
            }

            for (int i = 0; i < g.Efficiencies11.Count; i++)
            {
                using (var cmd = new SqlCommand(sql, conn, tran))
                {
                    cmd.Parameters.AddWithValue("@BatchId", batchId);
                    cmd.Parameters.AddWithValue("@GasName", DbValue(g.GasName));
                    cmd.Parameters.AddWithValue("@Concentration", DbValue(g.Concentration));
                    cmd.Parameters.AddWithValue("@Index", i);
                    cmd.Parameters.AddWithValue("@Value", g.Efficiencies11[i]);

                    cmd.ExecuteNonQuery();
                }
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
