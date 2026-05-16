using System;
using System.Data.SqlClient;

namespace WindowsFormsApp1.Data_Access.Page6
{
    internal static class P6Repository
    {
        public static void Insert(P6Batch batch)
        {

            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                using (var tran = conn.BeginTransaction())
                {
                    try
                    {
                        long batchId = InsertBatch(conn, tran, batch);

                        foreach (var item in batch.Items)
                        {
                            InsertItem(conn, tran, batchId, item);
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

        private static long InsertBatch(SqlConnection conn, SqlTransaction tran, P6Batch batch)
        {
            string sql = @"
                INSERT INTO P6_Batch
                (ReportNo, TestDate, UserName, SuppliedNO,CreatedAt)
                VALUES
                (@ReportNo, @TestDate, @UserName,@SuppliedNO, GETDATE());

                SELECT SCOPE_IDENTITY();
            ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@ReportNo", batch.ReportNo);
                cmd.Parameters.AddWithValue("@TestDate", batch.TestDate);
                cmd.Parameters.AddWithValue("@UserName", batch.UserName ?? "");
                cmd.Parameters.AddWithValue("@SuppliedNO", batch.SuppliedNO ?? "");
                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }

        private static void InsertItem(SqlConnection conn, SqlTransaction tran, long batchId, P6Item item)
        {
            string sql = @"
                INSERT INTO P6_Item
                (
                    BatchId,
                    Col1, Col2,
                    Spec1, Spec2,
                    Range1, Range2,
                    Result, Judgment,
                    Extra1, Extra2, Extra3
                )
                VALUES
                (
                    @BatchId,
                    @Col1, @Col2,
                    @Spec1, @Spec2,
                    @Range1, @Range2,
                    @Result, @Judgment,
                    @Extra1, @Extra2, @Extra3
                );
            ";

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);

                cmd.Parameters.AddWithValue("@Col1", item.Col1 ?? "");
                cmd.Parameters.AddWithValue("@Col2", item.Col2 ?? "");

                cmd.Parameters.AddWithValue("@Spec1", item.Spec1 ?? "");
                cmd.Parameters.AddWithValue("@Spec2", item.Spec2 ?? "");

                cmd.Parameters.AddWithValue("@Range1", item.Range1 ?? "");
                cmd.Parameters.AddWithValue("@Range2", item.Range2 ?? "");

                cmd.Parameters.AddWithValue("@Result", item.Result ?? "");
                cmd.Parameters.AddWithValue("@Judgment", item.Judgment ?? "");

                cmd.Parameters.AddWithValue("@Extra1", item.Extra1 ?? "");
                cmd.Parameters.AddWithValue("@Extra2", item.Extra2 ?? "");
                cmd.Parameters.AddWithValue("@Extra3", item.Extra3 ?? "");

                cmd.ExecuteNonQuery();
            }
        }
    }
}