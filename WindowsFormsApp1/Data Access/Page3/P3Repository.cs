using DocumentFormat.OpenXml.Office2010.Excel;
using Mysqlx.Crud;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO.Packaging;
using System.Security.Claims;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

public class Page3RowData
{
    public string SN { get; set; }
    public string Weight { get; set; }

    public string Length { get; set; }
    public string Width { get; set; }
    public string Height { get; set; }
    public string Diagonal { get; set; }

    public Dictionary<int, string> ControlValues { get; set; } = new Dictionary<int, string>();
}
public static class P3Repository
{
    public static void Insert(Page3ExportData data)
    {
        using (var conn = DbBootstrap.GetConnection())
        {
            conn.Open();

            using (var tran = conn.BeginTransaction())
            {
                try
                {
                    int batchId = InsertBatch(conn, tran, data);
                    InsertRows(conn, tran, batchId, data);

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

    private static int InsertBatch(SqlConnection conn, SqlTransaction tran, Page3ExportData data)
    {
            string sql = @"
            INSERT INTO P3_Batch
            (
                ArrivalDate,
                TestingDate,
                FilterReportNo,
                WorkOrder,
                PackageNo,
                CarbonLot,
                FilterMaterialNo,
                Customer,
                Model,
                ReFilterNo,
                Alarm,
                MA,
                MB,
                MC,
                UserName,
                CreatedAt
            )
            OUTPUT INSERTED.Id
            VALUES
            (
                @ArrivalDate,
                @TestingDate,
                @FilterReportNo,
                @WorkOrder,
                @PackageNo,
                @CarbonLot,
                @FilterMaterialNo,
                @Customer,
                @Model,
                @ReFilterNo,
                @Alarm,
                @MA,
                @MB,
                @MC,
                @UserName,
                GETDATE()
            ); ";
        using (var cmd = new SqlCommand(sql, conn, tran))
        {
            cmd.Parameters.AddWithValue("@TestingDate", DbValue(data.TestingDate));
            cmd.Parameters.AddWithValue("@ArrivalDate", DbValue(data.ArrivalDate));
            cmd.Parameters.AddWithValue("@CarbonLot", DbValue(data.CarbonLot));
            cmd.Parameters.AddWithValue("@FilterMaterialNo", DbValue(data.FilterMaterialNo));
            cmd.Parameters.AddWithValue("@FilterReportNo", DbValue(data.FilterReportNo));
            cmd.Parameters.AddWithValue("@WorkOrder", DbValue(data.WorkOrder));
            cmd.Parameters.AddWithValue("@PackageNo", DbValue(data.PackageNo));
            cmd.Parameters.AddWithValue("@Customer", DbValue(data.Customer));
            cmd.Parameters.AddWithValue("@Model", DbValue(data.Model));
            cmd.Parameters.AddWithValue("@ReFilterNo", DbValue(data.ReFilterNo));
            cmd.Parameters.AddWithValue("@Alarm", DbValue(data.Alarm));
            cmd.Parameters.AddWithValue("@MA", DbValue(data.MA));
            cmd.Parameters.AddWithValue("@MB", DbValue(data.MB));
            cmd.Parameters.AddWithValue("@MC", DbValue(data.MC));
            cmd.Parameters.AddWithValue("@UserName", DbValue(data.UserName));

            return Convert.ToInt32(cmd.ExecuteScalar());
        }
    }

    private static void InsertRows(SqlConnection conn, SqlTransaction tran, int batchId, Page3ExportData data)
    {
        string sql = @"
        INSERT INTO P3_Row
        (
            BatchId,
            RowNo,
            SN,
            Weight,
            Length,
            Width,
            Height,
            Diagonal,

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
            PressureDropSpec,
            PressureDrop
        )
        VALUES
        (
            @BatchId,
            @RowNo,
            @SN,
            @Weight,
            @Length,
            @Width,
            @Height,
            @Diagonal,

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
            @PressureDropSpec,
            @PressureDrop
        );";

        for (int i = 0; i < data.Rows.Count; i++)
        {
            var row = data.Rows[i];

            using (var cmd = new SqlCommand(sql, conn, tran))
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);
                cmd.Parameters.AddWithValue("@RowNo", i + 1);

                cmd.Parameters.AddWithValue("@SN", DbValue(row.SN));
                cmd.Parameters.AddWithValue("@Weight", DbValue(row.Weight));
                cmd.Parameters.AddWithValue("@Length", DbValue(row.Length));
                cmd.Parameters.AddWithValue("@Width", DbValue(row.Width));
                cmd.Parameters.AddWithValue("@Height", DbValue(row.Height));
                cmd.Parameters.AddWithValue("@Diagonal", DbValue(row.Diagonal));

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
                cmd.Parameters.AddWithValue("@PressureDropSpec", Ctrl(row, 53));
                cmd.Parameters.AddWithValue("@PressureDrop", Ctrl(row, 54));

                cmd.ExecuteNonQuery();
            }
        }
    }
    private static object DbValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return DBNull.Value;

        return value;
    }

    private static object Ctrl(Page3RowData row, int key)
    {
        if (row.ControlValues != null &&
            row.ControlValues.TryGetValue(key, out string value))
        {
            return DbValue(value);
        }

        return DBNull.Value;
    }
}