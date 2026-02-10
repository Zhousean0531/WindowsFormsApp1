using System.IO;
using System.Data.SQLite;

public static class DbBootstrap
{
    public static string ConnStr
    {
        get
        {
            return "Data Source=" +
                   Path.Combine(QCPathHelper.Data, "qc_data.db");
        }
    }

    public static void Init()
    {
        SQLiteConnection conn = null;
        SQLiteCommand cmd = null;

        try
        {
            conn = new SQLiteConnection(ConnStr);
            conn.Open();

            string sql = @"
            CREATE TABLE IF NOT EXISTS InspectionBatch (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                BatchNo TEXT,
                ProductType TEXT,
                InspectionDate TEXT,
                Inspector TEXT,
                MachineName TEXT,
                AppVersion TEXT,
                CreatedAt TEXT
            );";

            cmd = new SQLiteCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();

            if (conn != null)
                conn.Dispose();
        }
    }
}
