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
        using (var conn = new SQLiteConnection(ConnStr))
        {
            conn.Open();

            CreatePage1Tables(conn);
            CreatePage2Tables(conn);
            CreatePage3Tables(conn);
            CreatePage4Tables(conn);
            CreatePage5Tables(conn);
            CreatePage6Tables(conn);
        }
    }

    // ===============================
    // PAGE 1 濾網原料
    // ===============================
    private static void CreatePage1Tables(SQLiteConnection conn)
    {
        string sql = @"

        CREATE TABLE IF NOT EXISTS P1_Batch (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ArrivalDate TEXT,
            TestDate TEXT,
            ReportNo TEXT,
            MaterialType TEXT,
            Quantity REAL,
            ParticleAnalysis TEXT,
            CreatedAt TEXT,
            Username TEXT
        );

        CREATE TABLE IF NOT EXISTS P1_GasTest (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            BatchId INTEGER,
            GasName TEXT,
            Concentration REAL,
            Background REAL,
            FOREIGN KEY (BatchId) REFERENCES P1_Batch(Id)
        );

        CREATE TABLE IF NOT EXISTS P1_Sample (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            GasTestId INTEGER,
            InBatchNo TEXT,
            InternalBatchNo TEXT,
            Weight REAL,
            Density REAL,
            PressureDrop REAL,
            VOCIn REAL,
            VOCOut REAL,
            VOCOutgassing REAL,
            FOREIGN KEY (GasTestId) REFERENCES P1_GasTest(Id)
        );

        CREATE TABLE IF NOT EXISTS P1_Efficiency (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            SampleId INTEGER,
            SequenceIndex INTEGER,
            EfficiencyValue REAL,
            FOREIGN KEY (SampleId) REFERENCES P1_Sample(Id)
        );
        ";

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            cmd.ExecuteNonQuery();
        }
    }

    // ===============================
    // PAGE 2 濾網半成品
    // ===============================
    private static void CreatePage2Tables(SQLiteConnection conn)
    {
        string sql = @"

        CREATE TABLE IF NOT EXISTS P2_Batch (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ProductionDate TEXT,
            TestDate TEXT,
            WorkOrder TEXT,
            Material TEXT,
            MaterialBatchNo TEXT,
            TargetGsm REAL,
            Glue REAL,
            MaterialNo TEXT,
            Speed REAL,
            UpperTemp REAL,
            LowerTemp REAL,
            Pressure REAL,
            WindSpeed REAL,
            CarbonLine TEXT,
            CreatedAt TEXT,
            Username TEXT
        );

        CREATE TABLE IF NOT EXISTS P2_GasTest (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            BatchId INTEGER,
            GasName TEXT,
            Concentration REAL,
            Background REAL,
            FOREIGN KEY (BatchId) REFERENCES P2_Batch(Id)
        );

        CREATE TABLE IF NOT EXISTS P2_Sample (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            GasTestId INTEGER,
            Weight REAL,
            PressureDrop REAL,
            IsSelected INTEGER,
            FOREIGN KEY (GasTestId) REFERENCES P2_GasTest(Id)
        );

        CREATE TABLE IF NOT EXISTS P2_Efficiency (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            SampleId INTEGER,
            SequenceIndex INTEGER,
            EfficiencyValue REAL,
            FOREIGN KEY (SampleId) REFERENCES P2_Sample(Id)
        );
        ";

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            cmd.ExecuteNonQuery();
        }
    }

    // ===============================
    // PAGE 3 濾網成品
    // ===============================
    private static void CreatePage3Tables(SQLiteConnection conn)
    {
        string sql = @"

        CREATE TABLE IF NOT EXISTS P3_Batch (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ProductionDate TEXT,
            TestDate TEXT,
            ReportNo TEXT,
            CarbonOrder TEXT,
            PackageOrder TEXT,
            Customer TEXT,
            PartNo TEXT,
            Model TEXT,
            Type TEXT,
            RegenerationCount INTEGER,
            InitialEfficiency REAL,
            CreatedAt TEXT
        );

        CREATE TABLE IF NOT EXISTS P3_Sample (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            BatchId INTEGER,
            SerialNo TEXT,
            Weight REAL,
            Length REAL,
            Width REAL,
            Height REAL,
            Diagonal REAL,
            ParticleIn REAL,
            ParticleOut REAL,
            IPAIn REAL,
            IPAOut REAL,
            AcetoneIn REAL,
            AcetoneOut REAL,
            NontargetIn REAL,
            NontargetOut REAL,
            TVOC REAL,
            PressureSpec REAL,
            PressureDrop REAL,
            FOREIGN KEY (BatchId) REFERENCES P3_Batch(Id)
        );
        ";

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            cmd.ExecuteNonQuery();
        }
    }

    // ===============================
    // PAGE 4 濾筒原料
    // ===============================
    private static void CreatePage4Tables(SQLiteConnection conn)
    {
        string sql = @"

        CREATE TABLE IF NOT EXISTS P4_Batch (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ArrivalDate TEXT,
            TestDate TEXT,
            MaterialType TEXT,
            Quantity REAL,
            ParticleAnalysis TEXT,
            Moisture REAL,
            NButane REAL,
            IodineValue REAL,
            Ash REAL,
            CreatedAt TEXT,
            Username TEXT
        );

        CREATE TABLE IF NOT EXISTS P4_GasTest (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            BatchId INTEGER,
            GasName TEXT,
            Concentration REAL,
            Background REAL,
            FOREIGN KEY (BatchId) REFERENCES P4_Batch(Id)
        );

        CREATE TABLE IF NOT EXISTS P4_Sample (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            GasTestId INTEGER,
            InBatchNo TEXT,
            InternalBatchNo TEXT,
            Weight REAL,
            Density REAL,
            PressureDrop REAL,
            VOCIn REAL,
            VOCOut REAL,
            VOCOutgassing REAL,
            FOREIGN KEY (GasTestId) REFERENCES P4_GasTest(Id)
        );

        CREATE TABLE IF NOT EXISTS P4_Efficiency (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            SampleId INTEGER,
            SequenceIndex INTEGER,
            EfficiencyValue REAL,
            FOREIGN KEY (SampleId) REFERENCES P4_Sample(Id)
        );
        ";

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            cmd.ExecuteNonQuery();
        }
    }

    // ===============================
    // PAGE 5 濾筒成品
    // ===============================
    private static void CreatePage5Tables(SQLiteConnection conn)
    {
        string sql = @"

        CREATE TABLE IF NOT EXISTS P5_Batch (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            TestDate TEXT,
            ReportNo TEXT,
            PackageOrder TEXT,
            MaterialBatchNo TEXT,
            Customer TEXT,
            Model TEXT,
            Type TEXT,
            RegenerationCount INTEGER,
            InitialEfficiency REAL,
            CreatedAt TEXT,
            Username TEXT
        );

        CREATE TABLE IF NOT EXISTS P5_Sample (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            BatchId INTEGER,
            SerialNo TEXT,
            Weight REAL,
            ParticleIn REAL,
            ParticleOut REAL,
            IPAIn REAL,
            IPAOut REAL,
            AcetoneIn REAL,
            AcetoneOut REAL,
            NontargetIn REAL,
            NontargetOut REAL,
            TVOC REAL,
            PressureSpec REAL,
            PressureDrop REAL,
            FOREIGN KEY (BatchId) REFERENCES P5_Batch(Id)
        );
        ";

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            cmd.ExecuteNonQuery();
        }
    }

    // ===============================
    // PAGE 6 物料
    // ===============================
    private static void CreatePage6Tables(SQLiteConnection conn)
    {
        string sql = @"

        CREATE TABLE IF NOT EXISTS P6_Record (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            ReportNo TEXT,
            ArrivalDate TEXT,
            SampleDate TEXT,
            MaterialName TEXT,
            PartNo TEXT,
            Quantity REAL,
            SampleQuantity REAL,
            SpecValue REAL,
            MeasuredValue REAL,
            Remark TEXT,
            CreatedAt TEXT,
            Username TEXT
        );
        ";

        using (var cmd = new SQLiteCommand(sql, conn))
        {
            cmd.ExecuteNonQuery();
        }
    }
}
