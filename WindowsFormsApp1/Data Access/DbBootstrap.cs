using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;

public static class DbBootstrap
{
    public static void Init()
    {

        try
        {
            using (var conn = GetConnection())
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
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }
    }

    public static SqlConnection GetConnection()
    {
        string mode = ConfigurationManager.AppSettings["DbMode"];

        string connStr;

        if (mode == "PROD")
        {
            connStr = ConfigurationManager
                .ConnectionStrings["ProdDb"]
                .ConnectionString;
        }
        else
        {
            connStr = ConfigurationManager
                .ConnectionStrings["DevDb"]
                .ConnectionString;
        }

        return new SqlConnection(connStr);
    }
    public static bool TestConnection()
    {
        try
        {
            using (var conn = GetConnection())
            {
                conn.Open();
                return true;
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show("資料庫連線失敗\n" + ex.Message);
            return false;
        }
    }

    // ===== PAGE 1 =====
    private static void CreatePage1Tables(SqlConnection conn)
    {
        string sql = @"

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P1_Batch' AND xtype='U')
        CREATE TABLE P1_Batch (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            ArrivalDate DATE,
            TestingDate DATE,
            ReportNo NVARCHAR(50),
            Material NVARCHAR(50),
            Concentration DECIMAL(10,2),
            Background DECIMAL(10,2),
            QtyText NVARCHAR(100),
            ParticleAnalysis NVARCHAR(300),
            CreatedAt DATETIME,
            Username NVARCHAR(50)
        );
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P1_Sample' AND xtype='U')
        CREATE TABLE P1_Sample (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            BatchId INT,
            InBatchNo NVARCHAR(50),
            InternalBatchNo NVARCHAR(50),
            Weight DECIMAL(10,4),
            Density DECIMAL(10,4),
            PressureDrop INT,
            VOCIn DECIMAL(10,2),
            VOCOut DECIMAL(10,2),
            VOCOutgassing DECIMAL(10,2)
        );
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P1_Efficiency' AND xtype='U')
        CREATE TABLE P1_Efficiency (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            SampleId INT,
            SequenceIndex INT,
            EfficiencyValue DECIMAL(5,1)
        );
        ";

        new SqlCommand(sql, conn).ExecuteNonQuery();
    }

    // ===== PAGE 2 =====
    private static void CreatePage2Tables(SqlConnection conn)
    {
        string sql = @"

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P2_Batch' AND xtype='U')
        CREATE TABLE P2_Batch (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            ProductionDate DATETIME,
            TestDate DATETIME,
            WorkOrder NVARCHAR(50),
            Material NVARCHAR(50),
            MaterialBatchNo NVARCHAR(50),
            TargetGsm DECIMAL(10,2),
            Glue DECIMAL(10,2),
            MaterialNo NVARCHAR(50),
            Speed DECIMAL(10,2),
            UpperTemp DECIMAL(10,2),
            LowerTemp DECIMAL(10,2),
            Pressure DECIMAL(10,2),
            WindSpeed DECIMAL(10,2),
            CarbonLine NVARCHAR(50),
            CreatedAt DATETIME,
            Username NVARCHAR(50)
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P2_GasTest' AND xtype='U')
        CREATE TABLE P2_GasTest (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            BatchId INT,
            GasName NVARCHAR(50),
            Concentration DECIMAL(10,2),
            Background DECIMAL(10,2)
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P2_Sample' AND xtype='U')
        CREATE TABLE P2_Sample (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            GasTestId INT,
            Weight DECIMAL(10,2),
            PressureDrop DECIMAL(10,2),
            IsSelected BIT
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P2_Efficiency' AND xtype='U')
        CREATE TABLE P2_Efficiency (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            SampleId INT,
            SequenceIndex INT,
            EfficiencyValue FLOAT
        );
        ";

        new SqlCommand(sql, conn).ExecuteNonQuery();
    }

    // ===== PAGE 3 =====
    private static void CreatePage3Tables(SqlConnection conn)
    {
        string sql = @"

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P3_Batch' AND xtype='U')
        CREATE TABLE P3_Batch (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            ProductionDate DATETIME,
            TestDate DATETIME,
            ReportNo NVARCHAR(50),
            CarbonOrder NVARCHAR(50),
            PackageOrder NVARCHAR(50),
            Customer NVARCHAR(50),
            PartNo NVARCHAR(50),
            Model NVARCHAR(50),
            Type NVARCHAR(50),
            RegenerationCount INT,
            InitialEfficiency FLOAT,
            CreatedAt DATETIME
        );

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P3_Sample' AND xtype='U')
        CREATE TABLE P3_Sample (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            BatchId INT,
            SerialNo NVARCHAR(50),
            Weight FLOAT,
            Length FLOAT,
            Width FLOAT,
            Height FLOAT,
            Diagonal FLOAT,
            ParticleIn FLOAT,
            ParticleOut FLOAT,
            IPAIn FLOAT,
            IPAOut FLOAT,
            AcetoneIn FLOAT,
            AcetoneOut FLOAT,
            NontargetIn FLOAT,
            NontargetOut FLOAT,
            TVOC FLOAT,
            PressureSpec FLOAT,
            PressureDrop FLOAT
        );
        ";

        new SqlCommand(sql, conn).ExecuteNonQuery();
    }

    // ===== PAGE 4 =====
    private static void CreatePage4Tables(SqlConnection conn)
    {
        string sql = @"

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P4_Batch' AND xtype='U')
        CREATE TABLE P4_Batch (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            ReportNo NVARCHAR(50),
            Material NVARCHAR(50),
            MaterialNo NVARCHAR(50),
            ArrivalDate DATETIME,
            TestingDate DATETIME,
            QtyText NVARCHAR(50),
            Username NVARCHAR(50),
            CreatedAt DATETIME
        );
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P4_Particle' AND xtype='U')
        CREATE TABLE P4_Particle (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            BatchId INT,
            SizeName NVARCHAR(50),
            Percentage FLOAT
        );
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P4_Efficiency' AND xtype='U')
        CREATE TABLE P4_Efficiency (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            BatchId INT,
            GasName NVARCHAR(50),
            Concentration FLOAT,
            SequenceIndex INT,
            EfficiencyValue FLOAT
        );
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P4_Lot' AND xtype='U')
        CREATE TABLE P4_Lot (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            BatchId INT,
            LotNo NVARCHAR(50),
            LotFull NVARCHAR(50),
            Weight FLOAT,
            Density FLOAT,
            VocIn FLOAT,
            VocOut FLOAT,
            DeltaP FLOAT,
            Outgassing NVARCHAR(50),
            IsSelected BIT
        );
        ";

        new SqlCommand(sql, conn).ExecuteNonQuery();
    }

    // ===== PAGE 5 =====
    private static void CreatePage5Tables(SqlConnection conn)
    {
        string sql = @"

        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P5_Batch' AND xtype='U')
        CREATE TABLE P5_Batch (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            TestDate DATETIME,
            ReportNo NVARCHAR(50),
            PackageOrder NVARCHAR(50),
            MaterialBatchNo NVARCHAR(50),
            Customer NVARCHAR(50),
            Model NVARCHAR(50),
            Type NVARCHAR(50),
            RegenerationCount INT,
            InitialEfficiency FLOAT,
            CreatedAt DATETIME,
            Username NVARCHAR(50)
        );
        ";

        new SqlCommand(sql, conn).ExecuteNonQuery();
    }

    // ===== PAGE 6 =====
    private static void CreatePage6Tables(SqlConnection conn)
    {
        string sql = @"

    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P6_Batch' AND xtype='U')
    CREATE TABLE P6_Batch (
        Id INT IDENTITY(1,1) PRIMARY KEY,
        ReportNo NVARCHAR(50),
        TestDate DATETIME,
        UserName NVARCHAR(50),
        CreatedAt DATETIME
    );

    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P6_Item' AND xtype='U')
    CREATE TABLE P6_Item (
        Id INT IDENTITY(1,1) PRIMARY KEY,
        BatchId INT,

        Col1 NVARCHAR(100),
        Col2 NVARCHAR(100),

        Spec1 NVARCHAR(100),
        Spec2 NVARCHAR(100),

        Range1 NVARCHAR(100),
        Range2 NVARCHAR(100),

        Result NVARCHAR(100),
        Judgment NVARCHAR(100),

        Extra1 NVARCHAR(100),
        Extra2 NVARCHAR(100),
        Extra3 NVARCHAR(100)
    );
    ";

        new SqlCommand(sql, conn).ExecuteNonQuery();
    }
}