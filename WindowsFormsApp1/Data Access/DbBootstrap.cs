using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;

public static class DbBootstrap
{
    private static string _connStr;

    public static void Init()
    {
        try
        {
            _connStr = ResolveConnectionString();

            using (var conn = GetConnection())
            {
                conn.Open();

                CreateInstrumentTables(conn);
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

    private static string ResolveConnectionString()
    {
        string mode = ConfigurationManager.AppSettings["DbMode"];

        if (string.IsNullOrWhiteSpace(mode))
            mode = "DEV";

        mode = mode.Trim().ToUpper();

        if (mode == "PROD")
        {
            return ConfigurationManager
                .ConnectionStrings["ProdDb"]
                .ConnectionString;
        }

        return ConfigurationManager
            .ConnectionStrings["DevDb"]
            .ConnectionString;
    }

    public static SqlConnection GetConnection()
    {
        if (string.IsNullOrWhiteSpace(_connStr))
            _connStr = ResolveConnectionString();

        return new SqlConnection(_connStr);
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

    // ===== INSTRUMENTS =====
    private static void CreateInstrumentTables(SqlConnection conn)
    {
        string sql = @"
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Instruments' AND xtype='U')
CREATE TABLE Instruments (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    InstrumentName NVARCHAR(200),
    CalibrationDate DATETIME,
    ExpireDate DATETIME,
    CreatedAt DATETIME,
    UpdatedAt DATETIME,
    UpdatedBy NVARCHAR(50)
);
";

        new SqlCommand(sql, conn).ExecuteNonQuery();
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
    Username NVARCHAR(50),
    ReportNo NVARCHAR(50),
    FilterSize NVARCHAR(50)
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
    ArrivalDate DATETIME,
    TestingDate DATETIME,
    FilterReportNo NVARCHAR(50),
    WorkOrder NVARCHAR(50),
    PackageNo NVARCHAR(50),
    CarbonLot NVARCHAR(50),
    FilterMaterialNo NVARCHAR(50),
    Customer NVARCHAR(100),
    Model NVARCHAR(100),
    ReFilterNo NVARCHAR(50),
    Alarm NVARCHAR(100),
    MA NVARCHAR(50),
    MB NVARCHAR(50),
    MC NVARCHAR(50),
    UserName NVARCHAR(50),
    CreatedAt DATETIME
);

IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P3_Row' AND xtype='U')
CREATE TABLE P3_Row (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    BatchId INT,
    RowNo INT,
    SN NVARCHAR(50),
    Weight NVARCHAR(50),
    Length NVARCHAR(50),
    Width NVARCHAR(50),
    Height NVARCHAR(50),
    Diagonal NVARCHAR(50),

    ParticleIn NVARCHAR(50),
    ParticleOut NVARCHAR(50),
    ParticleDiff NVARCHAR(50),

    IPAIn NVARCHAR(50),
    IPAOut NVARCHAR(50),
    IPADiff NVARCHAR(50),

    AcetoneIn NVARCHAR(50),
    AcetoneOut NVARCHAR(50),
    AcetoneDiff NVARCHAR(50),

    NonTargetIn NVARCHAR(50),
    NonTargetOut NVARCHAR(50),
    NonTargetDiff NVARCHAR(50),

    TotalDiff NVARCHAR(50),
    PressureDropSpec NVARCHAR(50),
    PressureDrop NVARCHAR(50)
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
    CreatedAt DATETIME,
    Moisture NVARCHAR(50),
    Butane NVARCHAR(50),
    Ash NVARCHAR(50)
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
    ReportNo NVARCHAR(50),
    CylinderNo NVARCHAR(50),
    Customer NVARCHAR(100),
    FilterType NVARCHAR(50),
    TestDate DATETIME,
    UserName NVARCHAR(50),
    CreatedAt DATETIME,
    ReCylinderNo NVARCHAR(50),
    CarbonLot NVARCHAR(50)
);

IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='P5_Row' AND xtype='U')
CREATE TABLE P5_Row (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    BatchId INT,
    RowNo INT,
    SN NVARCHAR(50),
    Weight NVARCHAR(50),
    Efficiency NVARCHAR(50),

    ParticleIn NVARCHAR(50),
    ParticleOut NVARCHAR(50),
    ParticleDiff NVARCHAR(50),

    IPAIn NVARCHAR(50),
    IPAOut NVARCHAR(50),
    IPADiff NVARCHAR(50),

    AcetoneIn NVARCHAR(50),
    AcetoneOut NVARCHAR(50),
    AcetoneDiff NVARCHAR(50),

    NonTargetIn NVARCHAR(50),
    NonTargetOut NVARCHAR(50),
    NonTargetDiff NVARCHAR(50),

    TotalDiff NVARCHAR(50),
    PressureDrop NVARCHAR(50)
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