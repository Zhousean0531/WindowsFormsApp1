using System;
using System.Data;
using System.Data.SqlClient;

namespace WindowsFormsApp1.Data_Access.Query
{
    public static class QcQueryRepository
    {
        public static DataTable Query(QcQueryCriteria criteria)
        {
            var table = new DataTable();

            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                if (criteria.QueryKind == "原料")
                {
                    AppendRows(table, conn, BuildP1Sql(), cmd => AddCommonParameters(cmd, criteria, "RawMaterialType"));
                    AppendRows(table, conn, BuildP4Sql(), cmd => AddCommonParameters(cmd, criteria, "RawMaterialType"));
                }
                else if (criteria.QueryKind == "半成品")
                {
                    AppendRows(table, conn, BuildP2Sql(), cmd => AddCommonParameters(cmd, criteria, "SemiProductType"));
                }
                else if (criteria.QueryKind == "成品")
                {
                    AppendRows(table, conn, BuildP3Sql(), cmd => AddCommonParameters(cmd, criteria, "ProductNo"));
                    AppendRows(table, conn, BuildP5Sql(), cmd => AddCommonParameters(cmd, criteria, "ProductNo"));
                }
            }

            return table;
        }

        private static void AppendRows(
            DataTable table,
            SqlConnection conn,
            string sql,
            Action<SqlCommand> configure)
        {
            using (var cmd = new SqlCommand(sql, conn))
            {
                configure(cmd);

                var temp = new DataTable();
                using (var adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(temp);
                }

                if (temp.Rows.Count == 0)
                    return;

                table.Merge(temp, true, MissingSchemaAction.Add);
            }
        }

        private static void AddCommonParameters(SqlCommand cmd, QcQueryCriteria criteria, string filterName)
        {
            cmd.Parameters.Add("@StartDate", SqlDbType.Date).Value =
                criteria.StartDate.HasValue ? (object)criteria.StartDate.Value.Date : DBNull.Value;

            cmd.Parameters.Add("@EndDate", SqlDbType.Date).Value =
                criteria.EndDate.HasValue ? (object)criteria.EndDate.Value.Date : DBNull.Value;

            cmd.Parameters.Add("@Filter", SqlDbType.NVarChar, 100).Value =
                GetFilterValue(criteria, filterName);

            cmd.Parameters.Add("@SemiGsm", SqlDbType.NVarChar, 50).Value =
                criteria.SemiProductGsm ?? "";
        }

        private static string GetFilterValue(QcQueryCriteria criteria, string filterName)
        {
            if (filterName == "RawMaterialType")
                return criteria.RawMaterialType ?? "";

            if (filterName == "SemiProductType")
                return criteria.SemiProductType ?? "";

            if (filterName == "ProductNo")
                return criteria.ProductNo ?? "";

            return "";
        }

        private static string BuildP1Sql()
        {
            return @"
SELECT
    N'瞈曄雯??' AS [靘?],
    CONVERT(NVARCHAR(50), b.ArrivalDate, 120) AS [P1_Batch_ArrivalDate],
    CONVERT(NVARCHAR(50), b.TestingDate, 120) AS [P1_Batch_TestingDate],
    CONVERT(NVARCHAR(100), b.ReportNo) AS [P1_Batch_ReportNo],
    CONVERT(NVARCHAR(100), b.Material) AS [P1_Batch_Material],
    CONVERT(NVARCHAR(50), b.Concentration) AS [P1_Batch_Concentration],
    CONVERT(NVARCHAR(50), b.Background) AS [P1_Batch_Background],
    CONVERT(NVARCHAR(200), b.QtyText) AS [P1_Batch_QtyText],
    CONVERT(NVARCHAR(500), b.ParticleAnalysis) AS [P1_Batch_ParticleAnalysis],
    CONVERT(NVARCHAR(50), b.CreatedAt, 120) AS [P1_Batch_CreatedAt],
    CONVERT(NVARCHAR(100), b.Username) AS [P1_Batch_Username],
    CONVERT(NVARCHAR(100), s.InBatchNo) AS [P1_Sample_InBatchNo],
    CONVERT(NVARCHAR(100), s.SuppliedNO) AS [P1_Sample_SuppliedNO],
    CONVERT(NVARCHAR(50), s.Weight) AS [P1_Sample_Weight],
    CONVERT(NVARCHAR(50), s.Density) AS [P1_Sample_Density],
    CONVERT(NVARCHAR(50), s.PressureDrop) AS [P1_Sample_PressureDrop],
    CONVERT(NVARCHAR(50), s.VOCIn) AS [P1_Sample_VOCIn],
    CONVERT(NVARCHAR(50), s.VOCOut) AS [P1_Sample_VOCOut],
    CONVERT(NVARCHAR(50), s.VOCOutgassing) AS [P1_Sample_VOCOutgassing],
    e.EfficiencyAll AS [P1_Efficiency_All]
FROM P1_Batch b
LEFT JOIN P1_Sample s ON s.BatchId = b.Id
OUTER APPLY (
    SELECT STUFF((
        SELECT N' / ' +
               N'Seq=' + CONVERT(NVARCHAR(50), pe.SequenceIndex) +
               N',Value=' + ISNULL(CONVERT(NVARCHAR(50), pe.EfficiencyValue), '')
        FROM P1_Efficiency pe
        WHERE pe.SampleId = s.Id
        ORDER BY pe.SequenceIndex
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 3, '') AS EfficiencyAll
) e
WHERE (@StartDate IS NULL OR CONVERT(date, b.TestingDate) >= @StartDate)
  AND (@EndDate IS NULL OR CONVERT(date, b.TestingDate) <= @EndDate)
  AND (@Filter = '' OR ISNULL(b.Material, '') LIKE '%' + @Filter + '%')
ORDER BY b.TestingDate DESC, b.Id DESC, s.Id;";
        }

        private static string BuildP4Sql()
        {
            return @"
SELECT
    N'瞈曄???' AS [靘?],
    CONVERT(NVARCHAR(100), b.ReportNo) AS [P4_Batch_ReportNo],
    CONVERT(NVARCHAR(100), b.Material) AS [P4_Batch_Material],
    CONVERT(NVARCHAR(100), b.MaterialNo) AS [P4_Batch_MaterialNo],
    CONVERT(NVARCHAR(50), b.ArrivalDate, 120) AS [P4_Batch_ArrivalDate],
    CONVERT(NVARCHAR(50), b.TestingDate, 120) AS [P4_Batch_TestingDate],
    CONVERT(NVARCHAR(200), b.QtyText) AS [P4_Batch_QtyText],
    CONVERT(NVARCHAR(100), b.Username) AS [P4_Batch_Username],
    CONVERT(NVARCHAR(50), b.CreatedAt, 120) AS [P4_Batch_CreatedAt],
    CONVERT(NVARCHAR(100), b.Moisture) AS [P4_Batch_Moisture],
    CONVERT(NVARCHAR(100), b.Butane) AS [P4_Batch_Butane],
    CONVERT(NVARCHAR(100), b.Ash) AS [P4_Batch_Ash],
    CONVERT(NVARCHAR(100), l.LotNo) AS [P4_Lot_LotNo],
    CONVERT(NVARCHAR(100), l.LotFull) AS [P4_Lot_LotFull],
    CONVERT(NVARCHAR(50), l.Weight) AS [P4_Lot_Weight],
    CONVERT(NVARCHAR(50), l.Density) AS [P4_Lot_Density],
    CONVERT(NVARCHAR(50), l.VocIn) AS [P4_Lot_VocIn],
    CONVERT(NVARCHAR(50), l.VocOut) AS [P4_Lot_VocOut],
    CONVERT(NVARCHAR(50), l.DeltaP) AS [P4_Lot_DeltaP],
    CONVERT(NVARCHAR(100), l.Outgassing) AS [P4_Lot_Outgassing],
    e.EfficiencyAll AS [P4_Efficiency_All],
    p.ParticleAll AS [P4_Particle_All]
FROM P4_Batch b
LEFT JOIN P4_Lot l ON l.BatchId = b.Id
OUTER APPLY (
    SELECT STUFF((
        SELECT N' / ' +
               CASE
                   WHEN pe.SequenceIndex = (
                       SELECT MIN(pe2.SequenceIndex)
                       FROM P4_Efficiency pe2
                       WHERE pe2.BatchId = pe.BatchId
                         AND ISNULL(pe2.GasName, '') = ISNULL(pe.GasName, '')
                         AND ISNULL(CONVERT(NVARCHAR(50), pe2.Concentration), '') = ISNULL(CONVERT(NVARCHAR(50), pe.Concentration), '')
                   )
                   THEN N'Gas=' + ISNULL(pe.GasName, '') +
                        N',Conc=' + ISNULL(CONVERT(NVARCHAR(50), pe.Concentration), '') +
                        N','
                   ELSE N''
               END +
               N'Seq=' + CONVERT(NVARCHAR(50), pe.SequenceIndex) +
               N',Value=' + ISNULL(CONVERT(NVARCHAR(50), pe.EfficiencyValue), '')
        FROM P4_Efficiency pe
        WHERE pe.BatchId = b.Id
        ORDER BY pe.GasName, pe.Concentration, pe.SequenceIndex
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 3, '') AS EfficiencyAll
) e
OUTER APPLY (
    SELECT STUFF((
        SELECT N',' +
               ISNULL(pp.SizeName, '') +
               N':' +
               ISNULL(FORMAT(pp.Percentage, '0.###'), '')
        FROM P4_Particle pp
        WHERE pp.BatchId = b.Id
        ORDER BY pp.Id
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS ParticleAll
) p
WHERE (@StartDate IS NULL OR CONVERT(date, b.TestingDate) >= @StartDate)
  AND (@EndDate IS NULL OR CONVERT(date, b.TestingDate) <= @EndDate)
  AND (@Filter = '' OR ISNULL(b.Material, '') LIKE '%' + @Filter + '%')
ORDER BY b.TestingDate DESC, b.Id DESC, l.Id;";
        }

        private static string BuildP2Sql()
        {
            return @"
SELECT
    N'濾網半成品' AS [來源],
    CONVERT(NVARCHAR(50), b.ProductionDate, 120) AS [P2_Batch_ProductionDate],
    CONVERT(NVARCHAR(50), b.TestDate, 120) AS [P2_Batch_TestDate],
    CONVERT(NVARCHAR(100), b.WorkOrder) AS [P2_Batch_WorkOrder],
    CONVERT(NVARCHAR(100), b.Material) AS [P2_Batch_Material],
    CONVERT(NVARCHAR(100), b.MaterialNo) AS [P2_Batch_MaterialNo],
    CONVERT(NVARCHAR(100), b.BatchNo) AS [P2_Batch_BatchNo],
    CONVERT(NVARCHAR(50), b.TargetGsm) AS [P2_Batch_TargetGsm],
    CONVERT(NVARCHAR(100), b.Glue) AS [P2_Batch_Glue],
    CONVERT(NVARCHAR(50), b.Speed) AS [P2_Batch_Speed],
    CONVERT(NVARCHAR(50), b.UpperTemp) AS [P2_Batch_UpperTemp],
    CONVERT(NVARCHAR(50), b.LowerTemp) AS [P2_Batch_LowerTemp],
    CONVERT(NVARCHAR(50), b.Pressure) AS [P2_Batch_Pressure],
    CONVERT(NVARCHAR(50), b.WindSpeed) AS [P2_Batch_WindSpeed],
    CONVERT(NVARCHAR(100), b.CarbonLine) AS [P2_Batch_CarbonLine],
    CONVERT(NVARCHAR(50), b.CreatedAt, 120) AS [P2_Batch_CreatedAt],
    CONVERT(NVARCHAR(100), b.Username) AS [P2_Batch_Username],
    CONVERT(NVARCHAR(100), b.ReportNo) AS [P2_Batch_ReportNo],
    CONVERT(NVARCHAR(100), b.FilterSize) AS [P2_Batch_FilterSize],
    CONVERT(NVARCHAR(50), g.Id) AS [P2_GasTest_Id],
    CONVERT(NVARCHAR(100), g.GasName) AS [P2_GasTest_GasName],
    CONVERT(NVARCHAR(50), g.Concentration) AS [P2_GasTest_Concentration],
    CONVERT(NVARCHAR(50), g.Background) AS [P2_GasTest_Background],
    CONVERT(NVARCHAR(50), s.Id) AS [P2_Sample_Id],
    CONVERT(NVARCHAR(50), s.GasTestId) AS [P2_Sample_GasTestId],
    CONVERT(NVARCHAR(50), s.Weight) AS [P2_Sample_Weight],
    CONVERT(NVARCHAR(50), s.PressureDrop) AS [P2_Sample_PressureDrop],
    e.EfficiencyAll AS [P2_Efficiency_All]
FROM P2_Batch b
LEFT JOIN P2_GasTest g ON g.BatchId = b.Id
LEFT JOIN P2_Sample s ON s.GasTestId = g.Id
OUTER APPLY (
    SELECT STUFF((
        SELECT N' / ' +
               N'Seq=' + CONVERT(NVARCHAR(50), pe.SequenceIndex) +
               N',Value=' + ISNULL(CONVERT(NVARCHAR(50), pe.EfficiencyValue), '')
        FROM P2_Efficiency pe
        WHERE pe.SampleId = s.Id
        ORDER BY pe.SequenceIndex
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 3, '') AS EfficiencyAll
) e
WHERE (@StartDate IS NULL OR CONVERT(date, b.TestDate) >= @StartDate)
  AND (@EndDate IS NULL OR CONVERT(date, b.TestDate) <= @EndDate)
  AND (@Filter = '' OR LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Filter)))
  AND (
        @SemiGsm = ''
        OR (
            TRY_CONVERT(DECIMAL(10, 2), @SemiGsm) IS NOT NULL
            AND b.TargetGsm = TRY_CONVERT(DECIMAL(10, 2), @SemiGsm)
        )
      )
ORDER BY b.TestDate DESC, b.Id DESC, g.Id, s.Id;";
        }

        private static string BuildP3Sql()
        {
            return @"
SELECT
    N'濾網成品' AS [來源],
    CONVERT(NVARCHAR(50), b.ArrivalDate, 120) AS [P3_Batch_ArrivalDate],
    CONVERT(NVARCHAR(50), b.TestingDate, 120) AS [P3_Batch_TestingDate],
    CONVERT(NVARCHAR(100), b.FilterReportNo) AS [P3_Batch_FilterReportNo],
    CONVERT(NVARCHAR(100), b.WorkOrder) AS [P3_Batch_WorkOrder],
    CONVERT(NVARCHAR(100), b.PackageNo) AS [P3_Batch_PackageNo],
    CONVERT(NVARCHAR(100), b.CarbonLot) AS [P3_Batch_CarbonLot],
    CONVERT(NVARCHAR(100), b.FilterMaterialNo) AS [P3_Batch_FilterMaterialNo],
    CONVERT(NVARCHAR(100), b.Customer) AS [P3_Batch_Customer],
    CONVERT(NVARCHAR(100), b.Model) AS [P3_Batch_Model],
    CONVERT(NVARCHAR(100), b.ReFilterNo) AS [P3_Batch_ReFilterNo],
    CONVERT(NVARCHAR(100), b.Alarm) AS [P3_Batch_Alarm],
    CONVERT(NVARCHAR(100), b.MA) AS [P3_Batch_MA],
    CONVERT(NVARCHAR(100), b.MB) AS [P3_Batch_MB],
    CONVERT(NVARCHAR(100), b.MC) AS [P3_Batch_MC],
    CONVERT(NVARCHAR(100), b.UserName) AS [P3_Batch_UserName],
    CONVERT(NVARCHAR(50), b.CreatedAt, 120) AS [P3_Batch_CreatedAt],
    CONVERT(NVARCHAR(50), r.RowNo) AS [P3_Row_RowNo],
    CONVERT(NVARCHAR(100), r.SN) AS [P3_Row_SN],
    CONVERT(NVARCHAR(100), r.Weight) AS [P3_Row_Weight],
    CONVERT(NVARCHAR(100), r.Length) AS [P3_Row_Length],
    CONVERT(NVARCHAR(100), r.Width) AS [P3_Row_Width],
    CONVERT(NVARCHAR(100), r.Height) AS [P3_Row_Height],
    CONVERT(NVARCHAR(100), r.Diagonal) AS [P3_Row_Diagonal],
    CONVERT(NVARCHAR(100), r.ParticleIn) AS [P3_Row_ParticleIn],
    CONVERT(NVARCHAR(100), r.ParticleOut) AS [P3_Row_ParticleOut],
    CONVERT(NVARCHAR(100), r.ParticleDiff) AS [P3_Row_ParticleDiff],
    CONVERT(NVARCHAR(100), r.IPAIn) AS [P3_Row_IPAIn],
    CONVERT(NVARCHAR(100), r.IPAOut) AS [P3_Row_IPAOut],
    CONVERT(NVARCHAR(100), r.IPADiff) AS [P3_Row_IPADiff],
    CONVERT(NVARCHAR(100), r.AcetoneIn) AS [P3_Row_AcetoneIn],
    CONVERT(NVARCHAR(100), r.AcetoneOut) AS [P3_Row_AcetoneOut],
    CONVERT(NVARCHAR(100), r.AcetoneDiff) AS [P3_Row_AcetoneDiff],
    CONVERT(NVARCHAR(100), r.NonTargetIn) AS [P3_Row_NonTargetIn],
    CONVERT(NVARCHAR(100), r.NonTargetOut) AS [P3_Row_NonTargetOut],
    CONVERT(NVARCHAR(100), r.NonTargetDiff) AS [P3_Row_NonTargetDiff],
    CONVERT(NVARCHAR(100), r.TotalDiff) AS [P3_Row_TotalDiff],
    CONVERT(NVARCHAR(100), r.PressureDropSpec) AS [P3_Row_PressureDropSpec],
    CONVERT(NVARCHAR(100), r.PressureDrop) AS [P3_Row_PressureDrop]
FROM P3_Batch b
LEFT JOIN P3_Row r ON r.BatchId = b.Id
WHERE (@StartDate IS NULL OR CONVERT(date, b.TestingDate) >= @StartDate)
  AND (@EndDate IS NULL OR CONVERT(date, b.TestingDate) <= @EndDate)
  AND (@Filter = '' OR ISNULL(b.FilterMaterialNo, '') LIKE '%' + @Filter + '%')
ORDER BY b.TestingDate DESC, b.Id DESC, r.RowNo;";
        }

        private static string BuildP5Sql()
        {
            return @"
SELECT
    N'濾筒成品' AS [來源],
    CONVERT(NVARCHAR(100), b.ReportNo) AS [P5_Batch_ReportNo],
    CONVERT(NVARCHAR(100), b.CylinderNo) AS [P5_Batch_CylinderNo],
    CONVERT(NVARCHAR(100), b.Customer) AS [P5_Batch_Customer],
    CONVERT(NVARCHAR(100), b.FilterType) AS [P5_Batch_FilterType],
    CONVERT(NVARCHAR(100), b.MaterialNo) AS [P5_Batch_MaterialNo],
    p.ProductNo AS [P5_Mapped_ProductNo],
    CONVERT(NVARCHAR(50), b.TestDate, 120) AS [P5_Batch_TestDate],
    CONVERT(NVARCHAR(100), b.UserName) AS [P5_Batch_UserName],
    CONVERT(NVARCHAR(50), b.CreatedAt, 120) AS [P5_Batch_CreatedAt],
    CONVERT(NVARCHAR(100), b.ReCylinderNo) AS [P5_Batch_ReCylinderNo],
    CONVERT(NVARCHAR(100), b.CarbonLot) AS [P5_Batch_CarbonLot],
    CONVERT(NVARCHAR(50), r.RowNo) AS [P5_Row_RowNo],
    CONVERT(NVARCHAR(100), r.SN) AS [P5_Row_SN],
    CONVERT(NVARCHAR(100), r.Weight) AS [P5_Row_Weight],
    CONVERT(NVARCHAR(100), r.Efficiency) AS [P5_Row_Efficiency],
    CONVERT(NVARCHAR(100), r.ParticleIn) AS [P5_Row_ParticleIn],
    CONVERT(NVARCHAR(100), r.ParticleOut) AS [P5_Row_ParticleOut],
    CONVERT(NVARCHAR(100), r.ParticleDiff) AS [P5_Row_ParticleDiff],
    CONVERT(NVARCHAR(100), r.IPAIn) AS [P5_Row_IPAIn],
    CONVERT(NVARCHAR(100), r.IPAOut) AS [P5_Row_IPAOut],
    CONVERT(NVARCHAR(100), r.IPADiff) AS [P5_Row_IPADiff],
    CONVERT(NVARCHAR(100), r.AcetoneIn) AS [P5_Row_AcetoneIn],
    CONVERT(NVARCHAR(100), r.AcetoneOut) AS [P5_Row_AcetoneOut],
    CONVERT(NVARCHAR(100), r.AcetoneDiff) AS [P5_Row_AcetoneDiff],
    CONVERT(NVARCHAR(100), r.NonTargetIn) AS [P5_Row_NonTargetIn],
    CONVERT(NVARCHAR(100), r.NonTargetOut) AS [P5_Row_NonTargetOut],
    CONVERT(NVARCHAR(100), r.NonTargetDiff) AS [P5_Row_NonTargetDiff],
    CONVERT(NVARCHAR(100), r.TotalDiff) AS [P5_Row_TotalDiff],
    CONVERT(NVARCHAR(100), r.PressureDrop) AS [P5_Row_PressureDrop]
FROM P5_Batch b
LEFT JOIN P5_Row r ON r.BatchId = b.Id
CROSS APPLY (
    SELECT COALESCE(
        NULLIF(LTRIM(RTRIM(ISNULL(b.MaterialNo, ''))), ''),
        CASE
            WHEN UPPER(ISNULL(b.FilterType, '')) LIKE '%ACID%' THEN N'11A0C00Y000002'
            WHEN UPPER(ISNULL(b.FilterType, '')) LIKE '%BASE%' OR UPPER(ISNULL(b.FilterType, '')) LIKE '%MB%' THEN N'11B0B00Y000002'
            WHEN UPPER(ISNULL(b.FilterType, '')) LIKE '%DMS%' OR UPPER(ISNULL(b.FilterType, '')) LIKE '%MA%' THEN N'11D0S00Y000002'
            WHEN UPPER(ISNULL(b.FilterType, '')) LIKE '%TOC%' OR UPPER(ISNULL(b.FilterType, '')) LIKE '%MC%' THEN N'11T0C00Y000002'
            ELSE NULL
        END
    ) AS ProductNo
) p
WHERE (@StartDate IS NULL OR CONVERT(date, b.TestDate) >= @StartDate)
  AND (@EndDate IS NULL OR CONVERT(date, b.TestDate) <= @EndDate)
  AND (
        @Filter = ''
        OR ISNULL(p.ProductNo, '') LIKE '%' + @Filter + '%'
        OR ISNULL(b.FilterType, '') LIKE '%' + @Filter + '%'
      )
ORDER BY b.TestDate DESC, b.Id DESC, r.RowNo;";
        }
    }
}
