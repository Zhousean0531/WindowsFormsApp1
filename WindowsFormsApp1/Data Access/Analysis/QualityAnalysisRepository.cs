using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1.Data_Access.Analysis
{
    public static class QualityAnalysisRepository
    {
        public static QualityAnalysisResult Query(
            string material,
            string metricName,
            DateTime startDate,
            DateTime endDate)
        {
            using (var conn = DbBootstrap.GetConnection())
            {
                conn.Open();

                var setting = LoadSetting(conn, material, metricName);
                DateTime sigmaStart;
                DateTime sigmaEnd;
                ResolveSigmaRange(setting, startDate, endDate, out sigmaStart, out sigmaEnd);

                var points = LoadPoints(conn, material, metricName, startDate, endDate);
                var sigmaPoints = LoadPoints(conn, material, metricName, sigmaStart, sigmaEnd);

                var result = new QualityAnalysisResult
                {
                    Points = points,
                    SigmaPoints = sigmaPoints,
                    Setting = setting,
                    SigmaStartDate = sigmaStart,
                    SigmaEndDate = sigmaEnd
                };

                ApplyStats(result);
                return result;
            }
        }

        private static QualityAnalysisSetting LoadSetting(SqlConnection conn, string material, string metricName)
        {
            const string sql = @"
SELECT TOP 1
    Material,
    MetricName,
    SigmaMode,
    SigmaMonths,
    SigmaStartDate,
    SigmaEndDate,
    USL,
    LSL
FROM QualityAnalysisSetting
WHERE LTRIM(RTRIM(Material)) = LTRIM(RTRIM(@Material))
  AND LTRIM(RTRIM(MetricName)) = LTRIM(RTRIM(@MetricName));";

            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@Material", SqlDbType.NVarChar, 50).Value = material;
                cmd.Parameters.Add("@MetricName", SqlDbType.NVarChar, 50).Value = metricName;

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        return null;

                    return new QualityAnalysisSetting
                    {
                        Material = reader["Material"] as string,
                        MetricName = reader["MetricName"] as string,
                        SigmaMode = reader["SigmaMode"] as string,
                        SigmaMonths = ToNullableInt(reader["SigmaMonths"]),
                        SigmaStartDate = ToNullableDate(reader["SigmaStartDate"]),
                        SigmaEndDate = ToNullableDate(reader["SigmaEndDate"]),
                        USL = ToNullableDouble(reader["USL"]),
                        LSL = ToNullableDouble(reader["LSL"])
                    };
                }
            }
        }

        private static void ResolveSigmaRange(
            QualityAnalysisSetting setting,
            DateTime queryStart,
            DateTime queryEnd,
            out DateTime sigmaStart,
            out DateTime sigmaEnd)
        {
            sigmaStart = queryStart.Date;
            sigmaEnd = queryEnd.Date;

            if (setting == null)
                return;

            string mode = (setting.SigmaMode ?? "").Trim();

            if (mode.Equals("Custom", StringComparison.OrdinalIgnoreCase) &&
                setting.SigmaStartDate.HasValue &&
                setting.SigmaEndDate.HasValue)
            {
                sigmaStart = setting.SigmaStartDate.Value.Date;
                sigmaEnd = setting.SigmaEndDate.Value.Date;
                return;
            }

            if (mode.Equals("LastMonths", StringComparison.OrdinalIgnoreCase) &&
                setting.SigmaMonths.HasValue &&
                setting.SigmaMonths.Value > 0)
            {
                sigmaEnd = queryEnd.Date;
                sigmaStart = sigmaEnd.AddMonths(-setting.SigmaMonths.Value);
            }
        }

        private static List<QualityAnalysisPoint> LoadPoints(
            SqlConnection conn,
            string material,
            string metricName,
            DateTime startDate,
            DateTime endDate)
        {
            string sql = BuildPointsSql(metricName);
            var points = new List<QualityAnalysisPoint>();

            if (string.IsNullOrWhiteSpace(sql))
                return points;

            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@Material", SqlDbType.NVarChar, 100).Value = material;
                cmd.Parameters.Add("@StartDate", SqlDbType.Date).Value = startDate.Date;
                cmd.Parameters.Add("@EndDate", SqlDbType.Date).Value = endDate.Date;

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (reader["TestDate"] == DBNull.Value)
                            continue;

                        double value;
                        if (!TryParseNumber(reader["ValueText"], out value))
                            continue;

                        points.Add(new QualityAnalysisPoint
                        {
                            TestDate = Convert.ToDateTime(reader["TestDate"]).Date,
                            Value = value,
                            Source = reader["Source"] as string,
                            BatchId = ToNullableInt(reader["BatchId"]),
                            RowNo = ToNullableInt(reader["RowNo"])
                        });
                    }
                }
            }

            points = points
                .OrderBy(x => x.TestDate)
                .ThenBy(x => x.Source)
                .ThenBy(x => x.BatchId ?? 0)
                .ThenBy(x => x.RowNo ?? 0)
                .ToList();

            return ApplyPointSampling(points, metricName);
        }

        private static List<QualityAnalysisPoint> ApplyPointSampling(
            List<QualityAnalysisPoint> points,
            string metricName)
        {
            if (!IsCylinderFinishedSampledMetric(metricName))
                return points;

            var result = new List<QualityAnalysisPoint>();
            result.AddRange(points.Where(x => !IsSource(x, "P5")));

            var p5Groups = points
                .Where(x => IsSource(x, "P5"))
                .GroupBy(x => x.BatchId ?? -1)
                .OrderBy(g => g.Min(x => x.TestDate))
                .ThenBy(g => g.Key);

            foreach (var group in p5Groups)
            {
                var ordered = group
                    .OrderBy(x => x.RowNo ?? int.MaxValue)
                    .ThenBy(x => x.TestDate)
                    .ToList();

                for (int i = 0; i < ordered.Count; i++)
                {
                    if (ShouldKeepP5Point(ordered[i], i))
                        result.Add(ordered[i]);
                }
            }

            return result
                .OrderBy(x => x.TestDate)
                .ThenBy(x => x.Source)
                .ThenBy(x => x.BatchId ?? 0)
                .ThenBy(x => x.RowNo ?? 0)
                .ToList();
        }

        private static bool IsCylinderFinishedSampledMetric(string metricName)
        {
            string metric = (metricName ?? "").Trim();
            return metric == "PressureDrop" || metric == "Particle" || metric == "TVOC";
        }

        private static bool IsSource(QualityAnalysisPoint point, string source)
        {
            return string.Equals(point.Source, source, StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldKeepP5Point(QualityAnalysisPoint point, int index)
        {
            if (point.RowNo.HasValue && point.RowNo.Value > 0)
                return (point.RowNo.Value - 1) % 16 == 0;

            return index % 16 == 0;
        }

        private static string BuildPointsSql(string metricName)
        {
            switch ((metricName ?? "").Trim())
            {
                case "Weight":
                    return BuildWeightSql();
                case "PressureDrop":
                    return BuildPressureDropSql();
                case "Particle":
                    return BuildParticleSql();
                case "TVOC":
                    return BuildTvocSql();
                case "Efficiency":
                    return BuildEfficiencySql();
                default:
                    return "";
            }
        }

        private static string BuildWeightSql()
        {
            return @"
SELECT b.TestingDate AS TestDate, N'P1' AS Source, b.Id AS BatchId, s.Id AS RowNo, CONVERT(NVARCHAR(100), s.Weight) AS ValueText
FROM P1_Batch b
INNER JOIN P1_Sample s ON s.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P2' AS Source, b.Id AS BatchId, s.Id AS RowNo, CONVERT(NVARCHAR(100), s.Weight) AS ValueText
FROM P2_Batch b
INNER JOIN P2_GasTest g ON g.BatchId = b.Id
INNER JOIN P2_Sample s ON s.GasTestId = g.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P3' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.Weight) AS ValueText
FROM P3_Batch b
INNER JOIN P3_Row r ON r.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.FilterMaterialNo, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P4' AS Source, b.Id AS BatchId, l.Id AS RowNo, CONVERT(NVARCHAR(100), l.Weight) AS ValueText
FROM P4_Batch b
INNER JOIN P4_Lot l ON l.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P5' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.Weight) AS ValueText
FROM P5_Batch b
INNER JOIN P5_Row r ON r.BatchId = b.Id
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
WHERE (
        LTRIM(RTRIM(ISNULL(p.ProductNo, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
        OR LTRIM(RTRIM(ISNULL(b.FilterType, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
      )
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate;";
        }

        private static string BuildPressureDropSql()
        {
            return @"
SELECT b.TestingDate AS TestDate, N'P1' AS Source, b.Id AS BatchId, s.Id AS RowNo, CONVERT(NVARCHAR(100), s.PressureDrop) AS ValueText
FROM P1_Batch b
INNER JOIN P1_Sample s ON s.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P2' AS Source, b.Id AS BatchId, s.Id AS RowNo, CONVERT(NVARCHAR(100), s.PressureDrop) AS ValueText
FROM P2_Batch b
INNER JOIN P2_GasTest g ON g.BatchId = b.Id
INNER JOIN P2_Sample s ON s.GasTestId = g.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P3' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.PressureDrop) AS ValueText
FROM P3_Batch b
INNER JOIN P3_Row r ON r.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.FilterMaterialNo, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P4' AS Source, b.Id AS BatchId, l.Id AS RowNo, CONVERT(NVARCHAR(100), l.DeltaP) AS ValueText
FROM P4_Batch b
INNER JOIN P4_Lot l ON l.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P5' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.PressureDrop) AS ValueText
FROM P5_Batch b
INNER JOIN P5_Row r ON r.BatchId = b.Id
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
WHERE (
        LTRIM(RTRIM(ISNULL(p.ProductNo, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
        OR LTRIM(RTRIM(ISNULL(b.FilterType, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
      )
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate;";
        }

        private static string BuildParticleSql()
        {
            return @"
SELECT b.TestingDate AS TestDate, N'P3' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.ParticleDiff) AS ValueText
FROM P3_Batch b
INNER JOIN P3_Row r ON r.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.FilterMaterialNo, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P5' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.ParticleDiff) AS ValueText
FROM P5_Batch b
INNER JOIN P5_Row r ON r.BatchId = b.Id
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
WHERE (
        LTRIM(RTRIM(ISNULL(p.ProductNo, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
        OR LTRIM(RTRIM(ISNULL(b.FilterType, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
      )
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate;";
        }

        private static string BuildTvocSql()
        {
            return @"
SELECT b.TestingDate AS TestDate, N'P1' AS Source, b.Id AS BatchId, s.Id AS RowNo, CONVERT(NVARCHAR(100), s.VOCOutgassing) AS ValueText
FROM P1_Batch b
INNER JOIN P1_Sample s ON s.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P3' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.TotalDiff) AS ValueText
FROM P3_Batch b
INNER JOIN P3_Row r ON r.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.FilterMaterialNo, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P4' AS Source, b.Id AS BatchId, l.Id AS RowNo, CONVERT(NVARCHAR(100), l.Outgassing) AS ValueText
FROM P4_Batch b
INNER JOIN P4_Lot l ON l.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P5' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.TotalDiff) AS ValueText
FROM P5_Batch b
INNER JOIN P5_Row r ON r.BatchId = b.Id
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
WHERE (
        LTRIM(RTRIM(ISNULL(p.ProductNo, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
        OR LTRIM(RTRIM(ISNULL(b.FilterType, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
      )
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate;";
        }

        private static string BuildEfficiencySql()
        {
            return @"
SELECT b.TestingDate AS TestDate, N'P1' AS Source, b.Id AS BatchId, e.SequenceIndex AS RowNo, CONVERT(NVARCHAR(100), e.EfficiencyValue) AS ValueText
FROM P1_Batch b
INNER JOIN P1_Sample s ON s.BatchId = b.Id
INNER JOIN P1_Efficiency e ON e.SampleId = s.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P2' AS Source, b.Id AS BatchId, e.SequenceIndex AS RowNo, CONVERT(NVARCHAR(100), e.EfficiencyValue) AS ValueText
FROM P2_Batch b
INNER JOIN P2_GasTest g ON g.BatchId = b.Id
INNER JOIN P2_Sample s ON s.GasTestId = g.Id
INNER JOIN P2_Efficiency e ON e.SampleId = s.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestingDate AS TestDate, N'P4' AS Source, b.Id AS BatchId, e.SequenceIndex AS RowNo, CONVERT(NVARCHAR(100), e.EfficiencyValue) AS ValueText
FROM P4_Batch b
INNER JOIN P4_Efficiency e ON e.BatchId = b.Id
WHERE LTRIM(RTRIM(ISNULL(b.Material, ''))) = LTRIM(RTRIM(@Material))
  AND CONVERT(date, b.TestingDate) BETWEEN @StartDate AND @EndDate

UNION ALL

SELECT b.TestDate AS TestDate, N'P5' AS Source, b.Id AS BatchId, r.RowNo AS RowNo, CONVERT(NVARCHAR(100), r.Efficiency) AS ValueText
FROM P5_Batch b
INNER JOIN P5_Row r ON r.BatchId = b.Id
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
WHERE (
        LTRIM(RTRIM(ISNULL(p.ProductNo, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
        OR LTRIM(RTRIM(ISNULL(b.FilterType, ''))) LIKE '%' + LTRIM(RTRIM(@Material)) + '%'
      )
  AND CONVERT(date, b.TestDate) BETWEEN @StartDate AND @EndDate;";
        }

        private static void ApplyStats(QualityAnalysisResult result)
        {
            if (result.SigmaPoints == null || result.SigmaPoints.Count == 0)
                return;

            var values = result.SigmaPoints.Select(x => x.Value).ToList();
            double avg = values.Average();
            double variance = 0;

            if (values.Count > 1)
                variance = values.Sum(x => Math.Pow(x - avg, 2)) / (values.Count - 1);

            result.Average = avg;
            result.Sigma = Math.Sqrt(variance);
        }

        private static bool TryParseNumber(object value, out double number)
        {
            number = 0;

            string text = Convert.ToString(value)?.Trim();
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string upper = text.ToUpperInvariant();
            if (upper == "N.D." || upper == "ND" || upper == "N/A" || upper == "NA" || upper == "-")
                return false;

            text = text
                .Replace("，", ",")
                .Replace(",", "")
                .Replace("＜", "<")
                .Replace("＞", ">")
                .Replace("－", "-")
                .Replace("μ", "")
                .Replace("µ", "");

            var match = Regex.Match(text, @"[-+]?\d+(\.\d+)?");
            if (!match.Success)
                return false;

            return double.TryParse(
                match.Value,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out number);
        }

        private static int? ToNullableInt(object value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return Convert.ToInt32(value);
        }

        private static DateTime? ToNullableDate(object value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return Convert.ToDateTime(value);
        }

        private static double? ToNullableDouble(object value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return Convert.ToDouble(value);
        }
    }
}
