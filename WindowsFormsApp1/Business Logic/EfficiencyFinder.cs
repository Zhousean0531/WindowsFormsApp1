using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

public static class P3EfficiencyFinder
{
    public static Dictionary<string, string> FindMinEfficiencyByWorkOrderText(string text)
    {
        var result = new Dictionary<string, string>
        {
            { "MA", "" },
            { "MB", "" },
            { "MC", "" }
        };

        var minMap = new Dictionary<string, double>
        {
            { "MA", double.MaxValue },
            { "MB", double.MaxValue },
            { "MC", double.MaxValue }
        };

        var lots = SplitLots(text);

        if (lots.Count == 0)
            return result;

        using (var conn = DbBootstrap.GetConnection())
        {
            conn.Open();

            foreach (var lot in lots)
            {
                var records = FindP2EfficiencyBySingleLot(conn, lot);

                foreach (var item in records)
                {
                    string gasType = NormalizeGasType(item.GasType);

                    if (!minMap.ContainsKey(gasType))
                        continue;

                    double eff;

                    if (!double.TryParse(item.Efficiency, NumberStyles.Any, CultureInfo.InvariantCulture, out eff))
                        continue;

                    if (eff < minMap[gasType])
                    {
                        minMap[gasType] = eff;
                        result[gasType] = item.Efficiency;
                    }
                }
            }
        }

        return result;
    }

    private static List<string> SplitLots(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return new List<string>();

        return text
            .Split(new[] { '/', ',', '，', ';', '；', '\r', '\n', ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct()
            .ToList();
    }

    private static List<P2EfficiencyRecord> FindP2EfficiencyBySingleLot(SqlConnection conn, string lot)
    {
        var list = new List<P2EfficiencyRecord>();

        string sql = @"
            SELECT
                x.GasName AS GasType,
                x.EfficiencyValue AS Efficiency
            FROM
            (
                SELECT
                    g.GasName,
                    e.EfficiencyValue,
                    ROW_NUMBER() OVER (
                        PARTITION BY g.GasName
                        ORDER BY e.Id ASC
                    ) AS rn
                FROM P2_Batch b
                INNER JOIN P2_GasTest g ON g.BatchId = b.Id
                INNER JOIN P2_Sample s ON s.GasTestId = g.Id
                INNER JOIN P2_Efficiency e ON e.SampleId = s.Id
                WHERE b.WorkOrder = @WorkOrder
                  AND e.EfficiencyValue IS NOT NULL
            ) x
            WHERE x.rn = 1;
        ";

        using (var cmd = new SqlCommand(sql, conn))
        {
            cmd.Parameters.AddWithValue("@WorkOrder", lot);

            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    list.Add(new P2EfficiencyRecord
                    {
                        GasType = reader["GasType"] == DBNull.Value ? "" : reader["GasType"].ToString(),
                        Efficiency = reader["Efficiency"] == DBNull.Value ? "" : reader["Efficiency"].ToString()
                    });
                }
            }
        }

        return list;
    }

    private static string NormalizeGasType(string gasType)
    {
        if (string.IsNullOrWhiteSpace(gasType))
            return "";

        gasType = gasType.Trim().ToUpper();

        if (gasType == "SO2" || gasType == "H2S")
            return "MA";

        if (gasType == "NH3")
            return "MB";

        if (gasType == "TOLUENE" || gasType == "ACETONE" || gasType == "IPA")
            return "MC";

        if (gasType.Contains("MA"))
            return "MA";

        if (gasType.Contains("MB"))
            return "MB";

        if (gasType.Contains("MC"))
            return "MC";

        return "";
    }
    private class P2EfficiencyRecord
    {
        public string GasType { get; set; }
        public string Efficiency { get; set; }
    }
}
public static class EfficiencyFinder
{
    public static string FindMinEfficiencyByCarbonLot(
        string carbonLot,
        string filterType)
    {
        if (string.IsNullOrWhiteSpace(carbonLot) || string.IsNullOrWhiteSpace(filterType))
            return null;

        using (var conn = DbBootstrap.GetConnection())
        {
            conn.Open();

            string sql = @"
                SELECT MIN(TRY_CONVERT(float, e.EfficiencyValue)) AS MinEfficiency
                FROM P4_Batch b
                INNER JOIN P4_Lot l ON l.BatchId = b.Id
                INNER JOIN P4_Efficiency e ON e.BatchId = b.Id
                WHERE l.LotFull LIKE @CarbonLot + '#%'
                  AND (
                        (@FilterType = 'MA' AND e.GasName IN ('SO2', 'H2S'))
                     OR (@FilterType = 'MB' AND e.GasName = 'NH3')
                     OR (@FilterType = 'MC' AND e.GasName IN ('Toluene', 'Acetone', 'IPA'))
                  )
                  AND e.EfficiencyValue IS NOT NULL";

            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@CarbonLot", carbonLot.Trim());
                cmd.Parameters.AddWithValue("@FilterType", filterType.Trim());

                object value = cmd.ExecuteScalar();

                if (value == null || value == DBNull.Value)
                    return null;

                double eff = Convert.ToDouble(value);
                return eff.ToString("0.###");
            }
        }
    }
}