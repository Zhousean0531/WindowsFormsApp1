using System;
using System.Data.SqlClient;

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