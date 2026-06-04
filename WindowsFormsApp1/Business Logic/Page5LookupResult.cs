using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

public class Page5LookupResult
{
    public bool Found => HeaderValues.Any();

    public List<Dictionary<int, string>> Rows { get; set; }
        = new List<Dictionary<int, string>>();

    public Dictionary<string, string> HeaderValues { get; set; }
        = new Dictionary<string, string>();
}
public static class Page5LookupHelper
{
    public static Page5LookupResult SearchByCylinderNo(string cylinderNo)
    {
        var result = new Page5LookupResult();

        using (var conn = DbBootstrap.GetConnection())
        {
            conn.Open();

            // 1. 先找最新一筆 Batch
            int batchId;

            string batchSql = @"
                SELECT TOP 1
                    Id,
                    TestDate,
                    ReportNo,
                    CylinderNo,
                    Customer,
                    FilterType,
                    COALESCE(NULLIF(RawMaterialType, ''), FilterType) AS RawMaterialType,
                    ReCylinderNo,
                    CarbonLot
                FROM P5_Batch
                WHERE CylinderNo = @CylinderNo
                ORDER BY Id DESC";

            using (var cmd = new SqlCommand(batchSql, conn))
            {
                cmd.Parameters.AddWithValue("@CylinderNo", cylinderNo);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        return result;

                    batchId = Convert.ToInt32(reader["Id"]);
                    result.HeaderValues["TestDate"] = reader["TestDate"]?.ToString();
                    result.HeaderValues["ReportNo"] = reader["ReportNo"]?.ToString();
                    result.HeaderValues["CylinderNo"] = reader["CylinderNo"]?.ToString();
                    result.HeaderValues["Customer"] = reader["Customer"]?.ToString();
                    result.HeaderValues["FilterType"] = reader["FilterType"]?.ToString();
                    result.HeaderValues["RawMaterialType"] = reader["RawMaterialType"]?.ToString();
                    result.HeaderValues["ReCylinderNo"] = reader["ReCylinderNo"]?.ToString();
                    result.HeaderValues["CarbonLot"] = reader["CarbonLot"]?.ToString();
                }
            }

            // 2. 再查 Row
            string rowSql = @"
                SELECT
                    RowNo,
                    SN,
                    Weight,
                    Efficiency,

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
                    PressureDrop
                FROM P5_Row
                WHERE BatchId = @BatchId
                ORDER BY RowNo";

            using (var cmd = new SqlCommand(rowSql, conn))
            {
                cmd.Parameters.AddWithValue("@BatchId", batchId);
                bool efficiencySet = false;
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (!efficiencySet)
                        {
                            string eff = reader["Efficiency"]?.ToString();

                            if (!string.IsNullOrWhiteSpace(eff))
                            {
                                result.HeaderValues["Efficiency"] = eff;
                                efficiencySet = true;
                            }
                        }
                        var dict = new Dictionary<int, string>();
                        dict[1] = reader["SN"]?.ToString();
                        dict[2] = reader["Weight"]?.ToString();
                        dict[3] = reader["ParticleIn"]?.ToString();
                        dict[4] = reader["ParticleOut"]?.ToString();
                        dict[5] = reader["IPAIn"]?.ToString();
                        dict[6] = reader["IPAOut"]?.ToString();
                        dict[7] = reader["AcetoneIn"]?.ToString();
                        dict[8] = reader["AcetoneOut"]?.ToString();
                        dict[9] = reader["NonTargetIn"]?.ToString();
                        dict[10] = reader["NonTargetOut"]?.ToString();
                        dict[11] = reader["PressureDrop"]?.ToString();
                        result.Rows.Add(dict);
                    }
                }
            }
        }

        return result;
    }
}
