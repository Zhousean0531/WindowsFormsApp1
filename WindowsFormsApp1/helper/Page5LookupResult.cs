using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public class Page5LookupResult
{
    public bool Found => Rows.Any();

    // DGV 用
    public List<Dictionary<int, string>> Rows { get; set; }
        = new List<Dictionary<int, string>>();

    // 基本欄位用
    public Dictionary<string, string> HeaderValues { get; set; }
        = new Dictionary<string, string>();
}

public static class Page5LookupHelper
{
    public static Page5LookupResult SearchByCylinderNo(string cylinderNo)
    {
        string summaryPath = Path.Combine(
            Application.StartupPath,
            "總表.xlsx"
        );

        using (var wb = new XLWorkbook(summaryPath))
        {
            var ws = wb.Worksheet("濾筒");
            var matchedRows = ws.RowsUsed()
                .Where(r => r.Cell("W").GetString().Trim() == cylinderNo)
                .ToList();
            if (!matchedRows.Any())
                return new Page5LookupResult();
            var rowsWithBatchId = matchedRows
                .Where(r => !string.IsNullOrWhiteSpace(r.Cell("BD").GetString()))
                .ToList();
            List<IXLRow> latestRows;
            if (!rowsWithBatchId.Any())
            {
                latestRows = matchedRows;
            }
            else
            {
                var latestBatchId = rowsWithBatchId
                    .Max(r => r.Cell("BD").GetString().Trim());

                latestRows = rowsWithBatchId
                    .Where(r => r.Cell("BD").GetString().Trim() == latestBatchId)
                    .ToList();
            }
            var result = new Page5LookupResult();
            var headerRow = latestRows.First();
            string Get(string col) => headerRow.Cell(col).GetString().Trim();
            result.HeaderValues["U"] = Get("U");   // TestDate
            result.HeaderValues["V"] = Get("V");   // ReportNo
            result.HeaderValues["X"] = Get("X");
            result.HeaderValues["Y"] = Get("Y");   // Customer
            result.HeaderValues["AA"] = Get("AA");  // Type
            result.HeaderValues["AB"] = Get("AB");  // ReCylinderNo
            result.HeaderValues["AI"] = Get("AI");
            result.HeaderValues["AJ"] = Get("AJ");
            result.HeaderValues["AK"] = Get("AK");

            // ===== 4️⃣ DGV 欄位（只用最新那組）=====
            var map = new Dictionary<string, int>
            {
                ["AC"] = 1,
                ["AD"] = 2,
                ["AL"] = 3,
                ["AM"] = 4,
                ["AO"] = 5,
                ["AP"] = 6,
                ["AR"] = 7,
                ["AS"] = 8,
                ["AU"] = 9,
                ["AV"] = 10,
                ["BB"] = 11
            };

            foreach (var excelRow in latestRows)
            {
                var dict = new Dictionary<int, string>();
                foreach (var kv in map)
                {
                    dict[kv.Value] = excelRow.Cell(kv.Key).GetString();
                }
                result.Rows.Add(dict);
            }

            return result;
        }
    }
    private static bool TryGetCellDateTime(IXLCell cell, out DateTime dt)
    {
        dt = default;

        if (cell == null || cell.IsEmpty())
            return false;

        if (cell.DataType == XLDataType.DateTime)
        {
            dt = cell.GetDateTime();
            return true;
        }

        // 嘗試從文字解析
        return DateTime.TryParse(cell.GetString(), out dt);
    }

}

