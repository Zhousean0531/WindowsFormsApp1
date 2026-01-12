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
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
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
            var emptyTimeRows = matchedRows
                .Where(r => r.Cell("BD").IsEmpty())
                .ToList();

            // ===== 1️⃣ 找出最新的更新時間 =====
            List<IXLRow> latestRows;

            if (emptyTimeRows.Any())
            {
                // ★ 規則 1：只要有空時間 → 視為最新
                latestRows = emptyTimeRows;
            }
            else
            {
                // ★ 規則 2：全部都有時間 → 取時間最大的那一組
                var latestTime = matchedRows
                    .Max(r => r.Cell("BD").GetDateTime());

                latestRows = matchedRows
                    .Where(r => r.Cell("BD").GetDateTime() == latestTime)
                    .ToList();
            }
            var result = new Page5LookupResult();

            // ===== 3️⃣ HeaderValues（用最新那一筆）=====
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
}

