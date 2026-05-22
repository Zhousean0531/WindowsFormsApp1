using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using WindowsFormsApp1.Data_Access.Page1;

public static class Page1ExcelImporter
{
    public static int ImportFromExcel(string excelPath)
    {
        int importCount = 0;

        using (var wb = new XLWorkbook(excelPath))
        {
            var ws = wb.Worksheets.First();
            var batches = new Dictionary<string, P1Batch>();

            int row = 2;
            while (!IsEmptyRow(ws, row))
            {
                DateTime arrivalDate = GetDate(ws.Cell(row, 1), "ArrivalDate");
                DateTime testingDate = GetDate(ws.Cell(row, 2), "TestingDate");
                string reportNo = GetText(ws.Cell(row, 3));
                string material = GetText(ws.Cell(row, 4));
                string suppliedNo = GetText(ws.Cell(row, 5));
                string batchNo = GetText(ws.Cell(row, 6));
                decimal? vocIn = GetDecimal(ws.Cell(row, 7));
                decimal? vocOut = GetDecimal(ws.Cell(row, 8));
                decimal? outgassing = GetDecimal(ws.Cell(row, 9));
                decimal? pressureDrop = GetDecimal(ws.Cell(row, 10));
                decimal? eff0 = GetEfficiency(ws.Cell(row, 11));
                decimal? eff10 = GetEfficiency(ws.Cell(row, 12));
                string particleText = GetText(ws.Cell(row, 13));

                string key =
                    $"{arrivalDate:yyyyMMdd}|{testingDate:yyyyMMdd}|{reportNo}|{material}|{particleText}";

                if (!batches.TryGetValue(key, out P1Batch batch))
                {
                    batch = new P1Batch
                    {
                        ArrivalDate = arrivalDate,
                        TestingDate = testingDate,
                        ReportNo = reportNo,
                        Material = material,
                        Username = Environment.UserName,
                        ParticleSizePercentages = ParseParticleAnalysis(particleText),
                        Samples = new List<P1Sample>()
                    };

                    batches.Add(key, batch);
                }

                batch.Samples.Add(new P1Sample
                {
                    LotFull = batchNo,
                    SuppliedNO = suppliedNo,
                    VocIn = vocIn,
                    VocOut = vocOut,
                    Outgassing = outgassing,
                    DeltaP = pressureDrop,
                    IsSelected = eff0.HasValue || eff10.HasValue,
                    Efficiencies = BuildEfficiencies(eff0, eff10)
                });

                row++;
            }

            foreach (var batch in batches.Values)
            {
                P1Repository.Insert(batch);
                importCount++;
            }
        }

        return importCount;
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row)
    {
        for (int col = 1; col <= 13; col++)
        {
            if (!ws.Cell(row, col).IsEmpty())
                return false;
        }

        return true;
    }

    private static string GetText(IXLCell cell)
    {
        return cell?.GetString()?.Trim() ?? "";
    }

    private static DateTime GetDate(IXLCell cell, string fieldName)
    {
        if (cell != null && cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime();

        string text = GetText(cell);
        if (DateTime.TryParse(text, out DateTime date))
            return date;

        if (DateTime.TryParse(text.Replace(".", "/"), out date))
            return date;

        throw new FormatException($"{fieldName} format error: {text}");
    }

    private static decimal? GetDecimal(IXLCell cell)
    {
        if (cell == null || cell.IsEmpty())
            return null;

        if (cell.DataType == XLDataType.Number)
            return Convert.ToDecimal(cell.GetDouble());

        string text = GetText(cell);
        if (string.IsNullOrWhiteSpace(text) ||
            text.Equals("N.D.", StringComparison.OrdinalIgnoreCase) ||
            text == "-")
        {
            return null;
        }

        text = text.Replace("%", "").Trim();

        if (decimal.TryParse(text, out decimal value))
            return value;

        return null;
    }

    private static decimal? GetEfficiency(IXLCell cell)
    {
        decimal? value = GetDecimal(cell);
        if (!value.HasValue)
            return null;

        if (Math.Abs(value.Value) <= 1)
            return value.Value * 100m;

        return value;
    }

    private static List<decimal?> BuildEfficiencies(decimal? eff0, decimal? eff10)
    {
        if (!eff0.HasValue && !eff10.HasValue)
            return null;

        var efficiencies = Enumerable.Repeat<decimal?>(null, 11).ToList();
        efficiencies[0] = eff0;
        efficiencies[10] = eff10;
        return efficiencies;
    }

    private static Dictionary<string, double> ParseParticleAnalysis(string text)
    {
        var result = new Dictionary<string, double>();

        if (string.IsNullOrWhiteSpace(text))
            return result;

        string normalized = text.Replace("\r", "\n");
        var parts = normalized.Split(new[] { ',', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        foreach (string rawPart in parts)
        {
            string part = rawPart.Trim();
            var match = Regex.Match(part, @"^(?<name>.+?)\s*(?<value>-?\d+(\.\d+)?)\s*%?$");

            if (!match.Success)
                continue;

            string name = match.Groups["name"].Value.Trim();
            if (double.TryParse(match.Groups["value"].Value, out double percent))
                result[name] = percent;
        }

        return result;
    }
}
