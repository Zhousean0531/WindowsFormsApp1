using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page2;

public static class Page2ExcelImporter
{
    private const string SheetName = "濾網圓片";

    public static int ImportWithFileDialog(IWin32Window owner = null)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "Select Page2 Excel";
            ofd.Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*";
            ofd.Multiselect = false;

            DialogResult result = owner == null
                ? ofd.ShowDialog()
                : ofd.ShowDialog(owner);

            if (result != DialogResult.OK)
                return 0;

            return ImportFromExcel(ofd.FileName);
        }
    }

    public static int ImportFromExcel(string excelPath)
    {
        int importCount = 0;

        using (var wb = new XLWorkbook(excelPath))
        {
            var ws = wb.Worksheets
                .FirstOrDefault(x => string.Equals(x.Name, SheetName, StringComparison.OrdinalIgnoreCase));

            if (ws == null)
                throw new InvalidOperationException("Cannot find sheet: " + SheetName);

            var batches = new Dictionary<string, P2Batch>();
            var gasTests = new Dictionary<string, Dictionary<string, P2GasTest>>();

            int row = 2;

            while (!IsEmptyRow(ws, row))
            {
                DateTime? productionDate = GetDate(ws.Cell(row, 1));
                DateTime? testDate = GetDate(ws.Cell(row, 2)) ?? productionDate;
                string workOrder = GetText(ws.Cell(row, 3));
                string material = GetText(ws.Cell(row, 4));
                string batchNo = GetText(ws.Cell(row, 5));
                string materialNo = null;
                string productType = GetText(ws.Cell(row, 6));
                decimal? targetGsm = GetDecimal(ws.Cell(row, 7));
                string glue = GetText(ws.Cell(row, 8));
                decimal? speed = GetDecimal(ws.Cell(row, 9));
                decimal? upperTemp = GetDecimal(ws.Cell(row, 10));
                decimal? lowerTemp = GetDecimal(ws.Cell(row, 11));
                decimal? pressure = GetDecimal(ws.Cell(row, 12));
                decimal? windSpeed = GetDecimal(ws.Cell(row, 13));
                decimal? weight = GetDecimal(ws.Cell(row, 14));
                decimal? pressureDrop = GetDecimal(ws.Cell(row, 15));
                string gasName = GetText(ws.Cell(row, 16));
                decimal? concentration = GetDecimal(ws.Cell(row, 17));
                decimal? eff0 = GetEfficiency(ws.Cell(row, 18));
                decimal? eff10 = GetEfficiency(ws.Cell(row, 19));
                string carbonLine = GetText(ws.Cell(row, 20));

                if (string.IsNullOrWhiteSpace(gasName))
                {
                    row++;
                    continue;
                }

                string key =
                    $"{FormatDateKey(productionDate)}|{FormatDateKey(testDate)}|{workOrder}|{material}|{batchNo}|{materialNo}|{targetGsm}|{glue}|{speed}|{upperTemp}|{lowerTemp}|{pressure}|{windSpeed}|{productType}";

                if (!batches.TryGetValue(key, out P2Batch batch))
                {
                    batch = new P2Batch
                    {
                        ProductionDate = productionDate,
                        TestDate = testDate,
                        WorkOrder = workOrder,
                        Material = material,
                        MaterialNo = materialNo,
                        BatchNo = batchNo,
                        ProductDisplay = productType,
                        ProductSize = productType,
                        TargetGsm = targetGsm,
                        Glue = glue,
                        Speed = speed,
                        UpperTemp = upperTemp,
                        LowerTemp = lowerTemp,
                        Pressure = pressure,
                        WindSpeed = windSpeed,
                        CarbonLine = carbonLine,
                        ReportNo = null,
                        Username = Environment.UserName
                    };

                    batches.Add(key, batch);
                    gasTests.Add(key, new Dictionary<string, P2GasTest>(StringComparer.OrdinalIgnoreCase));
                }
                else
                {
                    FillMissingBatchValues(batch, testDate, carbonLine);
                }

                var batchGasTests = gasTests[key];

                if (!batchGasTests.TryGetValue(gasName, out P2GasTest gasTest))
                {
                    gasTest = new P2GasTest
                    {
                        GasName = gasName,
                        Concentration = concentration,
                        Background = 0m
                    };

                    batchGasTests.Add(gasName, gasTest);
                    batch.GasTests.Add(gasTest);
                }
                else if (!gasTest.Concentration.HasValue && concentration.HasValue)
                {
                    gasTest.Concentration = concentration;
                }

                var sample = new P2Sample
                {
                    Weight = weight,
                    PressureDrop = pressureDrop,
                    IsSelected = eff0.HasValue || eff10.HasValue
                };

                if (eff0.HasValue)
                    sample.EfficiencyPoints[0] = eff0.Value;

                if (eff10.HasValue)
                    sample.EfficiencyPoints[10] = eff10.Value;

                gasTest.Samples.Add(sample);

                row++;
            }

            foreach (var batch in batches.Values)
            {
                P2Repository.Insert(batch);
                importCount++;
            }
        }

        return importCount;
    }

    private static void FillMissingBatchValues(P2Batch batch, DateTime? testDate, string carbonLine)
    {
        if (!batch.TestDate.HasValue && testDate.HasValue)
        {
            batch.TestDate = testDate;
        }

        if (string.IsNullOrWhiteSpace(batch.CarbonLine) &&
            !string.IsNullOrWhiteSpace(carbonLine))
        {
            batch.CarbonLine = carbonLine;
        }
    }

    private static string BuildReportNo(DateTime? testDate)
    {
        if (!testDate.HasValue)
            return "";

        return "YS-Q-M20-" + testDate.Value.ToString("yyMMdd");
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row)
    {
        for (int col = 1; col <= 20; col++)
        {
            if (!ws.Cell(row, col).IsEmpty())
                return false;
        }

        return true;
    }

    private static string GetText(IXLCell cell)
    {
        if (cell == null || cell.IsEmpty())
            return "";

        if (cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime().ToString("yyyy.MM.dd");

        if (cell.DataType == XLDataType.Number)
            return cell.GetDouble().ToString("0.###", CultureInfo.InvariantCulture);

        return cell.GetString().Trim();
    }

    private static DateTime? GetDate(IXLCell cell)
    {
        if (cell == null || cell.IsEmpty())
            return null;

        if (cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime();

        string text = GetText(cell);

        if (DateTime.TryParse(text, out DateTime date))
            return date;

        if (DateTime.TryParse(text.Replace(".", "/"), out date))
            return date;

        return null;
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
            text.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
            text == "-")
        {
            return null;
        }

        text = text.Replace("%", "").Trim();

        if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal value))
            return value;

        if (decimal.TryParse(text, out value))
            return value;

        return null;
    }

    private static decimal? GetEfficiency(IXLCell cell)
    {
        decimal? value = GetDecimal(cell);

        if (!value.HasValue)
            return null;

        if (Math.Abs(value.Value) <= 1)
            value *= 100m;

        return Math.Round(value.Value, 1);
    }

    private static string FormatDateKey(DateTime? date)
    {
        return date.HasValue
            ? date.Value.ToString("yyyyMMdd")
            : "";
    }
}
