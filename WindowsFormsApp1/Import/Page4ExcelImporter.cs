using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page4;

public static class Page4ExcelImporter
{
    private const string SheetName = "濾筒原料";
    private const int FirstDataRow = 4;

    public static int ImportWithFileDialog(IWin32Window owner = null)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "選擇 Page4 匯入 Excel";
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

            var batches = new Dictionary<string, P4Batch>();
            var gasGroups = new Dictionary<string, Dictionary<string, P4EfficiencyGroup>>();

            string currentTestingDate = "";
            string currentArrivalDate = "";
            string currentMaterial = "";
            string currentGasName = "";

            int row = FirstDataRow;
            while (!IsEmptyRow(ws, row))
            {
                string testingDate = GetDateText(ws.Cell(row, 1));
                string arrivalDate = GetDateText(ws.Cell(row, 2));
                string material = GetText(ws.Cell(row, 3));

                if (!string.IsNullOrWhiteSpace(testingDate))
                    currentTestingDate = testingDate;

                if (!string.IsNullOrWhiteSpace(arrivalDate))
                    currentArrivalDate = arrivalDate;

                if (!string.IsNullOrWhiteSpace(material))
                    currentMaterial = material;

                testingDate = currentTestingDate;
                arrivalDate = currentArrivalDate;
                material = currentMaterial;

                string supplierLot = GetNullableText(ws.Cell(row, 4));
                string lotFull = GetNullableText(ws.Cell(row, 5));
                string lotRoot = GetLotRoot(lotFull);
                string key = $"{testingDate}|{arrivalDate}|{material}|{lotRoot}";

                if (!batches.TryGetValue(key, out P4Batch batch))
                {
                    batch = new P4Batch
                    {
                        ReportNo = null,
                        Material = material,
                        MaterialNo = null,
                        ArrivalDate = arrivalDate,
                        TestingDate = testingDate,
                        QtyText = null,
                        UserName = Environment.UserName,
                        SupplierLot = supplierLot,
                        FactoryLot = lotRoot
                    };

                    batches.Add(key, batch);
                    gasGroups.Add(key, new Dictionary<string, P4EfficiencyGroup>(StringComparer.OrdinalIgnoreCase));
                }
                else
                {
                    FillMissingBatchValues(batch, testingDate, arrivalDate, material, supplierLot, lotRoot);
                }

                string moisture = GetNullableText(ws.Cell(row, 8));
                string butane = GetNullableText(ws.Cell(row, 9));
                string ash = GetNullableText(ws.Cell(row, 10));

                if (string.IsNullOrWhiteSpace(batch.Moisture) && !string.IsNullOrWhiteSpace(moisture))
                    batch.Moisture = moisture;

                if (string.IsNullOrWhiteSpace(batch.Butane) && !string.IsNullOrWhiteSpace(butane))
                    batch.Butane = butane;

                if (string.IsNullOrWhiteSpace(batch.Ash) && !string.IsNullOrWhiteSpace(ash))
                    batch.Ash = ash;

                string particleText = GetText(ws.Cell(row, 18));
                if (batch.ParticleSizePercentages.Count == 0 && !string.IsNullOrWhiteSpace(particleText))
                    batch.ParticleSizePercentages = ParseParticleAnalysis(particleText);

                double? weight = GetDouble(ws.Cell(row, 6));
                double? density = GetDouble(ws.Cell(row, 7));
                double? vocIn = GetDouble(ws.Cell(row, 11));
                double? vocOut = GetDouble(ws.Cell(row, 12));
                string outgassing = GetOutgassing(ws.Cell(row, 13), vocIn, vocOut);
                double? deltaP = GetDouble(ws.Cell(row, 15));
                double? eff0 = GetEfficiency(ws.Cell(row, 16));
                double? eff10 = GetEfficiency(ws.Cell(row, 17));

                batch.Rows.Add(new P4Row
                {
                    LotNo = supplierLot,
                    LotFull = lotFull,
                    Weight = weight ?? 0,
                    Density = density ?? 0,
                    VocIn = vocIn ?? 0,
                    VocOut = vocOut ?? 0,
                    DeltaP = deltaP ?? 0,
                    Outgassing = outgassing,
                    IsSelected = eff0.HasValue || eff10.HasValue
                });

                string gasName = GetText(ws.Cell(row, 14));
                if (!string.IsNullOrWhiteSpace(gasName))
                    currentGasName = gasName;

                if (eff0.HasValue || eff10.HasValue)
                {
                    gasName = string.IsNullOrWhiteSpace(gasName) ? currentGasName : gasName;

                    if (string.IsNullOrWhiteSpace(gasName))
                        throw new InvalidOperationException($"濾筒原料第 {row} 列有效率資料，但 N 欄沒有測試氣體。");

                    var batchGasGroups = gasGroups[key];
                    if (!batchGasGroups.TryGetValue(gasName, out P4EfficiencyGroup group))
                    {
                        group = new P4EfficiencyGroup
                        {
                            GasName = gasName,
                            Concentration = null
                        };

                        batchGasGroups.Add(gasName, group);
                        batch.EfficiencyGroups.Add(group);
                    }

                    if (eff0.HasValue)
                    {
                        group.Eff0 = eff0;
                        group.EfficiencyPoints[0] = eff0.Value;
                    }

                    if (eff10.HasValue)
                    {
                        group.Eff10 = eff10;
                        group.EfficiencyPoints[10] = eff10.Value;
                    }
                }

                row++;
            }

            foreach (var batch in batches.Values)
            {
                P4Repository.Insert(batch);
                importCount++;
            }
        }

        return importCount;
    }

    private static void FillMissingBatchValues(
        P4Batch batch,
        string testingDate,
        string arrivalDate,
        string material,
        string supplierLot,
        string factoryLot)
    {
        if (string.IsNullOrWhiteSpace(batch.TestingDate) && !string.IsNullOrWhiteSpace(testingDate))
            batch.TestingDate = testingDate;

        if (string.IsNullOrWhiteSpace(batch.ArrivalDate) && !string.IsNullOrWhiteSpace(arrivalDate))
            batch.ArrivalDate = arrivalDate;

        if (string.IsNullOrWhiteSpace(batch.Material) && !string.IsNullOrWhiteSpace(material))
            batch.Material = material;

        if (string.IsNullOrWhiteSpace(batch.SupplierLot) && !string.IsNullOrWhiteSpace(supplierLot))
            batch.SupplierLot = supplierLot;

        if (string.IsNullOrWhiteSpace(batch.FactoryLot) && !string.IsNullOrWhiteSpace(factoryLot))
            batch.FactoryLot = factoryLot;
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row)
    {
        for (int col = 1; col <= 18; col++)
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

    private static string GetNullableText(IXLCell cell)
    {
        string text = GetText(cell);
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static string GetDateText(IXLCell cell)
    {
        if (cell != null && cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime().ToString("yyyy-MM-dd");

        string text = GetText(cell);
        if (string.IsNullOrWhiteSpace(text))
            return "";

        if (DateTime.TryParse(text, out DateTime date))
            return date.ToString("yyyy-MM-dd");

        if (DateTime.TryParse(text.Replace(".", "/"), out date))
            return date.ToString("yyyy-MM-dd");

        return text;
    }

    private static double? GetDouble(IXLCell cell)
    {
        if (cell == null || cell.IsEmpty())
            return null;

        if (cell.DataType == XLDataType.Number)
            return cell.GetDouble();

        string text = GetText(cell);
        if (IsEmptyValue(text))
            return null;

        text = text.Replace("%", "").Trim();

        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            return value;

        if (double.TryParse(text, out value))
            return value;

        return null;
    }

    private static double? GetEfficiency(IXLCell cell)
    {
        double? value = GetDouble(cell);
        if (!value.HasValue)
            return null;

        if (Math.Abs(value.Value) <= 1)
            value *= 100.0;

        return Math.Round(value.Value, 1);
    }

    private static string GetOutgassing(IXLCell cell, double? vocIn, double? vocOut)
    {
        string text = GetNullableText(cell);
        if (!string.IsNullOrWhiteSpace(text))
            return text;

        if (!vocIn.HasValue || !vocOut.HasValue)
            return null;

        double value = vocOut.Value - vocIn.Value;
        return value <= 0 ? "N.D." : value.ToString("0.0", CultureInfo.InvariantCulture);
    }

    private static bool IsEmptyValue(string text)
    {
        return string.IsNullOrWhiteSpace(text) ||
               text.Equals("N.D.", StringComparison.OrdinalIgnoreCase) ||
               text.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
               text == "-";
    }

    private static string GetLotRoot(string lotFull)
    {
        if (string.IsNullOrWhiteSpace(lotFull))
            return "";

        int index = lotFull.IndexOf('#');
        return index >= 0 ? lotFull.Substring(0, index) : lotFull;
    }

    private static Dictionary<string, double> ParseParticleAnalysis(string text)
    {
        var result = new Dictionary<string, double>();

        if (string.IsNullOrWhiteSpace(text))
            return result;

        string normalized = text
            .Replace("_x000D_", "\n")
            .Replace("\r", "\n");

        var parts = normalized.Split(new[] { ',', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        foreach (string rawPart in parts)
        {
            string part = rawPart.Trim();
            var match = Regex.Match(part, @"^(?<name>.+?)\s+(?<value>-?\d+(\.\d+)?)\s*%?$");

            if (!match.Success)
                continue;

            string name = match.Groups["name"].Value.Trim();
            if (double.TryParse(match.Groups["value"].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double percent))
                result[name] = percent;
        }

        return result;
    }
}
