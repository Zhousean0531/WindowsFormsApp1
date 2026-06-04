using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page5;

public static class Page5ExcelImporter
{
    private const string SheetName = "濾筒成品";
    private const int FirstDataRow = 2;

    public static int ImportWithFileDialog(IWin32Window owner = null)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "選擇 Page5 匯入 Excel";
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
        return ImportFromExcel(excelPath, SheetName);
    }

    public static int ImportFromExcel(string excelPath, string sheetName)
    {
        int importCount = 0;

        using (var wb = new XLWorkbook(excelPath))
        {
            var ws = FindWorksheet(wb, sheetName);

            if (ws == null)
                throw new InvalidOperationException("Cannot find sheet: " + sheetName);

            var batches = new Dictionary<string, Page5ExportData>();
            var rawEfficiencies = new Dictionary<string, string>();

            int row = FirstDataRow;
            while (!IsEmptyRow(ws, row))
            {
                string testDate = GetDateText(ws.Cell(row, 1));
                string filterType = GetNullableText(ws.Cell(row, 2));
                string materialNo = ResolveMaterialNoForImport(
                    filterType,
                    GetNullableText(ws.Cell(row, 3)),
                    row);
                string carbonLot = GetNullableText(ws.Cell(row, 4));
                string cylinderNo = GetNullableText(ws.Cell(row, 5));
                string rawEfficiency = GetEfficiencyText(ws.Cell(row, 9));

                string key = $"{testDate}|{filterType}|{materialNo}|{carbonLot}|{cylinderNo}";

                if (!batches.TryGetValue(key, out Page5ExportData data))
                {
                    data = new Page5ExportData
                    {
                        TestDate = testDate,
                        ReportNo = null,
                        CylinderNo = cylinderNo,
                        Customer = null,
                        FilterType = null,
                        RawMaterialType = filterType,
                        MaterialNo = materialNo,
                        ReCylinderNo = null,
                        CarbonLot = carbonLot,
                        UserName = Environment.UserName,
                        Rows = new List<Page5RowData>()
                    };

                    batches.Add(key, data);
                    rawEfficiencies.Add(key, rawEfficiency);
                }
                else
                {
                    FillMissingBatchValues(data, testDate, filterType, materialNo, carbonLot, cylinderNo);

                    if (string.IsNullOrWhiteSpace(rawEfficiencies[key]) &&
                        !string.IsNullOrWhiteSpace(rawEfficiency))
                    {
                        rawEfficiencies[key] = rawEfficiency;
                    }
                }

                data.Rows.Add(new Page5RowData
                {
                    SN = GetNullableText(ws.Cell(row, 6)),
                    Weight = GetNullableText(ws.Cell(row, 7)),
                    ControlValues = BuildControlValues(ws, row)
                });

                row++;
            }

            foreach (var kv in batches)
            {
                P5Repository.Insert(kv.Value, rawEfficiencies[kv.Key]);
                importCount++;
            }
        }

        return importCount;
    }

    private static IXLWorksheet FindWorksheet(XLWorkbook wb, string sheetName)
    {
        return wb.Worksheets
            .FirstOrDefault(x => string.Equals(x.Name, sheetName, StringComparison.OrdinalIgnoreCase));
    }

    private static void FillMissingBatchValues(
        Page5ExportData data,
        string testDate,
        string filterType,
        string materialNo,
        string carbonLot,
        string cylinderNo)
    {
        if (string.IsNullOrWhiteSpace(data.TestDate) && !string.IsNullOrWhiteSpace(testDate))
            data.TestDate = testDate;

        if (string.IsNullOrWhiteSpace(data.RawMaterialType) && !string.IsNullOrWhiteSpace(filterType))
            data.RawMaterialType = filterType;

        if (string.IsNullOrWhiteSpace(data.MaterialNo) && !string.IsNullOrWhiteSpace(materialNo))
            data.MaterialNo = materialNo;

        if (string.IsNullOrWhiteSpace(data.CarbonLot) && !string.IsNullOrWhiteSpace(carbonLot))
            data.CarbonLot = carbonLot;

        if (string.IsNullOrWhiteSpace(data.CylinderNo) && !string.IsNullOrWhiteSpace(cylinderNo))
            data.CylinderNo = cylinderNo;
    }

    private static string ResolveMaterialNoForImport(string filterType, string materialNoFromExcel, int row)
    {
        if (!string.IsNullOrWhiteSpace(materialNoFromExcel))
            return materialNoFromExcel.Trim();

        string materialNo = ResolveMaterialNoWithoutPrompt(filterType);

        if (materialNo == null)
        {
            throw new InvalidOperationException(
                $"濾筒成品第 {row} 列料號欄位空白，且原料種類「{filterType}」無法自動判斷料號。"
                + "大量匯入時請在 C 欄填入料號，例如 11A0C00Y000002、11B0B00Y000002、11T0C00Y000002、11D0S00Y000002。");
        }

        return materialNo;
    }

    private static string ResolveMaterialNoWithoutPrompt(string filterType)
    {
        string text = (filterType ?? "").Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(text))
            return null;

        var materialNos = new List<string>();

        foreach (string token in text.Split(new[] { '+', '/', ',', '，', '、' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string materialNo = ResolveSingleMaterialNoWithoutPrompt(token.Trim().ToUpperInvariant());

            if (materialNo == null)
                return null;

            if (!materialNos.Contains(materialNo))
                materialNos.Add(materialNo);
        }

        return materialNos.Count == 0 ? null : string.Join("+", materialNos);
    }

    private static string ResolveSingleMaterialNoWithoutPrompt(string type)
    {
        if (type == "ACID")
            return "11A0C00Y000002";

        if (type == "DMS")
            return "11D0S00Y000002";

        if (type == "MB" || type == "BASE")
            return "11B0B00Y000002";

        if (type == "MC" || type == "TOC")
            return "11T0C00Y000002";

        return null;
    }

    private static Dictionary<int, string> BuildControlValues(IXLWorksheet ws, int row)
    {
        return new Dictionary<int, string>
        {
            [38] = GetMeasurementText(ws.Cell(row, 10)),
            [39] = GetMeasurementText(ws.Cell(row, 11)),
            [40] = GetMeasurementText(ws.Cell(row, 12)),

            [41] = GetMeasurementText(ws.Cell(row, 13)),
            [42] = GetMeasurementText(ws.Cell(row, 14)),
            [43] = GetMeasurementText(ws.Cell(row, 15)),

            [44] = GetMeasurementText(ws.Cell(row, 16)),
            [45] = GetMeasurementText(ws.Cell(row, 17)),
            [46] = GetMeasurementText(ws.Cell(row, 18)),

            [47] = GetMeasurementText(ws.Cell(row, 19)),
            [48] = GetMeasurementText(ws.Cell(row, 20)),
            [49] = GetMeasurementText(ws.Cell(row, 21)),

            [50] = GetMeasurementText(ws.Cell(row, 22)),
            [54] = GetMeasurementText(ws.Cell(row, 23))
        };
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row)
    {
        for (int col = 1; col <= 23; col++)
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

        if (IsEmptyValue(text))
            return null;

        return text;
    }

    private static string GetMeasurementText(IXLCell cell)
    {
        string text = GetText(cell);

        if (IsBlankValue(text))
            return null;

        return text;
    }

    private static string GetDateText(IXLCell cell)
    {
        if (cell != null && cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime().ToString("yyyy-MM-dd");

        string text = GetText(cell);
        if (string.IsNullOrWhiteSpace(text))
            return null;

        if (DateTime.TryParse(text, out DateTime date))
            return date.ToString("yyyy-MM-dd");

        if (DateTime.TryParse(text.Replace(".", "/"), out date))
            return date.ToString("yyyy-MM-dd");

        return text;
    }

    private static string GetEfficiencyText(IXLCell cell)
    {
        string text = GetText(cell);

        if (IsEmptyValue(text))
            return null;

        text = text.Replace("%", "").Trim();

        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double value) ||
            double.TryParse(text, out value))
        {
            if (Math.Abs(value) <= 1)
                value *= 100.0;

            return Math.Round(value, 1).ToString("0.#", CultureInfo.InvariantCulture);
        }

        return text;
    }

    private static bool IsEmptyValue(string text)
    {
        return IsBlankValue(text) ||
               text.Equals("N.D.", StringComparison.OrdinalIgnoreCase) ||
               text.Equals("N/A", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsBlankValue(string text)
    {
        return string.IsNullOrWhiteSpace(text) || text == "-";
    }
}
