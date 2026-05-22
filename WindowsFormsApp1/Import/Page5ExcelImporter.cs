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
        int importCount = 0;

        using (var wb = new XLWorkbook(excelPath))
        {
            var ws = wb.Worksheets
                .FirstOrDefault(x => string.Equals(x.Name, SheetName, StringComparison.OrdinalIgnoreCase));

            if (ws == null)
                throw new InvalidOperationException("Cannot find sheet: " + SheetName);

            var batches = new Dictionary<string, Page5ExportData>();
            var rawEfficiencies = new Dictionary<string, string>();

            int row = FirstDataRow;
            while (!IsEmptyRow(ws, row))
            {
                string testDate = GetDateText(ws.Cell(row, 1));
                string filterType = GetNullableText(ws.Cell(row, 2));
                string carbonLot = GetNullableText(ws.Cell(row, 3));
                string cylinderNo = GetNullableText(ws.Cell(row, 4));
                string rawEfficiency = GetEfficiencyText(ws.Cell(row, 8));

                string key = $"{testDate}|{filterType}|{carbonLot}|{cylinderNo}";

                if (!batches.TryGetValue(key, out Page5ExportData data))
                {
                    data = new Page5ExportData
                    {
                        TestDate = testDate,
                        ReportNo = null,
                        CylinderNo = cylinderNo,
                        Customer = null,
                        FilterType = filterType,
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
                    FillMissingBatchValues(data, testDate, filterType, carbonLot, cylinderNo);

                    if (string.IsNullOrWhiteSpace(rawEfficiencies[key]) &&
                        !string.IsNullOrWhiteSpace(rawEfficiency))
                    {
                        rawEfficiencies[key] = rawEfficiency;
                    }
                }

                data.Rows.Add(new Page5RowData
                {
                    SN = GetNullableText(ws.Cell(row, 5)),
                    Weight = GetNullableText(ws.Cell(row, 6)),
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

    private static void FillMissingBatchValues(
        Page5ExportData data,
        string testDate,
        string filterType,
        string carbonLot,
        string cylinderNo)
    {
        if (string.IsNullOrWhiteSpace(data.TestDate) && !string.IsNullOrWhiteSpace(testDate))
            data.TestDate = testDate;

        if (string.IsNullOrWhiteSpace(data.FilterType) && !string.IsNullOrWhiteSpace(filterType))
            data.FilterType = filterType;

        if (string.IsNullOrWhiteSpace(data.CarbonLot) && !string.IsNullOrWhiteSpace(carbonLot))
            data.CarbonLot = carbonLot;

        if (string.IsNullOrWhiteSpace(data.CylinderNo) && !string.IsNullOrWhiteSpace(cylinderNo))
            data.CylinderNo = cylinderNo;
    }

    private static Dictionary<int, string> BuildControlValues(IXLWorksheet ws, int row)
    {
        return new Dictionary<int, string>
        {
            [38] = null, // Particle_in 匯入時固定寫 NULL
            [39] = GetNullableText(ws.Cell(row, 10)),
            [40] = GetNullableText(ws.Cell(row, 11)),

            [41] = GetNullableText(ws.Cell(row, 12)),
            [42] = GetNullableText(ws.Cell(row, 13)),
            [43] = GetNullableText(ws.Cell(row, 14)),

            [44] = GetNullableText(ws.Cell(row, 15)),
            [45] = GetNullableText(ws.Cell(row, 16)),
            [46] = GetNullableText(ws.Cell(row, 17)),

            [47] = GetNullableText(ws.Cell(row, 18)),
            [48] = GetNullableText(ws.Cell(row, 19)),
            [49] = GetNullableText(ws.Cell(row, 20)),

            [50] = GetNullableText(ws.Cell(row, 21)),
            [54] = GetNullableText(ws.Cell(row, 22))
        };
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row)
    {
        for (int col = 1; col <= 22; col++)
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
        return string.IsNullOrWhiteSpace(text) ||
               text.Equals("N.D.", StringComparison.OrdinalIgnoreCase) ||
               text.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
               text == "-";
    }
}
