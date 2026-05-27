using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

public static class Page3ExcelImporter
{
    private const string SheetName = "濾網成品";

    public static int ImportWithFileDialog(IWin32Window owner = null)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "選擇 Page3 匯入 Excel";
            ofd.Filter = "Excel 檔案 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有檔案 (*.*)|*.*";
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

            var batches = new Dictionary<string, Page3ExportData>();
            var alarms = new Dictionary<string, List<string>>();

            int row = 3;
            while (!IsEmptyRow(ws, row))
            {
                string arrivalDate = GetDateText(ws.Cell(row, 1));
                string testingDate = GetDateText(ws.Cell(row, 2));

                if (string.IsNullOrWhiteSpace(testingDate))
                    testingDate = arrivalDate;
                string reportNo = GetText(ws.Cell(row, 3));
                string carbonLot = GetText(ws.Cell(row, 4));
                string packageNo = GetText(ws.Cell(row, 5));
                string workOrder = GetText(ws.Cell(row, 6));
                string customer = GetText(ws.Cell(row, 7));
                string materialNo = GetText(ws.Cell(row, 8));
                string model = GetText(ws.Cell(row, 9));
                string reFilterNo = GetText(ws.Cell(row, 11));

                string key =
                    $"{testingDate}|{arrivalDate}|{reportNo}|{carbonLot}|{packageNo}|{workOrder}|{customer}|{materialNo}|{model}|{reFilterNo}";

                if (!batches.TryGetValue(key, out Page3ExportData data))
                {
                    data = new Page3ExportData
                    {
                        TestingDate = testingDate,
                        ArrivalDate = arrivalDate,
                        FilterReportNo = reportNo,
                        CarbonLot = carbonLot,
                        PackageNo = packageNo,
                        WorkOrder = workOrder,
                        Customer = customer,
                        FilterMaterialNo = materialNo,
                        Model = model,
                        ReFilterNo = reFilterNo,
                        Alarm = "N/A",
                        UserName = Environment.UserName
                    };

                    batches.Add(key, data);
                    alarms.Add(key, new List<string>());
                }

                string ma = GetEfficiencyText(ws.Cell(row, 22));
                string mb = GetEfficiencyText(ws.Cell(row, 23));
                string mc = GetEfficiencyText(ws.Cell(row, 24));

                if (string.IsNullOrWhiteSpace(data.MA) &&
                    !string.IsNullOrWhiteSpace(ma))
                    data.MA = ma;

                if (string.IsNullOrWhiteSpace(data.MB) &&
                    !string.IsNullOrWhiteSpace(mb))
                    data.MB = mb;

                if (string.IsNullOrWhiteSpace(data.MC) &&
                    !string.IsNullOrWhiteSpace(mc))
                    data.MC = mc;

                AddAlarms(ws, row, alarms[key]);

                data.Rows.Add(new Page3RowData
                {
                    SN = GetText(ws.Cell(row, 12)),
                    Weight = GetText(ws.Cell(row, 13)),
                    Length = GetText(ws.Cell(row, 18)),
                    Width = GetText(ws.Cell(row, 19)),
                    Height = GetText(ws.Cell(row, 20)),
                    Diagonal = GetText(ws.Cell(row, 21)),
                    ControlValues = BuildControlValues(ws, row)
                });

                row++;
            }

            foreach (var kv in batches)
            {
                if (alarms.TryGetValue(kv.Key, out List<string> alarmList) &&
                    alarmList.Count > 0)
                {
                    kv.Value.Alarm = string.Join("/", alarmList);
                }

                P3Repository.Insert(kv.Value);
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

    private static Dictionary<int, string> BuildControlValues(IXLWorksheet ws, int row)
    {
        string particleIn = GetText(ws.Cell(row, 25));
        string particleOut = GetText(ws.Cell(row, 26));
        string ipaIn = GetText(ws.Cell(row, 28));
        string ipaOut = GetText(ws.Cell(row, 29));
        string acetoneIn = GetText(ws.Cell(row, 31));
        string acetoneOut = GetText(ws.Cell(row, 32));
        string nonTargetIn = GetText(ws.Cell(row, 34));
        string nonTargetOut = GetText(ws.Cell(row, 35));
        string particleDiff = GetTextOrDiff(ws.Cell(row, 27), particleOut, particleIn);
        string ipaDiff = GetTextOrDiff(ws.Cell(row, 30), ipaOut, ipaIn);
        string acetoneDiff = GetTextOrDiff(ws.Cell(row, 33), acetoneOut, acetoneIn);
        string nonTargetDiff = GetTextOrDiff(ws.Cell(row, 36), nonTargetOut, nonTargetIn);
        string totalDiff = GetText(ws.Cell(row, 37));

        if (string.IsNullOrWhiteSpace(totalDiff) || IsFormulaText(totalDiff))
        {
            totalDiff = DiffUtil.GetSumDiff(
                ipaOut, acetoneOut, nonTargetOut,
                ipaIn, acetoneIn, nonTargetIn
            );
        }

        string pressureSpec;
        string pressureDrop;
        ResolvePressureDrop(ws, row, out pressureSpec, out pressureDrop);

        return new Dictionary<int, string>
        {
            [38] = particleIn,
            [39] = particleOut,
            [40] = particleDiff,

            [41] = ipaIn,
            [42] = ipaOut,
            [43] = ipaDiff,

            [44] = acetoneIn,
            [45] = acetoneOut,
            [46] = acetoneDiff,

            [47] = nonTargetIn,
            [48] = nonTargetOut,
            [49] = nonTargetDiff,

            [50] = totalDiff,
            [53] = pressureSpec,
            [54] = pressureDrop
        };
    }

    private static void ResolvePressureDrop(
        IXLWorksheet ws,
        int row,
        out string pressureSpec,
        out string pressureDrop)
    {
        string singleSpec = GetText(ws.Cell(row, 38));
        string singleDrop = GetText(ws.Cell(row, 39));
        string setSpec = GetText(ws.Cell(row, 40));
        string setDrop = GetText(ws.Cell(row, 41));

        if (!string.IsNullOrWhiteSpace(setSpec) ||
            !string.IsNullOrWhiteSpace(setDrop))
        {
            pressureSpec = setSpec;
            pressureDrop = setDrop;
            return;
        }

        pressureSpec = singleSpec;
        pressureDrop = singleDrop;
    }

    private static void AddAlarms(
        IXLWorksheet ws,
        int row,
        List<string> alarmList)
    {
        for (int col = 14; col <= 17; col++)
        {
            if (!IsChecked(ws.Cell(row, col)))
                continue;

            string header = GetText(ws.Cell(1, col))
                .Replace("\n", "")
                .Replace("(V)", "")
                .Trim();

            if (!string.IsNullOrWhiteSpace(header) &&
                !alarmList.Any(x => string.Equals(x, header, StringComparison.OrdinalIgnoreCase)))
            {
                alarmList.Add(header);
            }
        }
    }

    private static bool IsChecked(IXLCell cell)
    {
        if (cell == null || cell.IsEmpty())
            return false;

        if (cell.DataType == XLDataType.Boolean)
            return cell.GetBoolean();

        string text = GetText(cell).Trim();
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return !text.Equals("N/A", StringComparison.OrdinalIgnoreCase) &&
               !text.Equals("FALSE", StringComparison.OrdinalIgnoreCase) &&
               text != "-" &&
               text != "0";
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row)
    {
        for (int col = 1; col <= 41; col++)
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

        if (cell.DataType == XLDataType.Number)
            return cell.GetDouble().ToString("0.###", CultureInfo.InvariantCulture);

        if (cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime().ToString("yyyy.MM.dd");

        return cell.GetString().Trim();
    }

    private static string GetDateText(IXLCell cell)
    {
        if (cell != null && cell.DataType == XLDataType.DateTime)
            return cell.GetDateTime().ToString("yyyy.MM.dd");

        string text = GetText(cell);
        if (IsFormulaText(text))
            return "";

        if (DateTime.TryParse(text, out DateTime date))
            return date.ToString("yyyy.MM.dd");

        if (DateTime.TryParse(text.Replace(".", "/"), out date))
            return date.ToString("yyyy.MM.dd");

        return text;
    }

    private static string GetEfficiencyText(IXLCell cell)
    {
        string text = GetText(cell);

        if (IsEmptyText(text))
            return "";

        if (decimal.TryParse(text.Replace("%", ""), out decimal value))
        {
            if (Math.Abs(value) <= 1)
                value *= 100m;

            return value.ToString("0.#", CultureInfo.InvariantCulture);
        }

        return text;
    }

    private static string GetTextOrDiff(IXLCell diffCell, string outValue, string inValue)
    {
        string diff = GetText(diffCell);
        if (!string.IsNullOrWhiteSpace(diff) && !IsFormulaText(diff))
            return diff;

        return DiffUtil.GetDiff(outValue, inValue);
    }

    private static bool IsFormulaText(string value)
    {
        return !string.IsNullOrWhiteSpace(value) &&
               value.TrimStart().StartsWith("=", StringComparison.Ordinal);
    }

    private static bool IsEmptyText(string value)
    {
        return string.IsNullOrWhiteSpace(value) ||
               value.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("N.D.", StringComparison.OrdinalIgnoreCase) ||
               value == "-";
    }

}
