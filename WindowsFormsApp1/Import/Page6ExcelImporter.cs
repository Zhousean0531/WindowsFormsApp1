using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page6;

public static class Page6ExcelImporter
{
    private const string SheetName = "物料";

    public static int ImportWithFileDialog(IWin32Window owner = null)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "選擇 Page6 匯入 Excel";
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
        using (var wb = new XLWorkbook(excelPath))
        {
            var ws = FindWorksheet(wb, sheetName);
            if (ws == null)
                throw new InvalidOperationException("Cannot find sheet: " + sheetName);

            int headerRow = FindHeaderRow(ws);
            var columns = BuildColumnMap(ws, headerRow);
            int maxColumn = columns.Values.Max();

            var items = new List<P6Item>();
            DateTime? itemDate = null;

            int row = headerRow + 1;
            while (!IsEmptyRow(ws, row, maxColumn))
            {
                var item = new P6Item
                {
                    Col1 = GetText(ws.Cell(row, columns["Date"])),
                    Col2 = GetText(ws.Cell(row, columns["Material"])),
                    Spec1 = GetText(ws.Cell(row, columns["InQty"])),
                    Spec2 = GetText(ws.Cell(row, columns["InUnit"])),
                    Range1 = GetText(ws.Cell(row, columns["SampleQty"])),
                    Range2 = GetText(ws.Cell(row, columns["SampleUnit"])),
                    Result = GetText(ws.Cell(row, columns["Appearance"])),
                    Judgment = GetText(ws.Cell(row, columns["Spec"])),
                    Extra1 = GetText(ws.Cell(row, columns["Measured"])),
                    Extra2 = GetText(ws.Cell(row, columns["Judgement"])),
                    Extra3 = GetText(ws.Cell(row, columns["Note"])),
                    Supplier = GetText(ws.Cell(row, columns["Supplier"]))
                };

                if (!itemDate.HasValue)
                    itemDate = ParseDate(item.Col1);

                if (HasItemData(item))
                    items.Add(item);

                row++;
            }

            if (items.Count == 0)
                return 0;

            var batch = new P6Batch
            {
                ReportNo = FindReportNo(ws) ?? "",
                TestDate = itemDate ?? FindHeaderTestDate(ws) ?? DateTime.Today,
                UserName = Environment.UserName,
                Items = items
            };

            P6Repository.Insert(batch);
            return 1;
        }
    }

    private static IXLWorksheet FindWorksheet(XLWorkbook wb, string sheetName)
    {
        return wb.Worksheets
            .FirstOrDefault(x => string.Equals(x.Name, sheetName, StringComparison.OrdinalIgnoreCase));
    }

    private static int FindHeaderRow(IXLWorksheet ws)
    {
        for (int row = 1; row <= 20; row++)
        {
            string text = "";

            for (int col = 1; col <= 20; col++)
                text += GetText(ws.Cell(row, col));

            if (text.Contains("進貨日期") &&
                (text.Contains("料號") || text.Contains("物料") || text.Contains("品名")))
            {
                return row;
            }
        }

        return 1;
    }

    private static Dictionary<string, int> BuildColumnMap(IXLWorksheet ws, int headerRow)
    {
        int dateCol = FindColumn(ws, headerRow, 1, "進貨日期", "日期");
        int materialCol = FindColumn(ws, headerRow, 2, "料號/物料名稱", "料號", "物料名稱", "品名");
        int inQtyCol = FindColumn(ws, headerRow, 3, "進貨數量", "進貨量");
        int inUnitCol = FindUnitColumn(ws, headerRow, inQtyCol, 4);
        int sampleQtyCol = FindColumn(ws, headerRow, 5, "抽檢數", "抽樣數");
        int sampleUnitCol = FindUnitColumn(ws, headerRow, sampleQtyCol, 6);

        return new Dictionary<string, int>
        {
            ["Date"] = dateCol,
            ["Material"] = materialCol,
            ["InQty"] = inQtyCol,
            ["InUnit"] = inUnitCol,
            ["SampleQty"] = sampleQtyCol,
            ["SampleUnit"] = sampleUnitCol,
            ["Appearance"] = FindColumn(ws, headerRow, 7, "外觀"),
            ["Spec"] = FindColumn(ws, headerRow, 8, "規格值", "規格"),
            ["Measured"] = FindColumn(ws, headerRow, 9, "測量值", "測量"),
            ["Judgement"] = FindColumn(ws, headerRow, 10, "合格/不合格", "判定"),
            ["Note"] = FindColumn(ws, headerRow, 11, "備註"),
            ["Supplier"] = FindColumn(ws, headerRow, 12, "供應商")
        };
    }

    private static int FindColumn(IXLWorksheet ws, int headerRow, int fallback, params string[] names)
    {
        for (int col = 1; col <= 30; col++)
        {
            string header = GetText(ws.Cell(headerRow, col)).Replace("\r", "").Replace("\n", "");

            foreach (string name in names)
            {
                if (header.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    return col;
            }
        }

        return fallback;
    }

    private static int FindUnitColumn(IXLWorksheet ws, int headerRow, int afterColumn, int fallback)
    {
        for (int col = Math.Max(1, afterColumn + 1); col <= 30; col++)
        {
            string header = GetText(ws.Cell(headerRow, col)).Replace("\r", "").Replace("\n", "");
            if (header.IndexOf("單位", StringComparison.OrdinalIgnoreCase) >= 0)
                return col;
        }

        return fallback;
    }

    private static bool IsEmptyRow(IXLWorksheet ws, int row, int maxColumn)
    {
        for (int col = 1; col <= maxColumn; col++)
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

    private static bool HasItemData(P6Item item)
    {
        return !string.IsNullOrWhiteSpace(item.Col1) ||
               !string.IsNullOrWhiteSpace(item.Col2) ||
               !string.IsNullOrWhiteSpace(item.Spec1) ||
               !string.IsNullOrWhiteSpace(item.Spec2) ||
               !string.IsNullOrWhiteSpace(item.Range1) ||
               !string.IsNullOrWhiteSpace(item.Range2) ||
               !string.IsNullOrWhiteSpace(item.Result) ||
               !string.IsNullOrWhiteSpace(item.Judgment) ||
               !string.IsNullOrWhiteSpace(item.Extra1) ||
               !string.IsNullOrWhiteSpace(item.Extra2) ||
               !string.IsNullOrWhiteSpace(item.Extra3) ||
               !string.IsNullOrWhiteSpace(item.Supplier);
    }

    private static string FindReportNo(IXLWorksheet ws)
    {
        return FindHeaderValue(ws, "報告編號", "Report No");
    }

    private static DateTime? FindHeaderTestDate(IXLWorksheet ws)
    {
        string text = FindHeaderValue(ws, "抽檢日期", "Testing Date", "Test Date");
        return ParseDate(text);
    }

    private static string FindHeaderValue(IXLWorksheet ws, params string[] labels)
    {
        for (int row = 1; row <= 10; row++)
        {
            for (int col = 1; col <= 12; col++)
            {
                string text = GetText(ws.Cell(row, col));
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                bool matched = labels.Any(label =>
                    text.IndexOf(label, StringComparison.OrdinalIgnoreCase) >= 0);

                if (!matched)
                    continue;

                string inlineValue = ExtractAfterSeparator(text);
                if (!string.IsNullOrWhiteSpace(inlineValue))
                    return inlineValue;

                string nextCellValue = GetText(ws.Cell(row, col + 1));
                if (!string.IsNullOrWhiteSpace(nextCellValue))
                    return nextCellValue;
            }
        }

        return null;
    }

    private static string ExtractAfterSeparator(string text)
    {
        int index = text.LastIndexOf(':');
        if (index < 0)
            index = text.LastIndexOf('：');

        if (index < 0 || index == text.Length - 1)
            return "";

        return text.Substring(index + 1).Trim();
    }

    private static DateTime? ParseDate(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return null;

        if (DateTime.TryParse(text, out DateTime date))
            return date;

        if (DateTime.TryParse(text.Replace(".", "/"), out date))
            return date;

        return null;
    }
}
