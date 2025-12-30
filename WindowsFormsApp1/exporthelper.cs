using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Core;
using Org.BouncyCastle.Pqc.Crypto.Lms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1;
using YourNamespace;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeCore = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


public static class DiffUtil
{
    public static string GetSumDiff(string out1, string out2, string out3, string in1, string in2, string in3)
    {
        double o1, o2, o3, i1, i2, i3;
        bool p1 = double.TryParse(out1, out o1);
        bool p2 = double.TryParse(out2, out o2);
        bool p3 = double.TryParse(out3, out o3);
        bool p4 = double.TryParse(in1, out i1);
        bool p5 = double.TryParse(in2, out i2);
        bool p6 = double.TryParse(in3, out i3);

        if (!(p1 && p2 && p3 && p4 && p5 && p6))
            return "N.D.";

        double diff = (o1 + o2 + o3) - (i1 + i2 + i3);
        return diff <= 0 ? "N.D." : diff.ToString("0.###");
    }
    public static string GetDiff(string outVal, string inVal)
    {
        double o, i;
        if (double.TryParse(outVal, out o) && double.TryParse(inVal, out i))
        {
            double diff = o - i;
            return diff <= 0 ? "N.D." : diff.ToString("0.###");
        }
        return "N.D.";
    }
}
public static class EfficiencyFinder
{
    public static Dictionary<string, string> FindEfficiencyByTypeWithMin(IXLWorksheet ws, string carbonLot, List<string> targetTypes)
    {
        // 建立結果字典，先預設目標種類都為 "N/A"
        var result = new Dictionary<string, string>
        {
            { "MA", "N/A" },
            { "MB", "N/A" },
            { "MC", "N/A" }
        };

        foreach (var r in ws.RowsUsed())
        {
            int i,j,k;
            if (ws.Name == "濾網")
            {
                i = 21;j = 31;k = 32;
            }
            else
            {
                i = 5;j = 10;k = 15;
            }
            string existingLot = r.Cell(i).GetString().Trim();
            existingLot = existingLot.Length >= 14 ? existingLot.Substring(0, 14) : existingLot;
            string testItem = r.Cell(j).GetString().Trim();
            string efficiencyStr = r.Cell(k).GetString().Trim();

            if (existingLot.Equals(carbonLot.Trim(), StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(efficiencyStr, out double efficiency))
            {
                string type = null;
                if ((testItem == "SO2" || testItem == "H2S") && targetTypes.Contains("MA")) type = "MA";
                else if (testItem == "NH3" && targetTypes.Contains("MB")) type = "MB";
                else if ((testItem == "Toluene" || testItem == "Acetone" || testItem == "IPA") && targetTypes.Contains("MC")) type = "MC";

                if (type != null)
                {
                    if (result[type] == "N/A")
                        result[type] = efficiency.ToString("0.###");
                    else if (double.TryParse(result[type], out double existVal) && efficiency < existVal)
                        result[type] = efficiency.ToString("0.###");
                }
            }
        }
        return result;
    }
}
public static class ValidationHelper
{
    public static bool CheckRequiredFields(TabPage tab)
    {
        foreach (Control ctrl in tab.Controls)
        {
            if (ctrl is TextBox tb && string.IsNullOrWhiteSpace(tb.Text))
                return false;

            if (ctrl is ComboBox cb && string.IsNullOrWhiteSpace(cb.Text))
                return false;

            if (ctrl is DataGridView dgv)
            {
                bool hasData = dgv.Rows
                    .Cast<DataGridViewRow>()
                    .Any(row => !row.IsNewRow && row.Cells.Cast<DataGridViewCell>().Any(cell => cell.Value != null && cell.Value.ToString().Trim() != ""));

                if (!hasData)
                    return false;
            }
        }
        return true;
    }
}
public static class DictionaryExtensions
{
    public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dict, TKey key, TValue defaultValue = default)
    {
        return dict.TryGetValue(key, out TValue value) ? value : defaultValue;
    }
}
public static class ExportHelper
{
    public static Dictionary<string, string> CollectTabData(TabPage tabPage)
    {
        Dictionary<string, string> formData = new Dictionary<string, string>();
        formData["目前分頁"] = tabPage.Name;

        foreach (Control control in tabPage.Controls)
        {
            if (control is TextBox textBox)
            {
                if (textBox.Name == "FilterRawNumberBox")
                {
                    var parts = textBox.Text.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < parts.Length; i++)
                    {
                        formData[$"{textBox.Name}_{i + 1}"] = parts[i].Trim();
                    }
                }
                else
                {
                    formData[textBox.Name] = textBox.Text;
                }
            }
            else if (control is ComboBox comboBox)
            {
                formData[comboBox.Name] = comboBox.Text;
            }
            else if (control is DateTimePicker dateTimePicker)
            {
                formData[dateTimePicker.Name] = dateTimePicker.Value.ToString("yyyy.MM.dd");
            }
            else if (control is DataGridView dataGridView)
            {
                List<string> rows = new List<string>();
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        List<string> cells = new List<string>();
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            cells.Add(cell.Value?.ToString() ?? "");
                        }
                        rows.Add(string.Join(",", cells));
                    }
                }
                formData[dataGridView.Name] = string.Join(" | ", rows);
            }
        }
        return formData;
    }
}
public static class ExportHelper_Page1
{
    private static T FindCtrl<T>(Control root, string nameContains) where T : Control
    {
        foreach (Control c in root.Controls)
        {
            if (c is T t && c.Name.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) >= 0)
                return t;
            var child = FindCtrl<T>(c, nameContains); // 遞迴
            if (child != null) return child;
        }
        return null;
    }
    private static System.Collections.Generic.List<string> SplitStr(string raw)
    {
        var sep = new[] { '/', ',', '、', '|' };
        return (raw ?? "").Split(sep, StringSplitOptions.RemoveEmptyEntries)
                          .Select(s => s.Trim()).ToList();
    }
    private static System.Collections.Generic.List<double> SplitDouble(string raw)
    {
        return SplitStr(raw).Select(s => double.TryParse(s, out var v) ? v : 0).ToList();
    }
    private static bool AllSameLength(params System.Collections.IList[] lists)
    {
        if (lists == null || lists.Length == 0) return true;
        int n = lists[0].Count;
        return lists.All(x => x != null && x.Count == n);
    }
    private static string PickGas(string material)
    {
        var m = (material ?? "").Trim().ToUpperInvariant();
        if (m.Contains("SG")) return "Toluene";
        if (m == "SI013") return "SO2";
        if (m == "IKP101") return "H2S";
        return "TVOC";
    }
    private static string BuildMeshSummaryString(TabPage tab, string material)
    {
        var dgv = FindCtrl<DataGridView>(tab, "FilterRawParticleSizeBox");
        if (dgv == null) return "";

        double total = 0;
        var weights = new Dictionary<string, double>();
        int emptyCount = 0;

        foreach (DataGridViewRow r in dgv.Rows)
        {
            if (r.IsNewRow) continue;
            string label = Convert.ToString(r.Cells[0].Value)?.Trim();
            string valStr = Convert.ToString(r.Cells[1].Value)?.Trim();
            if (string.IsNullOrEmpty(label)) continue;

            if (string.IsNullOrEmpty(valStr))
            {
                emptyCount++;
                continue;
            }

            if (double.TryParse(valStr, out double w))
            {
                weights[label] = w;
                if (label.Contains("總重")) total = w;
            }
        }

        if (emptyCount > 1)
        {
            MessageBox.Show("粒徑重量最多只能有一格未填！請檢查。");
            throw new InvalidOperationException("Too many empty mesh weights");
        }

        // 自動計算中間 Mesh 重量
        if (weights.ContainsKey("總重"))
        {
            double sumKnown = 0;
            foreach (var kv in weights)
            {
                if (kv.Key != "總重") sumKnown += kv.Value;
            }

            if (weights.ContainsKey("<30 Mesh") && weights.ContainsKey(">60 Mesh") && !weights.ContainsKey("30~60 Mesh"))
            {
                double remain = total - sumKnown;
                if (remain > 0) weights["30~60 Mesh"] = remain;
            }
        }

        if (total == 0)
        {
            MessageBox.Show("請確認總重是否填寫！");
            return "";
        }

        var parts = new List<string>();
        double sumPct = 0;
        foreach (var kv in weights)
        {
            if (kv.Key == "總重") continue;
            double pct = kv.Value / total * 100;
            parts.Add($"{kv.Key} {pct:F1}%");
            sumPct += pct;
        }

        if (Math.Abs(sumPct - 100) > 2)
            MessageBox.Show($"粒徑分布總和為 {sumPct:F1}%，請確認是否正確。");

        return string.Join("; ", parts);
    }
    // 依 TabPage1 的命名讀效率區塊（濃度/背景/11 點讀值）
    private static (double conc, double bg, List<double> readings)GetEffBlockRaw(string gas, TabPage tab)
    {
        var tbConc = FindCtrl<TextBox>(tab, "FilterRawConcertrationBox");
        var tbBg = FindCtrl<TextBox>(tab, "FilterRawBackGroundBox");
        var tbVal = FindCtrl<TextBox>(tab, "FilterRawEffvalueBox");

        double.TryParse(tbConc?.Text, out double conc);
        double.TryParse(tbBg?.Text, out double bg);

        var readings = (tbVal?.Text ?? "")
            .Split(new[] { '\r', '\n', '/', ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => double.TryParse(s.Trim(), out var v) ? v : double.NaN)
            .Where(v => !double.IsNaN(v))
            .ToList();

        if (readings.Count < 11)
        {
            MessageBox.Show($"目前僅輸入 {readings.Count} 筆讀值，需為 11 筆 (0~10min)。請確認輸入是否完整。",
                            "讀值不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            throw new InvalidOperationException("readings < 11");
        }
        return (conc, bg, readings);
    }
    private static void InsertInspectorSignature(Excel.Worksheet ws)
    {
        // 直接指定簽名檔完整路徑（沒有空白）
        string signFile =
            @"C:\Users\User\OneDrive - 鈺祥企業股份有限公司\General - 鈺祥 - 品保實驗室\鈺祥CI繁體版\周崇賢.png";

        if (!File.Exists(signFile))
            return;

        // 假設「品保檢驗人員」在 C25，要把簽名放在右邊
        Excel.Range anchor = ws.Range["C25"];  // 如果實際是在 C26 就改成 "C26"

        float left = (float)(double)anchor.Left + (float)anchor.Width + 5;
        float top = (float)(double)anchor.Top + 2;

        ws.Shapes.AddPicture(
            signFile,
            MsoTriState.msoFalse,   // LinkToFile: false
            MsoTriState.msoCTrue,   // SaveWithDocument: true
            left,
            top,
            100,   // 寬度
            40     // 高度
        );
    }
    private static void ExportToExcel(string savePath, Page1ExportData d)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到公版檔：" + templatePath);
            return;
        }

        File.Copy(templatePath, savePath, true);

        var app = new Excel.Application();
        var wb = app.Workbooks.Open(savePath);
        var ws = (Excel.Worksheet)wb.Sheets[1];

        try
        {
            // ===== 標頭區 =====
            // 報告編號
            ws.Range["C4"].Value = d.ReportNo;

            // Material Name = 原料種類（不要再用 PickGas）
            ws.Range["E5"].Value = d.Material;

            // Delivery Date / Testing Date
            ws.Range["C5"].Value = d.ArrivalDate;
            ws.Range["C6"].Value = d.TestingDate;

            // Delivery Quantity = 進貨數量
            ws.Range["E6"].Value = d.QtyText;

            // ===== 多筆樣品：橫向填 C、D、E... =====
            const int COL_FIRST = 3;   // C 欄 = 3
            const int ROW_LOTNO = 10;  // Lot No. 行
            const int ROW_DENS = 13;  // Density 行
            const int ROW_DP = 14;  // Pressure Drop 行
            const int ROW_VIN = 15;  // VOC Inlet 行
            const int ROW_VOUT = 16;  // VOC Outlet 行
            const int ROW_VOUTG = 17;  // VOC Outgassing 行
            const int ROW_EFF = 18;  // Initial Efficiency(%) 行  (依你的公版調整)

            int count = d.LotFulls.Count;

            for (int i = 0; i < count; i++)
            {
                int col = COL_FIRST + i; // C=3, D=4, E=5...

                ws.Cells[ROW_LOTNO, col].Value = d.LotFulls[i];
                ws.Cells[ROW_DENS, col].Value = d.Densities[i].ToString("0.00");
                ws.Cells[ROW_DP, col].Value = d.DeltaPs[i].ToString("0.0");
                ws.Cells[ROW_VIN, col].Value = d.VocIns[i];
                ws.Cells[ROW_VOUT, col].Value = d.VocOuts[i];
                ws.Cells[ROW_VOUTG, col].Value = d.OutgassingList[i];

                // 初始效率：只有選中的那一筆填數字，其餘顯示 N.D.
                if (i == d.SelectedIndex)
                    ws.Cells[ROW_EFF, col].Value = d.Eff0;
                else
                    ws.Cells[ROW_EFF, col].Value = "N.D.";
            }
            wb.Save();
        }
        finally
        {
            wb.Close();
            app.Quit();
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);
        }
    }
    private class Page1ExportData
    {
        public string ReportNo { get; set; }        // 報告編號
        public string Material { get; set; }        // 原料種類/名稱
        public string ArrivalDate { get; set; }     // 到廠日期 yyyy.MM.dd
        public string TestingDate { get; set; }     // 測試日期 yyyy.MM.dd
        public string QtyText { get; set; }         // 進貨數量字串（457 kg/18袋）

        public List<string> LotFulls { get; set; }  // 每筆 B-2025...#16/#17/#19
        public List<double> Densities { get; set; }
        public List<double> DeltaPs { get; set; }
        public List<double> VocIns { get; set; }
        public List<double> VocOuts { get; set; }
        public List<string> OutgassingList { get; set; } // 每筆 outgassing 顯示 (N.D. / 數值)

        public int SelectedIndex { get; set; }   // 使用者選哪一筆壓損做效率
        public string Eff0 { get; set; }   // 初始效率(%)，只對 SelectedIndex 那一格填
    }
    private static (string conc, string eff0, string eff10)ComputeEfficiency(double conc, double bg, List<double> r)
    {
        if (conc <= 0 || r == null || r.Count < 11) return ("", "", "");
        double e0 = (conc + bg - r[0]) / conc * 100.0;
        double e10 = (conc + bg - r[10]) / conc * 100.0;
        return (conc.ToString("F0"), e0.ToString("F1"), e10.ToString("F1"));
    }
    public static void Handle(TabPage tab)
    {
        try
        {

            // ===== (A) 取共用欄位 =====
            var arrivePicker = FindCtrl<DateTimePicker>(tab, "FilterRawArriveDateBox");
            var testPicker = FindCtrl<DateTimePicker>(tab, "FilterRawTestDateBox");
            var materialBox = FindCtrl<ComboBox>(tab, "FilterRawTypeBox");
            var lotBox = FindCtrl<TextBox>(tab, "FilterRawBatchNOBox");
            var qtyBox1 = FindCtrl<TextBox>(tab, "FilterRawQtyWeightBox");
            var qtyBox2 = FindCtrl<TextBox>(tab, "FilterRawQuantityBox");
            string qty = "";
            // 兩格都當字串拿，空就不顯示單位
            var q1txt = (qtyBox1?.Text ?? "").Trim();
            var q2txt = (qtyBox2?.Text ?? "").Trim();

            if (q1txt != "" || q2txt != "")
            {
                var left = q1txt != "" ? (q1txt + " kg") : "";
                var right = q2txt != "" ? (q2txt + "板") : "";
                qty = (left + (left != "" && right != "" ? "/" : "") + right);
            }
            if (arrivePicker == null || testPicker == null || materialBox == null)
            {
                MessageBox.Show("找不到必要控制項（到廠/測試日期或原料種類）。");
                return;
            }

            string arriveDate = arrivePicker.Value.ToString("yyyy.MM.dd");
            string testDate = testPicker.Value.ToString("yyyy.MM.dd");
            string material = (materialBox.Text ?? "").Trim();
            string lotNo = lotBox != null ? lotBox.Text.Trim() : "";
            var tbReportNo = FindCtrl<TextBox>(tab, "FilterRawReportNoTB");

            if (string.IsNullOrEmpty(lotNo))
                lotNo = string.Format("B-{0:yyyyMMdd}-001", arrivePicker.Value);

            // ===== (B) 解析多筆欄位（用 / 分隔）=====
            var nos = SplitStr(FindCtrl<TextBox>(tab, "FilterRawNumberBox")?.Text);
            var weights = SplitDouble(FindCtrl<TextBox>(tab, "FilterRawWeightBox")?.Text);
            var vocIn = SplitDouble(FindCtrl<TextBox>(tab, "FilterRawVOCsInletBox")?.Text);
            var vocOut = SplitDouble(FindCtrl<TextBox>(tab, "FilterRawVOCsOutletBox")?.Text);
            var deltas = SplitDouble(FindCtrl<TextBox>(tab, "FilterRawPressureBox")?.Text);
            int n = new[] { nos.Count, weights.Count, vocIn.Count, vocOut.Count, deltas.Count }.Min();

            if (n == 0) { MessageBox.Show("沒有可匯出的資料。"); return; }
            // ===== (C) 讓使用者選壓損資料列（用索引第幾筆） =====
            deltas = deltas.Take(n).ToList();
            // 計算密度（假設體積固定 50）
            double vol = 50.0;
            List<double> densities = weights.Select(w => vol > 0 ? w / vol : 0).ToList();

            // 準備選項顯示字串，例如「第1筆 (ΔP=15)」
            var deltaItems = new List<string>();
            for (int i = 0; i < n; i++)
            {
                deltaItems.Add($"{deltas[i]}");
            }
            int x = new[] { nos.Count, weights.Count, vocIn.Count, vocOut.Count, deltas.Count }.Min();
            nos = nos.Take(n).ToList();
            weights = weights.Take(n).ToList();
            vocIn = vocIn.Take(n).ToList();
            vocOut = vocOut.Take(n).ToList();
            deltas = deltas.Take(n).ToList();

            // 顯示索引選擇表單
            int selectedIdx = -1;
            using (var f = new Form2(deltas, "請選擇用哪一筆壓損"))
            {
                if (f.ShowDialog() == DialogResult.OK)
                    selectedIdx = f.SelectedIndex0;   // 0-based
                else { MessageBox.Show("未選擇資料列，已取消。"); return; }
            }
            if (selectedIdx < 0 || selectedIdx >= n)
            {
                MessageBox.Show("選擇的索引超出範圍。");
                return;
            }
            // ===== (D) 效率區塊（依原料種類決定氣體）=====
            string gas = PickGas(material);
            var effBlk = GetEffBlockRaw(gas, tab); // 讀 濃度/背景/11點
            try
            {
                effBlk = GetEffBlockRaw(gas, tab);
                if (effBlk.conc <= 0)
                {
                    MessageBox.Show("初始濃度需 > 0。");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取效率資料失敗：" + ex.Message);
                return;
            }
            var effRes = ComputeEfficiency(effBlk.conc, effBlk.bg, effBlk.readings); // (conc, eff0, eff10)
            // ===== (E) 粒徑分布字串 =====
            string meshSummary = BuildMeshSummaryString(tab, material);

            // ===== (F) 組列 → 準備寫入 Excel =====
            var rows = new List<Dictionary<string, string>>();
            for (int i = 0; i < n; i++)
            {
                double outMinusIn = vocOut[i] - vocIn[i];
                var row = new Dictionary<string, string>
                {
                    ["測試日期"] = testDate,
                    ["到廠日期"] = arriveDate,
                    ["原料種類"] = material,

                    // ✅ 每筆都自動組出 B-20251101-001#01/#02/#03
                    ["進廠批號"] = $"{lotNo}#{nos[i].PadLeft(2, '0')}",

                    // ✅ 進貨數字串，不動
                    ["進貨數量"] = qty,
                    ["原料編號"] = nos[i],
                    ["樣品重量"] = weights[i].ToString(),
                    ["密度"] = densities[i].ToString("F3"),
                    ["VOC Inlet"] = vocIn[i].ToString(),
                    ["VOC Outlet"] = vocOut[i].ToString(),
                    ["OutMinusIn"] = outMinusIn.ToString("F1"),
                    ["壓損(pa)"] = deltas[i].ToString(),
                    ["測試氣體"] = gas,
                    ["粒徑分布"] = meshSummary
                };
                if (i == selectedIdx)
                {
                    row["初始濃度(ppb)"] = $"{effRes.conc}";
                    row["初始效率(%)"] = $"{effRes.eff0}";
                    row["十分鐘效率(%)"] = $"{effRes.eff10}";
                }
                else
                {
                    row["初始濃度(ppb)"] = "";
                    row["初始效率(%)"] = "";
                    row["十分鐘效率(%)"] = "";
                }

                rows.Add(row);
            }

            // ===== (G) 寫入 Excel（欄位對應照你習慣，可調整）=====
            string filePath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "總表.xlsx"
            );
            if (!System.IO.File.Exists(filePath))
            {
                MessageBox.Show("找不到總表.xlsx，請確認路徑。");
                return;
            }

            using (var wb = new ClosedXML.Excel.XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("濾網"); // 你的原料頁
                                             // 以 B 欄（測試日期）自第 5 列往下找最後一筆
                var used = ws.Column(2).Cells(c => c.Address.RowNumber >= 5 && !c.IsEmpty());
                int rowIdx = (used.LastOrDefault()?.Address.RowNumber ?? 4) + 1;

                foreach (var d in rows)
                {
                    ws.Cell(rowIdx, 1).Value = d["測試日期"];         // 測試日期
                    ws.Cell(rowIdx, 2).Value = d["到廠日期"];         // 到廠日期
                    ws.Cell(rowIdx, 3).Value = d["原料種類"];         // 原料種類
                    ws.Cell(rowIdx, 4).Value = d["進廠批號"];         // 進廠批號
                    ws.Cell(rowIdx, 5).Value = d["進貨數量"];         // 進貨數量
                    ws.Cell(rowIdx, 8).Value = d["樣品重量"];         // 測試品重量
                    ws.Cell(rowIdx, 9).Value = d["密度"];             // 密度
                    ws.Cell(rowIdx, 11).Value = d["VOC Inlet"];        // inlet
                    ws.Cell(rowIdx, 12).Value = d["VOC Outlet"];       // outlet
                    ws.Cell(rowIdx, 13).Value = d["OutMinusIn"];       // outlet-inlet(≤5)
                    ws.Cell(rowIdx, 14).Value = d["壓損(pa)"];         // 壓損(pa)
                    ws.Cell(rowIdx, 15).Value = d["初始效率(%)"];      // 初始效率
                    ws.Cell(rowIdx, 16).Value = d["十分鐘效率(%)"];    // 10分鐘效率
                    ws.Cell(rowIdx, 18).Value = d["粒徑分布"];         // 粒徑分布
                    rowIdx++;
                }
                wb.Save();
            }

            // ===== (H) 組出「單張報告」需要的資料（用使用者選的那一筆）=====
            // 每筆 LotNo 完整字串
            var lotFulls = new List<string>();
            var outStrings = new List<string>();
            for (int i = 0; i < n; i++)
            {
                string no = nos[i].PadLeft(2, '0');
                lotFulls.Add($"{lotNo}#{no}");

                double diff = vocOut[i] - vocIn[i];
                outStrings.Add(diff <= 0 ? "N.D." : diff.ToString("F1"));
            }

            // 報告編號 TextBox（上面其實已經先抓過一次，可以直接用那個 tbReportNo）
            var tbReportNo2 = FindCtrl<TextBox>(tab, "FilterRawReportNoTB");
            string reportNo = tbReportNo2?.Text.Trim() ?? "";

            // 組 Page1ExportData（用「List 版本」）
            var reportData = new Page1ExportData
            {
                ReportNo = reportNo,
                Material = material,
                ArrivalDate = arriveDate,
                TestingDate = testDate,
                QtyText = qty,

                LotFulls = lotFulls,
                Densities = densities,
                DeltaPs = deltas,
                VocIns = vocIn,
                VocOuts = vocOut,
                OutgassingList = outStrings,

                SelectedIndex = selectedIdx,
                Eff0 = effRes.eff0
            };

            // ===== (I) 另存新檔：檔名預設「報告編號_原料種類_(yyMMdd到廠).xlsx」=====
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                var arriveShort = arrivePicker.Value.ToString("yyMMdd");
                sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                sfd.FileName = $"{reportNo}_{material}({arriveShort}到廠).xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    ExportToExcel(sfd.FileName, reportData);
                    MessageBox.Show("匯出完成。");
                }
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }

}
public static class ExportHelper_Page2
{
    // ====== 公用：遞迴抓控制項 ======
    private static IEnumerable<Control> Walk(Control parent)
    {
        foreach (Control c in parent.Controls)
        {
            yield return c;
            foreach (var cc in Walk(c)) yield return cc;
        }
    }
    private static IEnumerable<TextBox> GetAllTextBoxes(Control parent)
        => Walk(parent).OfType<TextBox>();
    private static IEnumerable<CheckBox> GetAllCheckBoxes(Control parent)
        => Walk(parent).OfType<CheckBox>();
    private static Control FindFirst(Control parent, string nameContains)
        => Walk(parent).FirstOrDefault(c => c.Name.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) >= 0);
    private static TextBox FindTextBox(Control parent, string nameContains)
        => FindFirst(parent, nameContains) as TextBox;
    private static string GetText(Control parent, string keyword)
    {
        foreach (Control c in parent.Controls)
        {
            // 遞迴搜尋子控制項（例如 Panel、GroupBox）
            string result = GetText(c, keyword);
            if (!string.IsNullOrEmpty(result)) return result;

            // TextBox
            if (c is TextBox txt && c.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                return txt.Text.Trim();

            // ComboBox
            if (c is ComboBox combo && c.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                return combo.Text.Trim();

            // ✅ DateTimePicker
            if (c is DateTimePicker dtp && c.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                return dtp.Value.ToString("yyyy.MM.dd"); // 你也可以改成 yyyy/MM/dd 或 yyyy-MM-dd
        }
        return "";
    }
    // ====== 解析工具 ======
    private static List<double> ParseDoubleList(string raw)
    {
        // 支援 / 、, 、空白、中文頓號等分隔
        if (string.IsNullOrWhiteSpace(raw)) return new List<double>();
        var parts = Regex.Split(raw, @"[\/,、\s]+")
                         .Where(s => !string.IsNullOrWhiteSpace(s));
        var list = new List<double>();
        foreach (var p in parts)
        {
            if (double.TryParse(p, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ||
                double.TryParse(p, NumberStyles.Any, CultureInfo.CurrentCulture, out v))
            {
                list.Add(v);
            }
        }
        return list;
    }

    private static List<double> ParseReading11(string raw)
    {
        // 預期 11 點 (0~10min)。不足就忽略，多了就取前 11。
        var list = ParseDoubleList(raw);
        return list.Take(11).ToList();
    }
    private class Page2ExportData
    {
        public string ReportNo;
        public string ProductionDate;
        public string TestDate;
        public string ProductName;
        public string ProductSize;
        public string ProductNo;
        public double PressureDrop;
        public string DemandWeight;
        public string WindSpeed;
        public string TestGsm; 
        public string TestGas;
        public string Concentration;
        public string Eff0;
        public double SelectedDeltaP;
    }
    // ====== 效率計算 ======
    private class EffRow
    {
        public string Gas;                // 氣體名稱（SO2/H2S/NH3/...）
        public double Concentration;      // 濃度
        public double Background;         // 背景
        public List<double> Readings11;   // 11 點讀值
        public List<double> Efficiency11; // 11 點效率(%)
        public double Eff1st;             // 第 1 點效率(%)
        public double Eff10th;            // 第 10 點效率(%)  (索引 9)
    }
    private static void ExportToExcel_Page2(string savePath, Page2ExportData d)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_SemiFinished_Template.xlsx"
        );

        File.Copy(templatePath, savePath, true);

        var app = new Excel.Application();
        var wb = app.Workbooks.Open(savePath);
        var ws = (Excel.Worksheet)wb.Sheets[1];

        try
        {
            ws.Range["C6"].Value = d.ReportNo; //報告編號
            ws.Range["F7"].Value = d.ProductName; //
            ws.Range["F8"].Value = d.ProductSize; //成品尺寸
            ws.Range["F9"].Value = d.ProductNo; //產品編號
            ws.Range["H6"].Value = d.TestDate; //測試日期
            ws.Range["H10"].Value = d.TestGsm;
            ws.Range["F14"].Value = d.PressureDrop; //樣品壓損
            ws.Range["F13"].Value = d.WindSpeed; //風速
            ws.Range["F11"].Value = d.TestGas; //測試氣體
            ws.Range["F12"].Value = d.Concentration; //測試濃度
            ws.Range["F16"].Value = d.Eff0; //初始效率
            wb.Save();
        }
        finally
        {
            wb.Close();
            app.Quit();
        }
    }

    private static EffRow BuildEffRow(string gas, double conc, double bg, List<double> readings)
    {
        var eff = new List<double>();
        if (conc <= 0) conc = 1; // 防止除以 0
        foreach (var r in readings)
            eff.Add(((conc - bg - r) / conc) * 100.0);

        return new EffRow
        {
            Gas = gas,
            Concentration = conc,
            Background = bg,
            Readings11 = readings,
            Efficiency11 = eff,
            Eff1st = eff.Count > 0 ? eff[0] : 0,
            Eff10th = eff.Count > 9 ? eff[9] : (eff.Count > 0 ? eff.Last() : 0)
        };
    }

    // ====== 主流程：收集 → 驗證 → 讓使用者選重量 → 產出〈總表〉列資料 ======
    public static void Handle(TabPage tab)
    {
        try
        {
            // ① 抓欄位資料
            string reportNo = GetText(tab, "FilterInProcessReportNOTB");
            string productionDate = GetText(tab, "FilterInProcessProductionBox");
            string testDate = GetText(tab, "FilterInProcessTestBox");
            string productType = GetText(tab, "FilterInProcessTypeBox");

            string carbonOrder =
                    GetText(tab, "FilterInProcessCarbonOrderBox") +
                    "_" +
                    GetText(tab, "FilterInProcessOrderBox");
            string ProductSize = GetText(tab, "FilterSizeTB");
            string demandWeight = GetText(tab, "FilterInProcessgsmBox");
            string glue = GetText(tab, "FilterInProcessGileBox");
            string speed = GetText(tab, "FilterInProcessSpeedBox");
            string ovenUp = GetText(tab, "FilterInProcessUpperBox");
            string ovenDown = GetText(tab, "FilterInProcessLowerBox");
            string pressure = GetText(tab, "FilterInProcessPressureBox");
            string windSpeed = GetText(tab, "FilterInProcessWindBox");
            string carbonInfo = GetText(tab, "FilterInProcessCarbonInfoBox"); 
            string prodNo = GetText(tab, "FilterInProcessOrderBox");
            string productNumber = "";
            string carbonNo = GetText(tab, "FilterInProcessCarbonOrderBox");
            string mmddProduction = "";
            if (DateTime.TryParseExact(
                    productionDate,
                    "yyyy.MM.dd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime dt))
            {
                mmddProduction = $"({dt:MMdd}生產)";
            }
            if (!string.IsNullOrWhiteSpace(carbonNo) && !string.IsNullOrWhiteSpace(prodNo))
                productNumber = $"{carbonNo}_{prodNo}";
            else
                productNumber = carbonNo + prodNo;
            string productName = $"{productType} {demandWeight}gsm{mmddProduction}";
            // ② 拆分重量 / 壓損
            List<double> weights = ParseDoubleList(GetText(tab, "FilterInProcessTestGsmBox"));
            List<double> deltas = ParseDoubleList(GetText(tab, "FilterInProcessPressureDropBox"));

            if (weights.Count == 0 || deltas.Count == 0)
            {
                MessageBox.Show("請確認重量與壓損欄位皆已輸入數值。", "資料錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (weights.Count != deltas.Count)
            {
                MessageBox.Show($"重量筆數 ({weights.Count}) 與 壓損筆數 ({deltas.Count}) 不一致，請確認。",
                                "資料數量不一致", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ③ 讓使用者選擇要使用的壓損
            int n = Math.Min(weights.Count, deltas.Count);
            weights = weights.Take(n).ToList();
            deltas = deltas.Take(n).ToList();

            if (n == 0)
            {
                MessageBox.Show("沒有可用的測試筆數，請確認重量與壓損輸入。");
                return;
            }

            // 用 Form2 讓使用者選第幾筆
            int selectedIdx = -1;
            using (var f = new Form2(deltas, "請選擇用哪一筆壓損"))
            {
                if (f.ShowDialog() == DialogResult.OK)
                    selectedIdx = f.SelectedIndex0;   // 0-based 索引
                else
                {
                    MessageBox.Show("未選擇資料列，已取消。");
                    return;
                }
            }

            if (selectedIdx < 0 || selectedIdx >= n)
            {
                MessageBox.Show("選擇的索引超出範圍。");
                return;
            }
            double selectedTestGsm = weights[selectedIdx];
            // ④ 勾選氣體與效率
            var checkedBoxes = GetAllCheckBoxes(tab).Where(chk => chk.Checked).ToList();
            if (checkedBoxes.Count == 0)
            {
                MessageBox.Show("請至少選擇一項測試氣體。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ------------------------------------------------------------
            // ④ 建立氣體效率結果 (effList)
            // ------------------------------------------------------------
            var effList = new List<Dictionary<string, string>>();

            foreach (var chk in checkedBoxes)
            {
                // 從 CheckBox 名稱中取出氣體名稱
                // 假設 CheckBox 名稱結構為：FilterIn + gas + CheckBox
                string chkName = chk.Name ?? "";
                string gas = chkName.Replace("FilterIn", "")
                                    .Replace("CheckBox", "")
                                    .Trim();

                if (string.IsNullOrEmpty(gas))
                {
                    MessageBox.Show("找不到氣體名稱，請確認控制項命名規則。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }

                // 找對應的濃度 / 背景 / 讀值欄位
                TextBox tbConc = FindTextBox(tab, "FilterIn" + gas + "ConcentrationBox");
                TextBox tbBg = FindTextBox(tab, "FilterIn" + gas + "BackgroundBox");
                TextBox tbVal = FindTextBox(tab, "FilterIn" + gas + "ValueBox");

                double conc = tbConc != null && double.TryParse(tbConc.Text, out double c) ? c : 0;
                double bg = tbBg != null && double.TryParse(tbBg.Text, out double b) ? b : 0;
                double val = tbVal != null && double.TryParse(tbVal.Text, out double v) ? v : 0;

                if (conc == 0)
                {
                    MessageBox.Show($"氣體 {gas} 的濃度為 0，請確認輸入正確。", "資料錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }

                // 計算效率（示意）
                double eff1 = ((conc - bg - val) / conc) * 100;
                double eff10 = eff1 - 0.2; // 假設十分鐘效率略降

                effList.Add(new Dictionary<string, string>
                {
                    ["氣體名稱"] = gas,
                    ["初始濃度"] = conc.ToString(),
                    ["初始效率"] = eff1.ToString("F2"),
                    ["十分鐘效率"] = eff10.ToString("F2")
                });
            }

            if (effList.Count == 0)
            {
                MessageBox.Show("未找到有效氣體資料，請確認輸入。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ⑤ 整合所有資料 → 準備匯出到總表
            //    只有「被選擇的壓損」那一列填 AF~AH；其餘列留空
            if (selectedIdx < 0 || selectedIdx >= deltas.Count)
            {
                MessageBox.Show("選擇的壓損索引無效。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<Dictionary<string, string>> rowsForMaster = new List<Dictionary<string, string>>();
            int pairCount = Math.Min(weights.Count, deltas.Count);

            for (int i = 0; i < pairCount; i++)
            {
                foreach (var eff in effList) // 每個勾選的氣體各一列
                {
                    var row = new Dictionary<string, string>
                    {
                        ["生產日期"] = productionDate,
                        ["測試日期"] = testDate,
                        ["碳線工單"] = carbonOrder,
                        ["需求克數"] = demandWeight,
                        ["膠粉"] = glue,
                        ["速度"] = speed,
                        ["烘爐溫度(上)"] = ovenUp,
                        ["烘爐溫度(下)"] = ovenDown,
                        ["壓力(kg)"] = pressure,
                        ["風速(m/s)"] = windSpeed,

                        ["測重(gsm)"] = weights[i].ToString(),
                        ["壓差(pa)"] = deltas[i].ToString(),

                        ["測試氣體"] = eff["氣體名稱"], // AE
                        ["碳線資訊"] = carbonInfo      // AI
                    };

                    if (i == selectedIdx)
                    {
                        // ✅ 只有被選擇的壓損那一列填入 AF~AH
                        row["初始濃度(ppb)"] = eff["初始濃度"];   // AF
                        row["初始效率(%)"] = eff["初始效率"];   // AG
                        row["十分鐘效率(%)"] = eff["十分鐘效率"]; // AH
                    }
                    else
                    {
                        // 其它列不帶 AF~AH（留空）
                        row["初始濃度(ppb)"] = "";
                        row["初始效率(%)"] = "";
                        row["十分鐘效率(%)"] = "";
                    }

                    rowsForMaster.Add(row);
                }
            }

            // ⑥ 寫入 Excel 總表
            string filePath = @"C:\Users\User\Desktop\總表.xlsx";
            if (!File.Exists(filePath))
            {
                MessageBox.Show("找不到總表.xlsx，請確認路徑正確。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var selectedEff = effList.First(); // 或你要哪一個氣體
            var reportData = new Page2ExportData
            {
                ReportNo = reportNo, //報告編號
                ProductionDate = mmddProduction, //生產日期
                TestDate = testDate, //測試日期
                ProductSize = ProductSize, //成品尺寸
                ProductNo = productNumber,  //產品編號
                WindSpeed = windSpeed, //風速
                TestGas = selectedEff["氣體名稱"],
                Concentration = selectedEff["初始濃度"],
                Eff0 = selectedEff["初始效率"],
                SelectedDeltaP = deltas[selectedIdx],
                TestGsm = selectedTestGsm.ToString(),
                PressureDrop = deltas[selectedIdx],
                ProductName = productName,
            };

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel (*.xlsx)|*.xlsx";
                sfd.FileName = $"{reportData.ReportNo}_{productType}_{demandWeight}_{productNumber}{mmddProduction}.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    ExportToExcel_Page2(sfd.FileName, reportData);
                }
            }
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("濾網");
                var usedCells = ws.Column(19).Cells(c => c.Address.RowNumber >= 5 && !c.IsEmpty());
                int row = (usedCells.LastOrDefault()?.Address.RowNumber ?? 4) + 1;

                foreach (var data in rowsForMaster)
                {
                    ws.Cell(row, "S").Value = data["生產日期"];
                    ws.Cell(row, "T").Value = data["測試日期"];
                    ws.Cell(row, "U").Value = data["碳線工單"];
                    ws.Cell(row, "V").Value = data["需求克數"];
                    ws.Cell(row, "W").Value = data["膠粉"];
                    ws.Cell(row, "X").Value = data["速度"];
                    ws.Cell(row, "Y").Value = data["烘爐溫度(上)"];
                    ws.Cell(row, "Z").Value = data["烘爐溫度(下)"];
                    ws.Cell(row, "AA").Value = data["壓力(kg)"];
                    ws.Cell(row, "AB").Value = data["風速(m/s)"];
                    ws.Cell(row, "AC").Value = data["測重(gsm)"];
                    ws.Cell(row, "AD").Value = data["壓差(pa)"];
                    ws.Cell(row, "AE").Value = data["測試氣體"];
                    ws.Cell(row, "AF").Value = data["初始濃度(ppb)"];
                    ws.Cell(row, "AG").Value = data["初始效率(%)"];
                    ws.Cell(row, "AH").Value = data["十分鐘效率(%)"];
                    ws.Cell(row, "AI").Value = data["碳線資訊"];
                    row++;
                }
                wb.Save();
            }

            MessageBox.Show("已匯入", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        { 
            MessageBox.Show($"匯出時發生錯誤：{ex.Message}", "例外", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
public static class ExportHelper_Page3
{
    public static void Handle(TabPage tab)
    {
        // 先檢查欄位是否有漏填
        if (!ValidationHelper.CheckRequiredFields(tab))
        {
            MessageBox.Show("請填寫所有欄位再執行", "缺少資料", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var formData = ExportHelper.CollectTabData(tab);
        var filePath = @"C:\Users\User\Desktop\總表.xlsx";
        var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet("濾網");
        string carbonLotRaw = formData.GetValueOrDefault("FilterCarbonLotBox", "");

        // 取得測試類型字串
        string typeCombined = GetTestTypesFromLots(ws, carbonLotRaw);

        // 寫入 AR 欄 (第44欄)


        // 先把 CarbonLot 拆分
        var carbonLots = carbonLotRaw.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(c => c.Trim())
                                     .ToList();

        // 取得 DataGridView
        var dgv = tab.Controls.Find("FilterBox", true).FirstOrDefault() as DataGridView;

        foreach (var carbonLot in carbonLots)
        {
            int row = ws.Column(36).CellsUsed().LastOrDefault()?.Address.RowNumber ?? 1;
            row++;  // 下一列

            // 基本欄位寫入
            ws.Cell(row, 36).Value = formData.GetValueOrDefault("FilterProductionBox", "");
            ws.Cell(row, 37).Value = formData.GetValueOrDefault("FilterTestDateBox", "");
            ws.Cell(row, 38).Value = formData.GetValueOrDefault("FilterReportBox", "");
            ws.Cell(row, 39).Value = carbonLot; // 用拆分後的 CarbonLot
            ws.Cell(row, 40).Value = formData.GetValueOrDefault("FilterPackageNOBox", "");
            ws.Cell(row, 41).Value = formData.GetValueOrDefault("FilterReportCustmorBox", "");
            ws.Cell(row, 42).Value = formData.GetValueOrDefault("FilterReportCustmorBox", "");
            ws.Cell(row, 43).Value = formData.GetValueOrDefault("FilterModelBox", "");
            ws.Cell(row, 44).Value = typeCombined;
            ws.Cell(row, 45).Value = formData.GetValueOrDefault("ReFilterNOBox", "");
            for (int i = 0; i < 4; i++)
            {
                ws.Cell(row, 48 + i).Value = "V";
            }

            // 尺寸欄位拆分 (51~54欄)
            if (formData.TryGetValue("FilterSizeBox", out string sizeRaw))
            {
                var sizes = sizeRaw.Split('/');
                for (int i = 0; i < 4; i++)
                {
                    string value = i < sizes.Length ? sizes[i].Trim() : "";
                    ws.Cell(row, 52 + i).Value = value;
                }
            }

            if (dgv != null && dgv.Rows.Count > 0)
            {
                foreach (DataGridViewRow dataRow in dgv.Rows)
                {
                    if (dataRow.IsNewRow) continue;

                    string GetCell(string colName) => dataRow.Cells[colName]?.Value?.ToString() ?? "0";
                    string Number = GetCell("生產序號");
                    string weight = GetCell("重量");

                    string p_in = GetCell("Particle_In");
                    string p_out = GetCell("Particle_Out");
                    string p_diff = DiffUtil.GetDiff(p_out, p_in);

                    string ipa_in = GetCell("IPA_In");
                    string ipa_out = GetCell("IPA_Out");
                    string ipa_diff = DiffUtil.GetDiff(ipa_out, ipa_in);

                    string ace_in = GetCell("Acetone_In");
                    string ace_out = GetCell("Acetone_Out");
                    string ace_diff = DiffUtil.GetDiff(ace_out, ace_in);

                    string nt_in = GetCell("Nontarget_In");
                    string nt_out = GetCell("Nontarget_Out");
                    string nt_diff = DiffUtil.GetDiff(nt_out, nt_in);
                    string TVOC = DiffUtil.GetSumDiff(ipa_out, ace_out, nt_out, ipa_in, ace_in, nt_in);

                    string drop = GetCell("Pressure_Drop");
                    ws.Cell(row, 46).Value = Number;
                    ws.Cell(row, 47).Value = weight;
                    ws.Cell(row, 59).Value = p_in;
                    ws.Cell(row, 60).Value = p_out;
                    ws.Cell(row, 61).Value = p_diff;
                    ws.Cell(row, 62).Value = ipa_in;
                    ws.Cell(row, 63).Value = ipa_out;
                    ws.Cell(row, 64).Value = ipa_diff;
                    ws.Cell(row, 65).Value = ace_in;
                    ws.Cell(row, 66).Value = ace_out;
                    ws.Cell(row, 67).Value = ace_diff;
                    ws.Cell(row, 68).Value = nt_in;
                    ws.Cell(row, 69).Value = nt_out;
                    ws.Cell(row, 70).Value = nt_diff;
                    ws.Cell(row, 71).Value = TVOC;

                    string filterModel = formData.GetValueOrDefault("FilterModelBox", "");
                    // 預設先給 "N/A"
                    ws.Cell(row, 72).Value = "N/A";  // BS
                    ws.Cell(row, 73).Value = "N/A";  // BT
                    ws.Cell(row, 74).Value = "N/A";  // BU
                    ws.Cell(row, 75).Value = "N/A";  // BV

                    if (IsSingleFilter(filterModel))
                    {
                        ws.Cell(row, 72).Value = GetCell("Pressure_Drop_Spec");  // BS
                        ws.Cell(row, 73).Value = GetCell("Pressure_Drop");       // BT
                    }
                    else if (IsSetFilter(filterModel))
                    {
                        ws.Cell(row, 74).Value = GetCell("Pressure_Drop_Spec");  // BU
                        ws.Cell(row, 75).Value = GetCell("Pressure_Drop");       // BV
                    }

                    string filterType = formData.GetValueOrDefault("FilterTypeBox", "");
                    List<string> targetTypes = filterType.Split('+').Select(s => s.Trim()).ToList();
                    var efficiencyMap = EfficiencyFinder.FindEfficiencyByTypeWithMin(ws, carbonLot, targetTypes);

                    ws.Cell(row, 56).Value = efficiencyMap["MA"];
                    ws.Cell(row, 57).Value = efficiencyMap["MB"];
                    ws.Cell(row, 58).Value = efficiencyMap["MC"];

                    row++;
                }
            }
        }

        wb.Save();
        MessageBox.Show("匯入完成！", "已成功匯入資料", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static bool IsSingleFilter(string model)
    {
        string[] singleTypes = { "6VS", "6VB", "4VS", "4VB", "Panel", "抽取式單片", "堆叠式單片" };
        return singleTypes.Contains(model);
    }

    private static bool IsSetFilter(string model)
    {
        string[] setTypes = { "蒸籠式", "抽取式整組", "堆疊式整組" };
        return setTypes.Contains(model);
    }
    private static string GetTestTypesFromLots(IXLWorksheet ws, string lotBoxRaw)
    {
        // 分割工單號
        var lots = lotBoxRaw.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(l => l.Trim())
                            .ToList();

        var types = new HashSet<string>(); // 用 HashSet 去除重複

        foreach (var lot in lots)
        {
            // 找到 U 欄 (21) = 工單號所在列
            var row = ws.Column(21).CellsUsed()
                        .FirstOrDefault(c => c.GetString().Trim().Equals(lot, StringComparison.OrdinalIgnoreCase))
                        ?.Address.RowNumber;

            if (row.HasValue)
            {
                string gas = ws.Cell(row.Value, 31).GetString().Trim(); // AE = 第31欄

                if (gas == "SO2" || gas == "H2S")
                    types.Add("MA");
                else if (gas == "NH3")
                    types.Add("MB");
                else if (gas == "Toluene" || gas == "IPA" || gas == "Acetone")
                    types.Add("MC");
            }
        }

        return string.Join("+", types); // 用 + 串起來
    }

}
public static class ExportHelper_Page4
{
    private static IEnumerable<T> Descendants<T>(Control root) where T : Control
    {
        foreach (Control c in root.Controls)
        {
            if (c is T t) yield return t;
            foreach (var d in Descendants<T>(c)) yield return d;
        }
    }

    public static Dictionary<string, string> CollectTabData_CylinderEff(TabPage tabPage)
    {
        Dictionary<string, string> formData = new Dictionary<string, string>();
        formData["目前分頁"] = tabPage.Name;

        // ✅ 找到 TableLayoutPanel
        var layoutPanel = tabPage.Controls.Find("CylinderRawEffPanel", true).FirstOrDefault() as TableLayoutPanel;

        if (layoutPanel != null)
        {
            // 尋找所有 CheckBox 控制項
            var allCheckBoxes = layoutPanel.Controls.OfType<CheckBox>();

            foreach (var checkBox in allCheckBoxes)
            {
                if (!checkBox.Checked) continue; // 跳過沒勾選的

                int row = layoutPanel.GetRow(checkBox);

                // 找出該行所有控制項
                var controlsInRow = layoutPanel.Controls
                    .Cast<Control>()
                    .Where(c => layoutPanel.GetRow(c) == row)
                    .ToList();

                // 嘗試從該列中找到唯一的氣體 Label
                string gas = controlsInRow
                    .OfType<Label>()
                    .Where(l => !string.IsNullOrWhiteSpace(l.Text)) // 避免空的
                    .FirstOrDefault()?.Text ?? $"Row{row}";

                // 依照名稱格式明確抓欄位
                var concentrationBox = controlsInRow
                    .OfType<TextBox>()
                    .FirstOrDefault(tb => tb.Name.Equals($"CylinderRaw{gas}ConcertrationBox", StringComparison.OrdinalIgnoreCase));

                var backgroundBox = controlsInRow
                    .OfType<TextBox>()
                    .FirstOrDefault(tb => tb.Name.Equals($"CylinderRaw{gas}BackGroundBox", StringComparison.OrdinalIgnoreCase));

                var readingBox = controlsInRow
                    .OfType<TextBox>()
                    .FirstOrDefault(tb => tb.Name.Equals($"CylinderRaw{gas}ValueBox", StringComparison.OrdinalIgnoreCase));

                // 取出值
                string concentration = concentrationBox?.Text ?? "";
                string background = backgroundBox?.Text ?? "";
                string reading = readingBox?.Text ?? "";
                // 加入 formData 字典
                formData[$"Efficiency_{gas}_濃度"] = concentration;
                formData[$"Efficiency_{gas}_背景"] = background;
                formData[$"Efficiency_{gas}_讀值"] = reading;
            }
        }
        return formData;
    }
    public static T FindCtrl<T>(Control root, string nameContains) where T : Control
    {
        if (root == null || string.IsNullOrEmpty(nameContains)) return null;

        foreach (Control c in root.Controls)
        {
            if (c is T t && c.Name.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) >= 0)
                return t;

            var child = FindCtrl<T>(c, nameContains); // 遞迴
            if (child != null) return child;
        }
        return null;
    }
    private static T FindByExactName<T>(Control root, string exactName) where T : Control
    {
        foreach (Control c in root.Controls)
        {
            if (c is T t && string.Equals(c.Name, exactName, StringComparison.OrdinalIgnoreCase))
                return t;
            var child = FindByExactName<T>(c, exactName);
            if (child != null) return child;
        }
        return null;
    }
    private static List<string> SplitStr(string raw)
    {
        var sep = new[] { '/', ',', '、', '|', ' ', '\r', '\n', ';' };
        return (raw ?? "").Split(sep, StringSplitOptions.RemoveEmptyEntries)
                          .Select(s => s.Trim()).ToList();
    }
    private static List<double> SplitDouble(string raw)
        => SplitStr(raw).Select(s => double.TryParse(s, out var v) ? v : double.NaN)
                        .Where(v => !double.IsNaN(v)).ToList();
    private static List<(string gas, (string conc, string eff0, string eff10) res)>
    CollectEfficiencyBlocksByGasNames(TabPage tab)
    {
        var result = new List<(string gas, (string conc, string eff0, string eff10) res)>();
        var errors = new List<string>();

        // 找出所有名稱符合 ^CylinderRaw(.+)CheckBox$ 的 CheckBox
        var allChecks = Descendants<CheckBox>(tab)
            .Where(c => Regex.IsMatch(c.Name, @"^CylinderRaw(.+)CheckBox$", RegexOptions.IgnoreCase))
            .ToList();
        // 只取 Checked = true
        var checkedOnes = allChecks.Where(c => c.Checked).ToList();
        if (checkedOnes.Count == 0) return result;
        foreach (var chk in checkedOnes)
        {
            var m = Regex.Match(chk.Name, @"^CylinderRaw(.+)CheckBox$", RegexOptions.IgnoreCase);
            if (!m.Success) continue;
            string gasKey = m.Groups[1].Value;

            // 試著讀「測試氣體的 TB」：CylinderRaw + 氣體名稱
            string gasTbName = "CylinderRaw" + gasKey;
            string gasText =
                (FindByExactName<TextBox>(tab, gasTbName)?.Text ??
                 FindByExactName<ComboBox>(tab, gasTbName)?.Text ??
                 chk.Text ?? gasKey)?.Trim();
            if (string.IsNullOrEmpty(gasText)) gasText = gasKey;

            // 組出三件組名稱
            string concName = "CylinderRaw" + gasKey + "ConcertrationBox";
            string bgName = "CylinderRaw" + gasKey + "BackGroundBox";
            string valName = "CylinderRaw" + gasKey + "ValueBox";

            // 取得資料
            var (c, b, r) = TryGetEffBlockExact(tab, concName, bgName, valName);

            // 驗證
            if (c <= 0) errors.Add($"[{gasText}] 濃度必須 > 0（{concName}）");
            if (FindByExactName<TextBox>(tab, bgName) != null &&
                !double.TryParse(FindByExactName<TextBox>(tab, bgName).Text, out _))
                errors.Add($"[{gasText}] 背景值不是有效數字（{bgName}）");
            if (r == null || r.Count != 11) errors.Add($"[{gasText}] 讀值需為 11 筆（{valName}）");

            if (errors.Count == 0)
            {
                // 計算效率
                var res = ComputeEfficiency(c, b, r); // (conc, eff0, eff10)
                result.Add((gasText, res));
            }
        }

        if (errors.Count > 0)
        {
            MessageBox.Show("效率設定有問題，請修正後再試：\n\n" + string.Join("\n", errors),
                            "效率資料不完整", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null; // 用 null 代表本次不可繼續
        }

        return result;
    }

    private static (double conc, double bg, List<double> readings) TryGetEffBlockExact(
    TabPage tab, string concName, string bgName, string valName)
    {
        var tbConc = FindByExactName<TextBox>(tab, concName) ?? FindCtrl<TextBox>(tab, concName);
        var tbBg = FindByExactName<TextBox>(tab, bgName) ?? FindCtrl<TextBox>(tab, bgName);
        var tbVal = FindByExactName<TextBox>(tab, valName) ?? FindCtrl<TextBox>(tab, valName);

        double conc = 0, bg = 0;
        double.TryParse(tbConc?.Text, out conc);
        double.TryParse(tbBg?.Text, out bg);

        // 你現有的 SplitDouble/ SplitStr 可沿用；這裡再保一次
        var readings = (tbVal?.Text ?? "")
            .Split(new[] { '\r', '\n', '/', ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => double.TryParse(s.Trim(), out var v) ? v : double.NaN)
            .Where(v => !double.IsNaN(v))
            .ToList();

        return (conc, bg, readings);
    }

    private static (string conc, string eff0, string eff10)
        ComputeEfficiency(double conc, double bg, List<double> r)
    {
        if (conc <= 0 || r == null || r.Count < 11) return ("", "", "");
        double e0 = (conc - bg - r[0]) / conc * 100.0;
        double e10 = (conc - bg - r[10]) / conc * 100.0;
        return (conc.ToString("F0"), e0.ToString("F1"), e10.ToString("F1"));
    }

    private static string BuildMeshSummaryString(TabPage tab, string material)
    {
        // 1) 取得 DGV（TabPage4 你用的是 CylinderRawMeshBox；也容忍包含 'Particle' 的名字）
        var dgv = FindByExactName<DataGridView>(tab, "CylinderRawMeshBox")
                  ?? FindByExactName<DataGridView>(tab, "Particle");
        if (dgv == null) return "";

        const int LABEL_COL = 0;  // 粒徑標籤欄
        const int VALUE_COL = 1;  // 重量欄

        // 2) 讀取資料：找出總重、已填的層、空白層
        double total = 0;
        bool hasTotal = false;

        var weights = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        var blankRows = new List<DataGridViewRow>();   // 記錄哪一列空白（之後補值用）

        foreach (DataGridViewRow r in dgv.Rows)
        {
            if (r.IsNewRow) continue;
            string label = Convert.ToString(r.Cells[LABEL_COL].Value)?.Trim();
            if (string.IsNullOrEmpty(label)) continue; // 沒標籤就跳過

            string sVal = Convert.ToString(r.Cells[VALUE_COL].Value)?.Trim();

            if (label == "總重")
            {
                hasTotal = double.TryParse(sVal, out total);
                continue;
            }

            if (double.TryParse(sVal, out var v))
            {
                // 有填數字 → 記錄
                weights[label] = v;
            }
            else
            {
                // 空白或非數字 → 等待補值
                blankRows.Add(r);
                // 先把這層放進 dict，給後面輸出排序/一致性用（值暫不設或設為 0 都行）
                if (!weights.ContainsKey(label)) weights[label] = 0;
            }
        }

        if (!hasTotal)
        {
            MessageBox.Show("請先在「總重」輸入有效數值。");
            return "";
        }

        // 3) 若剛好只有一層空白 → 以「總重 − 其他層總和」補上
        if (blankRows.Count == 1)
        {
            double sumKnown = weights.Where(kv => kv.Key != "總重").Sum(kv => kv.Value);
            double remain = total - sumKnown;

            // 微小誤差修正
            if (Math.Abs(remain) < 1e-9) remain = 0;

            if (remain < 0)
            {
                MessageBox.Show($"各層已填重量加總（{sumKnown:F3}）已大於總重（{total:F3}），請確認數值。");
                return "";
            }

            // 寫回空白那格
            var r = blankRows[0];
            r.Cells[VALUE_COL].Value = Math.Round(remain, 3);

            // 同步回字典
            string label = Convert.ToString(r.Cells[LABEL_COL].Value)?.Trim();
            if (!string.IsNullOrEmpty(label))
                weights[label] = remain;
        }
        else if (blankRows.Count > 1)
        {
            MessageBox.Show("有超過一格未輸入重量，無法由總重推回；請僅留一格空白再試。");
            return "";
        }
        // 若 0 格空白 → 直接進入百分比計算

        // 4) 計算百分比並組摘要字串
        var parts = new List<string>();
        double sumPct = 0;

        foreach (var kv in weights)
        {
            // 已排除「總重」，只對各層計算
            double pct = (total > 0) ? (kv.Value / total * 100.0) : 0.0;
            parts.Add($"{kv.Key} {pct:F1}%");
            sumPct += pct;
        }

        // 5) 檢查百分比總和（允許 ±2% 的誤差）
        if (Math.Abs(sumPct - 100.0) > 2.0)
        {
            MessageBox.Show($"粒徑分布百分比加總為 {sumPct:F1}%（應約為 100%），請確認各層數值與總重是否正確。");
        }

        // 6) 回傳摘要字串（寫回 Excel 用）
        return string.Join("; ", parts);
    }
    // 進貨數量字串（重量單位 IKP201/IKP205 → lb；其他 → kg；第二格是「包」）
    private static string BuildQuantityString(string material, string weightText, string packText)
    {
        string unit = (material ?? "").Trim().ToUpperInvariant() is string m
                      && (m.Contains("IKP201") || m.Contains("IKP205")) ? "lb" : "kg";

        var w = (weightText ?? "").Trim();
        var p = (packText ?? "").Trim();

        if (w == "" && p == "") return "";
        var left = w != "" ? $"{w} {unit}" : "";
        var right = p != "" ? $"{p}包" : "";
        return left + (left != "" && right != "" ? "/" : "") + right;
    }
    private class Page4ExportData
    {
        public string ReportNo { get; set; }
        public string Material { get; set; }
        public string ArrivalDate { get; set; }
        public string TestingDate { get; set; }
        public string QtyText { get; set; }

        public List<string> LotFulls { get; set; }
        public List<double> Densities { get; set; }
        public List<double> DeltaPs { get; set; }
        public List<double> VocIns { get; set; }
        public List<double> VocOuts { get; set; }
        public List<string> OutgassingList { get; set; }

        public int SelectedIndex { get; set; }
        public string Eff0 { get; set; }
    }
    private static void ExportToExcel(string savePath, Page4ExportData d)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_RawMaterial_Template.xlsx");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到公版檔：" + templatePath);
            return;
        }

        File.Copy(templatePath, savePath, true);

        var app = new Excel.Application();
        var wb = app.Workbooks.Open(savePath);
        var ws = (Excel.Worksheet)wb.Sheets[1];

        try
        {
            ws.Range["C4"].Value = d.ReportNo;
            ws.Range["E5"].Value = d.Material;
            ws.Range["C5"].Value = d.ArrivalDate;
            ws.Range["C6"].Value = d.TestingDate;
            ws.Range["E6"].Value = d.QtyText;

            const int COL_FIRST = 3;
            const int ROW_LOTNO = 10;
            const int ROW_DENS = 13;
            const int ROW_DP = 14;
            const int ROW_VIN = 15;
            const int ROW_VOUT = 16;
            const int ROW_VOUTG = 17;
            const int ROW_EFF = 18;

            for (int i = 0; i < d.LotFulls.Count; i++)
            {
                int col = COL_FIRST + i;

                ws.Cells[ROW_LOTNO, col].Value = d.LotFulls[i];
                ws.Cells[ROW_DENS, col].Value = d.Densities[i].ToString("0.00");
                ws.Cells[ROW_DP, col].Value = d.DeltaPs[i].ToString("0.0");
                ws.Cells[ROW_VIN, col].Value = d.VocIns[i];
                ws.Cells[ROW_VOUT, col].Value = d.VocOuts[i];
                ws.Cells[ROW_VOUTG, col].Value = d.OutgassingList[i];

                if (i == d.SelectedIndex)
                    ws.Cells[ROW_EFF, col].Value = d.Eff0;
                else
                    ws.Cells[ROW_EFF, col].Value = "N.D.";
            }

            wb.Save();
        }
        finally
        {
            wb.Close();
            app.Quit();
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);
        }
    }

    // ---------------- 主流程：TabPage4 ----------------
    public static void Handle(TabPage tab)
    {
        try
        {
            // ───────────────────────── (A) 基本欄位 ─────────────────────────
            var testPicker = FindByExactName<DateTimePicker>(tab, "CylinderRawTestDateBox") ?? FindCtrl<DateTimePicker>(tab, "CylinderRawTestDateBox");
            var arrivePicker = FindByExactName<DateTimePicker>(tab, "CylinderRawArriveDateBox") ?? FindCtrl<DateTimePicker>(tab, "CylinderRawArriveDateBox");
            var materialBox = FindByExactName<ComboBox>(tab, "CylinderRawTypeBox") ?? FindCtrl<ComboBox>(tab, "CylinderRawTypeBox");

            if (testPicker == null || arrivePicker == null || materialBox == null)
            {
                MessageBox.Show("找不到測試日期 / 到廠日期 / 原料種類");
                return;
            }

            string testDate = testPicker.Value.ToString("yyyy.MM.dd");
            string arriveDate = arrivePicker.Value.ToString("yyyy.MM.dd");
            string arriveDateKey = arrivePicker.Value.ToString("yyyyMMdd"); // 進廠批號用
            string material = (materialBox.Text ?? "").Trim();

            // 供應商批號（/ 分隔）
            var lots = SplitStr(FindByExactName<TextBox>(tab, "CylinderRawLotBox")?.Text
                             ?? FindCtrl<TextBox>(tab, "CylinderRawLotBox")?.Text);

            // 進貨數量（左：重量；右：包數）
            string qtyLeft = (FindByExactName<TextBox>(tab, "CylinderRawQtyWeightBox")?.Text
                            ?? FindCtrl<TextBox>(tab, "CylinderRawQtyWeightBox")?.Text ?? "").Trim();
            string qtyRight = (FindByExactName<TextBox>(tab, "CylinderRawQtyPackBox")?.Text
                            ?? FindCtrl<TextBox>(tab, "CylinderRawQtyPackBox")?.Text ?? "").Trim();

            bool useLb = material.IndexOf("IKP201", StringComparison.OrdinalIgnoreCase) >= 0
                      || material.IndexOf("IKP205", StringComparison.OrdinalIgnoreCase) >= 0;
            string unit = useLb ? "lb" : "kg";
            string qty = "";
            if (qtyLeft != "" || qtyRight != "")
            {
                string left = qtyLeft != "" ? $"{qtyLeft} {unit}" : "";
                string right = qtyRight != "" ? $"{qtyRight} 包" : "";
                qty = left + (left != "" && right != "" ? " / " : "") + right;
            }

            // ───────────────────────── (B) 多筆欄位解析 ─────────────────────────
            var nos = SplitStr(FindByExactName<TextBox>(tab, "CylinderRawNumberBox")?.Text
                                ?? FindCtrl<TextBox>(tab, "CylinderRawNumberBox")?.Text);
            var weights = SplitDouble(FindByExactName<TextBox>(tab, "CylinderRawWeightBox")?.Text
                                ?? FindCtrl<TextBox>(tab, "CylinderRawWeightBox")?.Text);
            var vocIn = SplitDouble(FindByExactName<TextBox>(tab, "CylinderRawVOCsInletBox")?.Text
                                ?? FindCtrl<TextBox>(tab, "CylinderRawVOCsInletBox")?.Text);
            var vocOut = SplitDouble(FindByExactName<TextBox>(tab, "CylinderRawVOCsOutletBox")?.Text
                                ?? FindCtrl<TextBox>(tab, "CylinderRawVOCsOutletBox")?.Text);
            var deltas = SplitDouble(FindByExactName<TextBox>(tab, "CylinderRawPressureBox")?.Text
                                ?? FindCtrl<TextBox>(tab, "CylinderRawPressureBox")?.Text);

            // 密度（以固定體積 50 計）
            const double VOL = 50.0;
            var densities = weights.Select(w => VOL > 0 ? w / VOL : 0).ToList();

            // ───────────────────────── (C) 筆數一致性檢查 ─────────────────────────
            bool ignoreLotEq = (lots.Count == 1 && lots[0] == "-");           // 批號輸入「-」就不檢查筆數一致
            var lengths = new Dictionary<string, int>
            {
                ["原料編號"] = nos.Count,
                ["測試品重量"] = weights.Count,
                ["VOCs Inlet"] = vocIn.Count,
                ["VOCs Outlet"] = vocOut.Count,
                ["壓損"] = deltas.Count
            };
            if (!ignoreLotEq) lengths["供應商批號"] = lots.Count;

            if (lengths.Values.Distinct().Count() > 1)
            {
                int max = lengths.Values.Max();
                var msgs = lengths.Where(kv => kv.Value != max)
                                  .Select(kv => $"{kv.Key} 筆數={kv.Value}")
                                  .ToList();
                MessageBox.Show("欄位筆數不一致，請確認：\n" + string.Join("\n", msgs));
                return;
            }

            int n = lengths.Values.FirstOrDefault();
            if (n <= 0) { MessageBox.Show("沒有可匯出的資料。"); return; }

            // 對齊 n
            nos = nos.Take(n).ToList();
            weights = weights.Take(n).ToList();
            vocIn = vocIn.Take(n).ToList();
            vocOut = vocOut.Take(n).ToList();
            deltas = deltas.Take(n).ToList();
            densities = densities.Take(n).ToList();
            lots = ignoreLotEq ? Enumerable.Repeat("-", n).ToList() : lots.Take(n).ToList();

            // ───────────────────────── (D) 讓使用者選「第幾筆壓損」 ─────────────────────────
            int selectedIdx = -1;
            using (var fSel = new Form2(deltas, "請選擇本次效率對應的壓損"))
            {
                if (fSel.ShowDialog() == DialogResult.OK) selectedIdx = fSel.SelectedIndex0;
                else { MessageBox.Show("已取消。"); return; }
            }
            if (selectedIdx < 0 || selectedIdx >= n)
            { MessageBox.Show("選擇的壓損索引超出範圍。"); return; }

            // ───────────────────────── (E) 含水率 / 灰分 / 正丁烷（勾選才詢問平均） ─────────────────────────
            double? moisture = null, ash = null, nbutane = null;

            var moistureCheck = FindByExactName<CheckBox>(tab, "CylinderRawMoistureCheck") ?? FindCtrl<CheckBox>(tab, "CylinderRawMoistureCheck");
            var ashCheck = FindByExactName<CheckBox>(tab, "CylinderRawAshCheck") ?? FindCtrl<CheckBox>(tab, "CylinderRawAshCheck");
            var nbutaneCheck = FindByExactName<CheckBox>(tab, "CylinderRawNButaneCheck") ?? FindCtrl<CheckBox>(tab, "CylinderRawNButaneCheck");

            if (moistureCheck?.Checked == true)
            {
                using (var f = new FormCalcMoistureAsh(CalcMode.Moisture))
                {
                    if (f.ShowDialog() == DialogResult.OK && f.AverageResult.HasValue) moisture = f.AverageResult.Value;
                    else { MessageBox.Show("未輸入含水率平均值"); return; }
                }
            }
            if (ashCheck?.Checked == true)
            {
                using (var f = new FormCalcMoistureAsh(CalcMode.Ash))
                {
                    if (f.ShowDialog() == DialogResult.OK && f.AverageResult.HasValue) ash = f.AverageResult.Value;
                    else { MessageBox.Show("未輸入灰分平均值"); return; }
                }
            }
            if (nbutaneCheck?.Checked == true)
            {
                using (var f = new FormCalcMoistureAsh(CalcMode.NButane))
                {
                    if (f.ShowDialog() == DialogResult.OK && f.AverageResult.HasValue) nbutane = f.AverageResult.Value;
                    else { MessageBox.Show("未輸入正丁烷平均值"); return; }
                }
            }

            // ───────────────────────── (F) 粒徑摘要（自動補唯一空白層） ─────────────────────────
            string meshSummary = BuildMeshSummaryString(tab, material);

            // ───────────────────────── (G) 效率（不限組數；依命名規則抓勾選的氣體） ─────────────────────────
            var effBlocks = CollectEfficiencyBlocksByGasNames(tab);    // List<(gas, (conc,eff0,eff10))>
            if (effBlocks == null) return;                             // 有錯已提示
            int effGroupCount = Math.Max(effBlocks.Count, 1);
            // ───── 多氣體提醒 ─────
            if (effBlocks.Count > 1)
            {
                var msg = "目前已勾選以下測試氣體，系統將產出多份 QC 報告：\n\n"
                          + string.Join("\n", effBlocks.Select(e => "• " + e.gas))
                          + $"\n\n共 {effBlocks.Count} 份報告，是否繼續？";

                var dr = MessageBox.Show(
                    msg,
                    "多氣體效率匯出確認",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (dr != DialogResult.Yes) return;
            }

            // ───────────────────────── (H) 組列 rows ─────────────────────────
            var rows = new List<Dictionary<string, string>>();

            for (int g = 0; g < effGroupCount; g++)
            {
                string gasName = (g < effBlocks.Count) ? effBlocks[g].gas : "";

                string conc = "", eff0 = "", eff10 = "";
                if (g < effBlocks.Count)
                {
                    var t = effBlocks[g].res; // (conc, eff0, eff10)
                    conc = t.Item1;
                    eff0 = t.Item2;
                    eff10 = t.Item3;
                }

                for (int i = 0; i < n; i++)
                {
                    string lotStr = lots[i];
                    string lotNoBase = $"B-{arrivePicker.Value:yyyyMMdd}-001";
                    string lotNoFull = $"{lotNoBase}#{nos[i].PadLeft(2, '0')}";

                    double omi = vocOut[i] - vocIn[i];
                    string outMinusInStr = (omi <= 0) ? "N.D." : omi.ToString("F1");

                    var row = new Dictionary<string, string>
                    {
                        ["測試日期"] = testDate,
                        ["到廠日期"] = arriveDate,
                        ["原料種類"] = material,
                        ["供應商批號"] = lotStr,
                        ["進廠批號"] = lotNoFull,
                        ["進貨數量"] = qty,

                        ["樣品重量"] = weights[i].ToString("F3"),
                        ["密度"] = densities[i].ToString("F3"),
                        ["含水率"] = moisture?.ToString("F2") ?? "",
                        ["灰分(%)"] = ash?.ToString("F2") ?? "",
                        ["正丁烷(%)"] = nbutane?.ToString("F2") ?? "",
                        ["VOC Inlet"] = vocIn[i].ToString("F1"),
                        ["VOC Outlet"] = vocOut[i].ToString("F1"),
                        ["OutMinusIn"] = outMinusInStr,
                        ["壓損(pa)"] = deltas[i].ToString("F0"),

                        ["粒徑分布"] = meshSummary,

                        // ✅ 每列都帶「測試氣體」名稱（方便過濾/查閱）
                        ["測試氣體"] = gasName,

                        // 先置空，必要時再覆寫
                        ["初始濃度(ppb)"] = "",
                        ["初始效率(%)"] = "",
                        ["十分鐘效率(%)"] = ""
                    };

                    // ✅ 只有使用者選的壓損那一列才填入效率與濃度
                    if (g < effBlocks.Count && i == selectedIdx)
                    {
                        row["初始濃度(ppb)"] = (i == selectedIdx) ? conc : "";
                        row["初始效率(%)"] = (i == selectedIdx) ? eff0 : "";
                        row["十分鐘效率(%)"] = (i == selectedIdx) ? eff10 : "";
                    }

                    rows.Add(row);
                }
            }
            // ───────────────────────── (I) 寫入 Excel 濾筒分頁 ─────────────────────────
            string filePath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "總表.xlsx"
            );
            using (var wb = new ClosedXML.Excel.XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("濾筒");
                if (ws == null) { MessageBox.Show("總表.xlsx 中找不到「濾筒」工作表"); return; }

                // 以 B 欄（到廠日期）自第 5 列往下找最後一筆
                var used = ws.Column(2).Cells(c => c.Address.RowNumber >= 5 && !c.IsEmpty());
                int rowIdx = (used.LastOrDefault()?.Address.RowNumber ?? 4) + 1;

                foreach (var d in rows)
                {
                    ws.Cell(rowIdx, 1).Value = d["測試日期"];
                    ws.Cell(rowIdx, 2).Value = d["到廠日期"];
                    ws.Cell(rowIdx, 3).Value = d["原料種類"];
                    ws.Cell(rowIdx, 4).Value = d["供應商批號"];
                    ws.Cell(rowIdx, 5).Value = d["進廠批號"];
                    ws.Cell(rowIdx, 6).Value = d["進貨數量"];
                    ws.Cell(rowIdx, 8).Value = d["樣品重量"];
                    ws.Cell(rowIdx, 9).Value = d["密度"];
                    ws.Cell(rowIdx, 10).Value = d["含水率"];
                    ws.Cell(rowIdx, 11).Value = d.ContainsKey("測試氣體") ? d["測試氣體"] : "";
                    ws.Cell(rowIdx, 12).Value = d["VOC Inlet"];
                    ws.Cell(rowIdx, 13).Value = d["VOC Outlet"];
                    ws.Cell(rowIdx, 14).Value = d["OutMinusIn"];
                    ws.Cell(rowIdx, 15).Value = d["壓損(pa)"];
                    ws.Cell(rowIdx, 16).Value = d["初始效率(%)"];
                    ws.Cell(rowIdx, 17).Value = d["十分鐘效率(%)"];
                    ws.Cell(rowIdx, 18).Value = d["初始濃度(ppb)"];
                    ws.Cell(rowIdx, 19).Value = d["粒徑分布"];
                    ws.Cell(rowIdx, 20).Value = d["灰分(%)"];
                    ws.Cell(rowIdx, 21).Value = d["正丁烷(%)"];
                    rowIdx++;
                }
                wb.Save();
            }
            // ─────────── (J) 多氣體 → 多份 QC Report（模板） ───────────

            // 組 LotFulls / Outgassing（與 Page1 一致）
            var lotFulls = new List<string>();
            var outStrings = new List<string>();

            for (int i = 0; i < n; i++)
            {
                lotFulls.Add($"B-{arrivePicker.Value:yyyyMMdd}-001#{nos[i].PadLeft(2, '0')}");

                double diff = vocOut[i] - vocIn[i];
                outStrings.Add(diff <= 0 ? "N.D." : diff.ToString("F1"));
            }

            // 報告編號（若沒有可為空）
            string reportNo =
                FindByExactName<TextBox>(tab, "CylinderRawReportNoTB")?.Text.Trim() ?? "";

            // 讓使用者選「輸出資料夾」（只選一次）
            string outputDir = "";
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "請選擇 QC 報告輸出資料夾";
                if (fbd.ShowDialog() != DialogResult.OK) return;
                outputDir = fbd.SelectedPath;
            }

            // 依每個氣體 → 產一份報告
            foreach (var eff in effBlocks)
            {
                string gasName = eff.gas;
                string eff0 = eff.res.eff0;

                var reportData = new Page4ExportData
                {
                    ReportNo = reportNo,
                    Material = material,
                    ArrivalDate = arriveDate,
                    TestingDate = testDate,
                    QtyText = qty,

                    LotFulls = lotFulls,
                    Densities = densities,
                    DeltaPs = deltas,
                    VocIns = vocIn,
                    VocOuts = vocOut,
                    OutgassingList = outStrings,

                    SelectedIndex = selectedIdx,
                    Eff0 = eff0
                };

                string arriveShort = arrivePicker.Value.ToString("MMdd");
                string safeGas = string.Concat(
                    gasName.Where(c => !Path.GetInvalidFileNameChars().Contains(c))
                );

                // 👉 只有一個氣體時，不加 _氣體
                string fileName;
                if (effBlocks.Count == 1)
                {
                    fileName = $"{reportNo}_{material}({arriveShort}到廠).xlsx";
                }
                else
                {
                    fileName = $"{reportNo}_{material}_{safeGas}({arriveShort}到廠).xlsx";
                }

                string fullPath = Path.Combine(outputDir, fileName);

                ExportToExcel(fullPath, reportData);
            }

            MessageBox.Show(
                $"QC 報表匯出完成，共產出 {effBlocks.Count} 份報告。",
                "匯出完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show("TabPage4 匯出錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page5
{
    public static void Handle(TabPage tab)
    {
        if (!ValidationHelper.CheckRequiredFields(tab))
        {
            MessageBox.Show("請填寫所有欄位再執行", "缺少資料", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var formData = ExportHelper.CollectTabData(tab);
        var filePath = @"C:\Users\User\Desktop\總表.xlsx";
        var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet("濾筒");
        int row = (ws.Column(38).CellsUsed().LastOrDefault()?.Address.RowNumber ?? 3) + 1;

        var dgv = tab.Controls.Find("CylinderBox", true).FirstOrDefault() as DataGridView;
        if (dgv != null && dgv.Rows.Count > 0)
        {
            var dataRows = dgv.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToList();

            if (dataRows.Count % 16 != 0)
            {
                MessageBox.Show("資料筆數異常，請確認資料！", "資料錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int groupStartRow = row;
            int rowIndex = 0;
            Dictionary<int, string> controlValues = new Dictionary<int, string>();

            foreach (var dataRow in dataRows)
            {
                string GetCell(string colName) => dataRow.Cells[colName]?.Value?.ToString() ?? "";
                bool isControlRow = rowIndex % 16 == 0;

                if (!isControlRow)
                {
                    string[] checkColumns = { "CYL_Particle_In", "CYL_Particle_out", "CYL_IPA_in", "CYL_IPA_out",
                                          "CYL_Acetone_In", "CYL_Acetone_out", "CYL_Nontarget_in", "CYL_Nontarget_out",
                                          "CYL_Pressure_Drop" };
                    foreach (var col in checkColumns)
                    {
                        if (!string.IsNullOrWhiteSpace(GetCell(col)))
                        {
                            MessageBox.Show($"資料錯誤！第 {rowIndex + 1} 列含有不應填寫的測試數據，請檢查。", "資料錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                ws.Cell(row, 19).Value = formData.GetValueOrDefault("CylinderTestDateBox", "");
                ws.Cell(row, 20).Value = formData.GetValueOrDefault("CylinderReportNOBox", "");
                ws.Cell(row, 21).Value = formData.GetValueOrDefault("CylinderNOBox", "");
                ws.Cell(row, 22).Value = formData.GetValueOrDefault("CylinderCustmorBox", "");
                ws.Cell(row, 23).Value = "CYL";
                ws.Cell(row, 24).Value = formData.GetValueOrDefault("CYLTypeBox", "");
                ws.Cell(row, 25).Value = formData.GetValueOrDefault("ReCylinderNOBox", "");

                for (int i = 0; i < 4; i++)
                {
                    ws.Cell(row, 28 + i).Value = "V";
                }

                if (isControlRow)
                {
                    controlValues.Clear();
                    controlValues[35] = GetCell("CYL_Particle_In");
                    controlValues[36] = GetCell("CYL_Particle_out");
                    controlValues[37] = DiffUtil.GetDiff(GetCell("CYL_Particle_out"), GetCell("CYL_Particle_In"));

                    controlValues[38] = GetCell("CYL_IPA_in");
                    controlValues[39] = GetCell("CYL_IPA_out");
                    controlValues[40] = DiffUtil.GetDiff(GetCell("CYL_IPA_out"), GetCell("CYL_IPA_in"));

                    controlValues[41] = GetCell("CYL_Acetone_In");
                    controlValues[42] = GetCell("CYL_Acetone_out");
                    controlValues[43] = DiffUtil.GetDiff(GetCell("CYL_Acetone_out"), GetCell("CYL_Acetone_In"));

                    controlValues[44] = GetCell("CYL_Nontarget_in");
                    controlValues[45] = GetCell("CYL_Nontarget_out");
                    controlValues[46] = DiffUtil.GetDiff(GetCell("CYL_Nontarget_out"), GetCell("CYL_Nontarget_in"));

                    controlValues[47] = DiffUtil.GetSumDiff(GetCell("CYL_IPA_out"), GetCell("CYL_Acetone_out"), GetCell("CYL_Nontarget_out"),
                                                             GetCell("CYL_IPA_in"), GetCell("CYL_Acetone_In"), GetCell("CYL_Nontarget_in"));

                    controlValues[51] = GetCell("CYL_Pressure_Drop");
                }

                foreach (var kvp in controlValues)
                {
                    ws.Cell(row, kvp.Key).Value = kvp.Value;
                }
                ws.Cell(row, 26).Value = GetCell("CYLSN");
                ws.Cell(row, 27).Value = GetCell("CYLWeight");
                ws.Cell(row, 48).Value = "N/A";
                ws.Cell(row, 49).Value = "N/A";
                ws.Cell(row, 50).Value = "130";
                string carbonLot = formData.GetValueOrDefault("CYLMaterialSNBox", "");
                string filterType = formData.GetValueOrDefault("CYLTypeBox", "");
                List<string> targetTypes = filterType.Split('+').Select(s => s.Trim()).ToList();
                var efficiencyMap = EfficiencyFinder.FindEfficiencyByTypeWithMin(ws, carbonLot, targetTypes);
                if (efficiencyMap.ContainsKey("MA")) ws.Cell(row, 32).Value = efficiencyMap["MA"];
                if (efficiencyMap.ContainsKey("MB")) ws.Cell(row, 33).Value = efficiencyMap["MB"];
                if (efficiencyMap.ContainsKey("MC")) ws.Cell(row, 34).Value = efficiencyMap["MC"];
                row++;
                rowIndex++;
            }
        }
        wb.Save();
        MessageBox.Show("匯入完成！", "已成功匯入資料", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
