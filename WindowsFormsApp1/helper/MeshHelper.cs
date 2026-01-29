using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1.Helpers
{
    public static class MeshHelper
    {
        // ★ 核心：解析 + 補算（唯一真實資料來源）
        public static MeshParseResult ParseFromTab(TabPage tab)
        {
            if (tab == null)
                return null;

            var dgv =
                ControlHelper.Find<DataGridView>(tab, "FilterRawParticleSizeBox")
                ?? ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");

            if (dgv == null)
                return null;

            var rows = dgv.Rows.Cast<DataGridViewRow>()
                        .Where(r => !r.IsNewRow)
                        .ToList();

            if (rows.Count == 0)
                return null;

            // 嘗試 1：尋找第一欄包含「總重」的列（較完整的 key/value 格式）
            double totalWeight = 0;
            bool totalFound = false;
            foreach (var r in rows)
            {
                string key = r.Cells.Count > 0 ? r.Cells[0].Value?.ToString()?.Trim() : null;
                string valText = r.Cells.Count > 1 ? r.Cells[1].Value?.ToString()?.Trim() : null;

                if (!string.IsNullOrWhiteSpace(key) && key.Contains("總重"))
                {
                    if (TryParseDouble(valText, out totalWeight))
                        totalFound = true;
                    break;
                }
            }

            // 嘗試 2：若沒找到，使用第二欄第一列作為總重（舊格式）
            if (!totalFound)
            {
                string totalText = rows[0].Cells.Count > 1
                    ? rows[0].Cells[1].Value?.ToString()?.Trim()
                    : null;

                if (!string.IsNullOrWhiteSpace(totalText) &&
                    TryParseDouble(totalText, out totalWeight))
                {
                    totalFound = true;
                }
            }

            if (!totalFound || totalWeight <= 0)
            {
                MessageBox.Show("找不到或解析不到總重（請確認格式）。");
                return null;
            }

            // 兩種資料來源的處理：
            // - 若表格以 key/value（第一欄為粒徑名稱）呈現，建構 weights dictionary 並補算空白項
            // - 否則，若表格以純第二欄數值呈現（第二列開始），則計算第二欄的百分比序列
            var weights = new Dictionary<string, double>();
            string emptyKey = null;
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                string key = r.Cells.Count > 0 ? r.Cells[0].Value?.ToString()?.Trim() : null;
                string valText = r.Cells.Count > 1 ? r.Cells[1].Value?.ToString()?.Trim() : null;

                if (string.IsNullOrWhiteSpace(key))
                    continue;

                if (key.Contains("總重"))
                    continue;

                if (string.IsNullOrWhiteSpace(valText))
                {
                    emptyKey = key;
                }
                else if (TryParseDouble(valText, out double v))
                {
                    weights[key] = v;
                }
            }

            var secondColPercentages = new List<double>();
            string summary = "";

            if (weights.Count == 0)
            {
                // fallback: treat second column as series (從第二列開始讀)
                var secondColValues = new List<double>();
                for (int i = 1; i < rows.Count; i++)
                {
                    var r = rows[i];
                    string valText = r.Cells.Count > 1 ? r.Cells[1].Value?.ToString()?.Trim() : null;
                    if (string.IsNullOrWhiteSpace(valText))
                        break;

                    if (TryParseDouble(valText, out double v))
                        secondColValues.Add(v);
                }

                if (secondColValues.Count == 0)
                    return null;

                secondColPercentages = secondColValues
                    .Select(w => w / totalWeight * 100.0)
                    .ToList();

                summary = string.Join(" , ", secondColPercentages.Select(p => $"{p:F1}%"));

                return new MeshParseResult
                {
                    Weights = null,
                    Percentages = null,
                    Summary = summary,
                    SecondColumnPercentages = secondColPercentages
                };
            }

            // 補算唯一空白粒徑
            if (!string.IsNullOrWhiteSpace(emptyKey))
            {
                double sum = weights.Values.Sum();
                double missing = totalWeight - sum;

                if (missing < 0)
                {
                    MessageBox.Show("已輸入的重量總和超過總重，請檢查資料。");
                    return null;
                }

                weights[emptyKey] = missing;
            }

            // 轉成百分比（完整結果）
            var percentages = weights.ToDictionary(
                kv => kv.Key,
                kv => kv.Value / totalWeight * 100.0
            );

            summary = string.Join(" , ",
                percentages.Select(kv => $"{kv.Key} {kv.Value:F1}%"));

            return new MeshParseResult
            {
                Weights = weights,
                Percentages = percentages,
                Summary = summary,
                SecondColumnPercentages = secondColPercentages
            };
        }

        // ★ 舊介面：只給顯示用（不影響其他 Page）
        public static string BuildMeshSummary(TabPage tab)
        {
            var r = ParseFromTab(tab);
            if (r == null)
                return "";

            if (r.Percentages != null && r.Percentages.Count > 0)
                return string.Join(" , ", r.Percentages.Select(kv => $"{kv.Key} {kv.Value:F1}%"));

            if (r.SecondColumnPercentages != null && r.SecondColumnPercentages.Count > 0)
                return string.Join(" , ", r.SecondColumnPercentages.Select(p => $"{p:F1}%"));

            return r.Summary ?? "";
        }

        private static bool TryParseDouble(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                return true;

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return true;

            return false;
        }
    }

    public class MeshParseResult
    {
        // 原始重量（key/value 格式時才有）
        public Dictionary<string, double> Weights { get; set; }

        // 百分比（key/value 格式）
        public Dictionary<string, double> Percentages { get; set; }

        // 第二欄純數值格式算出的百分比序列
        public List<double> SecondColumnPercentages { get; set; }

        // 統一摘要字串（可被 Helper 主動指定）
        public string Summary { get; set; }
    }
}
