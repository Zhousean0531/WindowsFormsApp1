using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public static class MeshHelper
{
    public static string BuildMeshSummary(TabPage tab, string material)
    {
        if (tab == null)
            return "";

        var dgv = ControlHelper.Find<DataGridView>(tab, "FilterRawParticleSizeBox")?? ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");
        if (dgv == null)
            return "";

        double totalWeight = 0;
        var weights = new Dictionary<string, double>();
        string emptyKey = null;

        foreach (DataGridViewRow row in dgv.Rows)
        {
            if (row.IsNewRow) continue;

            string key = row.Cells[0].Value?.ToString()?.Trim();
            string valText = row.Cells[1].Value?.ToString()?.Trim();

            if (string.IsNullOrWhiteSpace(key))
                continue;

            // ✅ 讀取「總重」
            if (key.Contains("總重"))
            {
                if (!double.TryParse(valText, out totalWeight))
                    totalWeight = 0;

                continue;
            }

            // 粒徑重量
            if (!string.IsNullOrWhiteSpace(valText))
            {
                if (double.TryParse(valText, out double v))
                    weights[key] = v;
            }
            else
            {
                // 記錄空白粒徑
                emptyKey = key;
            }
        }

        // 🔴 防呆：沒有總重就不能算百分比
        if (totalWeight <= 0)
        {
            MessageBox.Show("請先輸入總重，才能計算粒徑百分比");
            return "";
        }

        // ✅ 自動補算空白粒徑重量
        if (!string.IsNullOrWhiteSpace(emptyKey))
        {
            double sum = weights.Values.Sum();
            double missing = totalWeight - sum;

            if (missing < 0)
            {
                MessageBox.Show("粒徑重量加總大於總重，請確認資料");
                return "";
            }

            weights[emptyKey] = missing;
        }

        if (weights.Count == 0)
            return "";

        // ✅ 重量 → 百分比
        string summary = string.Join(" , ",
            weights.Select(kv =>
            {
                double percent = kv.Value / totalWeight * 100.0;
                return $"{kv.Key} {percent:F1}%";
            })
        );

        return summary;
    }
}
