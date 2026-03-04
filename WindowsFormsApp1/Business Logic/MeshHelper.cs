using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

public static class MeshGridHelper
{
    /// <summary>
    /// 重新計算 mesh，補算唯一空白列，
    /// 並將「百分比」回填到 DGV 第二欄
    /// </summary>
    public static bool RecalculateAndFill(DataGridView dgv)
    {
        if (dgv == null) return true;

        var rows = dgv.Rows
            .Cast<DataGridViewRow>()
            .Where(r => !r.IsNewRow)
            .ToList();

        if (rows.Count == 0) return true;

        // ───── 找總重 ─────
        double totalWeight = 0;
        bool totalFound = false;

        foreach (var r in rows)
        {
            string key = r.Cells[0].Value?.ToString()?.Trim();
            string valText = r.Cells[1].Value?.ToString()?.Trim();

            if (!string.IsNullOrWhiteSpace(key) && key.Contains("總重"))
            {
                if (TryParse(valText, out totalWeight))
                    totalFound = true;
                break;
            }
        }

        if (!totalFound || totalWeight <= 0)
            return false;

        // ───── 收集重量，找空白列 ─────
        int emptyRowIndex = -1;
        double sum = 0;

        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            string key = r.Cells[0].Value?.ToString()?.Trim();
            string valText = r.Cells[1].Value?.ToString()?.Trim();

            if (string.IsNullOrWhiteSpace(key) || key.Contains("總重"))
                continue;

            if (string.IsNullOrWhiteSpace(valText))
            {
                if (emptyRowIndex != -1)
                    return false; // 超過一列空白
                emptyRowIndex = i;
            }
            else
            {
                if (!TryParse(valText, out double v))
                    return false;
                sum += v;
            }
        }

        // ───── 補算空白列重量 ─────
        if (emptyRowIndex != -1)
        {
            double missing = totalWeight - sum;
            if (missing < 0) return false;

            rows[emptyRowIndex].Cells[1].Value =
                missing.ToString("F2", CultureInfo.InvariantCulture);

            sum += missing;
        }

        if (Math.Abs(sum - totalWeight) > 0.001)
            return false;

        // ───── 全部轉為百分比並回填 ─────
        foreach (var r in rows)
        {
            string key = r.Cells[0].Value?.ToString()?.Trim();
            string valText = r.Cells[1].Value?.ToString()?.Trim();

            if (string.IsNullOrWhiteSpace(key) || key.Contains("總重"))
                continue;

            if (!TryParse(valText, out double weight))
                continue;

            double percent = weight / totalWeight * 100.0;
            r.Cells[1].Value = percent.ToString("F1", CultureInfo.InvariantCulture);
        }

        return true;
    }

    private static bool TryParse(string text, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return
            double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value) ||
            double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }
}
