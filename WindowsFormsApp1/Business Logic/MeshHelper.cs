using System;
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

        // ───── 檢查空白列，只允許一格空白 ─────
        int emptyRowIndex = -1;

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
                if (!TryParse(valText, out double weight))
                    return false;

                if (weight < 0)
                    return false;
            }
        }

        // ───── 計算百分比 ─────
        double percentSum = 0;

        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];

            string key = r.Cells[0].Value?.ToString()?.Trim();
            string valText = r.Cells[1].Value?.ToString()?.Trim();

            if (string.IsNullOrWhiteSpace(key) || key.Contains("總重"))
                continue;

            // 空白列先跳過，最後用 100 - 其他百分比補
            if (i == emptyRowIndex)
                continue;

            if (!TryParse(valText, out double weight))
                return false;

            double percent = weight / totalWeight * 100.0;

            // 先四捨五入到小數 1 位，避免最後加總和畫面不一致
            double roundedPercent = Math.Round(percent, 1, MidpointRounding.AwayFromZero);

            r.Cells[1].Value = roundedPercent.ToString("F1", CultureInfo.InvariantCulture);

            percentSum += roundedPercent;
        }

        // ───── 空白列用 100 - 其他百分比補足 ─────
        if (emptyRowIndex != -1)
        {
            double missingPercent = 100.0 - percentSum;

            // 避免 -0.0
            if (Math.Abs(missingPercent) < 0.0001)
                missingPercent = 0;

            if (missingPercent < 0)
                return false;

            rows[emptyRowIndex].Cells[1].Value =
                missingPercent.ToString("F1", CultureInfo.InvariantCulture);
        }
        else
        {
            // 如果沒有空白列，就檢查加總是否接近 100
            if (Math.Abs(percentSum - 100.0) > 0.1)
                return false;
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
