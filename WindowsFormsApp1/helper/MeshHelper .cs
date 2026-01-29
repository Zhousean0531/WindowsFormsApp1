using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Globalization;

public static class MeshHelper
{
    public static MeshParseResult ParseFromTab(TabPage tab)
    {
        if (tab == null)
            return null;

        var dgv =
            ControlHelper.Find<DataGridView>(tab, "FilterRawParticleSizeBox")
            ?? ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");

        if (dgv == null)
            return null;

        // 取得非空的資料列（排除 NewRow）
        var rows = dgv.Rows.Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow)
                    .ToList();

        if (rows.Count == 0)
            return null;

        // 第二欄第一列為總重
        string totalText = rows[0].Cells.Count > 1
            ? rows[0].Cells[1].Value?.ToString()?.Trim()
            : null;

        if (string.IsNullOrWhiteSpace(totalText) ||
            !double.TryParse(totalText, NumberStyles.Any, CultureInfo.CurrentCulture, out double totalWeight) &&
            !double.TryParse(totalText, NumberStyles.Any, CultureInfo.InvariantCulture, out totalWeight))
        {
            MessageBox.Show("找不到或解析不到總重（請確認第二欄第一列為總重）");
            return null;
        }

        if (totalWeight <= 0)
        {
            MessageBox.Show("總重需為大於 0 的數值");
            return null;
        }

        var secondColWeights = new List<double>();

        // 從第二列開始讀取第二欄，遇到空白則停止
        for (int i = 1; i < rows.Count; i++)
        {
            var r = rows[i];
            string valText = r.Cells.Count > 1 ? r.Cells[1].Value?.ToString()?.Trim() : null;
            if (string.IsNullOrWhiteSpace(valText))
                break;

            if (double.TryParse(valText, NumberStyles.Any, CultureInfo.CurrentCulture, out double v) ||
                double.TryParse(valText, NumberStyles.Any, CultureInfo.InvariantCulture, out v))
            {
                secondColWeights.Add(v);
            }
            else
            {
                // 若解析失敗，可跳過或回報；此處跳過
                continue;
            }
        }

        // 計算百分比（每筆 / 總重 * 100）
        var percentages = secondColWeights.Select(w => w / totalWeight * 100.0).ToList();

        string summary = percentages.Count > 0
            ? string.Join(" , ", percentages.Select(p => $"{p:F1}%"))
            : "";

        return new MeshParseResult
        {
            Weights = null,
            Percentages = null,
            Summary = summary,
            SecondColumnPercentages = percentages
        };
    }
}

public class MeshParseResult
{
    public Dictionary<string, double> Weights { get; set; }
    public Dictionary<string, double> Percentages { get; set; }
    public string Summary { get; set; }

    // 新增：第二欄的百分比（對應 Excel B7,B8,...）
    public List<double> SecondColumnPercentages { get; set; } = new List<double>();
}
