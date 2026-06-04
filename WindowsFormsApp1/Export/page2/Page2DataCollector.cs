using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page2;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static WindowsFormsApp1.Helpers.GasMappingHelper;

public static class Page2DataCollector
{
    public static P2Batch Collect(TabPage tab)
    {
        string reportNo = ControlHelper.GetText(tab, "FilterInProcessReportNOTB");
        string productionDate = ControlHelper.GetText(tab, "FilterInProcessProductionBox");
        string testDate = ControlHelper.GetText(tab, "FilterInProcessTestBox");
         
        string productType = ControlHelper.GetText(tab, "FilterInProcessTypeBox");
        string gsm = ControlHelper.GetText(tab, "FilterInProcessgsmBox");

        string productDisplay = ControlHelper.GetText(tab, "FilterInProcessProductDisplayBox");
        string productSize = ControlHelper.GetText(tab, "FilterInProcessProductSizeBox");
        string rawOrder = ControlHelper.GetText(tab, "FilterInProcessCarbonOrderBox");
        string normalOrder = ControlHelper.GetText(tab, "FilterInProcessOrderBox");
        string orderType = ControlHelper.GetText(tab, "FilterInProcessOrderTypeBox");
        string materialNo = ControlHelper.GetText(tab, "FilterInProcessMaterialNo");
        string batchNo=ControlHelper.GetText(tab, "FilterBatchNOBox");
        string thickness = ControlHelper.GetText(tab, "FilterInProcessThicknessBox");
        string gile = ControlHelper.GetText(tab, "FilterInProcessGileBox");
        string speed = ControlHelper.GetText(tab, "FilterInProcessSpeedBox");
        string upper = ControlHelper.GetText(tab, "FilterInProcessUpperBox");
        string lower = ControlHelper.GetText(tab, "FilterInProcessLowerBox");
        string pressure = ControlHelper.GetText(tab, "FilterInProcessPressureBox");
        string carbonInfo = ControlHelper.GetText(tab, "FilterInProcessCarbonInfoBox");
        string wind = ControlHelper.GetText(tab, "FilterInProcessWindBox");

        string userName = Environment.UserName;
        string filterSize = ControlHelper.GetText(tab, "FilterSizeTB");

        string orderDisplay = ResolveOrderDisplay(rawOrder, normalOrder, orderType);

        if (!StringUtil.TrySplitDouble(ControlHelper.GetText(tab, "FilterInProcessTestGsmBox"), out var weights))
        {
            MessageBox.Show("重量欄位格式錯誤");
            return null;
        }

        if (!StringUtil.TrySplitDouble(ControlHelper.GetText(tab, "FilterInProcessPressureDropBox"), out var deltas))
        {
            MessageBox.Show("壓損欄位格式錯誤");
            return null;
        }

        if (!StringUtil.TrySplitDouble(thickness, out var thicknesses))
        {
            MessageBox.Show("厚度欄位格式錯誤");
            return null;
        }

        int n = weights.Count;
        if (n == 0 || deltas.Count != n)
        {
            MessageBox.Show("重量與壓損筆數需一致");
            return null;
        }

        bool hasThicknessInput = !string.IsNullOrWhiteSpace(thickness);
        if (hasThicknessInput && thicknesses.Count != n)
        {
            MessageBox.Show("厚度筆數需和重量 / 壓損筆數一致");
            return null;
        }

        DateTime? prodDt = TryParseDate(productionDate);
        DateTime? testDt = TryParseDate(testDate);

        if (string.IsNullOrWhiteSpace(productDisplay))
        {
            productDisplay = BuildProductDisplay(
                productType,
                gsm,
                testDt,
                reportNo
            );
        }

        var batch = new P2Batch
        {
            ReportNo = reportNo,
            ProductionDate = prodDt,
            TestDate = testDt,
            WorkOrder = orderDisplay,
            Material = productType,
            MaterialNo = materialNo,
            BatchNo=batchNo,
            ProductDisplay = productDisplay,
            ProductSize = productSize,
            TargetGsm = TryParse(gsm),
            Glue = gile,
            Speed = TryParse(speed),
            UpperTemp = TryParse(upper),
            LowerTemp = TryParse(lower),
            Pressure = TryParse(pressure),
            WindSpeed = TryParse(wind),
            CarbonLine = carbonInfo,
            Username = userName,
            FilterSize = filterSize
        };

        var panel = ControlHelper.Find<TableLayoutPanel>(
            tab,
            "FilterInProcessEffPanel"
        );

        var pageType = GasPageType.FilterInProcess;
        var usedSelectedIndexes = new Dictionary<int, string>();

        foreach (var chk in ControlHelper.FindAll<System.Windows.Forms.CheckBox>(tab))
        {
            if (!chk.Checked)
                continue;

            if (!(chk.Tag is string gasKey))
                continue;

            var gasCfg = GasMappingHelper.Get(gasKey);
            if (gasCfg == null)
                continue;

            if (!gasCfg.UiMap.TryGetValue(pageType, out var ui))
                continue;

            var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);

            if (tbConc == null || tbBg == null || tbVal == null)
            {
                MessageBox.Show($"{gasKey} 找不到濃度 / 背景 / 讀值欄位");
                return null;
            }

            if (!EfficiencyHelper.TryValidateInputs(
                gasKey,
                tbConc.Text,
                tbBg.Text,
                tbVal.Text,
                11,
                out double concDouble,
                out double bgDouble,
                out var readings,
                out string efficiencyError))
            {
                MessageBox.Show(efficiencyError);
                return null;
            }

            decimal conc = (decimal)concDouble;
            decimal bg = (decimal)bgDouble;

            var eff = EfficiencyHelper.Compute11Points(
                concDouble,
                bgDouble,
                readings
            );

            int selectedIdx = AskSelectedPressureIndex(
                deltas,
                gasKey,
                usedSelectedIndexes
            );

            if (selectedIdx < 0)
                return null;

            usedSelectedIndexes[selectedIdx] = gasKey;

            var gasTest = new P2GasTest
            {
                GasName = gasKey,
                Concentration = conc,
                Background = bg
            };

            for (int i = 0; i < n; i++)
            {
                var sample = new P2Sample
                {
                    Weight = (decimal)weights[i],
                    Thickness = hasThicknessInput ? (decimal)thicknesses[i] : (decimal?)null,
                    PressureDrop = (decimal)deltas[i],
                    IsSelected = (i == selectedIdx)
                };

                if (i == selectedIdx)
                {
                    sample.Efficiencies = eff.Efficiencies
                        .Select(x => (decimal)Math.Round(x, 1))
                        .ToList();
                }

                gasTest.Samples.Add(sample);
            }

            batch.GasTests.Add(gasTest);
        }

        if (batch.GasTests.Count == 0)
        {
            MessageBox.Show("請至少勾選一種氣體並填寫完整資料");
            return null;
        }

        return batch;
    }

    private static int AskSelectedPressureIndex(
        System.Collections.Generic.List<double> deltas,
        string gasKey,
        Dictionary<int, string> usedSelectedIndexes
    )
    {
        if (deltas == null || deltas.Count == 0)
            return -1;

        if (usedSelectedIndexes.Count >= deltas.Count)
        {
            MessageBox.Show("勾選的氣體數量超過可選擇的壓損樣品數量");
            return -1;
        }

        while (true)
        {
            using (var f = new Form2(deltas, $"請選擇 {gasKey} 用哪一筆壓損"))
            {
                if (f.ShowDialog() != DialogResult.OK)
                    return -1;

                int idx = f.SelectedIndex0;

                if (!usedSelectedIndexes.TryGetValue(idx, out string usedGas))
                    return idx;

                MessageBox.Show(
                    $"第 {idx + 1} 筆壓損已經給 {usedGas} 使用，請替 {gasKey} 選另一筆"
                );
            }
        }
    }

    private static string ResolveOrderDisplay(
    string rawOrder,
    string normalOrder,
    string orderType
)
    {
        rawOrder = (rawOrder ?? "").Trim();
        normalOrder = (normalOrder ?? "").Trim();
        orderType = (orderType ?? "").Trim();

        if (string.Equals(orderType, "台積堆疊式", StringComparison.OrdinalIgnoreCase))
        {
            return $"台積堆疊式({rawOrder})";
        }

        if (string.IsNullOrWhiteSpace(normalOrder) || normalOrder == "-")
            return rawOrder;

        if (string.IsNullOrWhiteSpace(rawOrder))
            return normalOrder;

        return $"{rawOrder}_{normalOrder}";
    }

    private static string BuildProductDisplay(
        string productType,
        string gsm,
        DateTime? testDate,
        string reportNo
    )
    {
        string mmdd = "";

        if (testDate.HasValue)
        {
            mmdd = testDate.Value.ToString("MMdd");
        }
        else if (!string.IsNullOrWhiteSpace(reportNo) && reportNo.Length > 6)
        {
            mmdd = reportNo.Substring(reportNo.Length - 6, 4);
        }

        return $"{productType} {gsm}gsm ({mmdd}生產)";
    }

    private static DateTime? TryParseDate(string value)
    {
        if (DateTime.TryParse(value, out DateTime result))
            return result;

        return null;
    }

    private static decimal? TryParse(string value)
    {
        if (decimal.TryParse(value, out decimal result))
            return result;

        return null;
    }
}
