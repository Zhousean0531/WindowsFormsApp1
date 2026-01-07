using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static WindowsFormsApp1.Helpers.GasMappingHelper;

public static class Page2DataCollector
{

    public static Page2ExportData Collect(TabPage tab)
    {
        // ───── 基本欄位 ─────
        string reportNo = ControlHelper.GetText(tab, "FilterInProcessReportNOTB");
        string productionDate = ControlHelper.GetText(tab, "FilterInProcessProductionBox");
        string testDate = ControlHelper.GetText(tab, "FilterInProcessTestBox");
        string productType = ControlHelper.GetText(tab, "FilterInProcessTypeBox");
        string gsm = ControlHelper.GetText(tab, "FilterInProcessgsmBox");
        string type = ControlHelper.GetText(tab, "FilterInProcessTypeBox");
        string prodMmdd = reportNo.Length > 6
    ? reportNo.Substring(reportNo.Length - 6, 4)
    : reportNo;
        string productDisplay =$"{type} {gsm}gsm ({prodMmdd}生產)";
        string filterSize = ControlHelper.GetText(tab, "FilterSizeTB");
        string carbonOrder =ControlHelper.GetText(tab, "FilterInProcessCarbonOrderBox");
        string order =ControlHelper.GetText(tab, "FilterInProcessOrderBox");
        string gile = ControlHelper.GetText(tab, "FilterInProcessGileBox");
        string speed = ControlHelper.GetText(tab, "FilterInProcessSpeedBox");
        string upper = ControlHelper.GetText(tab, "FilterInProcessUpperBox");
        string lower = ControlHelper.GetText(tab, "FilterInProcessLowerBox");
        string pressure = ControlHelper.GetText(tab, "FilterInProcessPressureBox");
        string carbonInfo = ControlHelper.GetText(tab, "FilterInProcessCarbonInfoBox");
        
        string orderDisplay;

        if (order == "-")
        {
            var result = MessageBox.Show(
                "此訂單是否為台積堆疊式成品？",
                "訂單確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                orderDisplay = $"台積堆疊式({carbonOrder})";
            }
            else
            {
                orderDisplay = carbonOrder;
            }
        }
        else
        {
            orderDisplay = $"{carbonOrder}_{order}";
        }
        string wind =ControlHelper.GetText(tab, "FilterInProcessWindBox");
        var weights = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "FilterInProcessTestGsmBox"));

        var deltas = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "FilterInProcessPressureDropBox"));

        int n = Math.Min(weights.Count, deltas.Count);
        if (n == 0)
        {
            MessageBox.Show("沒有可用的測試資料");
            return null;
        }

        weights = weights.Take(n).ToList();
        deltas = deltas.Take(n).ToList();

        // ───── 讓使用者選壓損 ─────
        int selectedIdx;
        using (var f = new Form2(deltas, "請選擇用哪一筆壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;

            selectedIdx = f.SelectedIndex0;
        }

        // ───── 多氣體效率計算 ─────
        var effGroups = new List<EfficiencyGroup>();
        var panel = ControlHelper.Find<TableLayoutPanel>(tab, "FilterInProcessEffPanel");
        var pageType = GasPageType.FilterInProcess;
        foreach (var chk in ControlHelper.FindAll<System.Windows.Forms.CheckBox>(tab))
        {

            if (!chk.Checked) continue;
            if (!(chk.Tag is string gasKey)) continue;

            var gasCfg = GasMappingHelper.Get(gasKey);
            if (gasCfg == null)
                continue;

            if (!gasCfg.UiMap.TryGetValue(pageType, out var ui))
                continue;

            var uiChk = ControlHelper.Find<System.Windows.Forms.CheckBox>(panel, ui.CheckBoxName); var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);

            if (!double.TryParse(tbConc.Text, out double conc) || conc <= 0)
            {
                MessageBox.Show($"{gasKey} 濃度需>0");
                return null;
            }

            double.TryParse(tbBg.Text, out double bg);

            List<double?> readings = EfficiencyHelper.ParseReadings(tbVal.Text);
            if (readings.Count < 11)
            {
                MessageBox.Show($"{gasKey} 需輸入至少 11 筆讀值");
                return null;
            }

            var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);

            effGroups.Add(new EfficiencyGroup
            {
                GasName = gasKey,
                Concentration = conc,
                Eff0 = eff.Eff0,
                Eff10 = eff.Eff10,
                Readings11 = eff.Readings,
                Efficiencies11 = eff.Efficiencies
            });
        }
        var allChks = ControlHelper.FindAll<System.Windows.Forms.CheckBox>(tab);
       
        if (effGroups.Count == 0)
        {
            MessageBox.Show("請至少勾選一種氣體並填寫完整資料");
            return null;
        }

        // ───── 回傳 DTO（多氣體版） ─────
        return new Page2ExportData
        {
            ReportNo = reportNo,
            ProductionDate = productionDate,
            TestDate = testDate,
            ProductType = productType,
            Gsm = gsm,
            Pressure=pressure,
            CarbonOrder = carbonOrder,
            TestWeights = weights,
            PressureDrops = deltas,
            SelectedIndex = selectedIdx,
            OrderDisplay = orderDisplay,
            FilterSize = filterSize,
            CarbonInfo= carbonInfo,
            ProductDisplay = productDisplay,
            Wind = wind,
            Gile =gile,
            Speed=speed,
            Upper=upper,
            Lower=lower,
            EfficiencyGroups = effGroups   // ← 關鍵：不再用單一 Efficiency
        };
    }
}
