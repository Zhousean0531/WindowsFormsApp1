using System;
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
        string productsize= ControlHelper.GetText(tab, "FilterInProcessProductSizeBox");
        string carbonOrder = ControlHelper.GetText(tab, "FilterInProcessOrderBox");
        string materialBatchNo = ControlHelper.GetText(tab, "FilterInProcessCarbonOrderBox");
        string materialNo = ControlHelper.GetText(tab, "FilterMaterialNOBox");
        string gile = ControlHelper.GetText(tab, "FilterInProcessGileBox");
        string speed = ControlHelper.GetText(tab, "FilterInProcessSpeedBox");
        string upper = ControlHelper.GetText(tab, "FilterInProcessUpperBox");
        string lower = ControlHelper.GetText(tab, "FilterInProcessLowerBox");
        string pressure = ControlHelper.GetText(tab, "FilterInProcessPressureBox");
        string carbonInfo = ControlHelper.GetText(tab, "FilterInProcessCarbonInfoBox");
        string wind = ControlHelper.GetText(tab, "FilterInProcessWindBox");
        string userName = Environment.UserName;
        string filtersize = ControlHelper.GetText(tab, "FilterSizeTB");
        var weights = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "FilterInProcessTestGsmBox"));
        var deltas = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "FilterInProcessPressureDropBox"));
        int n = Math.Min(weights.Count, deltas.Count);
        DateTime.TryParse(productionDate, out DateTime prodDt);
        DateTime.TryParse(testDate, out DateTime testDt);
        if (n == 0)
        {
            MessageBox.Show("沒有可用的測試資料");
            return null;
        }
        int selectedIdx;
        using (var f = new Form2(deltas, "請選擇用哪一筆壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;

            selectedIdx = f.SelectedIndex0;
        }
        var batch = new P2Batch
        {
            ReportNo = reportNo,
            ProductionDate = prodDt,
            TestDate = testDt,
            WorkOrder = carbonOrder,
            Material = productType,
            MaterialBatchNo = materialBatchNo,
            MaterialNo = materialNo,
            ProductDisplay=productDisplay,
            ProductSize= productsize,
            TargetGsm = TryParse(gsm),
            Glue = TryParse(gile),
            Speed = TryParse(speed),
            UpperTemp = TryParse(upper),
            LowerTemp = TryParse(lower),
            Pressure = TryParse(pressure),
            WindSpeed = TryParse(wind),
            CarbonLine = carbonInfo,
            Username = userName,
            FilterSize=filtersize
        };
        var panel = ControlHelper.Find<TableLayoutPanel>(tab, "FilterInProcessEffPanel");
        var pageType = GasPageType.FilterInProcess;
        foreach (var chk in ControlHelper.FindAll<System.Windows.Forms.CheckBox>(tab))
        {
            if (!chk.Checked) continue;
            if (!(chk.Tag is string gasKey)) continue;
            var gasCfg = GasMappingHelper.Get(gasKey);
            if (gasCfg == null) continue;
            if (!gasCfg.UiMap.TryGetValue(pageType, out var ui)) continue;
            var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);
            if (!decimal.TryParse(tbConc.Text, out decimal conc) || conc <= 0)
            {
                MessageBox.Show($"{gasKey} 濃度需>0");
                return null;
            }
            decimal.TryParse(tbBg.Text, out decimal bg);
            var readings = EfficiencyHelper.ParseReadings(tbVal.Text);
            if (readings.Count < 11)
            {
                MessageBox.Show($"{gasKey} 需輸入至少 11 筆讀值");
                return null;
            }
            var eff = EfficiencyHelper.Compute11Points((double)conc, (double)bg, readings);
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
                    PressureDrop = (decimal)deltas[i],
                    IsSelected = (i == selectedIdx)
                };
                if (i == selectedIdx)
                {
                    sample.Efficiencies = eff.Efficiencies.Select(x => (decimal)Math.Round(x, 1)).ToList();
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
    private static decimal? TryParse(string value)
    {
        if (decimal.TryParse(value, out decimal result))
            return result;
        return null;
    }
}