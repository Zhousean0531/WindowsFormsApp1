using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Export.Page1;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static QuantityHelper;

public static class Page1DataCollector
{
    public static Page1ExportData Collect(TabPage tab)
    {
        // ───── 基本欄位 ─────
        var arrivePicker = ControlHelper.Find<DateTimePicker>(tab, "FilterRawArriveDateBox");
        var testPicker = ControlHelper.Find<DateTimePicker>(tab, "FilterRawTestDateBox");
        var materialBox = ControlHelper.Find<ComboBox>(tab, "FilterRawTypeBox");
        var reportNoBox = ControlHelper.Find<TextBox>(tab, "FilterRawReportNoTB");

        var tbConc = ControlHelper.Find<TextBox>(tab, "FilterRawConcertrationBox");
        var tbBg = ControlHelper.Find<TextBox>(tab, "FilterRawBackGroundBox");
        var tbVal = ControlHelper.Find<TextBox>(tab, "FilterRawEffvalueBox");

        if (!double.TryParse(tbConc?.Text, out double conc) || conc <= 0)
            return null;

        double.TryParse(tbBg?.Text, out double bg);

        var readings = EfficiencyHelper.ParseReadings(tbVal.Text);
        var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);
        if (eff == null) return null;

        // ───── 粒徑（唯一來源）─────
        var meshResult = MeshHelper.ParseFromTab(tab);
        if (meshResult == null) return null;

        // ───── 多筆資料 ─────
        var nos = ParseHelper.SplitStr(ControlHelper.GetText(tab, "FilterRawNumberBox"));
        var weights = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawWeightBox"));
        var vocIns = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsInletBox"));
        var vocOuts = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsOutletBox"));
        var deltaPs = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawPressureBox"));

        int n = new[] { nos.Count, weights.Count, vocIns.Count, vocOuts.Count, deltaPs.Count }.Min();
        if (n <= 0) return null;

        // ───── 選壓損 ─────
        int selectedIndex;
        using (var f = new Form2(deltaPs, "請選擇用哪一筆壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;
            selectedIndex = f.SelectedIndex0;
        }

        // ───── 組資料 ─────
        var lotFulls = nos.Take(n)
            .Select(no => $"B-{arrivePicker.Value:yyyyMMdd}-001#{no.PadLeft(2, '0')}")
            .ToList();

        var densities = weights.Take(n).Select(w => w / 50.0).ToList();

        var outgassing = vocOuts.Zip(vocIns, (o, i) =>
        {
            double d = o - i;
            return d <= 0 ? "N.D." : d.ToString("F1");
        }).ToList();

        return new Page1ExportData
        {
            ReportNo = reportNoBox.Text.Trim(),
            Material = materialBox.Text.Trim(),
            MaterialNo = MaterialMasterHelper.Get(materialBox.Text)?.MaterialNo ?? "",
            ArrivalDate = arrivePicker.Value.ToString("yyyy.MM.dd"),
            TestingDate = testPicker.Value.ToString("yyyy.MM.dd"),
            QtyText = QuantityHelper.BuildQuantityText(
                ProductKind.Filter,
                materialBox.Text,
                ControlHelper.GetText(tab, "FilterRawQtyWeight"),
                ControlHelper.GetText(tab, "FilterRawQuantityBox")),

            LotFulls = lotFulls,
            Densities = densities,
            DeltaPs = deltaPs.Take(n).ToList(),
            VocIns = vocIns.Take(n).ToList(),
            VocOuts = vocOuts.Take(n).ToList(),
            OutgassingList = outgassing,
            SelectedIndex = selectedIndex,

            ParticleSizePercentages = meshResult.Percentages,

            Eff0 = eff.Eff0,
            Eff10 = eff.Eff10,
            Efficiencies11 = eff.Efficiencies,

            UserName = Environment.UserName
        };
    }
}
