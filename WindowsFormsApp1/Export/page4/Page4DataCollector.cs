using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page4;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static QuantityHelper;
using static WindowsFormsApp1.Helpers.GasMappingHelper;

public static class Page4DataCollector
{
    public static P4Batch Collect(TabPage tab)
    {
        var testPicker = ControlHelper.Find<DateTimePicker>(tab, "CylinderRawTestDateBox");
        var arrivePicker = ControlHelper.Find<DateTimePicker>(tab, "CylinderRawArriveDateBox");
        var materialBox = ControlHelper.Find<ComboBox>(tab, "CylinderRawTypeBox");
        var reportNoBox = ControlHelper.Find<TextBox>(tab, "CylinderRawReportNoTB");

        if (testPicker == null || arrivePicker == null || materialBox == null)
        {
            MessageBox.Show("缺少必要欄位");
            return null;
        }

        string testingDate = testPicker.Value.ToString("yyyy.MM.dd");
        string arrivalDate = arrivePicker.Value.ToString("yyyy.MM.dd");
        string material = materialBox.Text.Trim();
        string reportNo = reportNoBox.Text.Trim();
        string userName = Environment.UserName;
        string qtyWeight = ControlHelper.GetText(tab, "CylinderRawQtyWeightBox");
        string qtyPack = ControlHelper.GetText(tab, "CylinderRawQtyPackBox");

        string qtyText = QuantityHelper.BuildQuantityText(
            ProductKind.Cylinder,
            material,
            qtyWeight,
            qtyPack
        );
        var matInfo = MaterialMasterHelper.Get(material);
        string materialNo = matInfo?.MaterialNo ?? "";

        var lotNos = ParseHelper.SplitStr(ControlHelper.GetText(tab, "CylinderRawLotBox"));
        var numbers = ParseHelper.SplitStr(ControlHelper.GetText(tab, "CylinderRawNumberBox"));

        var weights = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawWeightBox"));
        var vocIns = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsInletBox"));
        var vocOuts = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsOutletBox"));
        var deltaPs = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawPressureBox"));

        int n = new[] { weights.Count, vocIns.Count, vocOuts.Count, deltaPs.Count }.Min();

        if (n == 0)
        {
            MessageBox.Show("沒有可用資料");
            return null;
        }

        weights = weights.Take(n).ToList();
        vocIns = vocIns.Take(n).ToList();
        vocOuts = vocOuts.Take(n).ToList();
        deltaPs = deltaPs.Take(n).ToList();

        var lotFulls = new List<string>();

        for (int i = 0; i < n; i++)
        {
            string num = i < numbers.Count ? numbers[i] : "";
            lotFulls.Add($"{numbers}#{num}");
        }

        const double VOL = 50.0;
        var densities = weights.Select(w => w / VOL).ToList();

        var outgassingList = vocOuts.Zip(vocIns, (o, i) =>
        {
            double diff = o - i;
            return diff <= 0 ? "N.D." : diff.ToString("F1");
        }).ToList();

        int selectedIndex;
        using (var f = new Form2(deltaPs, "請選擇壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;

            selectedIndex = f.SelectedIndex0;
        }

        var dgv = ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");

        var particlePercentages = new Dictionary<string, double>();

        foreach (DataGridViewRow row in dgv.Rows)
        {
            if (row.IsNewRow) continue;

            string key = row.Cells[0].Value?.ToString();
            string val = row.Cells[1].Value?.ToString();

            if (double.TryParse(val, out double percent))
                particlePercentages[key] = percent;
        }

        var effGroups = new List<EfficiencyGroup>();

        var panel = ControlHelper.Find<TableLayoutPanel>(tab, "CylinderRawEffPanel");
        var pageType = GasPageType.CylinderRawPage;

        foreach (CheckBox chk in ControlHelper.FindAll<CheckBox>(tab))
        {
            if (!chk.Checked) continue;
            if (!(chk.Tag is string gasKey)) continue;

            var gasCfg = GasMappingHelper.Get(gasKey);
            if (!gasCfg.UiMap.TryGetValue(pageType, out var ui)) continue;

            var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);

            if (!double.TryParse(tbConc.Text, out double conc)) continue;
            double.TryParse(tbBg.Text, out double bg);

            var readings = EfficiencyHelper.ParseReadings(tbVal.Text);
            if (readings.Count < 11) continue;

            var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);

            effGroups.Add(new EfficiencyGroup
            {
                GasName = gasKey,
                Concentration = conc,
                Eff0 = eff.Eff0,
                Eff10 = eff.Eff10,
                Efficiencies11 = eff.Efficiencies
            });
        }

        return new P4Batch
        {
            ReportNo = reportNo,
            Material = material,
            MaterialNo = materialNo,
            ArrivalDate = arrivalDate,
            TestingDate = testingDate,
            UserName = userName,
            QtyText= qtyText,
            LotNos = lotNos,
            LotFulls = lotFulls,

            Weights = weights,
            Densities = densities,

            VocIns = vocIns,
            VocOuts = vocOuts,
            DeltaPs = deltaPs,

            OutgassingList = outgassingList,
            SelectedIndex = selectedIndex,

            ParticleSizePercentages = particlePercentages,
            EfficiencyGroups = effGroups
        };
    }
}